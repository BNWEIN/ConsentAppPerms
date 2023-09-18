<#
.SYNOPSIS
Consents an app with application permissions in customer tenants

.DESCRIPTION
Note, leverages the SAM app model

.NOTES
Created by   : Roel van der Wegen
Date Coded   : August 2023
More info    : https://github.com/rvdwegen
#>

# Connecting to Azure Parameters
$CSPtenant = "yourtenantid"
$applicationID = "<enter application id here>"
$clientKey = "<enter client secret here>"
$PartnerrefreshToken = "<enter Refresh Token here>"
$AppId = '<Enter your Halo PSA CSP - Partner Center Connection (multitenant) app ID here>' # This is the app you want to consent in your customer tenants

# in 7.2 the progress on Invoke-WebRequest is returned to the runbook log output
$ProgressPreference = 'SilentlyContinue'

#region ############################## Functions ####################################

function Get-MicrosoftToken {
    Param(
        # Tenant Id
        [Parameter(Mandatory=$false)]
        [guid]$TenantId,

        # Scope
        [Parameter(Mandatory=$false)]
        [string]$Scope = 'https://graph.microsoft.com/.default',

        # ApplicationID
        [Parameter(Mandatory=$true)]
        [guid]$ApplicationID,

        # ApplicationSecret
        [Parameter(Mandatory=$true)]
        [string]$ApplicationSecret,

        # RefreshToken
        [Parameter(Mandatory=$true)]
        [string]$RefreshToken
    )

    if ($TenantId) {
        $Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    }
    else {
        $Uri = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    }

    # Define the parameters for the token request
    $Body = @{
        client_id       = $ApplicationID
        client_secret   = $ApplicationSecret
        scope           = $Scope
        refresh_token   = $RefreshToken
        grant_type      = 'refresh_token'
    }

    $Params = @{
        Uri = $Uri
        Method = 'POST'
        Body = $Body
        ContentType = 'application/x-www-form-urlencoded'
        UseBasicParsing = $true
    }

    try {
        $AuthResponse = (Invoke-WebRequest @Params).Content | ConvertFrom-Json
    } catch {
        throw "Authentication Error Occured $_"
    }

    return $AuthResponse
}

$commonTokenSplat = @{
    ApplicationID = $applicationID
    ApplicationSecret = $clientKey
    RefreshToken = $PartnerrefreshToken
}

try {
    if ($CSPtoken = (Get-MicrosoftToken @commonTokenSplat -TenantID $CSPtenant -Scope "https://graph.microsoft.com/.default").Access_Token) {
        $CSPheader = @{
            Authorization = 'bearer {0}' -f $CSPtoken
            Accept        = "application/json"
        }
    }

    # Get tenants. There are many ways to rome. This one will work in most cases
    $customertenants = (Invoke-RestMethod -Method GET -headers $CSPheader -Uri 'https://graph.microsoft.com/beta/contracts?$top=999').value

    # Get app details including permissions
    $AppDetails = (Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/applications(appId='$AppId')" -headers $CSPheader)
} catch {
    throw "$($_.Exception.Message)"
}

#region ############################## Loop through tenants ####################################

foreach ($tenant in $customertenants) {

    try {
        Write-Output "Processing tenant: $($Tenant.defaultDomainName) | $($tenant.customerId)"

        try {
            if ($token = (Get-MicrosoftToken @commonTokenSplat -TenantID $($tenant.customerId) -Scope "https://graph.microsoft.com/.default").Access_Token) {
                $header = @{
                    Authorization = 'bearer {0}' -f $token
                    Accept        = "application/json"
                }
            }
        } catch {
            throw "$($_.Exception.Message)"
        }

        try {
            # Check if there is a service principal for the app, if not create it
            if (!($svcPrincipal = (Invoke-RestMethod -Method "GET" -Headers $header -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$($AppDetails.appId)'").value)) {
                # Define values for the new svcPrincipal
                $newsvcPrincipalBody = @{
                    appId = $AppDetails.appId
                } | ConvertTo-Json

                # Create the svcPrincipal
                if ($svcPrincipal = (Invoke-RestMethod -Method "POST" -Headers $header -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body $newsvcPrincipalBody -ContentType "application/json")) {
                    #Write-Output "svcPrincipal id $($svcPrincipal.id) was created"
                } else {
                    throw "Failed to create svcPrincipal"
                }
            } else {
                $roles = (Invoke-RestMethod -Method GET -Headers $header -Uri "https://graph.microsoft.com/beta/servicePrincipals(appId='$($AppDetails.appId)')/appRoleAssignments").value
            }
        } catch {
            throw "$($_.Exception.Message)"
        }

        # Consent App permissions one by one
        foreach ($ResourceApp in $AppDetails.requiredResourceAccess) {
            $ApiApp = $null
            $ApiApp = (Invoke-RestMethod -Method GET -Headers $header -Uri "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$($ResourceApp.ResourceAppId)')")
            foreach ($grant in $ResourceApp.ResourceAccess) {
                if ($grant.Type -eq 'Role' -AND $roles.appRoleId -notcontains $grant.id) {
                    $NewMgServicePrincipalAppRoleAssignmentSplat = @{
                        principalId = $svcPrincipal.Id
                        resourceId = $ApiApp.Id
                        appRoleId = $grant.Id
                    } | ConvertTo-Json

                    #Write-Output "Consenting $($grant.Id) for $($ApiApp.displayName) on $($svcPrincipal.displayName) in $($Tenant.defaultDomainName)"

                    try {
                        $null = Invoke-RestMethod -Method POST -Headers $header -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($svcPrincipal.id)/appRoleAssignedTo" -Body $NewMgServicePrincipalAppRoleAssignmentSplat -ContentType "application/Json"
                    } catch {
                        Write-Warning "Something went wrong while consenting $($grant.Id) on $($svcPrincipal.displayName) in $($Tenant.defaultDomainName): $($_.Exception.Message)"
                    }
                }
            }
        }
    } catch {
        Write-Error "$($Tenant.defaultDomainName) | $($tenant.customerId) :$($_.Exception.Message) | Line: $($_.InvocationInfo.ScriptLineNumber)"
    }
}
#endregion

Write-Output "End of run"
