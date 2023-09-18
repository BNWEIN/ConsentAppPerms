<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER customerTenantId
GUID property

.NOTES
Created by   : Roel van der Wegen
Date Coded   : August 2023
More info    : https://github.com/rvdwegen
#>

# Connecting to Azure Parameters
$tenantID = "<enter your partner tenant ID here>"
$applicationID = "<Enter your SAM APP ID Here>"
$clientKey = "<Enter your SAM Secret here>"
$PartnerrefreshToken = "<Enter your refresh token here>"
$TenantMGMTAppId = '<Enter your Halo PSA CSP - Partner Center Connection (multitenant) app ID here>'

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
        Write-Error "Authentication Error Occured $_"
        return
    }

    return $AuthResponse
}

$commonTokenSplat = @{
    ApplicationID = $applicationID
    ApplicationSecret = $clientKey
    RefreshToken = $PartnerrefreshToken
}

# Authenticate to Microsoft Grpah
Write-Host "Authenticating to Microsoft Graph via REST method"
 
$url = "https://login.microsoftonline.com/$tenantId/oauth2/token"
$resource = "https://graph.microsoft.com/"
$restbody = @{
         grant_type    = 'client_credentials'
         client_id     = $applicationID
         client_secret = $clientKey
         resource      = $resource
}
     
 # Get the return Auth Token
 $token = Invoke-RestMethod -Method POST -Uri $url -Body $restbody

 # Set the baseurl to MS Graph-API (BETA API)
$baseUrl = 'https://graph.microsoft.com/beta'

if ($ogtoken = (Get-MicrosoftToken @commonTokenSplat -TenantID $tenantID -Scope "https://graph.microsoft.com/.default" -ErrorAction SilentlyContinue).Access_Token) {
    $ogheader = @{
        Authorization = 'bearer {0}' -f $ogtoken
        Accept        = "application/json"
    }
} else {
    throw "Unable to authenticate to Prime Tenant tenant."
}

# Build the Base URL for the API call
$url = $baseUrl + '/contracts?$top=999'

# Call the REST-API
$customertenants = Invoke-RestMethod -Method GET -headers $ogheader -Uri $url

$tenantMGMTAppDetails = (Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/applications(appId='$TenantMGMTAppId')" -headers $ogheader)

#endregion

#region ############################## Loop through tenants ####################################

foreach ($tenant in $customertenants.value) {

    try {
        #Write-Output "Processing tenant: $($Tenant.defaultDomainName) | $($tenant.TenantId)"

        if ($token = (Get-MicrosoftToken @commonTokenSplat -TenantID $($tenant.customerId) -Scope "https://graph.microsoft.com/.default" -ErrorAction SilentlyContinue).Access_Token) {
            $header = @{
                Authorization = 'bearer {0}' -f $token
                Accept        = "application/json"
            }
        } else {
            throw "Unable to authenticate to tenant: $($Tenant.defaultDomainName) | $($tenant.customerId)"
        }

        # Check if there is a service principal for the app, if not create it
        if (!($svcPrincipal = (Invoke-RestMethod -Method "GET" -Headers $header -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$($tenantMGMTAppDetails.appId)'").value)) {
            # Define values for the new svcPrincipal
            $newsvcPrincipalBody = @{
                appId = $tenantMGMTAppDetails.appId
            } | ConvertTo-Json

            # Create the svcPrincipal
            if ($svcPrincipal = (Invoke-RestMethod -Method "POST" -Headers $header -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body $newsvcPrincipalBody -ContentType "application/json")) {
                #Write-Output "svcPrincipal id $($svcPrincipal.id) was created"
            } else {
                Write-Warning "Failed to create svcPrincipal"
            }
        } else {
            $roles = (Invoke-RestMethod -Method GET -Headers $header -Uri "https://graph.microsoft.com/beta/servicePrincipals(appId='$($tenantMGMTAppDetails.appId)')/appRoleAssignments").value
        }

        # Consent App permissions one by one
        foreach ($ResourceApp in $tenantMGMTAppDetails.requiredResourceAccess) {
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
