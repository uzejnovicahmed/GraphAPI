
# PowerShell Script: Access Token Management with Automatic Renewal

This script demonstrates how to automate the process of managing and renewing an Azure AD access token in the background using a `System.Timers.Timer` event. The token is renewed a specified number of minutes before it expires, and the script checks the token's status at a defined interval.

## 1. Script Setup

### Step 1: Enable Debugging and Verbose Output

```powershell
$DebugPreference = "Continue"
$VerbosePreference = "Continue"
```

- Enables detailed debug and verbose output for easier debugging and troubleshooting.

### Step 2: Define Global Variables

```powershell
$global:accessToken = $null
$global:tokenExpiresAt = [datetime]::MinValue
$tokencheckinterval = 5000  # 5 seconds interval for checking token expiration (set higher in production)
$minutesbeforetokenexpires = 59  # Minutes before token expiration when it should be renewed
```

- `accessToken`: Holds the current access token.
- `tokenExpiresAt`: Stores the expiration time of the access token.
- `tokencheckinterval`: Interval (in milliseconds) to check the token status.
- `minutesbeforetokenexpires`: Defines how many minutes before expiration the token should be renewed.

### Step 3: Set Tenant and Client Credentials

```powershell
$TenantName = ""
$CLIENTID = ""
$CLIENTSECRET = ""
```

- Set the values for `TenantName`, `CLIENTID`, and `CLIENTSECRET` with your Azure AD details.

## 2. Function Definitions

### Step 4: `Request-AccessToken` Function

```powershell
function Request-AccessToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]$TenantName,
        [Parameter(Mandatory = $true)] [string]$ClientId,
        [Parameter(Mandatory = $true)] [string]$ClientSecret
    )

    $tokenBody = @{
        Grant_Type    = 'client_credentials'
        Scope         = 'https://graph.microsoft.com/.default'
        Client_Id     = $ClientId
        Client_Secret = $ClientSecret
    }

    Write-Debug "Requesting a new access token from Microsoft Identity Platform..."

    try {
        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody -ErrorAction Stop
    } catch {
        Write-Output "Error generating access token: $($_.Exception.Message)"
        Write-Output "Exception Details: $($_.Exception)"
        return $null
    }

    Write-Output "Successfully generated authentication token"
    return $tokenResponse
}
```

- Requests a new access token from Azure AD.
- Takes `TenantName`, `ClientId`, and `ClientSecret` as input parameters.

### Step 5: `Renew-AccessToken` Function

```powershell
function Renew-AccessToken {
    param (
        [string]$TenantName,
        [string]$ClientId,
        [string]$ClientSecret
    )

    Write-Debug "Attempting to renew the access token..."
    try {
        $tokenResponse = Request-AccessToken -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret

        if ($null -ne $tokenResponse) {
            $global:accessToken = $tokenResponse.access_token
            $global:tokenExpiresAt = (Get-Date).AddSeconds($tokenResponse.expires_in)
            Write-Output "Token renewed successfully. New expiration time: $global:tokenExpiresAt"
        } else {
            Write-Output "Failed to renew the access token. Response was null."
        }
    } catch {
        Write-Output "Error renewing the access token: $($_.Exception.Message)"
    }
}
```

- Renews the access token using the `Request-AccessToken` function.
- Updates the global variables `accessToken` and `tokenExpiresAt` upon successful renewal.

### Step 6: `Check-TokenExpiration` Function

```powershell
function Check-TokenExpiration {
    try {
        Write-Output "Checking token expiration at: $(Get-Date)..."

        # Check if the token is about to expire
        if ((Get-Date) -ge $global:tokenExpiresAt.AddMinutes(-$minutesbeforetokenexpires)) {
            Write-Output "Access token is expired or close to expiration. Renewing the token..."
            Renew-AccessToken -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret
        } else {
            Write-Output "Access token is still valid. Expires at: $global:tokenExpiresAt"
        }
    } catch {
        Write-Output "Error in Check-TokenExpiration function: $($_.Exception.Message)"
    }
}
```

- Checks whether the token is close to expiration based on the defined `$minutesbeforetokenexpires`.
- If close to expiration, it calls the `Renew-AccessToken` function.

## 3. Initialize and Manage Token

### Step 7: Request Initial Token

```powershell
$tokenResponse = Request-AccessToken -TenantName $TenantName -ClientId $ClientId -ClientSecret $ClientSecret

if ($tokenResponse -ne $null) {
    $global:accessToken = $tokenResponse.access_token
    $global:tokenExpiresAt = (Get-Date).AddSeconds($tokenResponse.expires_in)
    Write-Output "Initial token acquired successfully. Expires at: $global:tokenExpiresAt"
} else {
    Write-Output "Failed to acquire the initial access token. Exiting script."
    return
}
```

- Requests the initial token and updates the global variables.

## 4. Set Up and Start Timer

### Step 8: Configure and Start the Timer

```powershell
$timer = New-Object System.Timers.Timer
$timer.Interval = $tokencheckinterval
$timer.AutoReset = $true
$timer.Enabled = $true  # Enable the timer

try {
    $timerEvent = Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier "TokenCheck" -Action {
        Write-Output "Timer Event Triggered: $(Get-Date)"
        Check-TokenExpiration
    }
    Write-Output "Event registered successfully."
} catch {
    Write-Output "Failed to register the timer event: $($_.Exception.Message)"
}

# Start the timer and verify that it is running
$timer.Start()
if ($timer.Enabled) {
    Write-Output "Timer started successfully. Access token will be checked for renewal every 30 seconds."
} else {
    Write-Output "Failed to start the timer."
}
```

- Configures a `System.Timers.Timer` object to check the token’s expiration status at a defined interval.
- Registers the `Check-TokenExpiration` function to be called each time the timer elapses.

## 5. Main Script Execution

### Step 9: Main Script Logic

```powershell
for ($i = 1; $i -le 60; $i++) {
    Write-Output "[$(Get-Date)] - Main Script: Working... (Iteration $i)"
    Start-Sleep -Seconds 2
    Write-Output "The Global Access Token Expires At: $global:tokenExpiresAt"
}
```

- Simulates other tasks that the script might perform while the token is being managed in the background.

## 6. Cleanup and Exit

### Step 10: Cleanup Timer and Unregister Event

```powershell
Write-Output "Stopping the timer and unregistering the event..."
$timer.Stop()
Unregister-Event -SourceIdentifier "TokenCheck"
$timer.Dispose()
Write-Output "Timer stopped. Final Access Token Expiration: $global:tokenExpiresAt"
```

- Stops the timer and unregisters the event when the script completes.

## Summary

This script provides an automated way to manage and renew Azure AD access tokens using a background timer. It ensures that the token is always up-to-date during script execution and can be customized for different intervals and renewal settings.
