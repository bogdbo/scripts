Import-Module .\Microsoft.Exchange.WebServices.dll
Import-Module .\Microsoft.Exchange.WebServices.Auth.dll

Function Send-Appointment()
{
[CmdletBinding()]
    Param
    (
      [parameter(Mandatory=$true)]
      [Uri]$ExchangeServiceUrl,
      
      [parameter(Mandatory=$true)]
      [string[]]$Emails,

      [Parameter(Mandatory=$false)]
      [ValidateSet(
      "Exchange2007_SP1", 
      "Exchange2010", 
      "Exchange2010_SP1", 
      "Exchange2010_SP2", 
      "Exchange2013", 
      "Exchange2013_SP1")]
      [string]$ExchangeVersion = "Exchange2010_SP1",

      [parameter(Mandatory=$false)]
      [string]$Subject = "Powershell meeting",

      [parameter(Mandatory=$false)]
      [string]$Body = "This is a test, please ignore",

      [parameter(Mandatory=$false)]
      [DateTime]$Start = [DateTime]::Now.AddHours(1),

      [parameter(Mandatory=$false)]
      [DateTime]$End = [DateTime]::Now.AddHours(2),

      [parameter(Mandatory=$false)]
      [string]$Location = "No location",

      [switch]
      [bool]$IsAllDayEvent = $false
    )

    $credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials($Username,$Password,$Domain)

    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
    $service.Url = $ExchangeServiceUrl
    $service.PreAuthenticate = $true
    $service.SendClientLatencies = $true
    $service.EnableScpLookup = $false
    $cred = Get-Credential
    $service.Credentials = $cred.GetNetworkCredential()

    $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment($service)
    $appointment.Subject = $Subject
    $appointment.Body = $Body
    $appointment.Start = $Start
    $appointment.End = $End
    $appointment.Location = $Location
    $appointment.IsAllDayEvent = $IsAllDayEvent

    foreach($email in $Emails) {
        $attendee = New-Object Microsoft.Exchange.WebServices.Data.Attendee($email)
        $appointment.RequiredAttendees.Add($attendee)
    }

    $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)

    Write-Host 'Appointment sent.'
}

Send-Appointment -ExchangeServiceUrl "asmx service url" -Emails "email1", "email2"
