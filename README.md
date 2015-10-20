# pst-info
Returns a list of what Outlook PSTs are mapped within the currently logged in user's Outlook profile.

## Usage:
```
> . .\PST_Info.ps1
$CurrentUserProfiles = Start-PSTInfo
foreach ($OutlookProfile in $CurrentUserProfiles) {
  $PSTs = $OutlookProfile.PSTs
  $IsDefault = $OutlookProfile.IsDefault
  foreach ($PST in $PSTs) {
    Write-Host "PST Display Name: $($PST.DisplayName)" 
    Write-Host "PST File Path:    $($PST.FilePath)" 
    Write-Host "PST Profile Name: $($PST.ProfileName) (IsDefault: $($IsDefault))"
  }
}
```
