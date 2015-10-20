# Microsoft doesn't make it easy to see a list
# of mapped PSTs within an Outlook profile.
# This script goes through all the necessary nonsense
# that is requires to get a plain list of those files.

# Currently only works with the current user (HKEY_CURRENT_USER).
# TODO: Add support for any domain user, and look up user in HKEY_USERS.
  
Function Hex {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [int]$Value
  )
Begin {}
Process {
  '{0:X2}' -f $Value
}
End {}
}

Function Is-A-PST {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [Alias("Value")]
    [String]$PSTGuid
  )
Begin {}
Process {
  $IsAPST = $False;
  $PSTCheckFile = "00033009";
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$PSTGuid");
  $PSTGuidValue = $reg.GetValue($PSTCheckFile, @(0));
  $PSTCheck = 0;
  Foreach($x in $PSTGuidValue) {
    $PSTCheck += [int] $x;
  }
  # If the subkey "$PSTCheckFile" in "$PSTGuid", when summed, 
  # equals hex 20 (32 in decimal), then that key corresponds
  # to a PST. If not, it corresponds with something else. 
  If ($PSTCheck -eq 0x20) {
    $IsAPST = $True;
  }
  Return $IsAPST;
}
End {}
}

Function Get-PST-Location {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [Alias("Value")]
    [String]$PSTGuid
  )
Begin {}
Process {
  $PSTGuidLocation = "01023d00";
  $PSTLocation = [String]::Empty;
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$PSTGuid");
  $PSTGuidValue = $reg.GetValue($PSTGuidLocation, @(0));
  # Loop through the subkey $PSTGuidLocation in $PSTGuid
  # to get a new Guid value. This value contains the path 
  # to the PST Location.
  foreach($y in $PSTGuidValue) {
    $PSTLocation += (Hex $y);
  }
  Return $PSTLocation;
}
End {}
}

Function Get-PST-FileName {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [Alias("Value")]
    [String]$PSTGuid
  )
Begin {}
Process {
  $PSTFileName = [String]::Empty;
  $PSTFile = "001f6700";
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$PSTGuid");
  $PSTName = $reg.GetValue($PSTFile, @(0));
  foreach($z in $PSTName) {
    If ($z -gt 0) {
      $PSTFileName += [char] $z;
    }
  }
  Return $PSTFileName;
}
End {}
}

Function Get-PST-DisplayName {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [Alias("Value")]
    [String]$PSTGuid
  )
Begin {}
Process {
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$PSTGuid");
  # If Key does not exist, exit and return 0
  If ($reg -eq $null) {
    Return 0;
  }
  $PSTType = $reg.GetValue("001e3001", 0);
  $PSTDisplayName = [String]::Empty;
  If ($PSTType -eq 0) {
    # 2003/2007 (and 2010 and 2013 apparently ) file type
    $PSTNameBytes = $reg.GetValue("001f3001", @(0));
    foreach($q in $PSTNameBytes) {
      If ($q -ne 0) {
        $PSTDisplayName += [char] $q;
      }
    }
  } Else {
    # We have a 97/2002 pst file type
    $PSTDisplayName = $PSTType.ToString();
  }
  Return $PSTDisplayName;
}
End {}
}

Function Get-PSTs-For-Profile {
[CmdletBinding()] 
Param 
  ( 
    [Parameter( 
      ValueFromPipeline=$True, 
      ValueFromPipelineByPropertyName=$True, 
      Mandatory=$True
		)]
    [String]$ProfileName
  )
Begin {
  If ([String]::IsNullOrEmpty($ProfileName)) {
    Return $null;
  }
}
Process {
  $ProfilesRoot = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles";
  $KeyMaster = "9207f3e0a3b11019908b08002b2a56c2";
  $MasterConfig = "01023d0e";
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey("$ProfilesRoot\$ProfileName\$KeyMaster");
  $ReturnValue = $reg.GetValue($MasterConfig, @(0));
  $StrPSTGuid = [String]::Empty;
  $AllPSTs = New-Object System.Collections.Generic.List[PSObject];
  ForEach ($i in $ReturnValue) {
    $StrHexNumber = (Hex $i);
    $StrPSTGuid += $StrHexNumber;
    If ($StrPSTGuid.Length -eq 32) {
      If (Is-A-PST "$ProfilesRoot\$ProfileName\$StrPSTGuid") {
        $PSTLocation = Get-PST-Location "$ProfilesRoot\$ProfileName\$StrPSTGuid"
        $PSTFileName = Get-PST-FileName "$ProfilesRoot\$ProfileName\$PSTlocation"
        $PSTDisplayName = Get-PST-DisplayName "$ProfilesRoot\$ProfileName\$PSTlocation"
        $UserPST = New-Object PSObject -Property @{
          FilePath=$PSTFileName;
          DisplayName=$PSTDisplayName;
          ProfileName=$ProfileName;
        }
        $AllPSTs.Add($UserPST);
      }
      $StrPSTGuid = [String]::Empty;
    }
  }
  Return ($AllPSTs.ToArray());
}
End {}
}

Function Get-Default-Outlook-Profile {
  $ProfilesRoot = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles";
  $DefaultProfileString = "DefaultProfile";
  Try {
    $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($ProfilesRoot);
    $DefaultProfileName = $reg.GetValue($DefaultProfileString, 0).ToString();
    Return $DefaultProfileName;
  } Catch {
    # User does not have an Outlook Profile.
    # Write-Host "This user does not have an Outlook profile.";
    Return $null;
  }
}

Function Start-PSTInfo {
  $DefaultProfileName = Get-Default-Outlook-Profile
  $ProfilesRoot = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles";
  $reg = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($ProfilesRoot);
  $AllAccounts = New-Object System.Collections.Generic.List[PSObject]
  foreach ($Profile in $reg.GetSubKeyNames()) {
    $CurrentOutlookAccountPSTS = Get-PSTs-For-Profile -ProfileName "$Profile"
    $CurrentOutlookAccount = New-Object PSObject -Property @{
      PSTs=$CurrentOutlookAccountPSTS;
      ProfileName=$Profile;
      IsDefault=($DefaultProfileName -eq $Profile);
    }
    $AllAccounts.Add($CurrentOutlookAccount);
  }
  Return ($AllAccounts.ToArray());
}


