<#
    .SYNOPSIS
    Remove orphaned users and groups from legacy public folder ACLs 
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.3, 2020-12-15

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://scripts.granikos.eu

    .DESCRIPTION
    This scripts removes or updates users in legacy public folder ACLs

    - Remove users which haven been deleted in Active Directory and cannot be resolved. Such users show up in ACL as NT User:S-1-*
    - Clean up users and groups
    - Remove user, if object is disabled in Active Directory or is not mail enabled anymore
    - Replace user, if object still exists in AD and is mail enabled, but ACL notation is NT User:DOMAIN\User
    - Remove group, if group is not a mail enabled security group
    
    .NOTES 
    Requirements 
    - Windows Server 2008 R2  
    - Exchange Server 2007/2010
    - Exchange Management Shell

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    1.1     Fixed group replacement logic
    1.2     Script optimzation
    1.3     Updated public fodler handlins
	
    .PARAMETER RootPublicFolder
    Root public folder for recurse checkign of ACLs

    .PARAMETER PublicFolderServer
    Exchange public folder server to query and write to

    .PARAMETER ValidateOnly
    Only validate ACL, do not make any changes. Affects only ACL entries which are not "fully orphaned" users (S-1-*)

    .PARAMETER SkipOrphanedUserCleanup
    Skip orphaned users (SID only users) cleanup
    
    .PARAMETER Recurse
    Check public folder client permission recursively. Without this switch the script will only check the client 
    permissions in the folder defined in RootPublicFolder
     
    .EXAMPLE
    Validate ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX2010
    .\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX2010 -ValidateOnly -Recurse

    .EXAMPLE
    Clean ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701
    .\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701 -Recurse

#>
Param(
  [parameter(Mandatory=$true,HelpMessage='Root public folder for recurse checkign of ACLs')]
  [string]$RootPublicFolder,  
  [parameter(Mandatory=$true,HelpMessage='Public folder server')]
  [string]$PublicFolderServer,
  [switch]$ValidateOnly,
  [switch]$SkipOrphanedUserCleanup,
  [switch]$Recurse
)

Import-Module -Name ActiveDirectory

# Fetch public folders

if($Recurse) {
  Write-Host "Fetching public folders $($RootPublicFolder) recursive"

  $PublicFolders = Get-PublicFolder $RootPublicFolder -Recurse -ResultSize Unlimited -Server $PublicFolderServer -ErrorAction SilentlyContinue
}
else {
  Write-Host "Fetching child public folders in $($RootPublicFolder) only"

  $PublicFolders = Get-PublicFolder $RootPublicFolder -ResultSize Unlimited -Server $PublicFolderServer -GetChildren -ErrorAction SilentlyContinue
}

Write-Host "Found $(($PublicFolders | Measure-Object).Count) public folders"

$FileName = "C:\scripts\Clean-PublicFolderACL\ValidatedPublicFolderPermissions $(Get-Date -Format 'yyyy-MM-dd hh-mm-ss').txt"

Set-Content -Path $FileName -Force -Confirm:$false -Value 'Started!'

$Action = @()

if($PublicFolders -ne $null) {

  # Fetch public folder permissions only once
  $FullPublicFolderPermissions = $PublicFolders | Get-PublicFolderClientPermission

  # Clean orphaned users -> NT User:S-1-*
  if(-not ($SkipOrphanedUserCleanup)) {
    
    Write-Host 'Cleaning orphaned users'
        
    $FullPublicFolderPermissions | Where-Object {$_.User -like 'NT User:S-1-*'} | ForEach-Object { Remove-PublicFolderClientPermission -Identity $_.Identity -User $_.User -Access $_.AccessRights -Confirm:$false}
        
  }
  else {
    Write-Host 'Skipping orphaned user cleanup, but still counting number of orphaned users'
        
    $OrphanedUserCount = ($FullPublicFolderPermissions | Where-Object {$_.User -like 'NT User:S-1-*'} | Measure-Object).Count
        
    Write-Host "| ($($OrphanedUserCount)) ACL objects found"
        
  }

  if($ValidateOnly) {    
    Write-Host 'Checking old users - VALIDATE ONLY'
  }
  else {    
    Write-Host 'Checking old users - with REMOVE/REPLACE'        
  }
    
  # $PublicFolderPermissions = $PublicFolders | Get-PublicFolderClientPermission -Server $PublicFolderServer | Where-Object{$_.User -like "NT User:*"}
  $PublicFolderPermissions = $FullPublicFolderPermissions | Where-Object {($_.User -notlike 'NT User:S-1-*') -and ($_.User -like 'NT User:*')}

  foreach($Permission in $PublicFolderPermissions) {

    # Split ACL entry and get samAccountname only
    [string]$User = ($Permission.User -Replace 'NT User:','').Split('\')[1]
        
    # Cache current public folder identity
    $PFIdentity = $Permission.Identity
        
    try {
      $ADObject = Get-User -Identity $User -ErrorAction SilentlyContinue 
    }
    catch {}
        
    if($ADObject -eq $null) {
      try {
        $ADObject = Get-ADGroup $User -ErrorAction SilentlyContinue
      }
      catch {}    
    }
        
    if($ADObject -ne $null) {
      # AD Lookup success
            
      $Recipient = $null
            
      try {
        $Recipient = Get-Recipient $User -ErrorAction SilentlyContinue
      }
      catch { }

      # Write-Host "$($User) | $($ADObject.ObjectClass)"
                    
      if($ADObject.ObjectClass -contains 'user') {
        # USER
                
        if($User -ne '') {
                
          $ADUser = Get-ADUser -Identity $User
                    
        }
                
        if($ADUser.Enabled -eq $true) {
                

          Write-Host "|  Replace user $($Permission.User) in $($PFIdentity)" 

          $Action += "$($PFIdentity),Replace,$($Permission.User)"
                    
          if(!($ValidateOnly)) {
                                         
            #Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
            # Remove User, not only access rights
            Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Confirm:$false -Server $PublicFolderServer
                    
            if($Recipient -ne $null) {
                        
              Write-Host "|  Add user $($Permission.User) in $($PFIdentity)" 

              $Action += "$($PFIdentity),Add,$($Permission.User)"

              Add-PublicFolderClientPermission -Identity $Permission.Identity -User $User -AccessRights $Permission.AccessRights -Server $PublicFolderServer
            }
            else {

              Write-Host "|  Re-Adding FAILED for user $($Permission.User) in $($PFIdentity)" 
            }
          }
        }
        else {
                        
          Write-Host "|  Remove user $($Permission.User) in $($PFIdentity)"

          $Action += "$($PFIdentity),Remove,$($Permission.User)"
                    
          if(!($ValidateOnly)) {
                                                
            Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
          }
        }
      }
      else {
        # GROUP

        if($Recipient -ne $null) {
                    
          Write-Host "|  Replace group $($Permission.User) in $($PFIdentity)" 

          $Action += "$($PFIdentity),Replace,$($Permission.User)"
                    
          if(!($ValidateOnly)) {
                        
            Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
                                                
            Add-PublicFolderClientPermission -Identity $Permission.Identity -User $User -AccessRights $Permission.AccessRights -Server $PublicFolderServer
          }
        }
        else {
                
          Write-Host "| Remove group $($Permission.User) in $($PFIdentity)" 

          $Action += "$($PFIdentity),Remove,$($Permission.User)"

          if(!($ValidateOnly)) {
                                                
            Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
          }
        }
      }
    }
  }

  $Action | Out-File -FilePath $FileName -Force -Confirm:$false -Encoding utf8

}
else {
  Write-Warning -Message 'No or non-existing public folder specified'
}

Write-Host 'Done'