<#
.SYNOPSIS
Remove orphaned users and groups from legacy public folder ACLs 
   
Thomas Stensitzki
	
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
Version 1.1, 2016-11-30

Ideas, comments and suggestions to support@granikos.eu 
 
.LINK  
More information can be found at http://scripts.granikos.eu

.DESCRIPTION
This scripts removes or updates users in legacy public folder ACLs
    
.NOTES 
Requirements 
- Windows Server 2008 R2  
- Exchange Server 2007/2010
- Exchange Management Shell

Revision History 
-------------------------------------------------------------------------------- 
1.0     Initial community release 
1.1     Fixed group replacement logic
	
.PARAMETER RootPublicFolder
Root public folder for recurse checkign of ACLs

.PARAMETER PublicFolderServer
Exchange public folder server to query and write to

.PARAMETER ValidateOnly
Only validate ACL, do not make any changes. Affects only ACL entries which are not "fully orphaned" users (S-1-*)

.PARAMETER SkipOrphanedUserCleanup
Skip orphaned users (SID only users) cleanup  
 
.EXAMPLE
Validate ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701 -ValidateOnly

.EXAMPLE
Clean ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701

#>
Param(
    [parameter(Mandatory=$true,HelpMessage='Root public folder for recurse checkign of ACLs')]
        [string]$RootPublicFolder,  
    [parameter(Mandatory=$true,HelpMessage='Public folder server')]
        [string]$PublicFolderServer,
    [parameter(Mandatory=$false)]
        [switch]$ValidateOnly,
    [parameter(Mandatory=$false)]
        [switch]$SkipOrphanedUserCleanup
)

Import-Module ActiveDirectory

# Fetch public folders
Write-Host "Fetching Public Folders $($RootPublicFolder)"
$PublicFolders = Get-PublicFolder $RootPublicFolder -Recurse -ResultSize Unlimited -Server $PublicFolderServer -ErrorAction SilentlyContinue

if($PublicFolders -ne $null) {

    # Clean orphaned users
    if(-not ($SkipOrphanedUserCleanup)) {
    
        Write-Host 'Cleaning orphaned users'
        $PublicFolders | Get-PublicFolderClientPermission | Where-Object{$_.User -like "NT User:S-1-*"} | ForEach-Object {Remove-PublicFolderClientPermission -Identity $_.Identity -User $_.User -Access $_.AccessRights -Confirm:$false}
        
    }
    else {
        $OrphanedUserCount = ($PublicFolders | Get-PublicFolderClientPermission | Where-Object{$_.User -like "NT User:S-1-*"} | Measure-Object).Count
        Write-Host "Orphaned user cleanup skipped! ($($OrphanedUserCount)) ACL objects found"
    }

    if($ValidateOnly) {
        Write-Host 'Checking old users - VALIDATE ONLY'
    }
    else {
        Write-Host 'Checking old users - with REMOVE/REPLACE'
    }
    $PublicFolderPermissions = $PublicFolders | Get-PublicFolderClientPermission -Server $PublicFolderServer | Where-Object{$_.User -like "NT User:*"}


    foreach($Permission in $PublicFolderPermissions) {
        if(-Not ([string]$Permission.User).StartsWith('NT User:S-1-*')) {
            [string]$User = ($Permission.User -Replace 'NT User:','').Split('\')[1]
        }
        
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
            # Lookup success
            
            $Recipient = $null
            
            try {
                $Recipient = Get-Recipient $User -ErrorAction SilentlyContinue
            }
            catch { }

            Write-Verbose "$($User) | $($ADObject.ObjectClass)"
                    
            if($ADObject.ObjectClass -contains 'user') {
                # USER
                
                if($User -ne '') {
                
                    $ADUser = Get-ADUser -Identity $User
                    
                }
                else {
                    Write-Host "| SKIPPING user $($Permission.User) in $($PFIdentity)" 
                    break
                }
                
                if($ADUser.Enabled -eq $true) {
                
                    Write-Host "| Replace user $($Permission.User) in $($PFIdentity)" 
                    
                    if(!($ValidateOnly)) {
                        
                        Write-Verbose "|  Remove user $($Permission.User) in $($PFIdentity)" 
                        Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
                    
                        if($Recipient -ne $null) {
                        
                            Write-Verbose "|  Add user $($Permission.User) in $($PFIdentity)" 
                            Add-PublicFolderClientPermission -Identity $Permission.Identity -User $User -AccessRights $Permission.AccessRights -Server $PublicFolderServer
                        }
                        else {
                            Write-Verbose "|  Re-Adding FAILED for user $($Permission.User) in $($PFIdentity)" 
                        }
                    }
                }
                else {
                
                    Write-Host "| Remove user $($User) from $($PFIdentity)"    
                    
                    if(!($ValidateOnly)) {
                        
                        Write-Verbose "|  Remove user $($Permission.User) in $($PFIdentity)" 
                        
                        Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
                    }
                }
            }
            else {
                # GROUP

                if($Recipient -ne $null) {
                    
                    Write-Host "| Replace group $($Permission.User) in $($PFIdentity)" 
                    
                    if(!($ValidateOnly)) {
                    
                        Write-Verbose "|  Remove group $($Permission.User) in $($PFIdentity)" 
                        Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
                        
                        Write-Verbose "|  Add group $($Permission.User) in $($PFIdentity)" 
                        Add-PublicFolderClientPermission -Identity $Permission.Identity -User $User -AccessRights $Permission.AccessRights -Server $PublicFolderServer
                    }
                }
                else {
                
                    Write-Host "| Remove group $($Permission.User) in $($PFIdentity)" 

                    if(!($ValidateOnly)) {
                    
                        Write-Verbose "|  Remove group $($Permission.User) in $($PFIdentity)"
                        Remove-PublicFolderClientPermission -Identity $Permission.Identity -User $Permission.User -Access $Permission.AccessRights -Confirm:$false -Server $PublicFolderServer
                    }
                }

            }
        }
    }
}
else {
    Write-Warning 'No or non-existing public folder specified'
}

Write-Host "Done"