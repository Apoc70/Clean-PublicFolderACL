# Clean-PublicFolderACL.ps1

Remove orphaned users and groups from legacy public folder ACLs 

## Description

This scripts removes or updates users in legacy public folder ACLs. This reduces the likelihood of legacy public folder migration errors due to corrupted ACLs.

## Parameters

### RootPublicFolder

Root public folder for recurse checkign of ACLs

### PublicFolderServer

Exchange public folder server to query and write to

### ValidateOnly

Only validate ACL, do not make any changes. Affects only ACL entries which are not "fully orphaned" users (S-1-*)

### SkipOrphanedUserCheck

Skip orphaned users check

## Examples

``` PowerShell
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701 -ValidateOnly
```

Validate ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701

``` PowerShell
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701
```

Clean ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery

Find the script at TechNet Gallery

* [https://gallery.technet.microsoft.com/Remove-orphaned-users-and-bba62a39 ](https://gallery.technet.microsoft.com/Remove-orphaned-users-and-bba62a39)

## Credits

Written by: Thomas Stensitzki

Stay connected:

* My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
* Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
* LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
* Github: [https://github.com/Apoc70](https://github.com/Apoc70)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
* Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
* Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)