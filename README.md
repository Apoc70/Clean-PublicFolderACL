# Clean-PublicFolderACL.ps1 
Remove orphaned users and groups from legacy public folder ACLs 

##Description
This scripts removes or updates users in legacy public folder ACLs. This reduces the likelihood of legacy public folder migration errors due to corrupted ACLs.

##Inputs
PARAMETER RootPublicFolder
Root public folder for recurse checkign of ACLs

PARAMETER PublicFolderServer
Exchange public folder server to query and write to

PARAMETER ValidateOnly
Only validate ACL, do not make any changes. Affects only ACL entries which are not "fully orphaned" users (S-1-*)

PARAMETER SkipOrphanedUserCheck
Skip orphaned users check 

##Outputs
NONE.

##Examples
```
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701 -ValidateOnly
```
Validate ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701

```
.\Clean-PublicFolderACL.ps1 -RootPublicFolder "\MYPF" -PublicFolderServer EX200701
```
Clean ACLs on public folder \MYPF and all of it's child public folders on Exchange server EX200701

##TechNet Gallery
Find the script at TechNet Gallery
* 

##Credits
Written by: Thomas Stensitzki

Stay connected:

* My Blog: http://justcantgetenough.granikos.eu
* Twitter:	https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de