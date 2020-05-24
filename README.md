# Set-GuestUserPhoto.ps1

Set photo thumbnail for Azure AD Guest users

## Description

This script sets the AzureADThumbnailPhoto for guest users to a photo provided as jpg/png file. The file can be up to 100KB in size. This ist currently not checked in the script.

You can either update a single guest user or all guest users. When updating the user photo can choose to set the photo forcibly or only if there is no photo set.

The changes are written to a log file. The log file functions are part of the GlobalFunctions module.

## Requirements

- Utilizes the global function library, found here: [http://scripts.granikos.eu](http://scripts.granikos.eu)
- AzureAD V2 module (aka AzureAdPreview), found here: [https://go.granikos.eu/AzureADv2](https://go.granikos.eu/AzureADv2)

## Parameters

### UserPrincipalName

The UPN of an AzureAD guest user which is normally the guest users external email address

### FilePath

The full filepath to the jpg/png file that you want to set

### GuestUsersToSelect

Switch to select, if you want to set the photo for a single user or all users

- Single = just a single user
- All = all guest users in your tenant

### UpdateMode

The update mode for setting guest user pictures

- OverwriteIfPhotoExists = set the user photo regardless if there is an existing photo
- SetIfNoPhotoExists = set the user photo only, if no user photo exists

## Examples

### Example 1

Set the photo ExternalUser.png for all guest users, if no photo exists

``` PowerShell
.\Set-GuestUserPhoto.ps1 -FilePath 'D:\Photos\ExternalUser.png' -GuestUsersToSelect All -UpdateMode SetIfNoPhotoExists
```

### Example 2

Set the photo ExternalUser.png for guest user JohnDoe@varunagroup.de, if no photo exists

``` PowerShell
.\Set-GuestUserPhoto.ps1 -FilePath D:\Photos\ExternalUser.png -GuestUsersToSelect Single -UserPrincipalName JohnDoe@varunagroup.de
```

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)