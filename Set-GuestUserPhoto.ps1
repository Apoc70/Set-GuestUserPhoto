<#
    .SYNOPSIS

    Set photo thumbnail for Azure AD Guest users.
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, 2019-06-03

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://scripts.Granikos.eu
	
    .DESCRIPTION
    
    This script set the AzureADThumbnailPhoto for guest users to a photo provided as jpg/png file. 
    The file can be up to 100KB in size. This ist currently not checked in the script.
    You can either update a single guest user or all guest users. When updating the user photo can 
    choose to set the photo forcibly or only if there is no photo set.

    The changes are written to a log file. The log file functions are part of the GlobalFunctions module.

    .NOTES 
    Requirements 
    - Utilizes the global function library, found here: http://scripts.granikos.eu
    - AzureAD V2 module (aka AzureAdPreview), found here: https://go.granikos.eu/AzureADv2

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
    .PARAMETER UserPrincipalName
    The UPN of an AzureAD guest user which is normally the guest users external email address

    .PARAMETER FilePath
    The full filepath to the jpg/png file that you want to set

    .PARAMETER GuestUsersToSelect
    Switch to select, if you want to set the photo for a single user or all users
    Single = just a single user
    All = all guest users in your tenant

    .PARAMETER UpdateMode
    The update mode for setting guest user pictures
    OverwriteIfPhotoExists = set the user photo regardless if there is an existing photo
    SetIfNoPhotoExists = set the user photo only, if no user photo exists
  
    .EXAMPLE
    Set the photo ExternalUser.png for all guest users, if no photo exists

    .\Set-GuestUserPhoto.ps1 -FilePath D:\Photos\ExternalUser.png -GuestUsersToSelect All -UpdateMode SetIfNoPhotoExists

    .EXAMPLE
    Set the photo ExternalUser.png for guest user JohnDoe@varunagroup.de, if no photo exists

    .\Set-GuestUserPhoto.ps1 -FilePath D:\Photos\ExternalUser.png -GuestUsersToSelect Single -UserPrincipalName JohnDoe@varunagroup.de

#>
[CmdletBinding()]
Param(
  [string]$UserPrincipalName = '',
  [string]$FilePath = '',
  [ValidateSet('All','Single')] # Available modes for selecting guest user target objects, default: SINGLE
  [string] $GuestUsersToSelect = 'Single',
  [ValidateSet('OverwriteIfPhotoExists','SetIfNoPhotoExists')] # 
  [string] $UpdateMode = 'SetIfNoPhotoExists'
)

# Some variables to declare
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

function Import-RequiredModules {
  <#
    .SYNOPSIS
    Import required PowerShell modules. If the modules are not available, script execution fails.

    .EXAMPLE
    Import-RequiredModules

    .INPUTS
    None

    .OUTPUTS
    None
  #>

  # Import central logging functions 
  if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
    Import-Module -Name GlobalFunctions
  }
  else {
    Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
    Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
    exit
  }
  
  # Import required PowerShell module for Azure AD
  if($null -ne (Get-Module -Name AzureADPreview -ListAvailable).Version) {
        
    # Import the most recent module, if module exists in multiple versions
    Get-Module -Name AzureADPreview -ListAvailable | Select-Object -First 1 | Import-Module
  }
  else {
    # Ooops
    Write-Warning -Message 'Unable to load AzureADPreview (Azure AD Version 2) PowerShell module.'
    exit
  }
}

function Test-PhotoPath {
  [boolean]$PhotoPathValid = $false

  if(($FilePath -ne '') -and (Test-Path -Path $FilePath)) {
    $PhotoPathValid = $true
  }   
  else {
    $message = ("The provided FilePath '{0}' is not valid." -f $FilePath)
    $logger.Write($message,2)
    Write-Warning -Message $message
    exit 
  } 
  $PhotoPathValid
}

function Set-AzureADUserPhoto {
  [CmdletBinding()]
  param(
    [Microsoft.Open.AzureAD.Model.User]$UserObject
  )
  if($UserObject.UserType -eq 'Guest') {
    if($UpdateMode -eq 'OverwriteIfPhotoExists') {
      # set thumbnail photo regardless, if there is an axisting photo
      Set-AzureADUserThumbnailPhoto -ObjectId $UserObject.ObjectId -FilePath $FilePath
      $logger.Write(('AzureADUserThumbnailPhoto set for user {0}' -f $UserObject.MailNickName))
    }
    elseif($UpdateMode -eq 'SetIfNoPhotoExists') {
      $ThumbnailPhoto = $null
      try {
        $ThumbnailPhoto = Get-AzureADUserThumbnailPhoto -ObjectId $UserObject.ObjectId -ErrorAction SilentlyContinue
        $logger.Write(('AzureADUserThumbnailPhoto NOT set for user {0}' -f $UserObject.MailNickName))
      }
      catch {
        $ThumbnailPhoto = $null
      }
      if($null -eq $ThumbnailPhoto) {
        Set-AzureADUserThumbnailPhoto -ObjectId $UserObject.ObjectId -FilePath $FilePath
        $logger.Write(('AzureADUserThumbnailPhoto set for user {0}' -f $UserObject.MailNickName))
      }
    }
  }
}

## MAIN #########################

# Import required modules first
Import-RequiredModules

# create new logger
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Purge() # Purge files based on file retention setting
$logger.Write('Script started')
$logger.Write(('UpdateMode: {0} | GuestUsersToSelect: {1}' -f ($UpdateMode), ($GuestUsersToSelect)))

switch($GuestUsersToSelect) {

  'All' {
    # Update all guest user objects in tenant
    if(Test-PhotoPath) {
      # Issue #1 implement automation capabilities
      $null = Connect-AzureAD

      # fetch all guest users from Azure AD
      $AllGuestUsers = Get-AzureADUser -Filter "UserType eq 'Guest'"
            
      $logger.Write(('Found {0} Azure AD guest users ' -f ($AllGuestUsers | Measure-Object).Count))

      foreach($AzureADUser in $AllGuestUsers) {
        # Set thumbnail photo for user
        Set-AzureADUserPhoto -UserObject $AzureADUser
      }
    }
    else {
      $logger.Write(('Testing file path {0} failed' -f $FilePath),2)
    }
  }

  default {
    if($UserPrincipalName -ne '') {
      if(Test-PhotoPath) {
        # Issue #1 implement automation capabilities
        Connect-AzureAD
        
        # fetch single user from Azure AD
        $AzureADUser = Get-AzureADUser -SearchString $UserPrincipalName.ToLower()
        
        # Set thumbnail photo for user
        Set-AzureADUserPhoto -UserObject $AzureADUser 
      }
      else {
        $logger.Write(('Testing file path {0} failed' -f $FilePath),2)
      }
    }
    else {
      $message = 'You have tried to update a single guest user object, but the UserPrincipal parameter is not set. Please try again.'
      Write-Warning -Message $message
      $logger.Write($message,2)
      exit
    }
  }
}

$logger.Write('Script finished')