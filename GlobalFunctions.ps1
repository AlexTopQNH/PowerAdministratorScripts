# =+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=
# Script      : GeneralFunctions_CreateHostNamedSiteCollections.ps1
# Date        : 28-11-2016
# Version     : 1.10
# Author      : Alex Top, Lise Pijl, Jev Suchoi
# Description : General functions for CreateHostNamedSiteCollections
# Parameters  :
#               Geen
#
# Changes     :   
#               1.0 - 06-01-2016 - AT Initiele versie 
#               1.1 - 26-08-2016 - LP Added cache cleaning functions, create content database function
#               1.2 - 12-09-2016 - AT Read XML aangepast
#               1.3 - 22-09-2016 - LP GF_RDW-CreateContentDB now sets permissions for McAfee service account
#               1.4 - 26-09-2016 - LP GF_RDW-CreateContentDB now sets correct permissions for admin group
#               1.5 - 11-10-2016 - AT GF_RDW-CreateContentDB now sets correct permissions for McAfee account (was Data_Reader; moet zijn SPDataAccess)
#               1.6 - 18-10-2016 - JS GF_DeleteSiteCollectionTermGroup now deletes the terms based on relationship honoring pinning and reuse of terms
#               1.7 - 20-10-2016 - JS GF_DeleteSiteCollection gewijzigd zodat deze niet automatisch GF_DeleteSiteCollectionTermGroup aanroept
#               1.8 - 28-10-2016 - AT GF_RDW-Check-Status-Service toegevoegd
#               1.9 - 24-11-2016 - AT GF_DeleteGlobalTermGroup toegevoegd
#               1.10 - 28-11-2016 - AT GF_Convert-Size toegevoegd en GF_CheckConfigQuotaTemplate geupdate
#
# =+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=

#========================BEGIN GENERIC FUNCTIONS========================================

Function RDW-Write-Host {
    param ([switch]$ErrorFlag)

    $outstr="$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args"
    If ($ErrorFlag) {
	  Write-Host $outstr -Foregroundcolor $ColorError
	} Else {
	  Write-Host $outstr
    }
    $outstr | Out-File $LogFileName -Append -Encoding UTF8
}

Function RDW-Start-Script {
  RDW-Write-Host "Start script $($MyInvocation.scriptname)"
}

Function RDW-Stop-Script {
  RDW-Write-Host "End script $($MyInvocation.scriptname)"
}

Function Read-IniFile {
  param ([string]$IniFile)

  If ($IniFile -eq '') {
    $IniFile = $MyInvocation.ScriptName -replace ".ps1",".ini"
	If ((Get-Item $IniFile -ErrorAction SilentlyContinue) -eq $null) {
	  RDW-Write-Host "No inifile specified and default inifile [$IniFile] not found. Continue without reading inifile"
      Return
    }
  }
  RDW-Write-Host "Reading from inifile $IniFile"
  Get-Content $IniFile | Foreach-Object {
    $line = $_.Split(" ",2)
    If (($line[0] -ne '') -and ($line[0] -ne $null) -and ($line[1] -ne '') -and ($line[1] -ne $null) -and ($line[0][0] -ne '#')) {
	  If ($IniFileHideValue) {
        RDW-Write-Host "Setup variable $($line[0]) (Value not shown because of IniFileHideValue)"
	  } Else {
        RDW-Write-Host "Setup variable $($line[0]) with [$($line[1])]"
      }
      Invoke-Expression "`$global:$($line[0])=$($line[1])"
    }
  }
}

Function Read-XMLFile {
    param ([string]$xmlfile)

    If ($xmlfile -eq '') {
        #$xmlfile = "D:\Beheer\Scripts\SP2013\GlobalFunctions.xml"
        $xmlfile = $MyInvocation.ScriptName -replace ".ps1",".xml"
        If ((Get-Item $xmlfile -ErrorAction SilentlyContinue) -eq $null) {
	        RDW-Write-Host "No xmlfile specified and default xmlfile [$xmlFile] not found. Continue without reading xmlfile"
            Return
        }
    }
    RDW-Write-Host "Reading from xmlfile [$xmlfile]"
    [xml]$Global:XMLConfiguration = Get-Content $xmlfile

    Write-Host "Setup variable [`$XMLContent] with [$($XMLConfiguration.InnerXml)]"
}

Function Read-CfgFile {
    $CfgFile = $MyInvocation.ScriptName -replace ".ps1",".cfg"
	If ((Get-Item $CfgFile -ErrorAction SilentlyContinue) -eq $null) {
	  RDW-Write-Host "No cfgfile specified and default cfgfile [$CfgFile] not found. Continue without reading cfgfile"
      Return
    }

  RDW-Write-Host "Reading from cfgfile [$CfgFile]"
  Get-Content $CfgFile | Foreach-Object {
    $line = $_.Split(" ",2)
    If (($line[0] -ne '') -and ($line[0] -ne $null) -and ($line[1] -ne '') -and ($line[1] -ne $null) -and ($line[0][0] -ne '#')) {
	  If ($IniFileHideValue) {
        RDW-Write-Host "Setup variable $($line[0]) (Value not shown because of IniFileHideValue)"
	  } Else {
        RDW-Write-Host "Setup variable $($line[0]) with [$($line[1])]"
      }
      Invoke-Expression "`$global:$($line[0])=$($line[1])"
    }
  }
}

If ( (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null ) {
    RDW-Write-Host "Begin loading the Microsoft.SharePoint.PowerShell Snapin"
    Add-PSSnapin Microsoft.SharePoint.PowerShell
    RDW-Write-Host "Finished loading the Microsoft.SharePoint.PowerShell Snapin"
}
#========================END GENERIC FUNCTIONS==========================================

#========================START GLOBAL FUNCTIONS=========================================

Function GF_AddSharePointGroup {
    param (
        [string[]]$GroupName = "", 
        [string[]]$GroupDescription = "",
        $GroupPermission = "",
        $SiteCollectionURL = "",
        $GroupOwner = ""
    )

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Add SharePoint Group"

    # Get the web object that requires the new groups
    RDW-Write-Host "`tGet the web object of [$SiteCollectionURL]"
    $SiteCol = Get-SPWeb $SiteCollectionURL

    RDW-Write-Host "`tMake SPUser Object form Member [$GroupOwner]"
    $SPGroupOwner = Get-SPUser -Web $SiteCollectionURL -Identity $GroupOwner

    # Check if the group already exists
    RDW-Write-Host "`tCheck if the group [$GroupName] already exists"
    $CheckGroupName = $SiteCol.SiteGroups[$GroupName]
    If ([string]::IsNullOrWhiteSpace($CheckGroupName)) {
        # Ensure Group/User is part of site collection users beforehand and add them if needed
        $EnsuredUser = $SiteCol.EnsureUser($SPGroupOwner)

        # Create the SharePoint Group – Group Name, Group Owner, Group Member, Group Description. Can’t add AD group yet…
        RDW-Write-Host "`tCreate the SharePoint Group [$GroupName]"
        RDW-Write-Host "`t  with grouppermission [$GroupPermission]"
        $NewSPGroup = $SiteCol.SiteGroups.Add($GroupName, $SPGroupOwner, $SPGroupOwner, $GroupDescription)
        $NewSPGroup = $SiteCol.SiteGroups[$GroupName]
        $NewSPGroup.Update()

        GF_AddSharePointGroupPermissions $GroupPermission $SiteCol $NewSPGroup
        
        RDW-Write-Host "`tRelease the site object"
        $SiteCol.Dispose()

        RDW-Write-Host "`tGroup [$($NewSPGroup.Name)] created!"
    }
    Else {
        RDW-Write-Host "`tGroup [$GroupName] already exists"
    }

    RDW-Write-Host "END Function Add SharePoint Group"
    RDW-Write-Host ""

    return $NewSPGroup
}

Function GF_AddDefaultSharePointGroups {
    param (
        $SiteCollectionURL = "", 
        [object]$XMLDefaultSharePointGroups
    )

    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Add Default SharePoint Groups"
    
    $SPweb = Get-SPWeb $SiteCollectionURL
    $primaryOwner = $SPweb.Site.Owner
    $secondaryOwner = ""

    $SPweb.CreateDefaultAssociatedGroups($primaryOwner, $secondaryOwner, $SPweb.Title)
    $SPweb.Update()

    RDW-Write-Host "`tEND Function Add Default SharePoint Groups"
    RDW-Write-Host ""
    return $status
}

Function GF_AddSharePointGroupPermissions {
    param (
        $GroupPermission = "",
        $SiteCollection, 
        $Group
    )

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Add SharePoint Group Permissions"

    Try{
        # Assign the Group permission
        RDW-Write-Host "`tAssign the Group permission [$GroupPermission]"
        $GroupAssignment = New-Object Microsoft.SharePoint.SPRoleAssignment($Group)
        $GroupRole = $SiteCollection.RoleDefinitions[$GroupPermission]
        $GroupAssignment.RoleDefinitionBindings.Add($GroupRole)
        $SiteCollection.RoleAssignments.Add($GroupAssignment)
        RDW-Write-Host "`tGroup permission [$GroupPermission] assigned"
    }
    Catch {
        RDW-Write-Host "`tAssigning Group Permission failed!"
    }

    RDW-Write-Host "END Function Add SharePoint Group Permissions"
    RDW-Write-Host ""
}

Function GF_AddMembersToGroup {
    param (
        $Group, 
        $Members = @(),
        $SiteCollectionURL
    )

    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Add Members to SharePoint Group"

    RDW-Write-Host "`tGroupName is [$($Group.Name)]"
    RDW-Write-Host "`tMembers are [$Members]"
    RDW-Write-Host "`tSiteCollectionURL is [$SiteCollectionURL]"

    # Get the web object that requires the new groups
    $SiteCol = Get-SPWeb $SiteCollectionURL
    RDW-Write-Host "`tFound SiteCollection: [$($SiteCol.URL)]"

    # Add the AD Group/User to the group, can’t be done during group creation when using Powershell otherwise errors so is done now.
    RDW-Write-Host "`tAdd the Members [$Members] "
    RDW-Write-Host "`t  to the SharePointGroup [$($Group.Name)]"
    Foreach ($Member in $Members) {
        RDW-Write-Host "`tAdd Member [$Member]"
        $User = $SiteCol.EnsureUser($Member)

        Set-SPUser -Identity $User -Web $SiteCollectionUrl -Group $Group
        RDW-Write-Host "`t  Member added with DisplayName [$($User.DisplayName)]"
        RDW-Write-Host "`t    and UserLogin [$($User.UserLogin)]"
    }

    RDW-Write-Host "`tRelease the site object"
    $SiteCol.Dispose()

    RDW-Write-Host "END Function Add Members to SharePoint Group"
    RDW-Write-Host ""
    return $status

}

Function GF_AddAdministratorsToSiteCollection {
    param (
        $Administrators = @(),
        $SiteCollectionURL
    )

    RDW-Write-Host "`tAdministrators are [$Administrators]"
    RDW-Write-Host "`tSiteCollectionURL is [$SiteCollectionURL]"

    $SiteColAdmins = Get-SPSite -Identity $SiteCollectionURL | % {$_.RootWeb.SiteAdministrators}
    RDW-Write-Host "`tThe currents SiteCollection Administators are:"

    Foreach ($SiteColAdmin in $SiteColAdmins) {
        RDW-Write-Host "`t    SiteCollections Administrator [$SiteColAdmin]"
    }
        
    # Add the AD Group/User to the AdministratorGroup
    RDW-Write-Host "`tAdd the Administrators [$Administrators] to the SiteCollection"
    Foreach ($Administrator in $Administrators) {
        RDW-Write-Host "`tAdd Administrator [$Administrator]"
        $site =  Get-SPSite -Identity $SiteCollectionURL
        RDW-Write-Host "`tFound SiteCollection: [$SiteCollectionURL]"
        $web = $site.RootWeb;
        if($web.IsRootWeb) {
            $NewAdmin = $web.EnsureUser($Administrator)
            $NewAdmin.IsSiteAdmin = $true
            $NewAdmin.Update()
            RDW-Write-Host "`t  Administrator added with DisplayName [$($NewAdmin.DisplayName)]"
            RDW-Write-Host "`t    and UserLogin [$($NewAdmin.UserLogin)]"
        }
        Else {
            RDW-Write-Host "`t  FAILED to add Administrator with DisplayName [$($NewAdmin.DisplayName)]"
            RDW-Write-Host "`t    and UserLogin [$($NewAdmin.UserLogin)]"
        }
    }

    return 0
}

Function GF_AddSharePointGroupWithMembers {
    param (
        [string[]]$GroupName = "", 
        [string[]]$GroupDescription = "",
        $GroupPermission = "",
        $SiteCollectionURL = "",
        $GroupOwner = "", 
        $GroupMembers = @(),
        $status = 0
    )

    $CreateSPGroup = GF_AddSharePointGroup $GroupName `
                        $GroupDescription `
                        $GroupPermission `
                        $SiteCollectionURL `
                        $GroupOwner

    $svcstatus = GF_AddMembersToGroup $CreateSPGroup `
                        $GroupMembers `
                        $SiteCollectionURL
    if ($svcstatus -ne 0) { $status=1}
    return $status
}

Function GF_CheckFarmQuotaTemplatesAgainstConfigFile {
    Param(
        [object]$XMLConfigQuotaTemplates,
        $status = 0
    )

    # Declaration variables
     [Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint") 

    # Get a reference to the content service
    $SPFarmQuotaTemplates = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.QuotaTemplates

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Check Farm Quota Templates against Default Quota Templates in Config File"
    RDW-Write-Host "`t========================================================================================"
    Foreach ($SPFarmQuotaTemplate in $SPFarmQuotaTemplates){
        RDW-Write-Host "`t  Current Farm Template Name is: [$($SPFarmQuotaTemplate.Name)]"
        RDW-Write-Host "`t  BEGIN Check if Template [$($SPFarmQuotaTemplate.Name)] is in Config File"
        $TemplateExist = $XMLConfigQuotaTemplates.SelectNodes('QuotaTemplate[@Name="' + $SPFarmQuotaTemplate.Name + '"]')
        If ($TemplateExist.Name -eq $null) {
            RDW-Write-Host "`t    ==========================================================================================="
            RDW-Write-Host "`t    Template [$($SPFarmQuotaTemplate.Name)] should not be in Farm QuotaTemplate Collection"
            RDW-Write-Host "`t    Remove Template [$($SPFarmQuotaTemplate.Name)] through Central Administration or"
            RDW-Write-Host "`t      add Template [$($SPFarmQuotaTemplate.Name)] to the default QuotaTemplate Collection (XML)"
            RDW-Write-Host "`t    ==========================================================================================="
            $status = 2
        }
        Else {
            RDW-Write-Host "`t    ...Template [$($SPFarmQuotaTemplate.Name)] does exist in Config File"
            $status = 2
        }
        RDW-Write-Host "`t  END Check if Template [$($SPFarmQuotaTemplate.Name)] is in Config File"
        RDW-Write-Host ""
    }
    RDW-Write-Host "`t========================================================================================"
    RDW-Write-Host "`tEND Function Check Farm Quota Templates against Default Quota Templates in Config File"
    return $status
}

Function GF_CheckConfigQuotaTemplate {
    Param(
        [object]$XMLConfigQuotaTemplates,
        $status = 0
    )

    # Declaration variables
     [Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint") 

    # Get a reference to the content service
    $SPFarmQuotaTemplates = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.QuotaTemplates

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Check Default Quota Templates in Config File against Farm Quota Templates "
    RDW-Write-Host "`t========================================================================================="

    # Loop throught XMLQuotaTemplates and check if they already exists
    RDW-Write-Host "`t  Check if default XMLQuotaTemplates are registerd in the farm:"
    Foreach ($XMLConfigQuotaTemplate in $XMLConfigQuotaTemplates.ChildNodes) {
        RDW-Write-Host "`t    BEGIN Current Config QuotaTemplate is: [$($XMLConfigQuotaTemplate.Name)]"
        If ($SPFarmQuotaTemplates[$XMLConfigQuotaTemplate.Name] -eq $null) {
            # Default XMLQuotaTemplate does not exists in farm
            RDW-Write-Host "`t`tDefault XMLQuotaTemplate is not registerd in farm"

            # Create QuotaTemplate
            # Instantiate an instance of an SPQuotaTemplate class #
            $SPQuotaTemplate = New-Object Microsoft.SharePoint.Administration.SPQuotaTemplate

            # Declaration variables
            [int64]$XMLStorageMaximumLevelMB = [int64]$XMLConfigQuotaTemplate.StorageMaximumLevelGB *1024 *1024 *1024
            [int64]$XMLStorageWarningLevelMB = [int64]$XMLConfigQuotaTemplate.StorageWarningLevelGB *1024 *1024 *1024

            # Set the Properties #
            RDW-Write-Host "`t`tAdd Quota Template [$($XMLConfigQuotaTemplate.Name)] with:"
            RDW-Write-Host "`t`t  Maximum Level Storage: [$XMLStorageMaximumLevelMB)]"
            RDW-Write-Host "`t`t  Warning Level Storage: [$XMLStorageWarningLevelMB)]"
            RDW-Write-Host "`t`t  Maximum Level UserCode: [$($XMLConfigQuotaTemplate.UserCodeMaximumLevel)]"
            RDW-Write-Host "`t`t  Warning Level UserCode: [$($XMLConfigQuotaTemplate.UserCodeWarningLevel)]"

            $SPQuotaTemplate.Name = $XMLConfigQuotaTemplate.Name
            $SPQuotaTemplate.StorageMaximumLevel = [int64]$XMLStorageMaximumLevelMB
            $SPQuotaTemplate.StorageWarningLevel = [int64]$XMLStorageWarningLevelMB
            $SPQuotaTemplate.UserCodeMaximumLevel = [double]$XMLConfigQuotaTemplate.UserCodeMaximumLevel
            $SPQuotaTemplate.UserCodeWarningLevel = [double]$XMLConfigQuotaTemplate.UserCodeWarningLevel

            RDW-Write-Host "`t`tQuota Template created!"

            # Get an Instance of the SPWebService Class 
            $Service = [Microsoft.SharePoint.Administration.SPWebService]::ContentService

            # Use the Add() method to add the quota template to the collection #
            RDW-Write-Host "`t`tAdding the [$($XMLDefaultQuotaTemplate.Name)] Quota Template to the Quota Templates Collection"
            $Service.QuotaTemplates.Add($SPQuotaTemplate)
            RDW-Write-Host "`t`tQuota Template [$($XMLDefaultQuotaTemplate.Name)] added to the Quota Templates Collection"
            RDW-Write-Host

            # Call the Update() method to commit the changes #
            $svcstatus = $Service.Update()
            if ($svcstatus -ne $null) { $status=1}
        }
        Else {
            # Default XMLQuotaTemplate exists in farm
            RDW-Write-Host ""
            RDW-Write-Host "`t`tBEGIN Check the Config of Quota Templates in the Farm against the Config File"
            RDW-Write-Host "`t`t====================================================================================="

            # Declaration variables
            [int64]$XMLStorageMaximumLevelMB = GF_Convert-Size -From GB -To Bytes -Value $XMLConfigQuotaTemplate.StorageMaximumLevelGB
            [int64]$XMLStorageWarningLevelMB = GF_Convert-Size -From GB -To Bytes -Value $XMLConfigQuotaTemplate.StorageWarningLevelGB


            # Check if QuotaTemplate in the Farm is conform Config Quota Template File           
            If ($($SPFarmQuotaTemplates[$XMLConfigQuotaTemplate.Name].StorageMaximumLevel) -eq $XMLStorageMaximumLevelMB -and `
                $($SPFarmQuotaTemplates[$XMLConfigQuotaTemplate.Name].StorageWarningLevel) -eq $XMLStorageWarningLevelMB -and `
                $($SPFarmQuotaTemplates[$XMLConfigQuotaTemplate.Name].UserCodeMaximumLevel) -eq $($XMLConfigQuotaTemplate.UserCodeMaximumLevel) -and `
                $($SPFarmQuotaTemplates[$XMLConfigQuotaTemplate.Name].UserCodeWarningLevel) -eq $($XMLConfigQuotaTemplate.UserCodeWarningLevel)) {
                RDW-Write-Host "`t`t  QuotaTemplate in the Farm is EQUAL to the ConfigQuotaTemplate"
            }
            Else {
                RDW-Write-Host ""
                RDW-Write-Host "`t`t  QuotaTemplate in the Farm is NOT EQUAL to the ConfigQuotaTemplate"
                
                GF_UpdateQuotaTemplate $XMLConfigQuotaTemplate
            }
            
            RDW-Write-Host ""
            RDW-Write-Host "`t`t====================================================================================="
            RDW-Write-Host "`t`tEND Check the Config of Quota Templates in the Farm against the Config File"
            
        }
        RDW-Write-Host "`t    END Current Config QuotaTemplate is: [$($XMLConfigQuotaTemplate.Name)]"
        RDW-Write-Host""
    }
    RDW-Write-Host "`tEND Function Check Default Quota Templates in Config File against Farm Quota Templates "
    RDW-Write-Host "`t========================================================================================="
    RDW-Write-Host ""

    return $status
}

Function GF_Convert-Size
{            
    [cmdletbinding()]            
    param(            
        [validateset("Bytes","KB","MB","GB","TB")]
        [string]$From,
        [validateset("Bytes","KB","MB","GB","TB")]
        [string]$To,
        [Parameter(Mandatory=$true)]
        [double]$Value,
        [int]$Precision = 4
    )

    switch($From)
    {
        "Bytes" {$value = $Value }
        "KB" {$value = $Value * 1024 }
        "MB" {$value = $Value * 1024 * 1024}
        "GB" {$value = $Value * 1024 * 1024 * 1024}
        "TB" {$value = $Value * 1024 * 1024 * 1024 * 1024}
    }
            
    switch ($To)
    {
        "Bytes" {return $value}            
        "KB" {$Value = $Value/1KB}            
        "MB" {$Value = $Value/1MB}            
        "GB" {$Value = $Value/1GB}            
        "TB" {$Value = $Value/1TB}            
    }            
            
    return [Math]::Round($value,$Precision,[MidPointRounding]::AwayFromZero)

}            
   

Function GF_UpdateQuotaTemplate {
    Param(
        [object]$XMLConfigQuotaTemplate,
        $status = 0
    )

    # Declaration variables
    [int]$status=0
    [Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint") 
    [int64]$XMLStorageMaximumLevelBytes = [math]::Round([double]$XMLConfigQuotaTemplate.StorageMaximumLevelGB *1024 *1024 *1024)
    [int64]$XMLStorageWarningLevelBytes = [math]::Round([double]$XMLConfigQuotaTemplate.StorageWarningLevelGB *1024 *1024 *1024)

    RDW-Write-Host ""
    RDW-Write-Host "`t`t  BEGIN Function Update Quota Template"
    RDW-Write-Host "`t`t  =================================="
    RDW-Write-Host "`t`t    The FarmQuotaTemplate to update is: [$($XMLConfigQuotaTemplate.Name)] with:"
    RDW-Write-Host "`t`t      XMLConfigQuotaTemplate.StorageMaximumLevelBytes is: [$XMLStorageMaximumLevelBytes]"
    RDW-Write-Host "`t`t      XMLConfigQuotaTemplate.StorageWarningLevelBytes is: [$XMLStorageWarningLevelBytes]"
    RDW-Write-Host "`t`t      XMLConfigQuotaTemplate.UserCodeWarningLevel is : [$($XMLConfigQuotaTemplate.UserCodeWarningLevel)]"
    RDW-Write-Host "`t`t      XMLConfigQuotaTemplate.UserCodeMaximumLevel is : [$($XMLConfigQuotaTemplate.UserCodeMaximumLevel)]"

    Try {
        # Get a reference to the content service
        $service = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
        $QTemp = $service.QuotaTemplates[$XMLConfigQuotaTemplate.Name]
        $QTemp.Name = $XMLConfigQuotaTemplate.Name
        $QTemp.StorageWarningLevel = [int64]$XMLStorageWarningLevelBytes
        $QTemp.StorageMaximumLevel = [int64]$XMLStorageMaximumLevelBytes
        $QTemp.UserCodeWarningLevel = [int32]$XMLConfigQuotaTemplate.UserCodeWarningLevel
        $QTemp.UserCodeMaximumLevel = [int32]$XMLConfigQuotaTemplate.UserCodeMaximumLevel
        $service.Update()
    }
    Catch {
        RDW-Write-Host "`t`t    Something went wrong with updating FarmQuotaTemplate"
        $status = 1
    }

    RDW-Write-Host "`t`t  =================================="
    RDW-Write-Host "`t`t  END Function Upate Quota Template"
    return $status
}

Function GF_DeleteQuotaTemplate {
    Param(
        [object]$QuotaTemplate,
        $svcstatus = "",
        $status = 0
    )

    # Declaration variables
    [int]$status=0
    [Void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint") 
    $TemplateName = $QuotaTemplate.Name
    # Get a reference to the content service
    $service = [Microsoft.SharePoint.Administration.SPWebService]::ContentService

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Delete Quota Template"

    If ($service.QuotaTemplates[$TemplateName] -ne $null) {
        # Set the Properties #
        RDW-Write-Host "`tDelete Quota Template [$TemplateName]"
        $svcstatus = $Service.QuotaTemplates.Delete($TemplateName)
        if ($svcstatus -ne 0) { $status=1}

        # Update the service
        $Service.Update()
        RDW-Write-Host "`tQuota Template [$TemplateName] deleted"
    }
    Else {
        RDW-Write-Host "`tQuota Template [$TemplateName] not found!"
    }

    RDW-Write-Host "END Function Delete Quota Template"
    RDW-Write-Host ""
    return $status
}

Function GF_CheckUserDefinedSiteCollection {
    Param(
        [object]$SiteCollections
    )

    [array]$NonUserDefinedSiteCollections = "CENTRALADMIN", "SPSMSITEHOST", "SPSPERS"
    $OwnerAlias = Get-SPUser -Web $SiteCollections[0].Url | Where {$_.UserLogin -like "i:0#.w|*\Colla2Install*"}

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Check Primary SiteCollections Owners"
    RDW-Write-Host "`tOverall SiteCollection Primary SiteCollection Owner is: [$OwnerAlias]"

    Foreach ($SiteCollection in $SiteCollections) {
        # Check if SiteCollection is a User Defined SiteCollection
        $SiteCol = Get-SPWeb $SiteCollection.Url
        If ($NonUserDefinedSiteCollections -contains $($SiteCol.WebTemplate)) {
            RDW-Write-Host "`t  SiteCollection [$($SiteCol.Url)] is a Non User Defined SiteCollection"
            RDW-Write-Host "`t  No need to check the Primary SiteCollection Administrator"
        }
        Else {
            # Get the current OwnerAlias
            $GETSPSite = Get-SPSite -Identity $SiteCollection.ID
            $SiteCollectionOwner = $GETSPSite.Owner.UserLogin
            $SiteCollectionURL = $GETSPSite.Url
            RDW-Write-Host "`t  For SiteCollection [$SiteCollectionURL]"
            RDW-Write-Host "`t    the current owner is [$SiteCollectionOwner]"

            If ($SiteCollectionOwner -ne $OwnerAlias) {
                RDW-Write-Host "`t  Begin Details========================================================="
                RDW-Write-Host ""
                RDW-Write-Host "`t  Current SiteCollectionOwner is not the default SiteCollectionOwner"
                RDW-Write-Host "`t  Change it to [$OwnerAlias]"
            
                # Set Default OwnerAlias
                Set-SpSite -Identity $SiteCollection.ID -OwnerAlias $OwnerAlias
                $status=2
                RDW-Write-Host ""
                RDW-Write-Host "`t  End Details==========================================================="
            }
            Else {
                RDW-Write-Host "`t  Current SiteCollectionOwner is the default SiteCollectionOwner"
                RDW-Write-Host ""
            }
        }
        $SiteCol.close()
    }
    RDW-Write-Host "END Function Check Primary SiteCollections Owners"
    RDW-Write-Host ""
    return $status

}

Function GF_AddQuotaTemplateToSiteCollection {
    Param(
        [string]$SiteCollectionURL,
        [string]$QuotaTemplate
    )
    
    # Declare variables
    $svcstatus = ""
    $status = 0
    
    $QuotaTemplateToSet = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.QuotaTemplates |  Where-Object {$_.Name -eq $QuotaTemplate}  
    RDW-Write-Host "`tAdd QuotaTemplate [$($QuotaTemplateToSet.Name)]"
    RDW-Write-Host "`t  to SiteCollection [$SiteCollectionURL]"
    $svcstatus = Set-SPSite -Identity $SiteCollectionURL -QuotaTemplate $QuotaTemplateToSet

    if ($svcstatus -ne $null) { $status=1}
    
    RDW-Write-Host "`tQuotaTemplate Added!"
    return $status

}

Function GF_AddRoleDefinition {
    Param(
        [object]$XMLRoleDefinition,
        [object]$SPSiteCollection
    )

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Add SharePoint Role Definition"
    RDW-Write-Host "`t  Check if Role Definition [$($XMLRoleDefinition.Name)] already exist"
    # returns all the subwebs in a given site collection
    $SPWeb = $SPSiteCollection | Get-SPWeb

    # $RoleDefinition is 'Machtigingsniveaus'
    If($SPWeb.RoleDefinitions[$XMLRoleDefinition.Name] -eq $null) {
        RDW-Write-Host "`t  SharePoint Role Definition [$($XMLRoleDefinition.Name)] doesn't exist!"
        RDW-Write-Host "`t  Create SharePoint Role Definition [$($XMLRoleDefinition.Name)]"
        RDW-Write-Host "`t   with BasePermissions [$($XMLRoleDefinition.BasePermission)]"
        # Create Role Definition
        $SPRoleDefinition = New-Object Microsoft.SharePoint.SPRoleDefinition
        # Set Role Definition Name
        $SPRoleDefinition.Name = $XMLRoleDefinition.Name
        # Set Role Definition Description
        $SPRoleDefinition.Description = $XMLRoleDefinition.Description
        
        $SPRoleDefinition.BasePermissions = $XMLRoleDefinition.BasePermission

        # Add RoleDefinitions to SiteCollection
        $SPWeb.RoleDefinitions.Add($SPRoleDefinition)

        $SPWeb.Dispose()
        $SPSiteCollection.Dispose()
        RDW-Write-Host "`t  SharePoint Role Definition [$($XMLRoleDefinition.Name)] created"
    }
    Else {
        RDW-Write-Host "`t  SharePoint Role Definition [$($XMLRoleDefinition.Name)] already exist!"
    }    

    RDW-Write-Host "`tEND Function Add SharePoint Role Definition"
    RDW-Write-Host ""
    return $XMLRoleDefinition.Name
}

Function GF_UpdateRoleDefinition {
    Param(
        [object]$XMLRoleDefinition,
        [object]$SPSiteCollection
    )

    RDW-Write-Host ""
    RDW-Write-Host "BEGIN Function Update Role Definition"


    RDW-Write-Host "END Function Update Role Definition Level"
    RDW-Write-Host ""
}

Function GF_AddCustomSiteCollectionProperties {
    Param(
        [string]$SiteCollectionURL,
        [array]$INIPropertyNames,
        [array]$INIPropertyValues
    )

    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Add Custom SiteCollection Properties"

    $myWeb = Get-SPWeb -Identity $SiteCollectionURL
    $myWeb.AllowUnsafeUpdates = "true"

    # Check if PropertyNames en PropertieValues are the samen count
    If ($INIPropertyNames.Count -ne $INIPropertyValues.Count ) {
        throw "PropertyNames en PropertyValues in .ini file have not the samen count!!"
    }

    # Check if Custom Property already exist
    For ($i = 0; $i -lt $INIPropertyNames.Count; $i++) {
        RDW-Write-Host "`t  Check if Custom SiteCollection Property [$($INIPropertyNames[$i])] already exist"
        [string]$PropertyName = $INIPropertyNames[$i]
        [string]$PropertyValue = $INIPropertyValues[$i]
        $SPPropertyName = $myWeb.Properties[$PropertyName]
        
        If ($SPPropertyName -ne $null) {
            RDW-Write-Host "`t   SPPropertyValue [$PropertyName] exsist with value [$SPPropertyName]"
            #$myWeb.DeleteProperty($PropertyName)
        }
        Else {
            RDW-Write-Host "`t   SPPropertyValue [$PropertyName] doesn't exist!"
            RDW-Write-Host "`t   Create Custom Propertie [$PropertyName] with value [$PropertyValue]"

            $myWeb.AllProperties.Add($PropertyName, $PropertyValue)

            RDW-Write-Host "`t   Custom Propertie [$PropertyName] created with value [$PropertyValue]"
        }
    }

    $myWeb.Update()
    $myWeb.AllowUnsafeUpdates = "false"
    $myWeb = $null
    return $status
 }

Function GF_DeleteSiteCollectionTermGroup {
    Param(
        [string]$SiteCollectionURL
    )
 
    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Delete Site Collection TermGroup"
    RDW-Write-Host "`t with URL [$SiteCollectionURL]"

    RDW-Write-Host "`t  Connect to SiteCollection"
    $SiteColl = Get-SPSite $SiteCollectionURL
    #$SiteColl = Get-SPSite "https://intranet.ot.tld/sites/test01"

    # Connect to Term Store in the Managed Metadata Service Application
    RDW-WRite-Host "`t  Connect to site with MMS service connection"
    $taxonomySession = Get-SPTaxonomySession -site $SiteColl
    
    RDW-Write-Host "`t  Connect to Term Store in the Managed Metadata Service Application"
    $termStore = $taxonomySession.TermStores[$taxonomySession.DefaultSiteCollectionTermStore.Name]
    $newSiteGrp = $termStore.GetSiteCollectionGroup($SiteColl)

    if($newSiteGrp -ne $null)
    {
        # Get the current count of terms in the Orphaned Terms
        $BeforeDeletingTermOrphanedTerms = $termStore.OrphanedTermsTermSet.Terms.Count
        RDW-Write-Host "`t  Current terms in the Orphaned Terms before deleting is: [$BeforeDeletingTermOrphanedTerms]"

        RDW-Write-Host "`t  Get the SiteCollectionTermStoreGroup"
        # the option $false means the no Group is added. When you add option $true, TermStoreGroup 
        # for SiteCollection is created.
        $termStoreGroup = $termStore.GetSiteCollectionGroup($SiteColl,$false)

        $allTerms = $null
        [int]$ctr = 0
        if($termStoreGroup -ne $null)
        {
            if($termStoreGroup.TermSets.Count -ne 0)
            {
                if($termStoreGroup.TermSets.Count -gt 0)
                {
                    #Build a collection of all terms in the current SiteGroup
                    ForEach($termSet in $termStoreGroup.TermSets)
                    {
                        $allTerms += $termSet.GetAllTerms()
                    }
                }

                else
                {
                    $termStoreGroup.TermSets | foreach{ $_.Delete() }
                }
            }

            RDW-Write-Host "Nr of terms pending for deletion: " $allTerms.Count
            
            <#        
                Function to delete terms by recursion, making sure the pinned root terms get deleted first
                and not their parents, otherwise orphaned items will occur
            #>
            while($allTerms.Count -gt 0)
            {
                # Get all Terms that do not have children and are not reused or pinned
                $notParentNotReusedTerms = $allTerms | Where-Object { $_.Terms.Count -eq 0 -and $_.IsReused -eq $false }

                # Delete them
                $notParentNotReusedTerms | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name}
                $termStore.CommitAll()

                # Filter out the just deleted terms from our collection
                $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }

                # Get all not Source Terms that are not pinned
                $notSourceTerm = $allTerms | Where-Object { $_.IsSourceTerm -eq $false -and $_.IsPinned -eq $false }

                # Delete them
                $notSourceTerm | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name }
                $termStore.CommitAll()

                # Filter out the just deleted terms from our collection
                $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }

                #Get all non Source Terms that are Pinned Root
                $pinnedRoot = $allTerms | Where-Object { $_.IsSourceTerm -eq $false -and $_.IsPinnedRoot -eq $true }

                # Delete them
                $pinnedRoot | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name }
                $termStore.CommitAll()

                # Filter out the just deleted terms from our collection
                $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }
            }

            RDW-Write-Host "`t  Nr of terms pending for deletion: [$($allTerms.Count)]"

            # Delete all TermSets, these should be empty by now
            $termStoreGroup.TermSets | foreach { $_.Delete() }

            # Finally delete the Term Group
            $termStoreGroup.Delete()
            $termStore.CommitAll()
        }

    
        # Get the current count of terms in the Orphaned Terms
        $AfterDeletingTermOrphanedTerms = $termStore.OrphanedTermsTermSet.Terms.Count
        RDW-Write-Host "`t  Current terms in the Orphaned Terms after deleting is: [$AfterDeletingTermOrphanedTerms]"

        If ($AfterDeletingTermOrphanedTerms -gt $BeforeDeletingTermOrphanedTerms) {
            RDW-Write-Host "========================================================"
            RDW-Write-Host ""
            Write-Host "$(Get-Date -format 'yyyy-dd-MM HH:mm:ss')" -NoNewLine
            Write-Host -ForegroundColor Red "`tLET OP: There were Terms (re)used in the Global TermStore."
            Write-Host "$(Get-Date -format 'yyyy-dd-MM HH:mm:ss')" -NoNewLine
            Write-Host -ForegroundColor Red "`tLET OP: These Terms are now Orphaned!"
            Write-Host "$(Get-Date -format 'yyyy-dd-MM HH:mm:ss')" -NoNewLine
            Write-Host -ForegroundColor Red "`tLET OP: Go to Central Admin Managed MetaData Service"
            Write-Host "$(Get-Date -format 'yyyy-dd-MM HH:mm:ss')" -NoNewLine
            Write-Host -ForegroundColor Red "`tLET OP:   and decide what to do with this terms."
            RDW-Write-Host ""
            RDW-Write-Host "========================================================"
        }

        $termStore.CommitAll()

        RDW-Write-Host "`tEND Function Delete Site Collection TermGroup"
        RDW-Write-Host "`t with URL [$SiteCollectionURL]"
        RDW-Write-Host ""   
    }

    else
    {
        RDW-Write-Host "`t No Site Collection Term Group found, nothing to delete"
    }

    return $status
}

Function GF_DeleteGlobalTermGroup {
    Param(
        $termStoreGroup
    )
 
    $status = 0

    if($termStoreGroup -ne $null)
    {
        if($termStoreGroup.TermSets.Count -ne 0)
        {
            if($termStoreGroup.TermSets.Count -gt 0)
            {
                #Build a collection of all terms in the current SiteGroup
                ForEach($termSet in $termStoreGroup.TermSets)
                {
                    $allTerms += $termSet.GetAllTerms()
                }
            }

            else
            {
                Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
                Write-Host ""
                Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
                Write-Host -Foregroundcolor $ColorAttention "About to START Delete TermGroup [$termStoreGroup]"
                Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
                Write-Host -Foregroundcolor $ColorAttention "Are you sure (Y/N) " -NoNewline
                If ((Read-Host).ToUpper() -ne "Y")
                {
	                Throw "Delete TermGroup [$termStoreGroup] cancelled by user"
	            }
                $termStoreGroup.TermSets | foreach{ $_.Delete() }
            }
        }

        RDW-Write-Host "Nr of terms pending for deletion: " $allTerms.Count
            
        <#        
            Function to delete terms by recursion, making sure the pinned root terms get deleted first
            and not their parents, otherwise orphaned items will occur
        #>

        Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
        Write-Host ""
        Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
        Write-Host -Foregroundcolor $ColorAttention "About to START Delete TermGroup [$termStoreGroup]"
        Write-Host "$(Get-Date -format 'yyyy-MM-dd HH:mm:ss') $args" -NoNewline
        Write-Host -Foregroundcolor $ColorAttention "Are you sure (Y/N) " -NoNewline
        If ((Read-Host).ToUpper() -ne "Y")
        {
	        Throw "Delete TermGroup [$termStoreGroup] cancelled by user"
	    }

        while($allTerms.Count -gt 0)
        {
            # Get all Terms that do not have children and are not reused or pinned
            $notParentNotReusedTerms = $allTerms | Where-Object { $_.Terms.Count -eq 0 -and $_.IsReused -eq $false }

            # Delete them
            $notParentNotReusedTerms | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name}
            $termStore.CommitAll()

            # Filter out the just deleted terms from our collection
            $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }

            # Get all not Source Terms that are not pinned
            $notSourceTerm = $allTerms | Where-Object { $_.IsSourceTerm -eq $false -and $_.IsPinned -eq $false }

            # Delete them
            $notSourceTerm | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name }
            $termStore.CommitAll()

            # Filter out the just deleted terms from our collection
            $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }

            #Get all non Source Terms that are Pinned Root
            $pinnedRoot = $allTerms | Where-Object { $_.IsSourceTerm -eq $false -and $_.IsPinnedRoot -eq $true }

            # Delete them
            $pinnedRoot | ForEach-Object { $_.Delete(); RDW-Write-Host "Deleted Term: " $_.Name }
            $termStore.CommitAll()

            # Filter out the just deleted terms from our collection
            $allTerms = $allTerms | Where-Object { $_.TermSet -ne $null }
        }

        RDW-Write-Host "`t  Nr of terms pending for deletion: [$($allTerms.Count)]"

        # Delete all TermSets, these should be empty by now
        $termStoreGroup.TermSets | foreach { $_.Delete() }

        # Finally delete the Term Group
        $termStoreGroup.Delete()
        $termStore.CommitAll()
    }
}

Function GF_DeleteSiteCollection {
    Param(
        [string]$SiteCollectionURL
    )
 
    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Delete Site Collection"
    RDW-Write-Host "`t with URL [$SiteCollectionURL]"

    <#
        Confirm:$False is Without confirmation
        Without -GradualDelete to make sure that  the site gets deleted right away
        instead of going trough Deleted ites (Trash Bin)
    #>
    Remove-SPSite -Identity "$SiteCollectionUrl" -Confirm:$False 

    RDW-Write-Host "`tEND Function Delete Site Collection"
    RDW-Write-Host "`t with URL [$SiteCollectionURL]"
    RDW-Write-Host ""

    return $status
}

Function GF_ChangeHNSCURL {
    Param(
        [string]$CurrentSiteCollectionURL,
        [string]$NewSiteCollectionURL
    )
 
    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Change URL of HNSC"
    RDW-Write-Host "`t with Current URL [$CurrentSiteCollectionURL]"
    RDW-Write-Host "`t to URL [$NewSiteCollectionURL]"

    RDW-Write-Host "`t Get Site object"
    $site = Get-SPSite -Identity $CurrentSiteCollectionURL

    RDW-Write-Host "`t Change URL"
    $uri = New-Object System.Uri($NewSiteCollectionURL)
    $site.Rename($uri)
    RDW-Write-Host "`t URL changed!"

    RDW-Write-Host "`t Refresh Sites in Configuration Database"
    ((Get-SPSite $NewSiteCollectionURL).contentdatabase).RefreshSitesInConfigurationDatabase

    RDW-Write-Host "`tEND Function Change URL of HNSC"
    RDW-Write-Host "`tThe new URL is [$NewSiteCollectionURL]"
    RDW-Write-Host ""

    return $status

}

Function GF_Try-RDW-Start-ExeFile {
    param (
        [string]$Cmd
    )

    [int]$stat=0
    try {
        GF_RDW-Start-ExeFile $Cmd
    }
    catch {
        RDW-Write-Host ERROR : $_.Exception.message
        $stat=1
    }
}

Function GF_RDW-Start-ExeFile {
  param([string]$Cmd)

  RDW-Write-Host "Starting Exe [$Cmd]"
  Try {
    $ps = new-object System.Diagnostics.Process
    $ps.StartInfo.Filename = $Cmd
    $ps.StartInfo.RedirectStandardOutput = $True
    $ps.StartInfo.UseShellExecute = $false
    $ps.StartInfo.Arguments = "/passive /norestart"
    $tmp = $ps.start()
    $ps.StandardOutput.ReadToEnd().Split("`n") | ForEach-Object {RDW-Write-Host $_}
    $tmp = $ps.WaitForExit()
    RDW-Write-Host "Ending Exe [$Cmd] with Exitcode [$($ps.ExitCode)]"

    $PSExitCode = $($ps.ExitCode)
    If ($($ps.ExitCode) -ne "0") {
        switch ($PSExitCode) { 
            # ExitCode 17025 mains already installed; not real error..
            "17025" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [Already installed]"
                Return 0
            }
            # ExitCode 17028 mains No product installed for contained patch..
            "17028" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [No product installed for contained patch]"
                Return 0
            }
            # ExitCode 17031 mains Invalid baseline which implies that it cannot locate the MS Office Installation media
            "17031" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [Invalid baseline which implies that it cannot locate the MS Office Installation media]"
                Return 0
            }
        }
    }
    Else {
        Return $ps.ExitCode
    }
  }
  Catch {
    $ErrorMessage = $_.Exception.Message
    #$FailedItem = $_.Exception.ItemName
    Throw "Failed to execute [$Cmd] with ErrorMessage [$ErrorMessage]!"
  }
}

Function GF_Try-RDW-Stop-Service {
    param(
        [string]$svcName
    )

    [int]$stat=0
    try {
        GF_RDW-Stop-Service $svcName
    }
    catch {
        RDW-Write-Host ERROR : $_.Exception.message
        $stat=1
    }
    return $stat
}

Function GF_RDW-Stop-Service {
    param(
        [string]$svcName,
        [string]$serverName = $env:COMPUTERNAME
    )

    
    $svc = Get-WMIObject Win32_Service -ComputerName $serverName | Where-Object {$_.Name -eq $svcName}
    RDW-Write-Host "Starting stopping service [$svcName] from server $servername with Start mode [$($svc.startmode)] State [$($svc.state)]..."
    
    if ( $svc -ne $null ) {
        RDW-Write-Host "  Start mode [$($svc.startmode)] State [$($svc.state)]"
        if ($svc.startmode -ne "Disabled") {
            if ($svc.state -eq "Running") {
                $dummy=$svc.stopservice()
                $svc = Get-WMIObject Win32_Service -ComputerName $serverName| Where-Object {$_.Name -eq $svcName}
                RDW-Write-Host "  Start mode [$($svc.startmode)] State [$($svc.state)]"
 
                # wait? for state stopped
                [int]$i = 1
                while ( ($i –le 15) -AND ($svc.state -ne "Stopped") ) {
                    start-sleep -seconds 5
                    $svc = Get-WMIObject Win32_Service -ComputerName $serverName| Where-Object {$_.Name -eq $svcName}
                    RDW-Write-Host "  retry [$i] Start mode [$($svc.startmode)] State [$($svc.state)]"
                    $i= $i+1
                }
                if ($svc.state -eq "Stopped") {
                    RDW-Write-Host "Service [$svcName] stopped." 
                }
                else {
                    # Dit moet een throw worden
                    Throw "Service [$svcName] NOT stopped. Last state [$($svc.state)]" 
                }
            }
            else {
                RDW-Write-Host "  Service [$svcName] not running." 
            }
        }
        else { 
            RDW-Write-Host "Service [$svcName] disabled." 
        } 
    }
    else { 
        Throw "Service [$svcName] not found." 
    } 
}

Function GF_Try-RDW-Start-Service {
    param(
        [string]$svcName
    )

    [int]$stat=0
    try {
        GF_RDW-Start-Service $svcName
    }
    catch {
        RDW-Write-Host ERROR : $_.Exception.message
        $stat=1
    }
    return $stat
}

Function GF_RDW-Start-Service {
    param (
        [string]$svcName,
        [string]$serverName = $env:COMPUTERNAME
    )

    $svc = Get-WMIObject Win32_Service -ComputerName $serverName | Where-Object {$_.Name -eq $svcName}
    RDW-Write-Host "Starting starting service [$svcName] from server $serverName with Start mode [$($svc.startmode)] State [$($svc.state)]..."

    if ( $svc -ne $null ) {
        RDW-Write-Host "  Start mode [$($svc.startmode)] State [$($svc.state)]"
        if ($svc.startmode -eq "Auto") {
            if ($svc.state -ne "Running") {
                $dummy=$svc.startservice()
                $svc = Get-WMIObject Win32_Service -ComputerName $serverName| Where-Object {$_.Name -eq $svcName}
                RDW-Write-Host "  Start mode [$($svc.startmode)] State [$($svc.state)]"
 
                # wait? for state Running
                [int]$i = 1
                while ( ($i –le $CheckRetriesServiceState) -AND ($svc.state -ne "Running") ) {
                    start-sleep -seconds $CheckIntervalServicecState
                    $svc = Get-WMIObject Win32_Service -ComputerName $serverName| Where-Object {$_.Name -eq $svcName}
                    RDW-Write-Host "  retry [$i] Start mode [$($svc.startmode)] State [$($svc.state)]"
                    $i= $i+1
                }
                if ($svc.state -eq "Running") {
                    RDW-Write-Host "Service [$svcName] started." 
                }
                else {
                    # Dit moet een throw worden
                    Throw "Service [$svcName] NOT started. Last state [$($svc.state)]" 
                }
            }
            else {
                RDW-Write-Host "Service [$svcName] already running." 
            }
        }
        else { 
            RDW-Write-Host "Service [$svcName]. Not in Auto mode" 
        } 
    }
    else { 
        Throw "Service [$svcName] not found." 
    } 
}

Function GF_RDW-SQLExecuteNonQuery {
  param([string]$SQLServerInstance,$Database,$Query)

  # Create SqlConnection object and define connection string
  $con = New-Object System.Data.SqlClient.SqlConnection
  $con.ConnectionString = "Server=$SQLServerInstance;Database=$Database;Integrated Security=true"
  $con.Open()

  # Create SqlCommand object, define command text, and set the connection
  $cmd = New-Object System.Data.SqlClient.SqlCommand
  $cmd.CommandText = $Query
  $cmd.Connection = $con

  $result = $cmd.ExecuteNonQuery()

  $con.Close()
}

Function GF_RDW-Start-ExeFile {
    param([string]$Cmd)

    RDW-Write-Host "Starting Exe [$Cmd]"
    $ps = new-object System.Diagnostics.Process
    $ps.StartInfo.Filename = $Cmd
    $ps.StartInfo.RedirectStandardOutput = $True
    $ps.StartInfo.UseShellExecute = $false
    $ps.StartInfo.Arguments = "/passive /norestart"
    $tmp = $ps.start()
    $ps.StandardOutput.ReadToEnd().Split("`n") | ForEach-Object {RDW-Write-Host $_}
    $tmp = $ps.WaitForExit()

    $PSExitCode = $($ps.ExitCode)

    If ($PSExitCode -ne "0") {
        RDW-Write-Host "PSExitCode is : [$PSExitCode]"
        switch ($PSExitCode) { 
            # ExitCode 17025 mains already installed; not real error..
            "17025" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [Already installed]"
                Return 0
            }
            # ExitCode 17028 mains No product installed for contained patch..
            "17028" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [No product installed for contained patch]"
                Return 0
            }
            # ExitCode 17031 mains Invalid baseline which implies that it cannot locate the MS Office Installation media
            "17031" {
                RDW-Write-Host "Exitcode [$($ps.ExitCode)] mains [Invalid baseline which implies that it cannot locate the MS Office Installation media]"
                Return 0
            }
            Default {
                Throw "Ending Exe [$Cmd] with Exitcode [$($ps.ExitCode)]"
            }
        }
    }
    Else {
        Return $PSExitCode
    }
}

Function GF_Try-RDW-Start-ExeFile {
  param([string]$Cmd) 

  [int]$stat=0
  try {
    GF_RDW-Start-ExeFile $Cmd
  }
  catch {
     RDW-Write-Host ERROR : $_.Exception.message
     $stat=1
  }
}

Function GF_RDW-BackUpSPSite {
    param(
        [string]$SiteCollectionUrl,
        [string]$BackUpLocation,
        [string]$SiteCollectionName
    )

    $status = 0

    $SiteCollectionName += ".bak"
    $fullBackupPath = $BackUpLocation.TrimEnd('\') + "\" +  $SiteCollectionName

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Create SiteCollection BackUp"
    RDW-Write-Host "`t for URL [$SiteCollectionUrl]"

    $svcstatus = Backup-SPSite -Identity $SiteCollectionUrl -Path $fullBackupPath -Force

    if ($svcstatus -ne $null) { $status=1}

    RDW-Write-Host "`tEND Function Create SiteCollection BackUp"
    RDW-Write-Host "`t with URL [$SiteCollectionUrl]"
    RDW-Write-Host ""

    return $status
}

Function GF_RDW-RestoreSPSite {
    param(
        [string]$SiteCollectionUrl,
        [string]$BackUpLocation,
        [string]$SiteCollectionName
    )

    $status = 0

    $SiteCollectionName += ".bak"
    #[string]$BackUpLocation +=  $SiteCollectionName
    $hostheaderwebapplicationurl = Get-SPWebApplication
    $hostheaderwebapplicationurl = $hostheaderwebapplicationurl.url

    RDW-Write-Host ""
    RDW-Write-Host "`tBEGIN Function Restore SiteCollection BackUp"
    RDW-Write-Host "`t for URL [$SiteCollectionUrl]"
    RDW-Write-Host "`t from location [$BackUpLocation]"
    RDW-Write-Host "`t with hostheader webapplication url [$hostheaderwebapplicationurl]"

    RDW-Write-Host "`t==SiteCollectionUrl is [$SiteCollectionUrl]"
    RDW-Write-Host "`t==BackUpLocation is [$BackUpLocation]"
    RDW-Write-Host "`t==hostheaderwebapplicationurl is [$hostheaderwebapplicationurl]"

    $svcstatus = Restore-SPSite -Identity $SiteCollectionUrl -Path $BackUpLocation -HostHeaderWebApplication $hostheaderwebapplicationurl -Force -Confirm:$false
    #$svcstatus = Restore-SPSite -Identity $SiteCollectionUrl -Path $BackUpLocation -Force -Confirm:$false

    if ($svcstatus -ne $null) { $status=1}

    RDW-Write-Host "`tEND Function Restore SiteCollection BackUp"
    RDW-Write-Host "`t with URL [$SiteCollectionUrl]"
    RDW-Write-Host ""

    return $status
}

Function GF_RPT_ContentDB_WriteXML {
    param(
        [string]$FilePath,
        [string]$StartElement
    )

    $status = 0

    RDW-Write-Host "  START Create XML file [$FilePath]"
    Try {
        # Create The Document
        $XmlWriter = New-Object System.Xml.XmlTextWriter($FilePath,$null)
    }
    catch {
        throw "Aanmaken xml file [$FilePath] is mislukt."
    }
        
    RDW-Write-Host "  ..file created!"
    RDW-Write-Host "  END Create XML file [$FilePath]"

    # Set The Formatting
    $XmlWriter.Formatting = "Indented"
    $XmlWriter.Indentation = "4"

    # Write The XML Decleration
    $XmlWriter.WriteStartDocument()

    # Write Root Element
    $XmlWriter.WriteStartElement($StartElement)

    RDW-Write-Host ""
    RDW-Write-Host "  START Writing ContentDB XML File [$FilePath]"
    $SPContentDBS = Get-SPContentDatabase
    try {
        foreach ($SPContentDB in $SPContentDBS) {
            $XmlWriter.WriteStartElement("ContentDatabase")
            $XmlWriter.WriteElementString("Name", $SPContentDB.Name)
            $XmlWriter.WriteElementString("NumberOfSiteCollections", $SPContentDB.CurrentSiteCount)
            
            $TotalSizeNeeded = 0
            $QuotaTemplates = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.quotatemplates
            $SiteColls = Get-SPSite -Limit All -ContentDatabase $SPContentDB.Name
            $XmlWriter.WriteStartElement("SiteCollections")
            foreach ($SiteColl in $SiteColls) {
                $XmlWriter.WriteStartElement("SiteCollection")
                $XmlWriter.WriteElementString("URL", $SiteColl.Url)
                
                $UsageStorageGB = $SiteColl.Usage.Storage/1GB
                $UsageStorageGB = "{0:n0}"-f $UsageStorageGB
                $XmlWriter.WriteElementString("TotalSizeUsed", $UsageStorageGB)

                $SPQuotaID = ($SiteColl.Quota).QuotaID
                $QuotaTemplateName = “No Template Applied”
                foreach($QuotaTemplate in $QuotaTemplates) {
                    If($QuotaTemplate.QuotaId -eq $SPQuotaID) {
                        $QuotaTemplateMaxSize = $QuotaTemplate.StorageMaximumLevel/1GB
                        $QuotaTemplateMaxSize = "{0:n0}"-f $QuotaTemplateMaxSize
                    }
                }
                $TotalSizeNeeded += $QuotaTemplateMaxSize
                
                $XmlWriter.WriteEndElement() # <-- Closing SiteCollection
            }
            $XmlWriter.WriteEndElement()# <-- Closing SiteCollections
            $XmlWriter.WriteElementString("TotalSizeReservedInGB", $TotalSizeNeeded)
            $XmlWriter.WriteEndElement() # <-- Closing ContentDatabase
        }
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        $status = 1
    }

    # Write Close Tag for Root Element
    $XmlWriter.WriteEndElement()# <-- Closing ContentDatabases

    # End the XML Document
    $XmlWriter.WriteEndDocument()

    # Finish The Document
    $XmlWriter.Finalize
    $XmlWriter.Flush()
    $XmlWriter.Close()
    RDW-Write-Host "  ..file writed"
    RDW-Write-Host "  END Writing ContentDB XML File [$FilePath]"

    return $status
}

Function GF_RPT_ContentDB_WriteHTML {
    param(
        [array]$NameOfReports,
        [string]$LocationConfigFile,
        [string]$LocationOfSourceXML
    )

    $status = 0

    RDW-Write-Host ""
    RDW-Write-Host "  START Create .html file(s) for:"

    foreach ($NameOfReport in $NameOfReports) {
        RDW-Write-Host "`t-$NameOfReport"
        $Name = $LocationConfigFile + $NameOfReport
        $NameXSLT = $Name + ".xslt"
        RDW-Write-Host "`tNameXSLT is : [$NameXSLT]"

        $FilePath = $LocationOfSourceXML
       
        RDW-Write-Host "`tName Of Source XML is : [$FilePath]"

        $NameHTML = $LocationReports + $NameOfReport + ".html"
        RDW-Write-Host "`tNameHTML is : [$NameHTML]"

        Try {
            $xslt = New-Object System.Xml.Xsl.XslCompiledTransform
            $xslt.Load($NameXSLT)
            $xslt.Transform($LocationOfSourceXML,$NameHTML)
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $status = 1
        }
        RDW-Write-Host ""
        
        # Add files to Document Library
        RDW-Write-Host "`tSTART Adding files to Document Library"
        Try {
            # GF_RPT_Add_To_DocLibrary NameHTML
            $svcstatus = GF_RPT_Add_To_DocLibrary $NameHTML
            If ($svcstatus -ne 0) { $status=1}
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $status = 1
        }
        RDW-Write-Host "`t..done"
        RDW-Write-Host "`tEND Adding files to Document Library"
        RDW-Write-Host "========================================================================"
        RDW-Write-Host ""
    }
    RDW-Write-Host "  ..done"
    RDW-Write-Host "  END Create .html file(s) for:"

    return $status
}

Function GF_RPT_Add_To_DocLibrary {
    param(
        [string]$LocationFile
    )

    $status = 0

    # Set Variables
    $WebURL = "https://intranet.ot.tld/sites/samenwerkingsplatform"
    $DocLibName = "Gedeelde%20%20documenten"
    $FilePath = $LocationFile

    # Get a variable that points to the folder 
    $Web = Get-SPWeb $WebURL 
    $List = $Web.GetFolder($DocLibName) 
    $Files = $List.Files 

    # Get just the name of the file from the whole path 
    $FileName = $FilePath.Substring($FilePath.LastIndexOf("\")+1) 

    # Load the file into a variable 
    $File= Get-ChildItem $FilePath

    # Upload it to SharePoint 
    $Files.Add($DocLibName +"/" + $FileName,$File.OpenRead(),$false) 
    $web.Dispose()

    return $status

}

Function GF_RDW-Clear-Cache {
    param (
        [string] $serverName = $env:COMPUTERNAME
    )
    RDW-Write-Host "Start clearing XML config cache from server" $serverName "..." 
    $server = Get-SPServer -Identity $serverName

    foreach($instance in $server.ServiceInstances) {
    # If the server has the timer service delete the XML files from the config cache
        if($instance.TypeName -eq "Microsoft SharePoint Foundation Timer") {

            RDW-Write-Host "  Deleting xml files from config cache on server" $serverName

            # Remove all xml files recursive on an UNC path
            $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\*.xml"
            Remove-Item -path $path -Force

            break
        }
    }
    RDW-Write-Host "..done" 
}

function GF_RDW-Clear-TimerCache {
    param(
        [string] $serverName = $env:COMPUTERNAME
    )

    $server = Get-SPServer -Identity $serverName
    RDW-Write-Host "Start clearing timer cache on server" $serverName"..." 

    [string] $path = ""

    # Iterate through each service in each server
  
    foreach($instance in $server.ServiceInstances) {
    # If the server has the timer service then force the cache settings to be refreshed
        if($instance.TypeName -eq "Microsoft SharePoint Foundation Timer") {
            
            RDW-Write-Host "  Clearing timer cache from server" $serverName

            # Clear the cache on an UNC path
            # 1 = refresh all cache settings
            $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\cache.ini"
            Set-Content -path $path -Value "1"
    
            break   
        }
    }
    RDW-Write-Host "..done" 
}

function GF_RDW-Clear-FarmCache {
    RDW-Write-Host "Start clearing farm cache..." 
    $farm = Get-SPServer | where {$_.Role -match "Application"}

    foreach ($server in $farm) {
        GF_RDW-Stop-Service -svcName "SPTimerv4" -serverName $server.Name   
    }
    
    foreach ($server in $farm) {
        GF_RDW-Clear-Cache -serverName $server.Name 
    }
    
    foreach ($server in $farm) {
        GF_RDW-Clear-TimerCache -serverName $server.Name 
    }


    foreach ($server in $farm) {
        GF_RDW-Start-Service -svcname "SPTimerv4" -serverName $server.Name
    }

    RDW-Write-Host "Finished clearing farm cache" 
}

Function GF_RDW-RunGradualSiteDeletion {
        $status = 0
        $timerJob = Get-SPTimerJob “job-site-deletion”
        $prevRunTime = $timerjob.LastRunTime
        RDW-Write-Host "Starting timerjob gradual site deletion. Previous starttime: $prevruntime" 

        [int] $counter = 0
   
        #Run job for the duration of 10 minutes
        Start-SPTimerJob $timerJob
        $timerJobSucceeded = $true
        while ($prevRunTime –eq $timerJob.LastRunTime -and ($counter -le 40)) {
	        Write-Host -NoNewLine “.”
	        $counter++;
	        Start-Sleep –Seconds 15

            If ($counter > 40) {
                RDW-Write-Host "Timerjob site deletion is taking more than 10 minutes to complete. Continuing script."
                $timerjobsucceeded = $false
                $status = 1
                break;
            }
        }

        If ($timerJobSucceeded) {
            Write-Host ""
            RDW-Write-Host “..done at" $timerjob.LastRunTime
        }
        RDW-Write-Host ""
        return $status
    }

Function GF_RDW-CreateContentDB {
        param (
         [Parameter(Mandatory=$true)] $DBName, 
         [Parameter(Mandatory=$true)] $DBServer,
         [Parameter(Mandatory=$true)] $WebApp,
         [Parameter(Mandatory=$true)] $AdminGroup,
         [Parameter(Mandatory=$true)] $McAfeeSA,
         [Parameter(Mandatory=$true)] $DBOwner
        )

        RDW-Write-Host "Creating contentdatabase $DBName on $DBServer for $WebApp using admin group $AdminGroup and installer account $DBOwner"
        New-SPContentDatabase -Name $DBName -DatabaseServer $DBServer -WebApplication $WebApp | Out-Null

        # Create SharePoint_Shell_Access Database Role if needed
        $Query = @"
IF NOT EXISTS (
SELECT 1
FROM sys.database_principals
WHERE name='SharePoint_Shell_Access' AND Type = 'R'
) BEGIN
  CREATE ROLE [SharePoint_Shell_Access]
END
"@
        RDW-Write-Host "Create SharePoint_Shell_Access Database Role if needed"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

        If ($AdminGroup -ne $null) {


        # Create Database user for Farm Admin Group if needed
        $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_principals
  WHERE name = '$AdminGroup' AND Type = 'G'
) BEGIN
  CREATE USER [$AdminGroup] FOR LOGIN [$AdminGroup]
END
"@
        RDW-Write-Host "Create Database user for Farm Admin Group if needed"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

        # Assign Sharepoint_Shell_Access Role to Farm Admin Group if needed
        $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_role_members drm
    JOIN sys.database_principals rp ON (drm.role_principal_id = rp.principal_id)
    JOIN sys.database_principals mp ON (drm.member_principal_id = mp.principal_id)
  WHERE rp.name = 'SharePoint_Shell_Access' AND mp.name = '$AdminGroup'
) BEGIN
  ALTER ROLE [SharePoint_Shell_Access] ADD MEMBER [$AdminGroup]
END
"@
        RDW-Write-Host "Assign Sharepoint_Shell_Access Role to Farm Admin Group if needed"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

        # Assign SPDataAccess Role to Farm Admin Group if needed
        $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_role_members drm
    JOIN sys.database_principals rp ON (drm.role_principal_id = rp.principal_id)
    JOIN sys.database_principals mp ON (drm.member_principal_id = mp.principal_id)
  WHERE rp.name = 'SPDataAccess' AND mp.name = '$AdminGroup'
) BEGIN
  ALTER ROLE [SPDataAccess] ADD MEMBER [$AdminGroup]
END
"@
        RDW-Write-Host "Assign SPDataAccess Role to Farm Admin Group if needed"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

        # Assign db_owner Role to Farm Admin Group 
        $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_role_members drm
    JOIN sys.database_principals rp ON (drm.role_principal_id = rp.principal_id)
    JOIN sys.database_principals mp ON (drm.member_principal_id = mp.principal_id)
  WHERE rp.name = 'db_owner' AND mp.name = '$AdminGroup'
) BEGIN
  ALTER ROLE [db_owner] ADD MEMBER [$AdminGroup]
END
"@
        RDW-Write-Host "Assign SPDataAccess Role to Farm Admin Group if needed"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query
        }


        If ($McAfeeSA -ne $null) {

          # Create Database user for Mcafee if needed
          $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_principals
  WHERE name = '$McAfeeSA' AND Type = 'U'
) BEGIN
  CREATE USER [$McAfeeSA] FOR LOGIN [$McAfeeSA]
END
"@
          RDW-Write-Host "Create Database user for McAfee service account if needed"
          GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

          # Assign db_datareader Role to Mcafee if needed
          $Query = @"
IF NOT EXISTS (
  SELECT 1
  FROM sys.database_role_members drm
    JOIN sys.database_principals rp ON (drm.role_principal_id = rp.principal_id)
    JOIN sys.database_principals mp ON (drm.member_principal_id = mp.principal_id)
  WHERE rp.name = 'SPDataAccess' AND mp.name = '$McAfeeSA'
) BEGIN
  ALTER ROLE [SPDataAccess] ADD MEMBER [$McAfeeSA]
END
"@
          RDW-Write-Host "Assign SPDataAccess Role to McAfee service account if needed"
          GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query
        }

        $Query = @"
EXEC dbo.sp_changedbowner @loginame = N'$DBOwner', @map = false
"@

        RDW-Write-Host "Set $DBOwner as owner of the database"
        GF_RDW-SQLExecuteNonQuery $DBServer $DBName $Query

        RDW-Write-Host "..done"

    }

Function GF_RDW-Check-Status-Service {
    param (
        [string]$svcName
    )
  
    RDW-Write-Host "Check status of service [$svcName]..."

    $svc = Get-WMIObject Win32_Service | Where-Object {$_.Name -eq $svcName}
    If ( $svc -ne $null ) {
        RDW-Write-Host "  Start mode [$($svc.startmode)] State [$($svc.state)]"
        If ($svc.startmode -ne "Disabled") {
            If ($svc.state -eq "Stopped") {
                RDW-Write-Host "Service [$svcName] has the stopped status."
                Return 0
            } Else {
                RDW-Write-Host "=============================================="
				RDW-Write-Host "Service [$svcName] has not the stopped status!"
				RDW-Write-Host "=============================================="
		        Return 1
        }
        } Else { 
            RDW-Write-Host "Service [$svcName] disabled."
	        Return 0
        }
    } Else { 
        Throw "Service [$svcName] not found." 
    } 
}

#========================END GLOBAL FUNCTIONS===========================================