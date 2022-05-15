<#
    Script: Module-SPOnlineInventory.ps1
    Author: Kris Seque (Github: https://github.com/RecursiveKea/SharePointOnlineInventory)    
#>

Param (
     [Parameter(Mandatory=$true)][String]$ModulePath
)

. ($ModulePath + "\Module-SPOnlineHelpers.ps1") -force;

#Requires -Modules PnP.PowerShell;
Add-Type -AssemblyName System.Web #Encoding XML

function Get-SPOnlineInventorySettings
{
    [cmdletbinding()]
	param (                                          
          [Parameter(Mandatory=$true)]$PnPCredentialParams
         ,[Parameter(Mandatory=$true)][String]$LogFilePrefix
         ,[Parameter(Mandatory=$true)][String]$ExportFolder
         ,[Parameter(Mandatory=$true)][String]$Delimiter
    );

    $strCurrentTimeZone = (Get-WmiObject win32_timezone).StandardName;
    $TimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById($strCurrentTimeZone);

    return (New-Object PSCustomObject -Property @{
        LogFilePrefix = $LogFilePrefix;
        ExportFolder = $ExportFolder;
        Delimiter = $Delimiter;
        PnPCredentialParams = $PnPCredentialParams;
        DateFormat = "dd-MMM-yyyy HH:mm:ss.fff";
        TimeZoneForExport = $TimeZone;
        ExcludedFieldList = @{            
            "_CopySource" = 1;
            "_EditMenuTableEnd" = 1;
            "_EditMenuTableStart" = 1;
            "_EditMenuTableStart2" = 1;
            "_HasCopyDestinations" = 1;
            "_ModerationComments" = 1;
            "_ModerationStatus" = 1;
            "_SharedFileIndex" = 1;
            "_SourceUrl" = 1;
            "_UIVersion" = 1;
            "_UIVersionString" = 1;
            "AppAuthor" = 1;
            "AppEditor" = 1;
            "BaseName" = 1;
            "CheckedOutTitle" = 1;
            "CheckedOutUserId" = 1;
            "Combine" = 1;
            "ContentTypeId" = 1;
            "ContentType" = 1;
            "Created_x0020_By" = 1;
            "Created_x0020_Date" = 1;
            "DocConcurrencyNumber" = 1;
            "DocIconEdit" = 1;
            "EncodedAbsUrl" = 1;
            "File_x0020_Size" = 1;
            "File_x0020_Type" = 1;
            "FileDirRef" = 1;
            "FileLeafRef" = 1;
            "FileRef" = 1;
            "FileSizeDisplay" = 1;
            "FolderChildCount" = 1;
            "FSObjType" = 1;
            "GUID" = 1;
            "HTML_x0020_File_x0020_Type" = 1;
            "InstanceID" = 1;
            "IsCheckedoutToLocal" = 1;
            "ItemChildCount" = 1;
            "Last_x0020_Modified" = 1;
            "LinkCheckedOutTitle" = 1;
            "LinkFilename" = 1;
            "LinkFilename2" = 1;
            "LinkFilenameNoMenu" = 1;
            "MetaInfo" = 1;
            "Modified_x0020_By" = 1;            
            "owshiddenversion" = 1;
            "ParentLeafName" = 1;
            "ParentVersionString" = 1;
            "PermMask" = 1;
            "ProgId" = 1;
            "RepairDocument" = 1;
            "ScopeId" = 1;
            "SelectFilename" = 1;
            "SelectTitle" = 1;
            "ServerUrl" = 1;
            "SortBehavior" = 1;
            "SyncClientId" = 1;
            "TemplateUrl" = 1;
            "UniqueId" = 1;
            "_VirusInfo" = 1;
            "_VirusVendorID" = 1;
            "_VirusStatus" = 1;
            "WorkflowInstanceID" = 1;
            "WorkflowVersion" = 1;
            "xd_ProgID" = 1;
            "xd_Signature" = 1;
            #"_dlc_DocId" = 1; #Commented out for checking if the document library has the DOC ID column
            "_dlc_DocIdPersistId" = 1;

            "_ComplianceFlags" = 1;
            "_ComplianceTag" = 1;
            "_ComplianceTagUserId" = 1;
            "_ComplianceTagWrittenTime" = 1;
            "_IsCurrentVersion" = 1;
            "_IsRecord" = 1;
            "_Level" = 1;
            "AccessPolicy" = 1;
            "ComplianceAssetId" = 1;
            "NoExecute" = 1;
            "OriginatorId" = 1;
            "PrincipalCount" = 1;
            "Restricted" = 1;
            "SMLastModifiedDate" = 1;
            "SMTotalFileCount" = 1;
            "SMTotalFileStreamSize" = 1;
            "SMTotalSize" = 1;
            "_CheckinComment" = 1;
            "_CommentCount" = 1;
            "_CommentFlags" = 1;
            "_Dirty" = 1;
            "_DisplayName" = 1;
            "_ExpirationDate" = 1;
            "_HasEncryptedContent" = 1;
            "_IpLabelAssignmentMethod" = 1;
            "_IpLabelHash" = 1;
            "_IpLabelId" = 1;
            "_IpLabelPromotionCtagVersion" = 1;
            "_LikeCount" = 1;
            "_ListSchemaVersion" = 1;
            "_Parsable" = 1;
            "_RmsTemplateId" = 1;
            "_StubFile" = 1;
            "A2ODMountCount" = 1;
            "BSN" = 1;
            "CheckoutUser" = 1;
            "ParentUniqueId" = 1;
            "StreamHash" = 1;
            "VirusStatus" = 1;
            "Attachments" = 1;
            "LinkTitle" = 1;
            "LinkTitle2" = 1;
            "LinkTitleNoMenu" = 1;
            "_ExtendedDescription" = 1;
            "_ShortcutSiteId" = 1;
            "_ShortcutUniqueId" = 1;
            "_ShortcutUrl" = 1;
            "_ShortcutWebId" = 1;
            "ContentVersion" = 1;
            "TriggerFlowInfo" = 1;
            
            "Edit" = 1;
            "DocIcon" = 1;
            "Order" = 1;
        };
        <# - Replaced by the property "IsSystemList"
        ExcludedLists = @{
            "masterpage" = 1; "design" = 1; "PublishedFeed" = 1; "theme" = 1; "wp" = 1; "users" = 1; "solutions" = 1; "Style Library" = 1; "lt" = 1; "appdata" = 1; "TaxonomyHiddenList" = 1; "ContentTypeSyncLog" = 1; "wfpub" = 1;
        }
        #>
        #ToDo: Add the switches: Include Schema XML, Include System Lists, Include System Fields
    });             
}


function Format-XmlFieldValueForExtract
{
	[cmdletbinding()]
	param ([Parameter(Mandatory=$true)]$XMLFieldVal);

	$theResult = [System.Web.HttpUtility]::HtmlEncode($XMLFieldVal);
	$theResult = $theResult.replace([Environment]::NewLine, " ").replace("`r", " ").replace("`n", " ");
	return $theResult;
}

function Format-CleanStringForExport
{
    [cmdletbinding()]
    param ([Parameter(Mandatory=$true)]$FieldVal);
    
    return ($FieldVal.replace([Environment]::NewLine, " ").replace("`r", " ").replace("`n", " "));
}

function Format-DateValueForExport
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$false)]$Date
    );

    if ($Date) {
        $TranDate = [System.TimeZoneInfo]::ConvertTimeFromUtc($Date, $InventorySettings.TimeZoneForExport);
        return $TranDate.ToString($InventorySettings.DateFormat)
    } else {
        return "";
    }
}

function Run-SPOnlineInventoryFullInventory
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$TenantUrl
        ,[switch]$WaitBetweenSites
        ,[Switch]$IncludeSchemaXML
        ,[Switch]$IncludeSystemLists
        ,[switch]$IncludeSystemFields
    );

    $CredentialParams = ($InventorySettings.PnPCredentialParams);
    Connect-PnPOnline -Url $TenantUrl -ErrorAction Stop @CredentialParams;

    $AllSites = (Get-PnPTenantSite);
    foreach ($currSite in $AllSites) {
        try {        
            if ($currSite.Template -ne "RedirectSite#0") {
                Run-SPOnlineInventoryFullInventoryForSite -InventorySettings $InventorySettings -SiteUrl ($currSite.Url) -IncludeSchemaXML:$IncludeSchemaXML -IncludeSystemLists:$IncludeSystemLists -IncludeSystemFields:$IncludeSystemFields;
            }
        } catch {
            Write-Host "ERROR:" -ForegroundColor Red;
            Write-Host $_;
            Write-Host "";
        }

        if ($WaitBetweenSites) {
            Start-Sleep -Seconds 60;
        }
    }
}


function Run-SPOnlineInventoryFullInventorySitesWebsLists
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$TenantUrl
        ,[switch]$WaitBetweenSites
        ,[Switch]$IncludeSchemaXML
        ,[Switch]$IncludeSystemLists
        ,[Switch]$IncludeSystemFields
    );

    $InvProps = @{};
    $InvProps["InventorySettings"] = $InventorySettings;
    $InvProps["SiteUrl"] = $null;
    $InvProps["InventoryWebs"] = $true;
    $InvProps["InventoryLists"] = $true;
    $InvProps["IncludeSchemaXML"] = $IncludeSchemaXML;
    $InvProps["IncludeSystemLists"] = $IncludeSystemLists;

    $CredentialParams = ($InventorySettings.PnPCredentialParams);
    Connect-PnPOnline -Url $TenantUrl -ErrorAction Stop @CredentialParams;

    $AllSites = (Get-PnPTenantSite);
    foreach ($currSite in $AllSites) {        
        if ($currSite.Template -ne "RedirectSite#0") {
            $InvProps.SiteUrl = ($currSite.Url);        
            Inventory-SPOnlineSiteCollection @InvProps;

            if ($WaitBetweenSites) {
                Start-Sleep -Seconds 60;
            }
        }
    }
}


function Run-SPOnlineInventoryPartialInventory
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$TenantUrl
        ,[switch]$WaitBetweenSites
        ,[Switch]$IncludeSchemaXML
        ,[Switch]$IncludeSystemLists
        ,[Switch]$IncludeSystemFields
    );

    $CredentialParams = ($InventorySettings.PnPCredentialParams);
    Connect-PnPOnline -Url $TenantUrl -ErrorAction Stop @CredentialParams;

    $AllSites = (Get-PnPTenantSite);
    foreach ($currSite in $AllSites) {        
        if ($currSite.Template -ne "RedirectSite#0") {
            $InvProps.SiteUrl = ($currSite.Url);        
            Run-SPOnlineInventoryPartialInventoryForSite -InventorySettings $InventorySettings -SiteUrl ($currSite.Url) -IncludeSchemaXML:$IncludeSchemaXML -IncludeSystemLists:$IncludeSystemLists -IncludeSystemFields:$IncludeSystemFields;

            if ($WaitBetweenSites) {
                Start-Sleep -Seconds 60;
            }
        }
    }
}

function Run-SPOnlineInventoryFullInventoryForSite
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$SiteUrl
        ,[Switch]$IncludeSchemaXML
        ,[Switch]$IncludeSystemLists        
        ,[switch]$IncludeSystemFields
    );

    $InvProps = @{};
    $InvProps["InventorySettings"] = $InventorySettings;
    $InvProps["SiteUrl"] = $SiteUrl;
    $InvProps["InventoryWebs"] = $true;
    $InvProps["InventorySiteGroups"] = $true;
    $InvProps["InventorySitePermissions"] = $true;
    $InvProps["InventorySiteCollectionFeatures"] = $true;
    $InvProps["InventorySiteFeatures"] = $true;
    $InvProps["InventorySiteApps"] = $true;
    $InvProps["InventoryListPermissions"] = $true;
    $InvProps["InventoryItemPermissions"] = $true;
    $InvProps["InventoryItems"] = $true;
    $InvProps["InventoryContentTypes"] = $true;
    $InvProps["InventoryLists"] = $true;
    $InvProps["InventoryListFields"] = $true;
    $InvProps["InventoryListViews"] = $true;
    $InvProps["InventoryListContentTypes"] = $true;
    $InvProps["InventoryWebParts"] = $true;
    $InvProps["IncludeSchemaXML"] = $IncludeSchemaXML;
    $InvProps["IncludeSystemLists"] = $IncludeSystemLists;

    Inventory-SPOnlineSiteCollection @InvProps;
}


function Run-SPOnlineInventoryPartialInventoryForSite
{
    [cmdletbinding()]
	param (                                        
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$SiteUrl
        ,[Switch]$IncludeSchemaXML
        ,[Switch]$IncludeSystemLists
        ,[switch]$IncludeSystemFields
    );

    $InvProps = @{};
    $InvProps["InventorySettings"] = $InventorySettings;
    $InvProps["SiteUrl"] = $SiteUrl;
    $InvProps["InventoryWebs"] = $true;
    <#
    $InvProps["InventorySiteGroups"] = $true;
    $InvProps["InventorySitePermissions"] = $true;
    $InvProps["InventorySiteCollectionFeatures"] = $true;
    $InvProps["InventorySiteFeatures"] = $true;
    $InvProps["InventorySiteApps"] = $true;
    $InvProps["InventoryListPermissions"] = $true;        
    $InvProps["InventoryContentTypes"] = $true;
    #>
    $InvProps["InventoryLists"] = $true;    
    $InvProps["InventoryListFields"] = $true;
    
    #$InvProps["InventoryListViews"] = $true;
    <#
    $InvProps["InventoryListContentTypes"] = $true;
    $InvProps["InventoryWebParts"] = $true;    
    #>
    $InvProps["IncludeSchemaXML"] = $IncludeSchemaXML;
    $InvProps["IncludeSystemLists"] = $IncludeSystemLists;
    $InvProps["IncludeSystemFields"] = $IncludeSystemFields;    

    Inventory-SPOnlineSiteCollection @InvProps;
}

function Inventory-SPOnlineSiteCollection
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$SiteUrl        
        ,[switch]$InventoryWebs
        ,[switch]$InventorySiteGroups
        ,[switch]$InventorySitePermissions
        ,[switch]$InventorySiteCollectionFeatures
        ,[switch]$InventorySiteFeatures
        ,[switch]$InventorySiteApps
        ,[switch]$InventoryListPermissions
        ,[switch]$InventoryItemPermissions
        ,[switch]$InventoryItems
        ,[switch]$InventoryContentTypes
        ,[switch]$InventoryLists
        ,[switch]$InventoryListFields
        ,[switch]$InventoryListViews
        ,[switch]$InventoryListContentTypes
        ,[switch]$InventoryWebParts
        ,[switch]$IncludeSchemaXML
        ,[switch]$IncludeSystemLists
        ,[switch]$IncludeSystemFields
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "SiteCollections.csv");
        $Delimiter = ($InventorySettings.Delimiter);
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID" + $Delimiter + "SiteGUID_WebGUID" + $Delimiter + "Url" + $Delimiter +  "IsHubSite" + $Delimiter +  "HubSiteId" + $Delimiter +  "MaxItemsPerThrottledOperation" + $Delimiter + 
                "Owner" + $Delimiter +  "RecycleBinItemCount" + $Delimiter + "LastContentModifiedDate" + $Delimiter + "ScriptingEnabled(DenyAddAndCustomizePages)" + $Delimiter + 
                "SharingCapability" + $Delimiter + "ShowPeoplePickerSuggestionsForGuestUsers" + $Delimiter + "Status" + $Delimiter + "Template" + $Delimiter + "StorageUsage" + $Delimiter + "LocaleID" + $Delimiter + "LocaleCode" + $Delimiter + "Locale" + $Delimiter +
                "CommentingOnSitePages" + $Delimiter + "ConditionalAccessPolicy" + $Delimiter + "ExternalUserExpirationInDays" + $Delimiter + "NumberOfWebs"
            );
        }

        $SiteProps = [String[]] @(
            "Id", "RootWeb", "CommentsOnSitePagesDisabled", "DisableFlows", "IsHubSite", "HubSiteId", "MaxItemsPerThrottledOperation", "Owner", "RecycleBin", "EventReceivers"            
        );

        $CredentialParams = ($InventorySettings.PnPCredentialParams)
        Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
        Write-Host ($Indentation + "Connected to Site Collection: " + $SiteUrl) -ForegroundColor Green;

        $SiteObj = (Get-PnPSite -Includes $SiteProps -ErrorAction Stop);
        $SiteGuid = ($SiteObj.ID.Guid.ToString());   
        $Web = (Get-PnPWeb -Includes ID,Url -ErrorAction Stop);
        $WebGuid = $Web.ID.Guid.ToString();

        $TenantSite = (Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop);
        $ClosedConnection = $false;    
    }
    PROCESS {
        $Locale = [System.Globalization.CultureInfo]::GetCultureInfo([int]$TenantSite.LocaleId);        

        $ExtractRows += (
            ($SiteGuid) +
            $Delimiter + ($SiteGuid + "_" + $WebGuid) +
            $Delimiter + ($SiteObj.Url) +
            $Delimiter + ($SiteObj.IsHubSite) +
            $Delimiter + ($SiteObj.HubSiteId.Guid.ToString()) +
            $Delimiter + ($SiteObj.MaxItemsPerThrottledOperation) +
            $Delimiter + ($SiteObj.Owner.LoginName) +
            $Delimiter + ($SiteObj.RecycleBin.Count.ToString()) +
            $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($TenantSite.LastContentModifiedDate)) +
            $Delimiter + ($TenantSite.DenyAddAndCustomizePages -eq [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]::Disabled) +
            $Delimiter + ($TenantSite.SharingCapability) +
            $Delimiter + ($TenantSite.ShowPeoplePickerSuggestionsForGuestUsers) +
            $Delimiter + ($TenantSite.Status) +
            $Delimiter + ($TenantSite.Template) +
            $Delimiter + ($TenantSite.StorageUsageCurrent) +
            $Delimiter + ($Locale.LCID.ToString()) +
            $Delimiter + ($Locale.Name) +
            $Delimiter + ($Locale.DisplayName) +
            $Delimiter + (-not $TenantSite.CommentsOnSitePagesDisabled) +
            $Delimiter + ($TenantSite.ConditionalAccessPolicy) +
            $Delimiter + ($TenantSite.ExternalUserExpirationInDays) +
            $Delimiter + ($TenantSite.WebsCount)
        );

        if ($InventorySiteCollectionFeatures) {
            Inventory-SPOnlineFeatures -InventorySettings $InventorySettings -Location "SiteCollection" -Indentation ($Indentation + "`t") -SkipConnection;
        }

        if (
            ($InventoryWebs) -or ($InventorySitePermissions) -or ($InventorySiteCollectionFeatures) -or ($InventorySiteFeatures) -or ($InventorySiteApps) -or ($InventoryListPermissions) -or ($InventoryItemPermissions) -or 
            ($InventoryContentTypes) -or ($InventoryLists) -or ($InventoryListFields) -or ($InventoryListViews) -or ($InventoryListContentTypes) -or ($InventoryWebParts) -or ($InventorySiteGroups) -or ($InventoryItems)
        ) {            
            $WebUrls = @();
            $WebUrls += ($Web.Url);

            $SubWebs = (Get-PnPSubWeb);
            foreach ($currSubWeb in $SubWebs) {
                $WebUrls += ($currSubWeb.Url);
            }

            Disconnect-PnPOnline;
            $ClosedConnection = $true;

            foreach ($cuurWebUrl in $WebUrls) {
                $InventoryWebParams = @{
                    InventorySettings = $InventorySettings;
                    SiteUrl = $cuurWebUrl;
                    InventorySitePermissions = $InventorySitePermissions;
                    InventorySiteCollectionFeatures = $InventorySiteCollectionFeatures;
                    InventorySiteFeatures = $InventorySiteFeatures;
                    InventorySiteApps = $InventorySiteApps;
                    InventoryListPermissions = $InventoryListPermissions;
                    InventoryItemPermissions = $InventoryItemPermissions;
                    InventoryContentTypes = $InventoryContentTypes;
                    InventoryLists = $InventoryLists;
                    InventoryListFields = $InventoryListFields;
                    InventoryListViews = $InventoryListViews;
                    InventoryListContentTypes = $InventoryListContentTypes;
                    InventoryWebParts = $InventoryWebParts;
                    IncludeSchemaXML = $IncludeSchemaXML;
                    InventorySiteGroups = $InventorySiteGroups;
                    InventoryItems = $InventoryItems;
                    IncludeSystemLists = $IncludeSystemLists;
                    IncludeSystemFields = $IncludeSystemFields;
                    Indentation = ($Indentation + "`t");
                }

                Inventory-SPOnlineWeb @InventoryWebParams;
            }
        }        

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if ($ClosedConnection -eq $false) {
            Disconnect-PnPOnline;
        }
    }    
}


function Inventory-SPOnlineWeb
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true)][String]$SiteUrl        
        ,[switch]$InventorySitePermissions
        ,[switch]$InventorySiteGroups
        ,[switch]$InventorySiteCollectionFeatures
        ,[switch]$InventorySiteFeatures
        ,[switch]$InventorySiteApps
        ,[switch]$InventoryListPermissions
        ,[switch]$InventoryItemPermissions
        ,[switch]$InventoryItems
        ,[switch]$InventoryContentTypes
        ,[switch]$InventoryLists
        ,[switch]$InventoryListFields
        ,[switch]$InventoryListViews
        ,[switch]$InventoryListContentTypes         
        ,[switch]$InventoryWebParts
        ,[switch]$IncludeSchemaXML
        ,[switch]$IncludeSystemLists
        ,[switch]$IncludeSystemFields
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + $Location + "Webs.csv");
        $Delimiter = ($InventorySettings.Delimiter);
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID" + $Delimiter + "WebGUID" + $Delimiter + "SiteGUID_WebGUID" + $Delimiter + "Url" + $Delimiter + "Title" + $Delimiter + "Description" + $Delimiter + "Theme" + $Delimiter + "WebIsRoot" + $Delimiter + "CreatedDate" +
                $Delimiter + "WebLastItemModifiedDate" + $Delimiter + "ParentSiteGUID_WebGUID" + $Delimiter + "CommentsOnSitePagesDisabled" + $Delimiter +  "FlowsEnabled" + $Delimiter + "EventReceiversCount" + $Delimiter + "RecycleBin" +
                $Delimiter + "TimeZoneID" + $Delimiter + "TimeZone" + $Delimiter + "TimeZoneOffset" + $Delimiter + "Crawled" + $Delimiter + "WebTemplate"
            );
        }

        $PropsToLoad = [String[]]@(
            "Url", "Title", "Description", "ThemeInfo", "Created", "LastItemModifiedDate", "ParentWeb", "CommentsOnSitePagesDisabled", "DisableFlows", "EventReceivers", "RecycleBin"
            ,"RegionalSettings", "RegionalSettings.TimeZone", "NoCrawl",  "WebTemplate"
        );

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams)
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            Write-Host ($Indentation + "Connected to Web: " + $SiteUrl);
        }
        
        $SiteObj = (Get-PnPSite -Includes RootWeb,ID -ErrorAction Stop);        
        $SiteGuid = ($SiteObj.ID.Guid.ToString());
        $Web = (Get-PnPWeb -Includes $PropsToLoad -ErrorAction Stop);
        $WebGuid = $Web.ID.Guid.ToString();
        $SiteGuid_WebGUID = ($SiteGuid + "_" + $WebGuid);
        $Timezones = (Get-PnPTimeZoneId);

        $ParentID = "";        
    }
    PROCESS {
        if ($web.ParentWeb.Id -ne $null) {
            $ParentID = ($SiteGuid + "_" + $web.ParentWeb.Id.Guid.ToString());
        }

        $TimeZoneSiteIsUsing = $null;
        foreach ($tz in $Timezones) {
            if ($tz.ID -eq ($web.RegionalSettings.TimeZone.Id)) {
                $TimeZoneSiteIsUsing = $tz;
            }
        }

        $ExtractRows += (
            $SiteGuid +
            $Delimiter + $WebGuid +
            $Delimiter + $SiteGuid_WebGUID +
            $Delimiter + ($Web.Url) +
            $Delimiter + ($Web.Title) +
            $Delimiter + (Format-CleanStringForExport -FieldVal ($Web.Description)) +
            $Delimiter + ($Web.ThemeInfo.AccessibleDescription) +
            $Delimiter + ($SiteObj.RootWeb.Url -eq $Web.Url) +
            $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($Web.Created)) +
            $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($Web.LastItemModifiedDate)) +
            $Delimiter + $ParentID +
            $Delimiter + ($web.CommentsOnSitePagesDisabled) +
            $Delimiter + (-not $web.DisableFlows) +
            $Delimiter + ($web.EventReceivers.Count.ToString()) +
            $Delimiter + ($web.RecycleBin.Count.ToString()) +
            $Delimiter + ($TimeZoneSiteIsUsing.Id) +
            $Delimiter + ($TimeZoneSiteIsUsing.Description) +
            $Delimiter + ($TimeZoneSiteIsUsing.Identifier) +
            $Delimiter + (-not $web.NoCrawl) +
            $Delimiter + ($web.WebTemplate)
        );

        if (($InventoryLists) -or ($InventoryListFields) -or ($InventoryListViews) -or ($InventoryListContentTypes) -or ($InventoryListPermissions) -or ($InventoryItemPermissions) -or ($InventoryItems) -or ($InventoryWebParts)) {
            $InventoryListsParams = @{
                InventorySettings = $InventorySettings;
                InventoryListFields = $InventoryListFields;
                InventoryListViews = $InventoryListViews;
                InventoryListContentTypes = $InventoryListContentTypes;                
                InventoryListPermissions = $InventoryListPermissions;
                InventoryItemPermissions = $InventoryItemPermissions;
                InventoryItems = $InventoryItems;                
                InventoryWebParts = $InventoryWebParts;
                SkipConnection = $true;
                SiteGUID_WebGUID = $SiteGuid_WebGUID;
                IncludeSchemaXML = $IncludeSchemaXML;
                IncludeSystemLists = $IncludeSystemLists;
                IncludeSystemFields = $IncludeSystemFields;
                Indentation = ($Indentation + "`t");
            };

            Inventory-SPOnlineLists @InventoryListsParams;
        }

        if ($InventoryContentTypes) {
            $InventoryContentTypeParams = @{
                InventorySettings = $InventorySettings;
                ObjectToGetContentTypesFor = $Web;
                Location = "Web";
                IncludeSchemaXML = $IncludeSchemaXML;
                SkipConnection = $true;
                SiteGUID_WebGUID = $SiteGuid_WebGUID;
                Indentation = ($Indentation + "`t");
            };

            Inventory-SPOnlineContentTypes @InventoryContentTypeParams;
        }

        $InventoryCommonParams = @{
            InventorySettings = $InventorySettings;
            SkipConnection = $true;
            SiteGUID_WebGUID = $SiteGuid_WebGUID;
            Indentation = ($Indentation + "`t");
        };

        if ($InventorySitePermissions) { Inventory-SPOnlinePermissions @InventoryCommonParams -Location "Web" -ObjectToGetPermissionsFor $web;}
        if ($InventorySiteGroups) { Inventory-SPOnlineGroups @InventoryCommonParams; }
        if ($InventorySiteFeatures) { Inventory-SPOnlineFeatures @InventoryCommonParams -Location "Web"; }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }   
}

function Inventory-SPOnlineLists 
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true, ParameterSetName = "SiteUrl")][String]$SiteUrl        
        ,[Parameter(Mandatory=$true, ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false, ParameterSetName = "SkipConnection")][String]$SiteGUID_WebGUID
        ,[switch]$InventoryListPermissions
        ,[switch]$InventoryItemPermissions
        ,[switch]$InventoryItems
        ,[switch]$InventoryWebParts
        ,[switch]$InventoryListFields
        ,[switch]$InventoryListViews
        ,[switch]$InventoryContentTypes
        ,[switch]$InventoryListContentTypes
        ,[switch]$IncludeSystemLists
        ,[switch]$IncludeSchemaXML
        ,[switch]$IncludeSystemFields
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );
    
    BEGIN {        
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "Lists.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "RootFolder" + $Delimiter + "InternalName" + $Delimiter + "DisplayName" +  $Delimiter + "Author" + $Delimiter + "BaseType" + $Delimiter + "ListTemplate" + $Delimiter + "AllowContentTypes" + $Delimiter + "Created" + $Delimiter + "LastItemModifiedDate" + $Delimiter + "EnableAttachments" + $Delimiter +
                ,"EnableFolderCreation" + $Delimiter + "EnableVersioning" + $Delimiter + "DefaultViewUrl" + $Delimiter + "Description" + $Delimiter + "Hidden" + $Delimiter + "IsSystemList" + $Delimiter + "HasUniqueRoleAssignments" + $Delimiter +
                ,"MajorVersionLimit" + $Delimiter + "MajorWithMinorVersionsLimit" + $Delimiter + "Crawled" + $Delimiter + "SiteGUID_WebGUID" + $Delimiter + "ListGUID" + $Delimiter + "SchemaXML" + $Delimiter + "SiteUrl" + $Delimiter + "SiteGUID_WebGUID_ListGUID" + $Delimiter + "ItemCount" + $Delimiter + "ValidationFormula"
            );
        }        

        $ListProps = [String[]] @(
            "RootFolder", "Author", "BaseType", "BaseTemplate", "AllowContentTypes", "Created", "LastItemModifiedDate", "EnableAttachments"
            ,"EnableFolderCreation", "EnableVersioning", "DefaultViewUrl", "Description", "Hidden", "IsSystemList", "HasUniqueRoleAssignments"
            ,"MajorVersionLimit", "MajorWithMinorVersionsLimit", "NoCrawl", "Title", "SchemaXML", "ItemCount", "ValidationFormula"
        );

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            Write-Host ("Connected to Site: " + $SiteUrl);
        }

        $WebObj = (Get-PnPWeb -Includes ID, Url, ServerRelativePath -ErrorAction Stop);
        $SiteUrl = $WebObj.Url.ToString();

        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID -ErrorAction Stop);                    
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $WebObj.ID.Guid.ToString();                                    
            
            $SiteGUID_WebGUID = ($SiteGUID + "_" + $WebGUID);
        }

        $Lists = (Get-PnPList -ErrorAction Stop);
        $TermSetDictionary = (Get-SPOnlineTermSets);
    }
    PROCESS {
        Write-Host ($Indentation + "Inventorying Lists: ");
        $Indentation += "`t";

        foreach ($currList in $Lists) {            
            $LoadListProps = (Get-SPOnlineHelperPnPProperty -ClientObject $currList -Property $ListProps);   
            $ListInternalName = ($currList.RootFolder.Name);
                        
            if (($IncludeSystemLists) -or ($currList.IsSystemList -eq $false)) { 
                
                $ListRootPath = (Get-SPOnlineHelperListRootPath -List $currList -WebRelativeUrl ($WebObj.ServerRelativePath.DecodedUrl));
                Write-Host ($Indentation + "-" + $currList.Title + " (" + $ListRootPath + ")");

                $ListSchemaXml = "Not Extracted";
                if ($IncludeSchemaXML) { $ListSchemaXml = (Format-XmlFieldValueForExtract -XMLFieldVal ($currList.SchemaXML)); }

                $SiteGUID_WebGUID_ListGUID = ($SiteGUID + "_" + $WebGUID + "_" + $currList.Id.Guid.ToString());

                $ExtractRows += (
                    ($ListRootPath) +   #Includes "Lists/{ListName}"
                    $Delimiter + ($ListInternalName) +   #List Internal Name
                    $Delimiter + ($currList.Title) +     #List Display Name
                    $Delimiter + ($currList.Author.LoginName)  + #Person/Account that created the list
                    $Delimiter + ($currList.BaseType) +          #Typically GenericList or DocumentLibrary
                    $Delimiter + ($currList.BaseTemplate) +      #List Template (eg Calendar)
                    $Delimiter + ($currList.AllowContentTypes) +
                    $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($currList.Created)) +
                    $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($currList.LastItemModifiedDate)) +
                    $Delimiter + ($currList.EnableAttachments) +
                    $Delimiter + ($currList.EnableFolderCreation) +
                    $Delimiter + ($currList.EnableVersioning) +
                    $Delimiter + ($currList.DefaultViewUrl) +
                    $Delimiter + (Format-CleanStringForExport -FieldVal ($currList.Description)) +
                    $Delimiter + ($currList.Hidden) +
                    $Delimiter + ($currList.IsSystemList) +
                    $Delimiter + ($currList.HasUniqueRoleAssignments) +
                    $Delimiter + ($currList.MajorVersionLimit) +
                    $Delimiter + ($currList.MajorWithMinorVersionLimit) +
                    $Delimiter + (-not $currList.NoCrawl) +                    
                    $Delimiter + $SiteGUID_WebGUID +
                    $Delimiter + ($currList.Id.Guid.ToString()) +
                    $Delimiter + $ListSchemaXml +
                    $Delimiter + $SiteUrl +
                    $Delimiter + $SiteGUID_WebGUID_ListGUID +
                    $Delimiter + ($currList.ItemCount.ToString()) +
                    $Delimiter + ($currList.ValidationFormula)
                );

                $InventoryCommonParams = @{
                    InventorySettings = $InventorySettings;                                        
                    SkipConnection = $true;
                    Indentation = ($Indentation + "`t");
                };

                if ($InventoryListFields) {
                    Inventory-SPOnlineListFields @InventoryCommonParams -List $currList -SiteGUID_WebGUID_ListGUID $SiteGUID_WebGUID_ListGUID -TermSetDictionary $TermSetDictionary -IncludeSchemaXML:$IncludeSchemaXML -IncludeSystemFields:$IncludeSystemFields;
                }

                if ($InventoryListContentTypes) {
                    Inventory-SPOnlineContentTypes @InventoryCommonParams -ObjectToGetContentTypesFor $currList -Location "List" -SiteGUID_WebGUID $SiteGUID_WebGUID;
                }

                if ($InventoryListPermissions) {
                    Inventory-SPOnlinePermissions @InventoryCommonParams -Location "List" -ObjectToGetPermissionsFor $currList -SiteGUID_WebGUID $SiteGUID_WebGUID;
                }

                if ($InventoryItems) {                    
                    Inventory-SPOnlineListItems @InventoryCommonParams -List $currList -SiteGUID_WebGUID_ListGUID $SiteGUID_WebGUID_ListGUID;
                }                

                if ($InventoryItemPermissions) {
                    Inventory-SPOnlineListItemPermissions @InventoryCommonParams -List $currList -SiteGUID_WebGUID $SiteGUID_WebGUID;
                }

                if ($InventoryListViews) {
                    Inventory-SPOnlineListViews @InventoryCommonParams -List $currList -SiteGUID_WebGUID_ListGUID $SiteGUID_WebGUID_ListGUID -IncludeSchemaXML:$IncludeSchemaXML;
                }

                if (($InventoryWebParts) -and ($currList.BaseType -eq "DocumentLibrary")) {
                    Inventory-SPOnlineWebParts @InventoryCommonParams -List $currList -SiteGUID_WebGUID_ListGUID $SiteGUID_WebGUID_ListGUID;
                }
            }
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;      
    }
    END {    
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }        
}


function Inventory-SPOnlineListFields
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$ListRootFolder
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")]$List
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID_ListGUID
        ,[Parameter(Mandatory=$false)][HashTable]$TermSetDictionary
        ,[switch]$IncludeSchemaXML
        ,[switch]$IncludeSystemFields
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "ListFields.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "DisplayName" + $Delimiter + "InternalName" + $Delimiter + "FieldGUID" + $Delimiter + "Description" + $Delimiter + "Type" + $Delimiter + "IsRequired" + $Delimiter + "DefaultValue" + $Delimiter + "IsHidden" +  $Delimiter + "ReadOnly" +
                $Delimiter + "SiteGUID_WebGUID_ListGUID" + $Delimiter + "SchemaXml" + $Delimiter + "SiteUrl" + $Delimiter + "ListInternalName" + $Delimiter + "ExtraInformation"
            );
        }

        $FieldProps = [string[]] @("InternalName", "Description", "TypeAsString", "Required", "DefaultValue", "Hidden", "ReadOnlyField");
        
        if (-not $SkipConnection) {
            Connect-PnPOnline -Url $SiteUrl -ClientId ($InventorySettings.ClientID) -ClientSecret ($InventorySettings.ClientSecret);
            $List = (Get-PnPList -Identity $ListRootFolder -ErrorAction Stop);
        }

        $WebObj = (Get-PnPWeb -Includes ID,Url,ServerRelativePath -ErrorAction Stop);
        $SiteUrl = $WebObj.Url.ToString();
        
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID_ListGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID);        
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $WebObj.ID.Guid.ToString();            
            $ListGUID = $List.Id.Guid.ToString();
            $SiteGUID_WebGUID_ListGUID = ($SiteGuid + "_" + $WebGuid + "_" + $ListGUID);
        }
        
        $LoadListFields = (Get-SPOnlineHelperPnPProperty -ClientObject $List -Property Fields);
        $ListRootPath = (Get-SPOnlineHelperListRootPath -List $currList -WebRelativeUrl ($WebObj.ServerRelativePath.DecodedUrl));
    }
    PROCESS {                
        Write-Host ($Indentation + "Inventorying List Fields");
        foreach ($currField in $List.Fields) {
            
            if (($IncludeSystemFields) -or ($InventorySettings.ExcludedFieldList[$currField.InternalName] -eq $null)) {
                $ListFieldSchemaXml = "Not Extracted";
                if ($IncludeSchemaXML) { $ListFieldSchemaXml = (Format-XmlFieldValueForExtract -XMLFieldVal ($currField.SchemaXML)); }
                                
                $LoadFieldProps = (Get-SPOnlineHelperPnPProperty -ClientObject $currField -Property $FieldProps);
                $ExtraInformation = "";
                if (($currField.TypeAsString -eq "Choice") -or ($currField.TypeAsString -eq "MultiChoice")) { 
                    $ExtraInformation = ("Choices: " + ($currField.Choices -join ";")); 
                } elseif ($currField.TypeAsString -eq "Calculated") { 
                    $ExtraInformation = ("Formula: " + $currField.Formula);                     
                } elseif (($currField.TypeAsString -eq "TaxonomyFieldTypeMulti") -or ($currField.TypeAsString -eq "TaxonomyFieldType")) { 
                    if ($TermSetDictionary -ne $null) {
                        $ExtraInformation = ("Term Set: " + $TermSetDictionary[$currField.TermSetId.Guid].Name);
                    }
                }

                if ($ExtraInformation -eq "-") {
                    $ExtraInformation = ";"
                }

                $ExtractRows += (
                    ($currField.Title) +   #Display Name
                    $Delimiter + ($currField.InternalName) +
                    $Delimiter + ($currField.Id.Guid.ToString()) +
                    $Delimiter + (Format-CleanStringForExport -FieldVal ($currField.Description)) + 
                    $Delimiter + ($currField.TypeAsString) +    #Text, Note, etc
                    $Delimiter + ($currField.Required) +
                    $Delimiter + ($currField.DefaultValue) +
                    $Delimiter + ($currField.Hidden) +
                    $Delimiter + ($currField.ReadOnlyField) +                    
                    $Delimiter + $SiteGUID_WebGUID_ListGUID +
                    $Delimiter + $ListFieldSchemaXml +
                    $Delimiter + $SiteUrl + 
                    $Delimiter + $ListRootPath +
                    $Delimiter + $ExtraInformation
                );
            }
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}


function Inventory-SPOnlineContentTypes
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl        
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID
        ,[Parameter(Mandatory=$true)]$ObjectToGetContentTypesFor
        ,[Parameter(Mandatory=$true)]$Location
        ,[switch]$IncludeSchemaXML
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + $Location + "ContentTypes.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (test-path $ExportFilePathAndName)) {
            $ExtractRows += (
                "Location" + $Delimiter + "ContentTypeID" + $Delimiter + "ContentTypeIdPath" + $Delimiter + "ContentTypeNamePath" + $Delimiter + "ContentTypeName" + $Delimiter + "IsDocumentSet" + 
                $Delimiter + "IsFolder" + $Delimiter + "SiteGUID_WebGUID" + $Delimiter + "ParentID" + $Delimiter + "UniqueContentTypeId"
            );
        }

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
        }
        
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID -ErrorAction Stop);        
            $SiteGuid = $SiteObj.ID.Guid.ToString();   
            $WebGuid = (Get-PnPWeb -ErrorAction Stop).ID.Guid.ToString();  
            $SiteGUID_WebGUID = ($SiteGuid + "_" + $WebGuid);                        
        }
        
        
        $ParentID = $SiteGUID_WebGUID;        
        if ($Location -eq "List") {
            $ObjectID = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetContentTypesFor -Property Id);
            $ParentID += ("_" + $ObjectToGetContentTypesFor.Id.Guid.ToString());
        }

        $ContentTypes = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetContentTypesFor -Property "ContentTypes");
    }
    PROCESS {
        Write-Host ($Indentation + "Inventorying " + $Location + " Content Types (" + $ObjectToGetContentTypesFor.ContentTypes.Count.ToString() + ")");
        foreach ($ct in $ObjectToGetContentTypesFor.ContentTypes) {
            $ContentTypeSchemaXml = "Not Extracted";
            if ($IncludeSchemaXML) { $ContentTypeSchemaXml = (Format-XmlFieldValueForExtract -XMLFieldVal ($ct.SchemaXML)); }
                        
            $ContentTypeTree = (Get-SPOnlineHelperContentTypeTree -ContentType $ct);
            $IsDocumentSet = ($ContentTypeTree.NamePath.StartsWith("Item > Folder > Document Collection Folder > Document Set"));
            $IsFolder = ($ContentTypeTree.NamePath.StartsWith("Item > Folder"));

            $ExtractRows += (
                $Location +
                $Delimiter + ($ct.Id.StringValue) +
                $Delimiter + ($ContentTypeTree.IDPath) +
                $Delimiter + ($ContentTypeTree.NamePath) +
                $Delimiter + ($ct.Name) +
                $Delimiter + ($IsDocumentSet) +
                $Delimiter + ($IsFolder) +
                $Delimiter + $SiteGUID_WebGUID +
                $Delimiter + $ParentID +
                $Delimiter + ($ParentID + "_" + ($ct.Id.StringValue))
            );
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }       
}



function Inventory-SPOnlinePermissions {
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl        
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID
        ,[Parameter(Mandatory=$true)]$ObjectToGetPermissionsFor
        ,[Parameter(Mandatory=$true)]$Location        
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {                        
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + $Location + "Permissions.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (test-path $ExportFilePathAndName)) {
            $ExtractRows += (
                "Location" + $Delimiter + "SiteGUID_WebGUID" + $Delimiter + "ParentID" + $Delimiter + "Title" + $Delimiter + "Url" + $Delimiter + 
                "Type" + $Delimiter + "MemberID" + $Delimiter + "MemberTitle" + $Delimiter + "MemberName" + $Delimiter + "Roles"                
            );
        }

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
        }        
        
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID -ErrorAction Stop);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = (Get-PnPWeb -ErrorAction Stop).ID.Guid.ToString(); 
            $SiteGUID_WebGUID = ($SiteGuid + "_" + $WebGuid);            
        }
        

        $ParentID = $SiteGUID_WebGUID;
        if ($Location -eq "List") {
            $ObjectID = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetPermissionsFor -Property Id);
            $ParentID += ("_" + $ObjectToGetPermissionsFor.Id.Guid.ToString());
        }

        $Title = "";
        $Url = "";

        if ($Location -eq "Item") {            
            $ItemProps = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetPermissionsFor -Property ParentList,FieldValuesAsText);
            $ListGUID = $ObjectToGetPermissionsFor.ParentList.Id.Guid.ToString();
            $ItemGUID = $ObjectToGetPermissionsFor["GUID"];
            $ParentID += ("_" + $ListGUID + "_" + $ItemGUID);

            $Title = ($ObjectToGetPermissionsFor["FileLeafRef"]);
            $Url = ($ObjectToGetPermissionsFor["FileRef"]);
        } else {
            $Title = ($ObjectToGetPermissionsFor.Title);
            if ($Location -eq "List") {
                $Url = ($ObjectToGetPermissionsFor.RootFolder.ServerRelativeUrl);
            } else {
                $Url = ($ObjectToGetPermissionsFor.ServerRelativeUrl);            
            }
        }
    }
    PROCESS {
        $UniqueRoleAssignments = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetPermissionsFor -Property HasUniqueRoleAssignments);
        if ($UniqueRoleAssignments -eq $true) {                    
            if ($ObjectToGetPermissionsFor.HasUniqueRoleAssignments) {
                $RoleAssignments = (Get-SPOnlineHelperPnPProperty -ClientObject $ObjectToGetPermissionsFor -Property RoleAssignments);
                Write-Host ($Indentation + "Inventorying " + $Location + " Permissions (" + $ObjectToGetPermissionsFor.RoleAssignments.Count.ToString() + ")");
                foreach ($ra in $RoleAssignments) {
                    $props = (Get-SPOnlineHelperPnPProperty -ClientObject $ra -Property Member,RoleDefinitionBindings);

                    $RoleDefinition = "";                
                    foreach ($roleDefinitionBinding in $ra.RoleDefinitionBindings) {
                        $RoleDefinition += ($roleDefinitionBinding.Name + ";");
                    }

                    $ExtractRows += (
                        $Location +
                        $Delimiter + $SiteGUID_WebGUID +
                        $Delimiter + $ParentID +
                        $Delimiter + $Title +
                        $Delimiter + $Url +
                        $Delimiter + ($ra.Member.GetType().Name) +
                        $Delimiter + ($ra.Member.ID) +
                        $Delimiter + ($ra.Member.Title) +
                        $Delimiter + ($ra.Member.LoginName) +
                        $Delimiter + ($RoleDefinition)
                    );
                }
            }

            $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
        }
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}


function Inventory-SPOnlineGroups
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl        
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID 
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""     
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "SiteGroups.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (test-path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID_WebGUID" + $Delimiter + "GroupID" + $Delimiter + "GroupLoginName" + $Delimiter + "GroupName" + $Delimiter + "UserID" + $Delimiter + "UserName" + $Delimiter + "Title"
            );
        }

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
        }
        
        $Web = (Get-PnPWeb -Includes ID,SiteGroups);
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $Web.ID.Guid.ToString();
            $SiteGUID_WebGUID = ($SiteGuid + "_" + $WebGuid);
        }                
    }
    PROCESS {
        
        Write-Host ($Indentation + "Inventorying Site Groups (" + $Web.SiteGroups.Count.ToString() + ")");
        foreach ($grp in $Web.SiteGroups) {
            
            $grpUsers = (Get-SPOnlineHelperPnPProperty -ClientObject $grp -Property Users);
            foreach ($usr in $grpUsers) {
                $ExtractRows += (
                    $SiteGUID_WebGUID +
                    $Delimiter + ($grp.ID.ToString()) +
                    $Delimiter + ($grp.LoginName) +
                    $Delimiter + ($grp.Title) +
                    $Delimiter + ($usr.ID.ToString()) +
                    $Delimiter + ($usr.LoginName) +
                    $Delimiter + ($usr.Title)
                );
            }
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}


function Inventory-SPOnlineFeatures
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID
        ,[Parameter(Mandatory=$false)][String]$Location = "SiteCollection"  
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""      
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + $Location + "Features.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID_WebGUID" + $Delimiter + "DefinitionId" + $Delimiter + "DisplayName" + $Delimiter + "Version" + $Delimiter + "Type"
            );
        }

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams)
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
        }
        
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {    
            $SiteObj = (Get-PnPSite -Includes ID -ErrorAction Stop);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $Web.ID.Guid.ToString();
            $SiteGUID_WebGUID = ($SiteGuid + "_" + $WebGuid);
        }

        $FeatureList = $null;
        if ($Location -eq "SiteCollection") {
            $FeatureList = ((Get-PnPSite -Includes Features -ErrorAction Stop).Features);
        } elseif (($Location -eq "Site") -or ($Location -eq "Web")) {
            $FeatureList = ((Get-PnPWeb -Includes Features -ErrorAction Stop).Features);
        } else {
            throw [System.InvalidOperationException]::new("Only Site Collection or Web are supported");
        }
        
        $FeatureProps = [string[]] @("DefinitionId", "DisplayName");
    }
    PROCESS {
        
        Write-Host ($Indentation + "Inventorying " + $Location + " Features (" + $FeatureList.Count.ToString() + ")");
        foreach ($feat in $FeatureList) {
            
            $featProps = (Get-SPOnlineHelperPnPProperty -ClientObject $feat -Property $FeatureProps);
            $AssemblyInfo = $feat.GetType().AssemblyQualifiedName;
            $AssemblyVersion = $AssemblyInfo.Substring($AssemblyInfo.indexOf("Version=") + 8);
            $AssemblyVersion = $AssemblyVersion.substring(0, $AssemblyVersion.indexOf(","));

            $ExtractRows += (
                ($SiteGuid + "_" + $WebGuid) +                
                $Delimiter + ($feat.DefinitionId) +
                $Delimiter + ($feat.DisplayName) +
                $Delimiter + ($AssemblyVersion) +
                $Delimiter + ($AssemblyInfo)
            );
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}


function Inventory-SPOnlineListItems
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$ListRootFolder
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID_ListGUID
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")]$List
        ,[switch]$InventoryItemPermissions
        ,[switch]$IncludeItemJSON #Not Implemented
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "ListItems.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        if (-not (Test-Path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID_WebGUID_ListGUID" + $Delimiter + "SiteGUID_WebGUID_ListGUID_ItemGUID" + $Delimiter + "ID" + $Delimiter + "Title" + $Delimiter + "ServerRelativeUrl" + $Delimiter + "ContentTypeId" + 
                $Delimiter + "SiteGUID_WebGUID_ListGUID_ContentTypeId" + $Delimiter + "Created" + $Delimiter + "Modified" + $Delimiter + "CreatedBy" + $Delimiter + "ModifiedBy" + $Delimiter + "FileSize" + $Delimiter + "Extension" + $Delimiter + "HasUniquePermissions" + $Delimiter + "VersionLabel"
            );
        }        

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams)
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            $List = (Get-PnPList -Identity $ListRootFolder -Includes ID,BaseType -ErrorAction Stop);            
        }

        $IsDocLibrary = ($List.BaseType.ToString() -eq "DocumentLibrary");        

        $WebObj = (Get-PnPWeb -Includes ID);
        $SiteUrl = $WebObj.Url.ToString();
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID_ListGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $WebObj.ID.Guid.ToString();
            $SiteGUID_WebGUID_ListGUID = ($SiteGuid + "_" + $WebGuid + "_" + $List.Id.Guid.ToString());
        }
    }
    PROCESS {
        Write-Host ($Indentation + "Inventorying List '" + $List.Title + "' Items");
        $ListDisplayFormUrl = ((Get-PnPProperty -ClientObject $List -Property DefaultDisplayFormUrl) + "?ID=");        
        
        $ListItemsWithUniquePermissions = @{};
        if ($List.Title -ne "User Information List") {
            $ListItemsWithUniquePermissionsREST = (Call-SPOnlineHelperListRestService -SiteUrl $SiteUrl -ListGUID ($List.Id.Guid) -FieldsCommaSeperated "ID,HasUniqueRoleAssignments,GUID");
            $ListItemsWithUniquePermissions = @{};
            foreach ($currItem in $ListItemsWithUniquePermissionsREST) {
                if ($currItem.HasUniqueRoleAssignments) {
                    $ListItemsWithUniquePermissions[$currItem.GUID] = 1;
                }
            }
        }

        $FieldsToQuery = @(
             "ID"
            ,"Name"
            ,"Title"
            ,"FileRef"
            ,"ContentTypeId"
            ,"Created"
            ,"Modified"
            ,"Author"
            ,"Editor"
            ,"_UIVersionString"
            ,"GUID"
            ,"File_x0020_Size"
        );
        $ViewFieldRefXml = "";
        foreach ($currField in $FieldsToQuery) {
            $ViewFieldRefXml += ("<FieldRef Name='" + $currField + "' />");
        }        

        $ListItemsQuery = (New-Object Microsoft.SharePoint.Client.CamlQuery);
        $ListItemsQuery.AllowIncrementalResults = $true;
        $ListItemsQuery.ViewXml = ("
            <View Scope='RecursiveAll'>        
                <ViewFields>" + $ViewFieldRefXml +" </ViewFields>
                <Query><OrderBy><FieldRef Name='ID' Ascending='True' /></OrderBy></Query>
                <RowLimit>1000</RowLimit>
                <Paged>True</Paged>
            </View> "
        );
        
        do {                                               
            $ListItems = $List.GetItems($ListItemsQuery);
            $List.Context.Load($ListItems);
            $List.Context.ExecuteQuery();

            foreach ($currItem in $ListItems) {                                            
                $ItemTitle = "-";
                $ItemUrl = "";
                $ItemExtension = "";
                if ($IsDocLibrary) { 
                    $ItemTitle = $currItem["FileLeafRef"]; 
                    $ItemUrl = ($currItem["FileRef"]);
                    $ItemExtension = [System.IO.Path]::GetExtension($ItemUrl);
                } else { 
                    $ItemTitle = $currItem["Title"]; 
                    $ItemUrl = ($ListDisplayFormUrl + $currItem.Id.ToString());
                }

                $itemGUID = $currItem["GUID"].ToString();
                $HasUniquePermissions = $false;
                if ($ListItemsWithUniquePermissions[$itemGUID] -eq 1) {
                    $HasUniquePermissions = $true;
                }

                $ContentTypeID = ($currItem["ContentTypeId"].StringValue);

                $ExtractRows += (
                    $SiteGUID_WebGUID_ListGUID +
                    $Delimiter + ($SiteGUID_WebGUID_ListGUID + "_" + $itemGUID) +
                    $Delimiter + ($currItem.Id.ToString()) +
                    $Delimiter + ($ItemTitle) +
                    $Delimiter + ($ItemUrl) +
                    $Delimiter + $ContentTypeID +
                    $Delimiter + ($SiteGUID_WebGUID_ListGUID + "_" + $ContentTypeID) +
                    $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($currItem["Created"])) +
                    $Delimiter + (Format-DateValueForExport -InventorySettings $InventorySettings -Date ($currItem["Modified"])) +
                    $Delimiter + ($currItem["Author"].LookupValue) +
                    $Delimiter + ($currItem["Editor"].LookupValue) +
                    $Delimiter + ($currItem["File_x0020_Size"]) +
                    $Delimiter + ($ItemExtension) +
                    $Delimiter + ($HasUniquePermissions) +
                    $Delimiter + ($currItem["_UIVersionString"])
                );            

                if (($ProcessedCounter % 500) -eq 0) {
                    $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
                    $ExtractRows = @();
                }
            }

            $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
            $ExtractRows = @();
            
            [System.Threading.Thread]::Sleep(1000); #Help prevent throttling
            $ListItemsQuery.ListItemCollectionPosition = ($ListItems.ListItemCollectionPosition);
        } while ($ListItemsQuery.ListItemCollectionPosition -ne $null);
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}


function Inventory-SPOnlineListItemPermissions
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$ListRootFolder
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false)]$SiteGUID_WebGUID
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")]$List
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {             
        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams)
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            $List = (Get-PnPList -Identity $ListRootFolder -Includes ID -ErrorAction Stop);            
        }
        
        $WebObj = (Get-PnPWeb -Includes ID,Url);
        $SiteUrl = $WebObj.Url.ToString();
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $WebObj.ID.Guid.ToString();
            $SiteGUID_WebGUID = ($SiteGuid + "_" + $WebGuid);            
        }        
    }
    PROCESS {
        $ListItemsWithUniquePermissions = @();
        if ($List.Title -ne "User Information List") {
            $ListItemsWithUniquePermissions = (Call-SPOnlineHelperListRestService -SiteUrl $SiteUrl -ListGUID ($List.Id.Guid) -FieldsCommaSeperated "ID,HasUniqueRoleAssignments,GUID");        
        }
        
        if ($ListItemsWithUniquePermissions.Count -gt 0) {
            
            Write-Host ($Indentation + "Inventorying List '" + $List.Title + "' Item Permissions");
            foreach ($itemFromRestCall in $ListItemsWithUniquePermissions) {
                if ($itemFromRestCall.HasUniqueRoleAssignments) { 
                    $ItemObj = (Get-PnPListItem -List $List -Id ($itemFromRestCall.ID));
                    Inventory-SPOnlinePermissions -InventorySettings $InventorySettings -ObjectToGetPermissionsFor $ItemObj -Location "Item" -Indentation ($Indentation + "`t") -SkipConnection -SiteGUID_WebGUID $SiteGUID_WebGUID;
                }
            }
        }
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;6
        }
    }
}


function Inventory-SPOnlineListViews
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$ListRootFolder
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection        
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")]$List
        ,[Parameter(Mandatory=$false)]$SiteGUID_WebGUID_ListGUID
        ,[switch]$IncludeSchemaXML
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "ListViews.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        
        
        if (-not (test-path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID_WebGUID_ListGUID" + $Delimiter + "Name" + $Delimiter + "UrlName" + $Delimiter + "Url" + $Delimiter +
                "RowLimit" + $Delimiter + "Paged" + $Delimiter + "Type" + $Delimiter + "Default"+ $Delimiter + "Query" + $Delimiter + "Fields" + $Delimiter + "SchemaXml"
            );
        }
        
        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            $List = (Get-PnPList -Identity $ListRootFolder -ErrorAction Stop);
        }

        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID_ListGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID -ErrorAction Stop);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = (Get-PnPWeb -ErrorAction Stop).ID.Guid.ToString();
            $ListGUID = $List.Id.Guid.ToString();
            $SiteGUID_WebGUID_ListGUID = ($SiteGuid + "_" + $WebGuid + "_" + $ListGUID);
        }

        $LoadListViews = (Get-SPOnlineHelperPnPProperty -ClientObject $List -Property "Views");
    }
    PROCESS {
        
        Write-Host ($Indentation + "Inventorying List '" + $List.Title + "' Views (" + $List.Views.Count.ToString() + ")");
        foreach ($currView in $List.Views) 
        {
            $ViewSchemaXml = ([xml]$currView.HtmlSchemaXml).View;
            $ViewInternalName = ($ViewSchemaXml.Url.Substring($ViewSchemaXml.Url.LastIndexOf("/") + 1).replace(".aspx", ""));

            $FieldNames = @();            
            foreach ($vf in ($ViewSchemaXml.ViewFields.FieldRef)) {
                $FieldNames += ($vf.Name);
            }

            $ViewSchemaXmlForExtract = "Not Extracted";
            if ($IncludeSchemaXML) {
                $ViewSchemaXmlForExtract = (Format-XmlFieldValueForExtract -XMLFieldVal ($currView.HtmlSchemaXml));
            }

            $ExtractRows += (
                $SiteGUID_WebGUID_ListGUID +
                $Delimiter + ($ViewSchemaXml.DisplayName) +
                $Delimiter + $ViewInternalName +
                $Delimiter + ($ViewSchemaXml.Url) +
                $Delimiter + ($currView.RowLimit.ToString()) +
                $Delimiter + ($currView.Paged) +
                $Delimiter + ($currView.ViewType) +
                $Delimiter + ($currView.DefaultView) +
                $Delimiter + ($currView.ViewQuery) +
                $Delimiter + ($FieldNames -join ";") +
                $Delimiter + $ViewSchemaXmlForExtract
            );
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}

function Inventory-SPOnlineWebParts
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)][PSCustomObject]$InventorySettings
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$SiteUrl
        ,[Parameter(Mandatory=$true,ParameterSetName = "SiteUrl")][String]$ListRootFolder
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")][switch]$SkipConnection
        ,[Parameter(Mandatory=$false,ParameterSetName = "SkipConnection")]$SiteGUID_WebGUID_ListGUID
        ,[Parameter(Mandatory=$true,ParameterSetName = "SkipConnection")]$List        
        ,[Parameter(Mandatory=$false)][String]$Indentation = ""
    );

    BEGIN {
        Write-Warning "Inventory-SPOnlineWebParts: Still in Development - doesn't provide complete list";

        $ExportFilePathAndName = ($InventorySettings.ExportFolder + "\" + $InventorySettings.LogFilePrefix + "WebParts.csv");
        $Delimiter = $InventorySettings.Delimiter;
        $ExtractRows = @();        

        if (-not (test-path $ExportFilePathAndName)) {
            $ExtractRows += (
                "SiteGUID_WebGUID_ListGUID_ItemGUID" + $Delimiter + "Url" + $Delimiter + "WebPartTitle" + $Delimiter + "InstanceId" + $Delimiter + "PropertiesJSON"
            );
        }

        if (-not $SkipConnection) {
            $CredentialParams = ($InventorySettings.PnPCredentialParams);
            Connect-PnPOnline -Url $SiteUrl -ErrorAction Stop @CredentialParams;
            $List = (Get-PnPList -Identity $ListRootFolder);
        }

        $WebObj = (Get-PnPWeb -Includes ID,ServerRelativeUrl);
        if ([String]::IsNullOrEmpty($SiteGUID_WebGUID_ListGUID)) {
            $SiteObj = (Get-PnPSite -Includes ID);
            $SiteGuid = $SiteObj.ID.Guid.ToString();
            $WebGuid = $WebObj.ID.Guid.ToString();
            $ListGUID = $List.Id.Guid.ToString();
            $SiteGUID_WebGUID_ListGUID = ($SiteGuid + "_" + $WebGuid + "_" + $ListGUID);
        }                
        
        $WebRelativeUrl = ($WebObj.ServerRelativeUrl);                
    }
    PROCESS {
        $LibraryInternalName = (Get-SPOnlineHelperPnPProperty -ClientObject $List -Property RootFolder).Name;
        if ($WebRelativeUrl -ne "/") { $LibraryInternalName = ("/" + $LibraryInternalName); }

        $ListItems = (Get-PnPListItem -List $List -Fields Id,FileRef,GUID,File_x0020_Type -PageSize 500);
        $ItemCount = 0;
        foreach ($currItem in $ListItems) {
            if ($currItem["File_x0020_Type"] -eq "aspx") { $ItemCount++; }
        }

        Write-Host ($Indentation + "Inventorying Web Parts " + $List.Title + " (" + $ItemCount + ")");
        foreach ($currItem in $ListItems) {
            if ($currItem["File_x0020_Type"] -eq "aspx") {
                <# ToDo: Handle classic experiece
                $currItemFile = (Get-SPOnlineHelperPnPProperty -ClientObject $currItem -Property File);
                $WebPartManager = $currItemFile.GetLimitedWebPartManager([System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared);

                $WebPartsForPage = (Get-SPOnlineHelperPnPProperty -ClientObject $WebPartManager -Property WebParts);
                foreach ($WebPart in $WebPartsForPage) {
                    $ExtractRows += (
                        ($FullUniqueId + "_" + $currItem["GUID"].ToString()) +
                        $Delimiter + ($currItem["FileRef"]) +
                        $Delimiter + ($WebPart.Title) +
                        $Delimiter + ($WebPart.GetType().ToString())                        
                    );
                }
                #>

                $FolderAndPageName = ($currItem["FileRef"].replace($WebRelativeUrl + $LibraryInternalName + "/", ""));
                $ModernWebPartsForPage = (Get-PnPClientSidePage -Identity $FolderAndPageName);

                foreach ($webPart in $ModernWebPartsForPage.Controls) {
                    $ExtractRows += (
                        ($SiteGUID_WebGUID_ListGUID + "_" + $currItem["GUID"].ToString()) +
                        $Delimiter + ($currItem["FileRef"]) +
                        $Delimiter + ($webPart.Title) +                        
                        $Delimiter + ($webPart.InstanceId) +                        
                        $Delimiter + ($webPart.PropertiesJson)
                    );
                }
            }
        }

        $ExtractRows | Out-File -FilePath $ExportFilePathAndName -Append -Encoding UTF8;
    }
    END {
        if (-not $SkipConnection) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue;
        }
    }
}