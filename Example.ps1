#1. Load Inventory Script
$ModulePath = "{SCRIPT PATH}";
. ($ModulePath + "\Module-SPOnlineInventory.ps1") -ModulePath $ModulePath;

#2. Set tenant Url
$TenantUrl = "https://{TENANT}.sharepoint.com";

#3. Get Credentials to run the inventory scripts with
$Tenant = "{TENANT}.onmicrosoft.com";
$AppCreds = (Get-SPOnlineHelperAppCredential -StoredCredentialName "{STORED CRED NAME}");

#Certificate Authentication
$PnPCredentialParams = @{
    ClientID = ($AppCreds.ClientID);
    Thumbprint = ($AppCreds.ClientSecret);
    Tenant = $Tenant;
};

#Client ID and Secret
<#
$PnPCredentialParams = @{
    ClientID = ($AppCreds.ClientID);
    ClientSecret = ($AppCreds.ClientSecret);
};

#Interactive
$PnPCredentialParams = @{
    Interactive = $true;
};
#>


#3. Inventory Settings
$InventorySettingsParams = @{
    PnPCredentialParams = $PnPCredentialParams;
    LogFilePrefix = "{LOG PREFIX}";
    ExportFolder = "{EXPORT FOLDER PATH}";
    Delimiter = "{#]";
};

$InventorySettings = (Get-SPOnlineInventorySettings @InventorySettingsParams);

#4. Remove existing CSV Extracts
Get-ChildItem -LiteralPath ($InventorySettings.ExportFolder) -Filter "*.csv" | Remove-Item;

#5. Run Inventory Command
Run-SPOnlineFullInventorySitesWebsLists -InventorySettings $InventorySettings -TenantUrl $TenantUrl;

#Example Function Calls
#Run-SPOnlineInventoryFullInventory -InventorySettings $InventorySettings -TenantUrl $TenantUrl -WaitBetweenSites;
#Run-SPOnlineFullInventorySitesWebsLists -InventorySettings $InventorySettings -TenantUrl $TenantUrl;
#Inventory-SPOnlineLists -InventorySettings $InventorySettings -SiteUrl ($TenantUrl + "/sites/{SITE}") -InventoryListFields -InventoryListViews;
#Inventory-SPOnlineListItems -InventorySettings $InventorySettings -SiteUrl ($TenantUrl + "/sites/{SITE}") -ListRootFolder "/Shared Documents"