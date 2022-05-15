function Get-SPOnlineHelperAppCredential
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$StoredCredentialName
    );
    
    $Cred = (Get-PnPStoredCredential -Name $StoredCredentialName);
    $ClientId = ($Cred.UserName);
    $ClientSecret = ([System.Net.NetworkCredential]::new("", $Cred.Password).Password);    
    return (New-Object PSCustomObject -Property @{ 
        ClientID = $ClientId; 
        ClientSecret = $ClientSecret;
    });
}

function Add-SPOnlineHelperAppCredential
{
    [cmdletbinding()]
    param (
         [Parameter(Mandatory=$true)][String]$StoredCredentialName
        ,[Parameter(Mandatory=$true)][String]$AppID 
        ,[Parameter(Mandatory=$true)][String]$AppSecret
    );
    
    Add-PnPStoredCredential -Name $StoredCredentialName -Username $AppID -Password (ConvertTo-SecureString -String $AppSecret -AsPlainText -Force);
}

function Get-SPOnlineHelperPnPProperty
{
    [cmdletbinding()]
	param (                        
          [Parameter(Mandatory=$true)]$ClientObject
         ,[Parameter(Mandatory=$true)][String[]]$Property         
    );   

    $TheResult = $null;

    $RetrievedProp = $false;
    $Counter = 0;
    $MAX_TRIES_TO_RETRIEVE_PROPERTY = 10;
    $Err = $null;        
    do {
        $Counter++;
        try {
            $TheResult = (Get-PnPProperty -ClientObject $ClientObject -Property $Property -ErrorAction Stop);
            $RetrievedProp = $true;
        } catch {
            #ToDo: Check only for throttling error
            Start-Sleep -Seconds ($Counter * 30); #If it errors sleep as likely being throttled            
        }
    }
    while (($RetrievedProp -eq $false) -and ($Counter -lt $MAX_TRIES_TO_RETRIEVE_PROPERTY))

    if ($RetrievedProp -eq $false) {
        throw [System.Exception]::new("Failed to retrieve property " + ($Property -join ";"));
    }
    
    return $TheResult;
}



function Call-SPOnlineHelperListRestService
{
    [cmdletbinding()]
	param (                        
          [Parameter(Mandatory=$true)]$SiteUrl
         ,[Parameter(Mandatory=$true)]$ListGUID  
         ,[Parameter(Mandatory=$true)]$FieldsCommaSeperated  
         ,[Parameter(Mandatory=$false)]$Filters
    );    

    $SelectFields = ("%24select=" + $FieldsCommaSeperated);
    $RestApiUrl = ($SiteUrl + "/_api/Web/Lists(guid'"  + $ListGUID + "')/Items?" + $SelectFields);
    if (-not [String]::IsNullOrEmpty($Filters)) { $RestApiUrl += ("&" + $Filters); }
    $RestApiUrl += "&%24Paged=TRUE";
    $RestApiUrl += "&%24Top=2000";    
    $ListItemData = @();

    do {        
        <#
        #$Result = Invoke-RestMethod -Method Get -WebSession $WebSession -Uri $RestApiUrl        
        $ResultTransformed = (($Result.ToString() -creplace "`"Id`"", "_ID") | ConvertFrom-JSON); #Rest API returns ID as "ID" and "Id" when ID is selected :<
        foreach ($val in $ResultTransformed.d.results) {
            $currItem = (New-Object PSCustomObject -Property @{ });
            foreach ($kv in $val.PSObject.Properties) {
                if ((-not $currItem.($kv.Name)) -and ($kv.Name -ne "__metadata")) {
                    $currItem | Add-Member -MemberType NoteProperty -Name ($kv.Name) -Value ($kv.Value)
                }
            }            

            $ListItemData += $currItem;
        }
        
        $RestApiUrl = ($ResultTransformed.d.__next);
        #>

        $Result = (Invoke-PnPSPRestMethod -Url $RestApiUrl -ErrorAction Stop); 
        if ($Result.Value) {
            $ListItemData += ($Result.Value);
        }

        $RestApiUrl = ($Result.'odata.nextLink');                
    }
    while ($RestApiUrl -ne $null)

    return $ListItemData;
}

function Get-SPOnlineHelperListRootPath
{
    param(
         [Parameter(Mandatory=$true)]$List
        ,[Parameter(Mandatory=$true)]$WebRelativeUrl
    );

    $RelativeUrl = (Get-SPOnlineHelperPnPProperty -ClientObject ($List.RootFolder) -Property ServerRelativePath);                                
    if ($WebRelativeUrl -eq "/") {
        $RelativeUrl = ($RelativeUrl.DecodedUrl.ToString().Substring(1)); #Removes the leading '/'
    } else {
        $RelativeUrl = ($RelativeUrl.DecodedUrl.ToString().Replace($WebRelativeUrl + "/", ""));
    }

    return $RelativeUrl;
}



function Process-SPOnlineTermStoreTerm {
    param(
         [Parameter(Mandatory=$true)]$TermStore        
        ,[Parameter(Mandatory=$true)]$Group
        ,[Parameter(Mandatory=$true)]$GroupGUID
        ,[Parameter(Mandatory=$true)]$TermSet
        ,[Parameter(Mandatory=$true)]$TermSetGUID
        ,[Parameter(Mandatory=$true)]$Path
        ,[Parameter(Mandatory=$true)]$Term        
        ,[Parameter(Mandatory=$true)]$Depth
        ,[Parameter(Mandatory=$true)]$Tree
    );

    $TermProps = (Get-SPOnlineHelperPnPProperty -ClientObject $Term -Property Description, Terms, Labels, CustomProperties);

    $theResult = @();
    $theLabels = @();
    $Tree += ("|" + $Term.Name);

    foreach ($l in $Term.Labels) {
        $theLabels += $l.Value;
    }
    
    $theCustomProperties = @{};
    foreach($customProperty in $Term.CustomProperties.GetEnumerator()){
        $theCustomProperties.Add($customProperty.Key, $customProperty.Value);
    }

    $theResult += (New-Object PSCustomObject -Property @{
        TermStore = $TermStore;
        Group = $Group;
        GroupGUID = $GroupGUID
        TermSet = $TermSet;
        TermSetGUID = $TermSetGUID;
        Path = $Path;
        Name = $Term.Name;
        GUID = $Term.Id.Guid;
        Description = $Term.Description;
        AvailableForTagging = $Term.IsAvailableForTagging;
        CreatedDate = $Term.CreatedDate;
        LastModifiedDate = $Term.LastModifiedDate;
        Labels = $theLabels;
        CustomProperties = $theCustomProperties;
        ChildTerms = ($Term.Terms.Count);
        Depth = $Depth;
        Tree = $Tree;
    });

    if ($Depth -eq 1) {
        $Path = ($Term.Name);
    } else {
        $Path = ($Path + " > " + $Term.Name);
    }

    foreach ($childTerm in $Term.Terms) {                        
        $ChildTermData = (Process-SPOnlineTermStoreTerm -TermStore $TermStore -Group $Group -GroupGUID $GroupGUID -TermSet $TermSet -TermSetGUID $TermSetGUID -Path $Path -Term $childTerm -Depth ($Depth + 1) -Tree $Tree);
        foreach ($ctd in $ChildTermData) {
            $theResult += $ctd;
        }
    }

    return $theResult;
}


function Get-SPOnlineTermSetDictionary
{
    param(
         [Parameter(Mandatory=$true)]$SiteUrl
        ,[Parameter(Mandatory=$false)]$SpecificGroup
        ,[Parameter(Mandatory=$false)]$SpecificTermSet
    );

    $theResult = @{};    

    $session = Get-PnPTaxonomySessionl
    foreach ($TermStore in $session.TermStores) {
    
        Write-host ($TermStore.Name);
        $AllGroups = (Get-PnPProperty -ClientObject $TermStore -Property Groups);
        foreach ($Group in $AllGroups) {
            
            if (([String]::IsNullOrEmpty($SpecificGroup)) -or ($Group.Name -eq $SpecificGroup)) {
                
                Write-host ("`t" + $Group.Name);
                $AllTermSets = (Get-PnPProperty -ClientObject $Group -Property TermSets);
                foreach ($TermSet in $AllTermSets) {

                    if (([String]::IsNullOrEmpty($SpecificTermSet)) -or ($TermSet.Name -eq $SpecificTermSet)) {
                        
                        $TermTree = ($Group.Name + "|" + $TermSet.Name);
                        Write-host ("`t`t" + $TermSet.Name);
                        $AllTerms = (Get-PnPProperty -ClientObject $TermSet -Property Terms);
                        foreach ($Term in $TermSet.Terms) {
                               
                            $TermInfo = (Process-SPOnlineTermStoreTerm -TermStore ($TermStore.Name) -Group ($Group.Name) -GroupGUID ($Group.Id) -TermSet ($TermSet.Name) -TermSetGUID ($TermSet.Id) -Path "" -Term $Term -Depth 1 -Tree $TermTree);
                            foreach ($t in $TermInfo) {
                                $theResult[$t.GUID] += $t;
                            }
                        }
                    }
                }
            }
        }
    }

    return $theResult
}


function Get-SPOnlineTermSets
{
    $theResult = @{};    

    $session = Get-PnPTaxonomySession;
    foreach ($TermStore in $session.TermStores) {
    
        Write-host ($TermStore.Name);
        $AllGroups = (Get-PnPProperty -ClientObject $TermStore -Property Groups);
        foreach ($Group in $AllGroups) {
            
            if (([String]::IsNullOrEmpty($SpecificGroup)) -or ($Group.Name -eq $SpecificGroup)) {
                
                Write-host ("`t" + $Group.Name);
                $AllTermSets = (Get-PnPProperty -ClientObject $Group -Property TermSets);
                foreach ($TermSet in $AllTermSets) {
                    $theResult[$TermSet.ID.ToString()] = $TermSet;
                }
            }
        }
    }

    return $theResult
}



function Get-SPOnlineHelperContentTypeTree
{
    [cmdletbinding()]
	param (                                
         [Parameter(Mandatory=$true)]$ContentType
    );

    $ContentTypeTree = (New-Object PSCustomObject -Property @{
        IdPath = "";
        NamePath = "";
    })
    
    $ParentContentType = (Get-SPOnlineHelperPnPProperty -ClientObject $ContentType -Property "Parent")
    if ($ParentContentType.Name -ne "System") {
        $ParentContentTypeTree = (Get-SPOnlineHelperContentTypeTree -ContentType $ParentContentType);        
        $ContentTypeTree.NamePath = $ParentContentTypeTree.NamePath + " > ";
        $ContentTypeTree.IDPath = $ParentContentTypeTree.IDPath + " > ";
    }

    $ContentTypeTree.NamePath += ($ContentType.Name);
    $ContentTypeTree.IDPath += ($ContentType.Id.StringValue);    
    return $ContentTypeTree;
}