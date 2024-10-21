param(
[parameter(Mandatory=$true, Position=1)]
[string]$ClientId,
[parameter(Mandatory=$true, Position=1)]
[string]$ClientSecret,
[parameter(Mandatory=$true, Position=1)]
[string]$URL,
[parameter(Mandatory=$true, Position=2)]
[string]$WebResourceFolderPath,
[parameter(Mandatory=$true, Position=3)]
[string]$InputFilePath,
[parameter(Mandatory=$true, Position=4)]
[string]$SolutionName
)
 
$Module = Get-InstalledModule -Name "Microsoft.Xrm.Data.Powershell" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
if($Module -ne $null)
{
    Write-host "Module is already installed!!!!, `nModule details:";
    $Module | Format-Table -Wrap 
}
else
{
    Write-host "Module is not installed,`nInstalling Module('Microsoft.Xrm.Data.Powershell')...";
    Install-Module -Name Microsoft.Xrm.Data.Powershell -Force -Scope CurrentUser
    Write-host "Module('Microsoft.Xrm.Data.Powershell') is installed successfully....";
}
 
$Conn = Get-CrmConnection -ConnectionString "AuthType=ClientSecret;Url=$URL;Timeout=02:00:00; Domain=; ClientId=$ClientId;ClientSecret=$ClientSecret" -MaxCrmConnectionTimeOutMinutes 200 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
 
# Show the Connection is estalbished with target instane
if($Conn.IsReady){
Write-Host "Service connection is successfully established to '"$Conn.ConnectedOrgFriendlyName"'`n"
}
else
{
    Write-error "Service connection is not established !!!";
    return;
}
 
function CreateWebResource{
    param(
        [parameter(Mandatory=$true, Position=1)]
        [string]$Name,
        [parameter(Mandatory=$true, Position=2)]
        [string]$DisplayName,
        [parameter(Mandatory=$false, Position=3)]
        [string]$Description,
        [parameter(Mandatory=$true, Position=4)]
        [int]$WRType,
        [parameter(Mandatory=$true, Position=5)]
        [string]$Content,
        [parameter(Mandatory=$false, Position=6)]
        [int]$LanguageCode,
        [parameter(Mandatory=$true, Position=7)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$Conn
    )
    $webSrc =  New-Object Microsoft.Xrm.Sdk.Entity("webresource");
    $webSrc.Attributes["name"] = $Name;
    $webSrc.Attributes["displayname"] = $DisplayName;
    if($Description -ne "")
    {
    $webSrc.Attributes["description"] = $Description;
    }
    $webSrc.Attributes["webresourcetype"] = [Microsoft.Xrm.Sdk.OptionSetValue] $WRType;
    $webSrc.Attributes["languagecode"] = $LanguageCode;
    # Set the content of the webresource languagecode
    $webSrc.Attributes["content"] = $Content;
    Write-Verbose $Content.Length       
    $WebResourceID=$Conn.Create($webSrc);
 
    $addWebResourceReq = new-object Microsoft.Crm.Sdk.Messages.AddSolutionComponentRequest;
    $addWebResourceReq.ComponentType = 61;
    $addWebResourceReq.ComponentId = $WebResourceID;
    $addWebResourceReq.SolutionUniqueName = "$SolutionName";
    $Conn.Execute($addWebResourceReq)
 
    
 
 
}
 
function UpdateWebResource{
    param(
        [parameter(Mandatory=$true, Position=1)]
        [string]$WRId,
        [parameter(Mandatory=$true, Position=2)]
        [string]$DisplayName,
        [parameter(Mandatory=$false, Position=3)]
        [string]$Description,
        [parameter(Mandatory=$true, Position=3)]
        [string]$Content,
        [parameter(Mandatory=$false, Position=4)]
        [int]$LanguageCode,
        [parameter(Mandatory=$true, Position=5)]
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$Conn
    )
 
    $webSrc =  New-Object Microsoft.Xrm.Sdk.Entity("webresource");
    $webSrc.Id = $WRId;
    $webSrc.Attributes["displayname"] = $DisplayName;
    if($Description -ne "")
    {
    $webSrc.Attributes["description"] = $Description;
    }
    $webSrc.Attributes["languagecode"] = $LanguageCode;
 
    # update the content of the webresource
    $webSrc.Attributes["content"] = $Content;
    Write-Verbose $Content.Length
    $Conn.Update($webSrc);
 
    $addWebResourceReq = new-object Microsoft.Crm.Sdk.Messages.AddSolutionComponentRequest;
    $addWebResourceReq.ComponentType = 61;
    $addWebResourceReq.ComponentId = $WRId;
    $addWebResourceReq.SolutionUniqueName = "$SolutionName";
    $Conn.Execute($addWebResourceReq)
    
 
}
$WebResourceTypes =@{
HTM = 1;
HTML = 1;
CSS = 2;    
JS = 3; 
XML = 4;            
PNG = 5;
JPG = 6;
GIF = 7;
XAP = 8;
XSL = 9;
ICO = 10;
SVG = 11;
RESX = 12
}
 
Measure-Command{
[xml]$WebResourcesList = Get-Content $InputFilePath
 
 $SummaryTable = @() 
 
foreach($Webresource in $WebResourcesList.CRMWebResources.CRMWebResource)
{
    $RepoContent = "";
    $FilePath = $Webresource.Include 
    $DisplayName = $Webresource.DisplayName
    $UniqueName = $Webresource.UniqueName
    $Description = @{$true=$Webresource.Description ;$false=""}[$Webresource.Description -ne $null] 
    Write-Host "Validating if the webresource '$UniqueName' exists in "$Conn.ConnectedOrgFriendlyName" environment"
 
    $WebResourceType = $Webresource.WebResourceType
    # get content from repo
    $RepoFilePath = "$WebResourceFolderPath\$FilePath";
    $RepoFilePath  = [System.Web.HttpUtility]::UrlDecode($RepoFilePath);
    
   $IsFileExists = Test-Path -Path $RepoFilePath;
    if($IsFileExists){
    $RepoContent = Get-Content "$RepoFilePath" # -Raw -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $Extension = [System.IO.Path]::GetExtension($RepoFilePath).replace(".",""); 
   if($RepoContent -ne $null -and $RepoContent -ne ""){
 
    $base64Converted= [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($RepoFilePath))
     $base64String -eq $base64Converted 
$WRFetch = @"
<fetch>
<entity name="webresource" >
<attribute name="organizationid" />
<attribute name="languagecode" />
<attribute name="description" />
<attribute name="webresourceidunique" />
<attribute name="content" />
<attribute name="solutionid" />
<attribute name="componentstate" />
<attribute name="displayname" />
<attribute name="name" />
<attribute name="webresourceid" />
<attribute name="webresourcetype" />
<attribute name="organizationidname" />
<attribute name="webresourcetypename" />
<filter>
<condition attribute="name" operator="eq" value="$UniqueName" />
</filter>
</entity>
</fetch>
"@
    $WRFetchExpr = New-Object  Microsoft.Xrm.Sdk.Query.FetchExpression($WRFetch);
 
    # Retrieve webresource from instance
 
    $WebResourceRecords = $Conn.RetrieveMultiple($WRFetchExpr);
 
    if($WebResourceRecords.Entities.Count -gt 0)
    {
        Write-Host "Web Resource ('$UniqueName') exists in "$Conn.ConnectedOrgFriendlyName" environment!!"
        $WRId = $WebResourceRecords.Entities[0].Id;
        $WRName = $WebResourceRecords.Entities[0].Attributes["name"];
        $Base64SourceContent = $WebResourceRecords.Entities[0].Attributes["content"];
 
        if($base64Converted.ToString().Length -gt 160)
        {
       # Write-Host "Source "$base64Converted.Length": Last 150 characters of base64String for checking:"
       # Write-Host $base64Converted.ToString().Substring($base64Converted.Length-150)
        }
        if($Base64SourceContent.ToString().Length -gt 160)
        {
       # Write-Host "Target "$Base64SourceContent.Length": Last 150 characters of base64String for checking:"
       # Write-Host $Base64SourceContent.ToString().Substring($Base64SourceContent.Length-150)
        }
 
        # condition to always override the script and web resource
 
        Write-Host "Vailidating if Web Resource ('$UniqueName') has any changes to update..."
 
         if( $base64Converted.ToString() -ne $Base64SourceContent.ToString())
        {
            try
            {
             Write-Host "Changes found...."
             Write-Host "Updating Web Resource ('$UniqueName') ...."
             $response = UpdateWebResource -WRId $WRId -DisplayName $DisplayName -Description $Description  -Content $base64Converted -LanguageCode "1033" -Conn $Conn
             Write-Host "Web Resource ('$UniqueName') is updated successfully."
             Write-Host "Web Resourse "$DisplayName" added to the solution "$SolutionName" !!"
             $SummaryTable += "$UniqueName Updated"
            }
            catch{
            Write-Warning "Unable to update the webresource ('$UniqueName')"
            Write-Warning $_.Exception.Message
            }
        }
        else
        {
            Write-Host "No changes found in $UniqueName'....skipping to the next webresource!!"
            $SummaryTable += "$UniqueName Skipped"
        }
    }
    else
    { 
        Write-Host "Web Resource ('$UniqueName') does not exist in "$Conn.ConnectedOrgFriendlyName" environment!!"
        $Wrtype = [int]$WebResourceTypes.($Extension)
        try{
        Write-Host "Creating Web Resource ('$UniqueName') ......"
        $response =  CreateWebResource -Name "$UniqueName" -DisplayName "$DisplayName" -Description $Description  -WRType $Wrtype -Content $base64Converted -LanguageCode "1033" -Conn $Conn
        Write-Host "Web Resource ('$UniqueName') has been created successfully"
        Write-Host "Web Resourse "$DisplayName" added to the solution "$SolutionName" !!"
        $SummaryTable += "$UniqueName Created"
        }
        catch{
        Write-Warning "Unable to create the webresource ('$UniqueName')"
        Write-Warning $_.Exception.Message
        }
    }
    }
    else 
    {
         Write-Warning "Webresource file is empty, No code/data is present !!!`nFile Path: $RepoFilePath "
         Write-Warning "Skipping the web resource ('$UniqueName')."
    }
    }
    else
    {
        Write-Warning "File path does not exists !!!"
    }
 
    Write-Host "`n###############-------------------------------------------------------------------------------------###############`n"
}
 
Write-Host "Summary Table:`n"
 
Write-Host "_______________________________________________________________________________"
Write-Host "| Web Resource Name                             | Status                      |"
Write-Host "|_______________________________________________|_____________________________|"
 
 
 
$SummaryTable | ForEach-Object { 
    $columns = $_ -split ' '
    $webResourceName = $columns[0].PadRight(45)
    $status = $columns[1].PadRight(27)
    Write-Host "| $webResourceName | $status |"
    Write-Host "|_______________________________________________|_____________________________|"
}
 
 
 
 
#Write-Host "Time taken for webreources deployment is:`n"
 
}
