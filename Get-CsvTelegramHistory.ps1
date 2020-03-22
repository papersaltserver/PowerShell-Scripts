[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $exportFolder=$null
)
function Test-ExportFolderPath {
    param (
        $exportFolder
    )
    Add-Type -AssemblyName "System.Windows.Forms"
    if($null -eq $exportFolder -or -not (Test-Path $exportFolder)){
        $foldername = [System.Windows.Forms.FolderBrowserDialog]::new()
        $foldername.Description = "Select a folder containing Telegram export"
        $foldername.RootFolder = "MyComputer"
        
        if($foldername.ShowDialog() -eq "OK")
        {
            $exportFolder = $foldername.SelectedPath
        }
    }
    
    return $exportFolder
    
}
function Get-ExportedMessagesFiles {
    param (
        $exportFolder
    )
    (Get-ChildItem -Path $exportFolder\* -Include "messages*.html").FullName
}
function Get-ChatName {
    param (
        [string]
        $firstFile
    )
    Add-Type -Path ".\HtmlAgilityPack.dll"
    $htmlDocument = [HtmlAgilityPack.HtmlDocument]::new()
    $fs = [System.IO.File]::OpenRead($firstFile)
    $htmlDocument.Load($fs)
    $headerDiv = $htmlDocument.DocumentNode.Descendants("div") | Where-Object{$_.Attributes["class"].Value -like "page_header"}
    $headerDiv.InnerText.Replace("`r","").Replace("`n","").Trim()
}
function Get-MessagesFromFile {
    param (
        $exportFileName
    )
    Add-Type -Path ".\HtmlAgilityPack.dll"
    $htmlDocument = [HtmlAgilityPack.HtmlDocument]::new()
    $fs = [System.IO.File]::OpenRead($exportFileName)
    $htmlDocument.Load($fs)
    $messages = foreach($div in $htmlDocument.DocumentNode.Descendants("div")){
        if($div.Attributes["class"].Value.Contains("message") -and $div.Attributes["class"].Value.Contains("default")){
            $messageDate = ($div.Descendants("div") | Where-Object{$_.Attributes["class"].Value.Contains("date")}).Attributes["title"].Value
            $messageId = $div.Attributes["id"].Value
            $currentUserDiv = $div.Descendants("div") | Where-Object{$_.Attributes["class"].Value.Contains("from_name")}
            if($null -ne $currentUserDiv){
                $lastUser = $currentUserDiv.InnerText.Trim()
            }
            #$messageBody = ($div.Descendants("div") | Where-Object {$_.Attributes["class"].Value.Contains("body")}).ChildNodes | Where-Object{$_.NodeType -eq "Element"} | Select-Object -Last 1
            $messageBody = ($div.ChildNodes | Where-Object{$_.NodeType -eq "Element" -and $_.Attributes["class"].Value -like "*body*"}).ChildNodes | Where-Object{$_.NodeType -eq "Element"} | Select-Object -Last 1
            $messageType = $messageBody.Attributes["class"].Value
            if($messageType -like "*text*"){
                $messageText = $messageBody.InnerText.Replace("`r","").Replace("`n","").Trim()
            }elseif ($messageType -like "*media_wrap*") {
                $hrefElement = $messageBody.Descendants("a") | Where-Object{$true}
                if($null -ne $hrefElement){
                    $messageText = $hrefElement.Attributes["href"].Value
                }else {
                    $messageText = $messageBody.InnerText.Replace("`r","").Replace("`n","").Trim()
                }
            }elseif($messageType -like "*forwarded*"){
                $messageText = $messageBody.InnerText.Replace("`r","").Replace("`n","").Trim()
            }else{
                Write-Host $messageType
                Write-Host $messageBody.InnerText.Replace("`r","").Replace("`n","").Trim()
                $messageText = ""
            }
            [PSCustomObject]@{
                Id = $messageId
                Date = $messageDate
                User = $lastUser
                Type = $messageType
                Text = $messageText
            }
        }
        
    }
    $messages
}
function Get-AgitityPack {
    if(-not(Test-Path ".\HtmlAgilityPack.dll")){
        $agPackUrl = 'https://www.nuget.org/api/v2/package/HtmlAgilityPack/'
        $output = (Resolve-Path .\).Path+"\HtmlAgilityPack.zip"
        $webCl = [System.Net.WebClient]::new()
        $webCl.DownloadFile($agPackUrl, $output)
        Add-Type -Assembly System.IO.Compression.FileSystem
        $zip = [IO.Compression.ZipFile]::OpenRead($output)
        $entries=$zip.Entries | Where-Object {$_.FullName -like 'lib/Net45/HtmlAgilityPack.dll'}
        [IO.Compression.ZipFileExtensions]::ExtractToFile( $entries, (Resolve-Path .).Path+"\" + $entries.Name)
        $zip.Dispose()
        Remove-Item $output -Force
    }
}


Get-AgitityPack
Add-Type -Path ".\HtmlAgilityPack.dll"
$exportFolder = Test-ExportFolderPath $exportFolder
$exportFiles = @()
$exportFiles += Get-ExportedMessagesFiles $exportFolder
$csvName = Get-ChatName $exportFiles[0]
if(Test-Path "$exportFolder\$csvName.csv"){
    $answer = Read-Host "File Exist, Continue? Y/N: "
    if($answer -ne "Y"){
        Write-Output "Exiting..."
        exit 0
    }
}
foreach($ef in $exportFiles){
    Get-MessagesFromFile $ef | Export-Csv "$exportFolder\$csvName.csv" -Delimiter "," -NoTypeInformation -Append -Encoding unicode
}