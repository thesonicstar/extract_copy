# extract_copy
Powershell Script to upzip a specific file type and copy that to another location

#ERROR REPORTING ALL
Set-StrictMode -Version latest

#----------------------------------------------------------
#STATIC VARIABLES
#----------------------------------------------------------
$search = ".jpg"                       #file type
$dest   = "D:\####\####\####\"               #destination
$zips   = "D:\####\####\####\"           #Source

#----------------------------------------------------------
#FUNCTION GetZipFileItems
#----------------------------------------------------------
Function GetZipFileItems
{
  Param([string]$zip)
  
  $split = $split.Split(".")
  $dest = $dest + "\" + $split[0]
  If (!(Test-Path $dest))
  {
    Write-Host "Created folder : $dest"
    $strDest = New-Item $dest -Type Directory
  }

  $shell   = New-Object -Com Shell.Application
  $zipItem = $shell.NameSpace($zip)
  $items   = $zipItem.Items()
  GetZipFileItemsRecursive $items
}

#----------------------------------------------------------
#FUNCTION GetZipFileItemsRecursive
#----------------------------------------------------------
#$dest   = "D:\CDMS\NMT\Laserail 3000\_2016\FTP Data\New Folder\"
Function GetZipFileItemsRecursive
{
  Param([object]$items)

  ForEach($item In $items)
  {
    If ($item.GetFolder -ne $Null)
    {
      GetZipFileItemsRecursive $item.GetFolder.items()
    }
    $strItem = [string]$item.Name
    If ($strItem -Like "*$search*")
    {
      If ((Test-Path ($dest + "\" + $strItem)) -eq $False)
      {
        Write-Host "Copied file : $strItem from zip-file : $zipFile to destination folder"
        $shell.NameSpace($dest).CopyHere($item)
      }
      Else
      {
        Write-Host "File : $strItem already exists in destination folder"
      }
    }
  }
}

#----------------------------------------------------------
#FUNCTION GetZipFiles
#----------------------------------------------------------
Function GetZipFiles
{
  $zipFiles = Get-ChildItem -Path $zips -Recurse -Filter "*.zip" | % { $_.DirectoryName + "\$_" }
  
  ForEach ($zipFile In $zipFiles)
  {
    $split = $zipFile.Split("\")[-1]
    Write-Host "Found zip-file : $split"
    GetZipFileItems $zipFile
  }
}
#RUN SCRIPT 
GetZipFiles
#Finished

#----------------------------------------------------------
# Search and Copy .jpg file
#----------------------------------------------------------
#Searches below directory for .JPG files and copies files written in the last 720 minutes

Get-ChildItem -Path "D:\####\###\####\" -Filter *.jpg* -Recurse| ? {$_.LastWriteTime -gt (Get-Date).AddMinutes(-720)} |
#Copy .tpe files to below location
Copy-Item -Destination "\\####\###\####\"
