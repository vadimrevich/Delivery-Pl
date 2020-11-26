
#******************************************************************************
#
# Create-NitVBSRunOnceKeyHKCU( Path, VBSCmd )
# This Function Creates a strBatCmd Key at the Registry Node
# HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce
# with Value "wscript.exe //B //Nologo  " & strPath & "\" & strBatCmd
# After overload it is run this command
# The Script No Run at Elevated Mode. Further the Rights can be Elevated.
#
# PARAMETERS:   Path -- The Path to strBatCmd
#               VBSCmd -- a Script File with instructions
#               (Windows Script Shell)
#
# RETURNS:      0 -- if Success
#               1 -- on Error takes place
#*****************************************************************************/

function Create-NitVBSRunOnceKeyHKCU{
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$VBSCmd
        )
    $constRunOnce = "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
    $constRunBat = "wscript.exe //B //Nologo "
    $strKey = $VBSCmd
    $strValue = $constRunBat + '"' + $Path + '\' + $VBSCmd + '"'
    Set-ItemProperty -Path $constRunOnce -Name $strKey -Value $strValue
    Write-Output( 0 )
}

#******************************************************************************
#
# New-NitTwoCascadeFolders( Disk, Folder1, Folder2 )
#
# This Function Creates on Disk strDisk the Folders strFolder1 and strFolder2
#
# PARAMETERS:   Disk -- the Disk Drive. Must be Exist!
#               Folder1 -- the first Folder on strDisk. May be Exist.
#               Folder2 -- the Second Folder on strDisk. May be Exist.
#
# RETURNS:      true if Success Create or Folder Exist
#               false if Cascade Can#t Create
#
#******************************************************************************

function New-NitTwoCascadeFolders{
    param(
        [Parameter(Mandatory)][string]$Disk,
        [Parameter(Mandatory)][string]$Folder1,
        [Parameter(Mandatory)][string]$Folder2
        )
    $strDisk = $Disk + "\"
    $blnEnableCreated = $true;
    $blnDrv = Test-Path -Path $strDisk
    #if($blnDrv){"Drive $strDisk Exist $blnDrv"}
    #else { "Drive $strDisk not Exist $blnDrv" }
    if ( $blnDrv -and $blnEnableCreated) {
        $blnEnableCreated = $true
        #"Drive Exist"
    }
    else {
        $blnEnableCreated = $false
    }
    $strPath = $strDisk + $Folder1
    $blnFolderExist = Test-Path -Path $strPath
    if ($blnEnableCreated -and -not $blnFolderExist) {
        $temp = New-Item -Path $strDisk -Name $Folder1 -ItemType "directory"
        $blnCreated = Test-Path -Path $strPath
        if ($blnCreated) {
            $blnEnableCreated = $true
        #    "Path $strPath is Created"
        }
        else {
            $blnEnableCreated = $false
            "Path $strPath is not Created"
        }
    }
    $strPath1 = $strPath
    $strPath = $strPath + "\" + $Folder2
    $blnFolderExist = Test-Path -Path $strPath
    if ($blnEnableCreated -and -NOT $blnFolderExist) {
        $temp = New-Item -Path $strPath1 -Name $Folder2 -ItemType "directory"
        $blnCreated = Test-Path -Path $strPath
        if ($blnCreated) {
            $blnEnableCreated = $true
        #    "Folder $strPath is Created"
        }
        else {
            $blnEnableCreated = $false
            "Folder $strPath is not  Created"
        }
    }
   Write-Output( $blnEnableCreated )
}

#******************************************************************************
#
# Invoke-NitVBSScript( Path, VBSCmd )
# This Function Run a strVBS File
# with Command "cscript //NoLogo " & strPath & "\" & strVBS
#
# PARAMETERS:   Path -- The Path to strVBS
#               VBSCmd -- a VBS File with instructions
#               (Windows Scripts Shell)
#				intTimeOut -- Estimated Time for Running (ms)
#
# RETURNS:      NONE
#
#******************************************************************************

function Invoke-NitVBSScript {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$VBSCmd
        )
    #$constRunVBS = " //B //NoLogo "
    $constRunVBS = " //NoLogo "
    $constOpt = ""
    $vbsApp = "& " + $env:SystemRoot + "\system32\" + "cscript.exe" + $constRunVBS + '"' + $Path + "\" + $VBSCmd + '"' + $constOpt
    Invoke-Expression $vbsApp
}


#******************************************************************************
#
# Download-NitFileFromInt01( Name, URL, Path )
# This Function Upload the File strFile from URL on HTTP/HTTPS Protocols
# and Save it on Local Computer to Path strPath
# Function Uses "BitsAdmin.exe" Funktion for Load File
#
# PARAMETERS:   Name -- a File to be Downloaded (only name and extension)
#               URL -- an URL of the web-site, from which the File
#               is Downloaded
#               Path -- a Place in a Windows Computer (Full path without slash)
#               in which the File is Downloaded
#				intTimeOut -- estimated Time for Download (ms)
#
# RETURNS:      0 -- If File is Normally Downloaded and Created
#               1 -- if File in Path strPath Can't Create
#
#******************************************************************************

function Download-NitFileFromInt01 {
    param (
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$URL,
        [Parameter(Mandatory)][string]$Path
    )
    $strFileUrl = $URL + $Name
    $strLocalFile = $Path + "\" + $Name
    $pathCMD = $env:SystemRoot + "\system32\"
    $strParam = "/Transfer STEA_TRANSFER /DOWNLOAD /Priority FOREGROUND " + $strFileUrl + ' "' + $strLocalFile + '"'
    if ( Test-Path -Path $Path ) {
        $app = "& " + $pathCMD + "bitsadmin.exe" + " " + $strParam
        Invoke-Expression $app
        if (Test-Path -Path $strLocalFile) {
            $intUploadFilesFromInt = 0
        }
        else {
            $intUploadFilesFromInt = 1
        }
    }
    else {
        $intUploadFilesFromInt = 1
    }
    Write-Output( $intUploadFilesFromInt )
}

#******************************************************************************
#
# Download-NitFileFromInt( Name, URL, Path )
# This Function Upload the File strFile from URL on HTTP/HTTPS Protocols
# and Save it on Local Computer to Path strPath
# Function Uses Objects "Microsoft.XMLHTTP" and "Adodb.Stream"
#
# PARAMETERS:   Name -- a File to be Downloaded (only name and extension)
#               URL -- an URL of the web-site, from which the File
#               is Downloaded
#               Path -- a Place in a Windows Computer (Full path without slash)
#               in which the File is Downloaded
#
# RETURNS:      0 -- If File is Normally Downloaded and Created
#               1 -- if File in Path strPath Can't Create
#
#******************************************************************************

function Download-NitFileFromInt {
    param (
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$URL,
        [Parameter(Mandatory)][string]$Path
    )
    $strFileUrl = $URL + $Name
    $strLocalFile = $Path + "\" + $Name
    #"URL: $strFileUrl, File: $strLocalFile"
    $WebClient = New-Object System.Net.WebClient
    if (Test-Path -Path $Path) {
        $WebClient.DownloadFile($strFileUrl, $strLocalFile)
        if (Test-Path -Path $strLocalFile) {
            $intUploadFilesFromInt = 0
        }
        else {
            $intUploadFilesFromInt = 1
        }
    }
    else {
        $intUploadFilesFromInt = 1
    }
    Write-Output( $intUploadFilesFromInt )
}

