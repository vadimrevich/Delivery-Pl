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
' SUBROUTINE ScriptTestRunKeyCopied
'
' This Subroutine Copy & Run NIT Update Script on Computer at Test Mode
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub ScriptTestRunKeyCopied()

        Dim local_Path 'Local Path to Command File with Drive Letter

        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2

        CreatedCascade STR_DISK, STR_FOLDER1, STR_FOLDER2
        Copy_VBS THREAD_VBS, local_Path
        CreatedRunOnceKey local_Path, THREAD_VBS

End Sub

' *****************************************************************************
'
' SUBROUTINE SimpleScriptTestRunKeyCopied
'
' This Subroutine Copy & Run NIT Update Script on Computer
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub SimpleScriptTestRunKeyCopied()

        Dim local_Path      'Local Path to Command File with Drive Letter
        Dim iTemp

        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2

        iTemp = CreateTwocascadeFolders(STR_DISK, STR_FOLDER1, STR_FOLDER2)
        Copy_VBS THREAD_VBS, local_Path
        iTemp = CreateRunOnceKey(local_Path, THREAD_VBS)

End Sub
