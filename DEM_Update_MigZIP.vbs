'#########################################################################################
'# Script to copy extra data from Persona Management to the Local Documents Directory
'#  Josh Spencer / Chris Halstead - VMware
'# There is NO support for this script - it is provided as is
'# 
'# Version 2.0 - July 15, 2020
'##########################################################################################

Option Explicit
Const Read = 1
Const Write = 2

Dim objFso, objFile
Dim objFolder, objSubFolders, objSubfolder
Dim flexEngine
Dim strTemp, strZip, strLog, strText, strReplace, strProfile, strPD

Set flexEngine = CreateObject("ImmidioFlexProfiles.Engine")
Set objFso = CreateObject("Scripting.FileSystemObject")

'************************** EDIT THESE BEFORE RUNNING *************************************
strProfile = "\\fs1\uem_profile" 'DEM profile archives path 
strTemp = "c:\temp" 'Local working directory
strLog = "C:\temp\zip-unzip.log" 'flexEngine logging
strPD = "D:\\Users" 'Replace D with the drive letter used for your persistent disk mapping. 
'Note: The double \\ above is intentional, don't modify this!
' *****************************************************************************************

Function Update(userpath)
	'--- Unpack the migration.zip ---
	' UnZip(destination folder, source archive, log file)
	Call flexEngine.UnZip(strTemp, strZip, strLog)

	'Read the REG file. Note the -1 is required as this is Unicode, not ASCII
	Set objFile = objFSO.OpenTextFile(strTemp + "\registry\flex profiles.reg", Read, False, -1)
	strText = objFile.ReadAll
	objFile.Close
	
	'Replace Persistent Disk drive letter.
	strReplace = Replace(strText, strPD, "C:\\Users", 1, -1, 1)
	
	'Write the updated REG file. Note the -1 is required as this is Unicode, not ASCII
	Set objFile = objFSO.OpenTextFile(strTemp + "\registry\flex profiles.reg", Write, False, -1)
	objFile.WriteLine strReplace
	objFile.Close

	'--- Rewrite the migration.zip ---
	' Zip(source folder, destination archive, log file, compress yes/no)
	Call flexEngine.Zip(strTemp, strZip, strLog, True)
	
End Function

'Execute script
Set objFolder = objFSO.GetFolder(strProfile)
Set objSubFolders = objFolder.subfolders
For each objSubfolder in objSubFolders
	If objFSO.FileExists(ObjSubfolder + "\archives\migration.zip") Then
		strZip = ObjSubfolder + "\archives\migration.zip"
		Call Update(strZip) 
	End If
next

wscript.echo "Done!"