# VBS
'Declare variables
Dim objFSO As Object
Dim objFile As Object
Dim strSourceFile As String
Dim strDestFile As String
Dim strDate As String
Dim strTime As String
Dim intDays As Integer

'Set variables
objFSO = CreateObject("Scripting.FileSystemObject")
strSourceFile = "C:\Users\Public\Desktop\myfile.txt"
strDestFile = "C:\Users\Public\Desktop\archives\myfile.txt"
intDays = 7

'Get date and time
strDate = FormatDateTime(Now(), "yyyy-mm-dd")
strTime = FormatDateTime(Now(), "hh:mm:ss")

'Check if file is older than specified number of days
If DateDiff("d", strDate, strDate - intDays) > 0 Then

'Move file
objFSO.MoveFile strSourceFile, strDestFile

'End if
End If 



-------------------------------------------****************************************_________________________________---------------------------------
'Declare variables
Dim objFSO As Object
Dim objFile As Object
Dim strSourceFile As String
Dim strDestFile As String
Dim strDate As String
Dim strTime As String

'Set variables
objFSO = CreateObject("Scripting.FileSystemObject")
strSourceFile = "C:\Users\Public\Desktop\myfile.txt"
strDestFile = "C:\Users\Public\Desktop\archives\myfile.txt"

'Copy file
objFSO.CopyFile strSourceFile, strDestFile

'Get date and time
strDate = FormatDateTime(Now(), "yyyy-mm-dd")
strTime = FormatDateTime(Now(), "hh:mm:ss")

'Send email
' Replace "yourname@email.com" with your email address
SendEmail("yourname@email.com", "File Copied", "The file myfile.txt was copied to the archives folder on " & strDate & " at " & strTime) 



*********************************************************************************************************************************************************************************
Option Explicit
dim objFolder,fso, rootdir, targdir,ArchWODir,objLog,intResultPDf,Currentdt,objErr,objEmail,flgEmail,objshl,shell,objEmailRecon 
dim targdirNID 
set fso = CreateObject("Scripting.FileSystemObject")


set objshl= createobject("wscript.shell")
'Location where original file will received 
rootdir = "E:\ApplicationFolder\Temp GESBI"     
'Location where we have to send files
targdir = "E:\CPS-FTP\NID\EXTERNAL\GSBI\IN\"
'targdirNID = "E:\ApplicationFolder\TempHSBC\A201\IN\"
'Location where we have to archive data
ArchWODir = "E:\ApplicationFolder\InternalAr\GSBI\"
flgEmail = "0"
Currentdt = Day(Date) & "-" & Month(Date) & "-" & Year(Date)
Set objFolder = fso.GetFolder(rootdir)
set objLog = fso.OpenTextfile("E:\BankingLog\AxisTrfr\GSBITrfrLog_" & Currentdt & ".txt",8,TRUE)
set objErr = fso.OpenTextfile("E:\Application\AxisTransfer\MailBody\GSBITrfr.txt",2,TRUE)
set objEmail = fso.OpenTextfile("E:\Application\AxisTransfer\Batch\GSBIEmail.bat",2,TRUE)
'set objEmailRecon = fso.OpenTextfile("E:\Application\AxisTransfer\Batch\EmailRecon.bat",2,TRUE)
'set objshl= createobject("wscript.shell")

objLog.WriteLine (Date & Time & " --------Script Started--------"  )
'call RecurseDir(rootdir)

dim tempdir, file, subdir,intResult,trDir,intResult1,intResult2,dir,i,intResult3

	'set TempDir = fso.GetFolder(dir)
	set TempDir = objFolder.Files
	'trDir = trdr
	'Wscript.Echo TempDir.files

	'Copy files from current directory to target
	For Each file in TempDir
	   'Wscript.Echo file.Name
	   'intResult = InStr(UCase(file.Name),"SYSCOM_EMB")
	   'intResult1 = InStr(UCase(file.Name),"SYSCOM_CARDMAILER")
	   'intResult2 = InStr(UCase(file.Name),"SYSCOM_RECON")
 	   'intResult3 = InStr(UCase(file.Name),"STAGE")
	   'Wscript.echo UCase(file.Name)
	   'Wscript.echo ("intResult:" & intResult)
	   'Wscript.echo ("intResult1:" & intResult1)
	  ' Wscript.echo ("intResult2:" & intResult2)
	   'if intResult <> 0 or intResult1 <> 0 or intResult2 <> 0 or intResult3 <> 0 then
	   	'WScript.Echo(file.Name)
	   	'WScript.Echo(file.path)
	   	'WScript.Echo(targdir)
	   	for i=1 to 10000
	   	next
	   	'objLog.WriteLine "FileName: " &  objFile.Name & "->" & strLine
	   			
	   	If FileExists(ArchWODir & "\" & file.Name ) Then
	   		'WScript.Echo "Does Exist"
	   		objLog.WriteLine (Date & Time & " file already exists at Archive Folder:-> " & file.path )
	   	Else
	   		'WScript.Echo "Does not exist"
	   		'objLog.WriteLine (Date & Time & " file Copied at Archive Folder:-> " & file.path )
	   		'file.copy(ArchWODir)
	   	End If
	   	If FileExists(targdir & "\" & file.Name ) Then
	   		objLog.WriteLine (Date & Time & " file already exists at Target Folder:-> " & file.path )
	   	Else
	   		'WScript.Echo "Does not exist"
	   		 file.move(targdir)
			 'file.move(targdirNID)

		  	objLog.WriteLine (Date & Time & " file copied Folder:-> " & file.Name )
		  	'here need to write the code for mails transfer.
		  	objErr.writeline (Date & Time & " File :-> " & file.Name )
		  	flgEmail = "1"
		End If
		'if intResult2 <> 0  then 
		'	objEmailRecon.writeline "C:\BLAT\Blat E:\Application\AxisTransfer\MailBody\Dummy.txt -subject " & Chr(34) & "AXIS Debit Recon Received : " & Chr(34) & " -to sec.nid.bcsc@morpho.com,mph.sec.nid.tpmteam.morpho.com@morpho.com -attach "  & Chr(34) & ArchWODir & file.Name & Chr(34) 
		'	objEmailRecon.Close
		'	objshl.run "E:\Application\AxisTransfer\Batch\EmailRecon.bat"
		'	'WScript.Echo ArchWODir &  file.Name 
 		'	'set shell=nothing  
		'End iF
	   			
	   			
	   'End if
	   'file.copy(targdir)
	Next

	'for each file in TempDir.Files

		'If LCase(file.Name) = sFileName Then WScript.Echo sFileName, "exists."
		
		'intResultPDf = InStr(UCase(file.Name),".PDF")
		
		

	'next
	'For each subfolder of current dir, copy files to target and recurse its subdirs
	'for each subdir in TempDir.subfolders
	'	call RecurseDir(subdir.path)
	'next

if flgEmail = "1" then
	'EMail sending through BLAT.....
	objEmail.writeline "C:\BLAT\Blat E:\Application\AxisTransfer\MailBody\GSBITrfr.txt -subject " & Chr(34) & "GSBI Files Received : " & Chr(34) & " -to sec.nid.bcsc@morpho.com,mph.sec.nid.tpmteam.morpho.com@morpho.com,himanshu.varshney@idemia.com,meetu.singh@idemia.com,vimalchandra.sharma@idemia.com"  
	'objEmail.writeline "pause"
	objEmail.Close
	objshl.run "E:\Application\AxisTransfer\Batch\GSBIEmail.bat"
 	set shell=nothing  
end if 

objLog.WriteLine (Date & Time & " --------Script End--------" )


function RecurseDir(dir)

	
	

end function

Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function


'intResult = InStr(strLine,"Err")
'		if intResult <> 0 then
'			WScript.Echo(strLine)
'			objErr.WriteLine "FileName: " &  objFile.Name & "->" & strLine
'		End if


