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
