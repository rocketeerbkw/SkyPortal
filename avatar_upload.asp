<!--#include file="config.asp" --><%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using this program, you are agreeing to the terms of the
'<> SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'******* Customizable Parameters: change them to your own use! ************

AVlogFlag = 0 ' 1 = yes; 0 = no;
AVlogFile = "memberAvatars.txt"
MemberAvDir = "files/members/"
'maximum size (bytes) of the file to be uploaded
sizeLimit = 50000 

 'The file types allowed to be uploaded -without dot- and separated by a comma.
extIsAllowed = Array("jpg","gif")

'************************* end customizable parameters ***********************************************
'*****************************************************************************************************
%>
<!--#include file="inc_functions.asp" --> 
<%
'set myX_Conn = Server.CreateObject("ADODB.Connection")
	'myX_Conn.Errors.Clear
	'myX_Conn.Open strConnString

on error resume next
Err.Clear
Response.Expires = 0
Response.Buffer = TRUE
Response.Clear

'security check
if session.Contents("AVloggedUser") = "" then

Session.Contents("AVmsgText") = "<b>Error!</b><br /><br />You are not allowed to upload files.<br />Please <A href=""register.asp?mode=register"">register</A> or <A href=""default.asp"">login</A>."
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")

else

QsiteUrl = Request.ServerVariables("HTTP_HOST")
siteUrl = strHomeURL
memID = getmemberid(session.contents("AVloggedUser"))

remotePath = MemberAvDir & memID
remotePathMapped = Server.MapPath(remotePath) & "\"

if bFso then

SET FSO = Server.CreateObject("Scripting.FileSystemObject")
If NOT FSO.FolderExists(remotePathMapped) Then
FSO.CreateFolder(remotePathMapped)
End If
Set fso = nothing

byteCount = Request.TotalBytes

RequestBin = Request.BinaryRead(byteCount)
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest RequestBin
contentType = UploadRequest.Item("file1").Item("ContentType")
filepathname = UploadRequest.Item("file1").Item("FileName")
if instr(filepathname,"<") = true or instr(filepathname,">") = true or instr(filepathname,"'") = true or instr(filepathname,"""") = true then
 closeAndGo("default.asp")
end if
filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
value = UploadRequest.Item("file1").Item("Value")

if trim(fileName)="" or LenB(value)=0 then 

Session.Contents("AVmsgText") = "<b>" & txtNoFile & "</b><br /><br />"
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")

else
ext = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"."))
allowed = false
for each extNA in extIsAllowed
if lcase(extNA) = lcase(ext) then
allowed = true
end if
next

if not allowed then 
txt = Date & "- " & txtAction & ": " & txtBadFileType & "(" & ext & ") - " & txtUsrName & ": " & session.contents("AVloggedUser") & " - " & txtFileName & ": " & fileName & " - " & txtUploaded & ": " & txtNo & ""
'logActivity(txt)

Session.Contents("AVmsgText") = "<b>" & txtFileNotAllowed & "</b> - <b>." & ext & "</b>"
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")

else
if LenB(value) > sizeLimit then 
txt = Date & "- " & txtAction & ": " & txtBadFileSize & "(" & LenB(value) & ") - " & txtUsrName & ": " & session.contents("AVloggedUser") & " - " & txtFileName & ": " & fileName & " - " & txtUploaded & ": " & txtNo & ""
'logActivity(txt)

Session.Contents("AVmsgText") = "<b>" & txtFileTooLg & " " & sizeLimit & "</b>."
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")

else
Err.Clear
fileName = memID & "." & ext
Set fso = Server.CreateObject("Scripting.FileSystemObject")

if fso.FileExists(remotePathMapped & filename) = true then
Set MyFile = fso.CreateTextFile(remotePathMapped & filename)
For i = 1 to LenB(value)
MyFile.Write chr(AscB(MidB(value,i,1)))
Next

MyFile.Close
Set fso = nothing

if err.number = 0 then
AVATAR_URL = ""
AVATAR_URL = Trim(siteUrl & remotePath & "/" & fileName)
Session.Contents("AVAvatarUrl") = AVATAR_URL
Session.Contents("AVfileName") = fileName
closeAndGo("cp_main.asp?cmd=2&mode=AvatarEditIt")
else
Session.Contents("AVmsgText") = "<b>Error!</b><br /><br /><b>" & err.description & "</b>."
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")
end if

else
Set MyFile = fso.CreateTextFile(remotePathMapped & filename)
For i = 1 to LenB(value)
MyFile.Write chr(AscB(MidB(value,i,1)))
Next

AVATAR_URL = ""
AVATAR_URL = Trim(siteUrl & remotePath & "/" & fileName)

MyFile.Close
Set fso = nothing

if err.number = 0 then
Session.Contents("AVAvatarUrl") = AVATAR_URL
Session.Contents("AVfileName") = fileName
closeAndGo("cp_main.asp?cmd=2&mode=AvatarEditIt")
else
Session.Contents("AVmsgText") = "<b>Error!</b><br />" & AVATAR_URL & "<br /><br />" & remotePath & "<br /><br />" & remotePathMapped & "<br /><b>" & err.description & "</b>."
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")
end if

end if
end if
end if
end if
end if

end if
if err.number <> 0 then
Session.Contents("AVloggedUser") = ""
Session.Contents("AVmsgText") = "<b>" & txtWasErr & "!</b><br />" & err.description & "<br />" & txtTryAgin & "..."
closeAndGo("cp_main.asp?cmd=3&mode=AvatarError")
end if


Function logActivity(txtToLog)
on error resume next
if logFlag = "1" then
if logFile = "" then
logFile = "memberAvatars.log"
end if
Set fsoLog = Server.CreateObject("Scripting.FileSystemObject")
Set logFile = fsoLog.OpenTextFile(remotePathMapped & logFile, 8, True)
logFile.WriteLine(txtToLog)
logFile.close
set fsoLog = nothing
end if
end function
'thanks to Aruba
Sub BuildUploadRequest(RequestBin)
PosBeg = 1
PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
boundaryPos = InstrB(1,RequestBin,boundary)
Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
Dim UploadControl
Set UploadControl = CreateObject("Scripting.Dictionary")
Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
Pos = InstrB(Pos,RequestBin,getByteString("name="))
PosBeg = Pos+6
PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
PosBound = InstrB(PosEnd,RequestBin,boundary)
If PosFile<>0 AND (PosFile<PosBound) Then
PosBeg = PosFile + 10
PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
UploadControl.Add "FileName", FileName
Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
PosBeg = Pos+14
PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
UploadControl.Add "ContentType",ContentType
PosBeg = PosEnd+4
PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
Else
Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
PosBeg = Pos+4
PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
End If
UploadControl.Add "Value" , Value
UploadRequest.Add name, UploadControl
BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
Loop
End Sub
Function getByteString(StringStr)
For i = 1 to Len(StringStr)
char = Mid(StringStr,i,1)
getByteString = getByteString & chrB(AscB(char))
Next
End Function
Function getString(StringBin)
getString =""
For intCount = 1 to LenB(StringBin)
getString = getString & chr(AscB(MidB(StringBin,intCount,1)))
Next
End Function%>