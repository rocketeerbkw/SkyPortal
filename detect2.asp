<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>SkyPortal component check</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
bDebug = false
if request.QueryString("sky") = "dogg" then
  bDebug = true
end if
dim bFso, bFPOK

function DetectDotNetComponent(DotNetResize)
  Dim DotNetImageComponent, ResizeComUrl, LastPath
	
	DotNetImageComponent = ""
	ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	LastPath = InStrRev(ResizeComUrl,"/")
	if LastPath > 0 then
		ResizeComUrl = left(ResizeComUrl,Lastpath)
	end if
	ResizeComUrl = ResizeComUrl & DotNetResize
	'Response.Write ResizeComUrl & "<br />"
	
	'Check for ASP.NET
	if DotNetCheckComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
	  DotNetImageComponent = "Microsoft.XMLHTTP"
	else
	  if DotNetCheckComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
		DotNetImageComponent = "Msxml2.ServerXMLHTTP"
	  else
		if DotNetCheckComponent("Msxml2.ServerXMLHTTP.4.0", ResizeComUrl) = true then
		  DotNetImageComponent = "Msxml2.ServerXMLHTTP.4.0"
		else
		  DotNetImageComponent = "ASP.NET not found"
		end if
	  end if
	end if
	on error goto 0
  
	DetectDotNetComponent = DotNetImageComponent
end function

function DotNetCheckComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
	'response.write("Checking "&DotNetObj&"<br />")
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
  	'response.write("Object "&DotNetObj&" created<br />")
    objHttp.open "GET", ResizeComUrl, false
		if err.number = 0 then
      objHttp.Send ""
			if (objHttp.status <> 200 ) then
				'Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br />"
				'Response.Write "Returned:<br />" & objHttp.responseText & "<br />"
				'Response.End
			end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
		end if
    Set objHttp = nothing
  End if
  on error goto 0
  DotNetCheckComponent = Detection
end function

function fsoCheck()
     on error resume next
     err.clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   bFSO = true
	   set fso = nothing
	 else 
	   bFSO = false
	 end if
     on error goto 0
	 fsoCheck = bFSO
end function

function getEmailComponents()
Dim arrComponent(10)
Dim arrValue(10)
Dim arrName(10)

' components
arrComponent(0) = "CDO.Message"
arrComponent(1) = "CDONTS.NewMail"
arrComponent(2) = "SMTPsvg.Mailer"
arrComponent(3) = "Persits.MailSender"
arrComponent(4) = "SMTPsvg.Mailer"
arrComponent(5) = "CDONTS.NewMail"
arrComponent(6) = "dkQmail.Qmail"
arrComponent(7) = "Geocel.Mailer"
arrComponent(8) = "iismail.iismail.1"
arrComponent(9) = "Jmail.smtpmail"
arrComponent(10) = "SoftArtisans.SMTPMail"

' component values
arrValue(0) = "cdosys"
arrValue(1) = "cdonts"
arrValue(2) = "aspmail"
arrValue(3) = "aspemail"
arrValue(4) = "aspqmail"
arrValue(5) = "chilicdonts"
arrValue(6) = "dkqmail"
arrValue(7) = "geocel"
arrValue(8) = "iismail"
arrValue(9) = "jmail"
arrValue(10) = "smtp"

' component names
arrName(0) = "CDOSYS (IIS 5/5.1/6)"
arrName(1) = "CDONTS (IIS 3/4/5)"
arrName(2) = "ASPMail"		'yes
arrName(3) = "ASPEMail"	'yes
arrName(4) = "ASPQMail"	'yes
arrName(5) = "Chili!Mail (Chili!Soft ASP)"
arrName(6) = "dkQMail"
arrName(7) = "GeoCel"
arrName(8) = "IISMail"
arrName(9) = "JMail"				
arrName(10) = "SA-Smtp Mail"

Response.Write("<ul>") & vbcrlf
'Response.Write("<option value=""none"" selected="selected"></option>") & vbcrlf
Dim i
for i=0 to UBound(arrComponent)
	if isInstalled(arrComponent(i)) then
	  Response.Write("<li>"  & arrName(i) &"</li>") & vbcrlf
	end if
next
Response.Write("</ul>") & vbcrlf
end function				'

Function isInstalled(obj)
	on error resume next
	installed = False
	Err = 0
	Dim chkObj
	Set chkObj = Server.CreateObject(obj)
	If 0 = Err Then installed = True
	Set chkObj = Nothing
	isInstalled = installed
	Err = 0
	on error goto 0
End Function

function imgCompCheck()  
  gotOne = false
  Response.Write("<ul>") & vbcrlf
    if det <> "" then
	  Response.Write("<li>Asp.Net - " & det & "</li>") & vbcrlf
	  gotOne = true
	else
	  Response.Write("<li>Asp.Net - <font color=""#FF0000""><b>Not available</b></font></li>") & vbcrlf
	end if
	if isInstalled("Persits.Jpeg") then
  	  Set Jpeg = Server.CreateObject("Persits.Jpeg")
	  Response.Write("<li>AspJpeg - v" & Jpeg.Version & "</li>") & vbcrlf
  	  Set Jpeg = nothing
	  gotOne = true
	else
	  Response.Write("<li>AspJpeg - <font color=""#FF0000""><b>Not available</b></font></li>") & vbcrlf
	end if
	if isInstalled("AspImage.Image") then
	  Set Jpeg = Server.CreateObject("AspImage.Image")
	  Response.Write("<li>AspImage - v" & Jpeg.Version & "</li>") & vbcrlf
	  Set Jpeg = nothing
	  gotOne = true
	else
	  Response.Write("<li>AspImage - <font color=""#FF0000""><b>Not available</b></font></li>") & vbcrlf
	end if
    Response.Write("</ul>") & vbcrlf
	if gotOne then
	  response.Write("<b>Image resizing will be available.</b>")
	else
	  Response.Write("<font color=""#FF0000""><b>Image resizing will not be available.</b></font>") & vbcrlf
	end if
end function

%>
<body>
 
<h1>SkyPortal Pre-Installation Check</h1>
<!-- Detecting Components:<br /><br /> -->
<% 
bFso = fsoCheck()
if bDebug then
  det = ""
else
  det = DetectDotNetComponent("includes/scripts/checkfordotnet.aspx") 
end if
%>
	  <hr>
	  <% checkServer() %>
</body>
</html>
<%


function testFolder(fldr)
		dim fs, f
     	set fs = CreateObject("Scripting.FileSystemObject")
		set f = fs.GetFolder(fldr)
		'tmpMSG = tmpMSG & "<li></li>"
		fParent = right(f.ParentFolder,len(f.ParentFolder)-instrrev(f.ParentFolder,"\"))
		fs.CreateFolder fldr & "\test1"
		If fs.FolderExists(fldr & "\test1") = true Then
		  fs.CreateFolder fldr & "\test1\test2"
		  If fs.FolderExists(fldr & "\test\test2") = true Then
			'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder created</li>"
			fs.DeleteFolder fldr & "\test\test2"
			If fs.FolderExists(fldr & "\test\test2") = true Then
				'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder not deleted</li>"
				tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
				tmpMSG = tmpMSG & "Please check the <b>""" & fParent & "/" & f.Name & """</b> folder<br />for <b>""Delete""</b> permissions"
				tmpMSG = tmpMSG & "<br />" & fParent & "/" & f.Name & "/test/test2 folder was not deleted"
				tmpMSG = tmpMSG & "</span></li>"
				boolPerm = false
			else
			  'tmpMSG = tmpMSG & "<li><b>""" & fParent & "/" & f.Name & "/test""</b> - correctly set</li>"
			  'boolPerm = true
			end if
			fs.DeleteFolder fldr & "\test"
			If fs.FolderExists(fldr & "\test") = true Then
				'tmpMSG = tmpMSG & "<li>/files/" & x.Name & "/test folder not deleted</li>"
				tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
				tmpMSG = tmpMSG & "Please check the <b>""" & fParent & "/" & f.Name & """</b> folder for <b>""Delete""</b> permissions"
				tmpMSG = tmpMSG & "<br />" & fParent & "/" & f.Name & "/test folder was not deleted"
				tmpMSG = tmpMSG & "</span></li>"
				boolPerm = false
			else
				'tmpMSG = tmpMSG & "<li><b>""" & fParent & "/" & f.Name & """</b> - correctly set</li>"
				'boolPerm = false
			end if
		  else ':: \test\test2 not created
			'response.Write("test folder not created<br />")
			tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
			tmpMSG = tmpMSG & "Please check <b>""" & fParent & "/" & f.Name & """</b> folder permissions"
			tmpMSG = tmpMSG & "<br />Make sure that the permissions apply to child folders.</span></li>"
			boolPerm = false
		  end if
		else
			'response.Write("test folder not created<br />")
			tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
			tmpMSG = tmpMSG & "Please check <b>""" & fParent & "/" & f.Name & """</b> folder permissions"
			tmpMSG = tmpMSG & "<br />Make sure that the permissions apply to child folders.</span></li>"
			boolPerm = false
		end if
		set fs = nothing
end function

function chkDB(ckDBpath)
	tmpMSG = ""
	boolPerm = true
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(ckDBpath) = true Then
	  set fo = fso.GetFile(ckDBpath)
	  if fo.Name = "sp_db2k3.mdb" then
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "<b>Please rename your Database from the default of sp_db2k3.mdb</b></span></li>"
	  else
		tmpMSG = tmpMSG & "<li><b>Database has been renamed!</b></li>"
		bDBOK = true
	  end if
	  set fo = nothing
	else
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "Database does not exist, Check that<br />"
		tmpMSG = tmpMSG & "your Database path is correct:<br />"
		tmpMSG = tmpMSG & "<b>" & ckDBpath & "</b>"
		tmpMSG = tmpMSG & "</span></li>"
	end if
	set fso = nothing
	chkDB = tmpMSG
end function

function chkPerm(ckFolder)
	tmpMSG = ""
	boolPerm = true
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(ckFolder) = true Then
	  set fo = fso.GetFolder(ckFolder) ':: get the "files" folder
	  testFolder(fo.Path)
	else
	  fso.CreateFolder(ckFolder)
	  If fso.FolderExists(ckFolder) = true Then
	    set fo = fso.GetFolder(ckFolder) ':: get the "files" folder
	    testFolder(fo.Path)
	  else
		tmpMSG = tmpMSG & "<li><span class=""fAlert"">"
		tmpMSG = tmpMSG & "<b>""" & ckFolder & """ does not exist</b>"
		tmpMSG = tmpMSG & "</span></li>"
		boolPerm = false
	  end if
	end if
	if boolPerm = true then
	  tmpMSG = tmpMSG & "<li><b>""/" & fo.Name & """ folder permissions are correctly set</b></li>"
	  bFPOK = true
	end if
	set fo = nothing
	set fso = nothing
	chkPerm = tmpMSG
end function

sub chkFolderPerms()
  response.Write("<ul>")
  response.Write(chkPerm(Server.MapPath("files")))
  response.Write(chkPerm(Server.MapPath("files/config")))
  response.Write(chkPerm(Server.MapPath("files/config/menu")))
end sub

function checkServer()
  'spThemeTitle = "SkyPortal Server Check"
  'spThemeBlock1_open(1)
  Response.Write("<table border=""0"" cellspacing=""3"" cellpadding=""0"">")
  Response.Write("<tr><td><br />")
  'response.Write("<h4>SkyPortal Pre-Installation Check:</h4>")
  response.Write("<ul>")
	  if bFso = false then
	    response.Write("<li><b>FileSystemObject is not available on this server</b></li>")
	  else
	    response.Write("<li><b>FileSystemObject is available on this server</b>")
	    response.Write("<br>Uploads will be available</li>")
		
		if strDBType = "access" then
		  'response.Write(chkDB(strDBPath))
		  bDBOK = true
		else
		  bDBOK = true
		end if
    response.Write("</ul>")
	response.Write("<hr><h4>Checking folder permissions</h4>")
  		chkFolderPerms()
	  end if
  response.Write("</ul>")
  if bFso then
	if bDBOK and bFPOK then
      'session.Contents("chkServer") = "OK"
	  'Response.Write("<input name=""submit"" type=""button"" value=""" & txtContinue & """ onclick=""serverOK();"" />")
	else
  	  response.Write("<b>You must correct the items highlighted above before installation can continue.</b>")
	end if
	response.Write("<hr><h4>The following compatable EMAIL components are available on this server</h4>")
	getEmailComponents()
	response.Write("<hr><h4>The following compatable IMAGE components are available on this server</h4>")
	imgCompCheck()
  else
    'session.Contents("chkServer") = "BAD"
	'session.Contents("setup") = ""
  	response.Write("<h4>SkyPortal can not be used on this server!</h4>")
  end if 
  'response.Write("<hr /><div style=""text-align:left;padding-left:50px;""></div>")
  Response.Write("</td></tr></table>")
  Response.Write("<p>&nbsp;</p>")
  Response.Write("<p>&nbsp;</p>")
  Response.Write("<p>&nbsp;</p>")
  Response.Write("<p>&nbsp;</p>")
  'spThemeBlock1_close(1)
  'closeAndGo("stop")
end function
%>
