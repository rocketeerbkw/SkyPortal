<!--#INCLUDE FILE="config.asp" --><%

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
curpagetype = "pictures"
uploadPg = true
sString = ""
dim sizeLimit 
filename = ""
%>
 
<!--#INCLUDE FILE="inc_functions.asp" --> 
<!--#INCLUDE file="includes/inc_clsUpload.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = "New Picture"
	PageAction = "Submitted<br />" 
	CurPageInfo = PageAction & PageName
end function

if trim(objUpload.Fields("cat").Value) = "" then
		Response.Redirect("pic.asp")
	else
		cat = cLng(trim(objUpload.Fields("cat").Value)) 'subcat ID
end if
if sString <> "" then
  'response.Write(sString)
end if
title = ChkString(objUpload.Fields("title").Value,"sqlstring")
desc = ChkString(ChkBadWords(objUpload.Fields("desc").Value),"sqlstring")
key = ChkString(objUpload.Fields("key").Value,"sqlstring")
url = ChkString(objUpload.Fields("url").Value,"url")
turl = ChkString(objUpload.Fields("turl").Value,"url")
poster = strDBNTUserName
copyright = ChkString(objUpload.Fields("copyright").Value,"sqlstring")
privatePic = objUpload.Fields("private").Value
today = strCurDateString
uLoad = filename
  'response.Write(":" & uLoad & ":")
  
  'arrUplds(i,1)

Set objUpload = Nothing

strSQL = "SELECT CAT_ID FROM pic_SUBCATEGORIES WHERE SUBCAT_ID = " & cat
dim rsCategories
set rsCategories = server.CreateObject("adodb.recordset")
rsCategories.Open strSQL, my_Conn
parent = rsCategories("CAT_ID")
rsCategories.Close
set rsCategories = nothing

	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'pictures'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	else
	  intAppID = 0
	end if

strSQL = "select UP_ACTIVE, UP_ALLOWEDGROUPS, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'pictures'"
dim rsUload
set rsUload = server.CreateObject("adodb.recordset")
rsUload.Open strSQL, my_Conn
uActive = rsUload("UP_ACTIVE")
uAllowed = rsUload("UP_ALLOWEDGROUPS")
extAllowed = rsUload("UP_ALLOWEDEXT")
rsUload.Close
set rsUload = nothing

if trim(uLoad) <> "" and bFso and strAllowUploads = 1 and uActive = 1 and hasAccess(uAllowed) then
	banner = remotePath & parent & "/"  & cat & "/" & uLoad
end if
if trim(uLoad) = "" then
	sString = ""
end if

isOK = false
if len(trim(url)) > 8 and trim(uLoad) = "" then 
	'extAllowed = "gif,jpg,jpeg,bmp,png"
    tmpExt = split(extAllowed,",")
	for ex = 0 to ubound(tmpExt)
    if lcase(right(url,3)) = lcase(tmpExt(ex)) then
	  isOK = true
	end if
	next
	if not isOK then
	  sString = sString & "<li>Please enter a valid extention.</li>"
	end if
end if

isOK = false
if len(trim(turl)) > 8 and trim(uLoad) = "" then 
	'extAllowed = "gif,jpg,jpeg,bmp,png"
    tmpExt = split(extAllowed,",")
	for ex = 0 to ubound(tmpExt)
    if lcase(right(turl,3)) = lcase(tmpExt(ex)) then
	  isOK = true
	end if
	next
	if not isOK then
	  sString = sString & "<li>Please enter a valid thumbnail extention.</li>"
	end if
end if

if len(trim(title)) = 0 then
	sString = sString & "<li>Please enter a title.</li>"
else
	Set RS=Server.CreateObject("ADODB.Recordset")

	strSql="Select TITLE from pic where TITLE='" & title & "'"
	RS.Open strSql, my_Conn

	if not rs.eof then
		sString = sString & "<li>This picture already exists in our database.</li>"
	end if
	RS.close
end if

if cat = "--Please select one--" then
	sString = sString & "<li>Please select category that match your picture.</li>"
end if

if len(trim(desc)) > 254 then
	sString = sString & "<li>" &len(trim(desc))&" characters. Your description is too long. 255 characters max.</li>"
end if

if (len(trim(url)) = 7 or trim(url) = "http://" or trim(URL)= "") and trim(uLoad) = "" then 
	sString = sString & "<li>You must enter a valid image.</li>"
end if
	
if len(trim(turl)) < 8 and trim(uLoad) = "" then 
	turl = url
end if
if trim(uLoad) <> "" then
  dnChk = false
  select case strImgComp
    case "aspnet"
  	  det = checkForDotNet("includes/scripts/checkfordotnet.aspx")
  	  if det <> "" then
		dnChk = true
  	  end if
	case "aspjpeg"
	  dnChk = true
  end select
  if bFso then
    mlastPos = InStrRev(banner,".")
    if mlastPos > 0 then
       mCurExt = mid(banner,mlastPos+1,Len(banner)-mlastPos)	
       mCurName = mid(banner,1,mlastPos-1)
  	   if dnChk then
	     turl = mCurName & "_sm." & mCurExt
	     url = mCurName & "_rs." & mCurExt
	     rurl = banner
  	   else
	     turl = mCurName & "." & mCurExt
	     url = mCurName & "." & mCurExt
	     rurl = banner
	   end if
    end if
  end if
end if

if sString = "" then
session.Contents("uploadType") = ""
if privatePic = 1 then
	owner = "|"&getMemberID(STRdbntUserName)&"|"
else 
	owner = "0"
end if

strSql = "INSERT INTO PIC"
strSql = strSql & "(TITLE"
strSql = strSql & ", KEYWORD"
strSql = strSql & ", CATEGORY"
strSql = strSql & ", DESCRIPTION"
strSql = strSql & ", POST_DATE"
strSql = strSql & ", PARENT_ID"
strSql = strSql & ", ACTIVE"
strSql = strSql & ", HIT"
strSql = strSql & ", URL"
strSql = strSql & ", TURL"
strSql = strSql & ", COPYRIGHT"
strSql = strSql & ", OWNER"
strSql = strSql & ", POSTER"
strSql = strSql & ") VALUES ("
strSql = strSql & "'" & replace(title, "''","'", 1, -1, 1) & "'"
strSql = strSql & ", " & "'" & key & "'"
strSql = strSql & ", " & "'" & cat & "'"
strSql = strSql & ", " & "'" & desc & "'"
strSql = strSql & ", " & "'" & today & "'"
strSql = strSql & ", " & "'" & parent & "'"
if hasAccess(1) then
strSql = strSql & ", " & "1"
else
strSql = strSql & ", " & "0"
end if
strSql = strSql & ", " & "0"
strSql = strSql & ", " & "'" & replace(url, "'","", 1, -1, 1) & "'"
strSql = strSql & ", " & "'" & replace(turl, "'","", 1, -1, 1) & "'"
strSql = strSql & ", " & "'" & copyright & "'"
strSql = strSql & ", " & "'" & owner & "'"
strSql = strSql & ", " & "'" & poster & "'"
strSql = strSql & ")"

	executeThis(strSQL)
		
	if trim(uLoad) <> "" and bFso and strAllowUploads = 1 and uActive = 1 and hasAccess(uAllowed) then
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(remotePath) & "\"
		if fso.FolderExists(dirPath & parent) = false then
			fso.CreateFolder(dirPath & parent)
		end if
		if fso.FolderExists(dirPath & parent & "\" & cat) = false then
			fso.CreateFolder(dirPath & parent & "\" & cat)
		end if
		if fso.FileExists(dirPath & uLoad) = true then
			fso.MoveFile dirPath & uLoad, dirPath & parent & "\" & cat & "\" & uLoad
		end if
    	mlastPos = InStrRev(uload,".")
    	if mlastPos > 0 then
       	  mCurExt = mid(uload,mlastPos+1,Len(uload)-mlastPos)	
       	  mCurName = mid(uload,1,mlastPos-1)
		  mResize = mCurName & "_rs." & mCurExt
	   	  mThumb = mCurName & "_sm." & mCurExt
    	end if
		if fso.FileExists(dirPath & mResize) = true then
			fso.MoveFile dirPath & mResize, dirPath & parent & "\" & cat & "\" & mResize
		end if
		if fso.FileExists(dirPath & mThumb) = true then
			fso.MoveFile dirPath & mThumb, dirPath & parent & "\" & cat & "\" & mThumb
		end if
	set fso = nothing
	end if
	  
	  if hasAccess(1) and intSubscriptions = 1 and strEmail = 1 then
	    'send subscriptions emails
	    eSubject = strSiteTitle & " - New Photo"
		eMsg = "A new photo has been submitted at " & strSiteTitle & vbCrLf
		eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		eMsg = eMsg & "You can view the new photos by visiting " & strHomeUrl & "pic.asp?cmd=3" & vbCrLf
	    sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
	  end if
	%> 
	<center>
	&nbsp;<% if hasAccess(1) then%>Your picture has been added to our database<%else%>Your picture has been accepted for review.<br>
	Please wait 1-3 days for your picture to be reviewed and added.<%end if%>
<table border=0 cellpadding=0 cellspacing=0 valign=top align=center width=100%>
	<tr>
		<td valign=top class="mainPgCol">
			<%
			intSkin = getSkin(intSubSkin,2)
			spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")%>
				<tr>
					<td valign="middle" width=100%>
						<table border="0" cellpadding="4" cellspacing="0" width="75%" align="center">
							<tr>
								<td valign=top width=30%>
									<b>Title:</b>&nbsp;
								</td>
								<td valign=top align=left width=70%>
									<%Response.write replace(replace(title, "''","'", 1, -1, 1), "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Description:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% Response.write replace(desc, "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Keywords:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% Response.write replace(key, "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>URL:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% Response.write replace(replace(url, "''","'", 1, -1, 1), "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Thumbnail URL:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% Response.write replace(replace(turl, "''","'", 1, -1, 1), "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Copyright:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% Response.write replace(copyright, "''","'", 1, -1, 1) %>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Private:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<% if privatePic = "1" then Response.write "on" else response.write "off" end if%>
								</td>
							</tr>
							<tr>
								<td valign=top>
									<b>Preview:</b>&nbsp;
								</td>
								<td valign=top align=left>
									<img src="<%= turl %>">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%Response.Write("</table>")
spThemeBlock1_close(intSkin)%>
		</td>
	</tr>
</table>
	<p><a href="pic.asp">Back to Picture Categories </a></p>   
	</center>
<meta http-equiv="Refresh" content="10; URL=pic.asp">
<%else
	'They have made an error, delete their upload, if there is one
	if bFso and strAllowUploads = 1 then
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(galleryDir) & "\" & uLoad
		if fso.FileExists(dirPath) = true then
			fso.DeleteFile dirPath
		end if
	set fso = nothing
	end if %>
<br>
<table border="0" cellpadding="0" cellspacing="0" valign="top" align="center" width="100%">
	<tr>
		<td valign=top class="mainPgCol">
			<%
			intSkin = getSkin(intSubSkin,2)
			spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")%>
				<tr>
					<td valign="middle" width=100%>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td valign=top align=center>
									<p align="center"><span class="fTitle">There Was A Problem.</span></p>
									<table align="center" border="0">
									  <tr>
									    <td>
										<ul>
										<% =sString %>
										</ul>
									    </td>
									  </tr>
									</table>
									<p align="center"><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			<%Response.Write("</table>")
spThemeBlock1_close(intSkin)%>
		</td>
	</tr>
</table>
<%end if%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
function checkForDotNet(DotNetFile)
  Dim DotNetComp, ResizeComUrl, LastPath
	DotNetComp = ""
	ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	LastPath = InStrRev(ResizeComUrl,"/")
	if LastPath > 0 then
		ResizeComUrl = left(ResizeComUrl,Lastpath)
	end if
	ResizeComUrl = ResizeComUrl & DotNetFile
	'Response.Write ResizeComUrl & "<br>"
	
	'Check for ASP.NET 1
	if chkDotNetComponent("Msxml2.ServerXMLHTTP.4.0", ResizeComUrl) = true then
		'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP.4.0<br>"
		DotNetComp = "DOTNET1"
	else
		if chkDotNetComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
			'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP<br>"
			DotNetComp = "DOTNET2"
		else
			if chkDotNetComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
				'Response.Write "FOUND: ASP.NET Microsoft.XMLHTTP<br>"
				DotNetComp = "DOTNET3"
			else
				'Response.Write "NOT FOUND: ASP.NET Server Component<br>"
			end if
		end if
	end if
	on error goto 0  
	checkForDotNet = DotNetComp
end function

function chkDotNetComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
	'response.write("Checking "&DotNetObj&"<br>")
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
  	'response.write("Object "&DotNetObj&" created<br>")
    objHttp.open "GET", ResizeComUrl, false
		if err.number = 0 then
      objHttp.Send ""
			if (objHttp.status <> 200 ) then
				'Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br>"
				'Response.Write "Returned:<br>" & objHttp.responseText & "<br>"
				'Response.End
			end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
		end if
    Set objHttp = nothing
  End if
  on error goto 0
 	'response.write("Detection is "&Detection&"<br>")
  chkDotNetComponent = Detection
end function
%>
