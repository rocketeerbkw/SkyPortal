<!--#include file="config.asp" -->
<!-- #include file="Lang/en/downloads_lang.asp" --><%
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
Server.ScriptTimeout = 3600
curpagetype = "downloads"
uploadPg = true
sString = ""
'filename = ""
'uLoad = ""
size=0
%> 
<!--#INCLUDE file="includes/inc_clsUpload.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" --> 
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="modules/downloads/dl_functions.asp" -->
<!-- #include file="modules/downloads/dl_custom.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
  
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = "New Download"
	PageAction = "Submitted<br />" 
	CurPageInfo = PageAction & PageName
end function

function formatSize(s)
  dim fSize
  if s > 1024 then
    fSize = round(s/1024)
	if fSize > 1024 then
      fSize = round(fSize/1024)
	  if fSize > 1024 then
	    fSize = round(fSize/1024)
	    fSize = fSize & " gb"
	  else
	    fSize = fSize & " mb"
	  end if
	else
	  fSize = fSize & " kb"
	end if
  else
    fSize = round(s) & " bytes"
  end if
  formatSize = fSize
end function

  setAppPerms CurPageType,"iName"

if not isObject(objUpload) then
	sString = sString & "<li>Your session has expired.</li>"
	sString = sString & "<li>You will need to refresh the submission page<br />to get the session back.</li>"
else
  
  if trim(objUpload.Fields("subcat").Value) = "" then
		Response.Redirect("dl.asp")
  else
		cat = cLng(objUpload.Fields("subcat").Value)
  end if

  today = strCurDateString
  uLoad = filename
  if trim(uLoad) = "" then
	sString = ""
  end if
  
  bSecCodeMatch = true
  if intSecCode <> 0 then
    fSecCode = ChkString(objUpload.Fields("secCode").Value,"sqlstring")
    if DoSecImage(fSecCode) then
      'Image matched their input 
      bSecCodeMatch = true
    else
      'Image did not match their input
      bSecCodeMatch = false
	  sString = sString & "<li>Your Security Code didn't match.</li>"
    end if
  end if
'response.End()
  if bSecCodeMatch then
	uLoad = filename
	name = ChkString(objUpload.Fields("name").Value,"sqlstring")
	URL = ChkString(objUpload.Fields("URL").Value,"url")
	key = ChkString(objUpload.Fields("key").Value,"sqlstring")
	'filesize = ChkString(objUpload.Fields("filesize").Value,"sqlstring")
	'filesize = formatSize(size)
  if left(URL,4) = "http" then
	filesize = ChkString(objUpload.Fields("size").Value,"sqlstring")
  else
    filesize = formatSize(size)
  end if
    'Response.Write "filesize: " & filesize
	'if len(filesize & "x") > 1 then
	  'filesize = size
	'end if
	sdesc = ChkString(objUpload.Fields("sdes").Value,"sqlstring")
	ldesc = ChkString(objUpload.Fields("Message").Value,"message")
	email = ChkString(objUpload.Fields("mail").Value,"url")
	uploader = ChkString(objUpload.Fields("uploader").Value,"sqlstring")
	'license = ChkString(objUpload.Fields("license").Value,"sqlstring")
	'language = ChkString(objUpload.Fields("language").Value,"sqlstring")
	'platform = ChkString(objUpload.Fields("platform").Value,"sqlstring")
	'publisher = ChkString(objUpload.Fields("publisher").Value,"sqlstring")
	'publisherURL = ChkString(objUpload.Fields("publisherURL").Value,"url")
	
	sT1Sql = ""
	sT2Sql = ""

  if txtDlLabel1 <> "" then
    sTDATA1 = ChkString(objUpload.Fields("TDATA1").Value,"sqlstring")
	sT1Sql = ", TDATA1"
	sT2Sql = ", '" & sTDATA1 & "'"
  end if
  if txtDlLabel2 <> "" then
    sTDATA2 = ChkString(objUpload.Fields("TDATA2").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2"
	sT2Sql = sT2Sql & ", '" & sTDATA2 & "'"
  end if
  if txtDlLabel3 <> "" then
    sTDATA3 = ChkString(objUpload.Fields("TDATA3").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3"
	sT2Sql = sT2Sql & ", '" & sTDATA3 & "'"
  end if
  if txtDlLabel4 <> "" then
    sTDATA4 = ChkString(objUpload.Fields("TDATA4").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4"
	sT2Sql = sT2Sql & ", '" & sTDATA4 & "'"
  end if
  if txtDlLabel5 <> "" then
    sTDATA5 = ChkString(objUpload.Fields("TDATA5").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5"
	sT2Sql = sT2Sql & ", '" & sTDATA5 & "'"
  end if
  if txtDlLabel6 <> "" then
    sTDATA6 = ChkString(objUpload.Fields("TDATA6").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6"
	sT2Sql = sT2Sql & ", '" & sTDATA6 & "'"
  end if
  if txtDlLabel7 <> "" then
    sTDATA7 = ChkString(objUpload.Fields("TDATA7").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7"
	sT2Sql = sT2Sql & ", '" & sTDATA7 & "'"
  end if
  if txtDlLabel8 <> "" then
    sTDATA8 = ChkString(objUpload.Fields("TDATA8").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8"
	sT2Sql = sT2Sql & ", '" & sTDATA8 & "'"
  end if
  if txtDlLabel9 <> "" then
    sTDATA9 = ChkString(objUpload.Fields("TDATA9").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9"
	sT2Sql = sT2Sql & ", '" & sTDATA9 & "'"
  end if
  if txtDlLabel10 <> "" then
    sTDATA10 = ChkString(objUpload.Fields("TDATA10").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10"
	sT2Sql = sT2Sql & ", '" & sTDATA10 & "'"
  end if

  Set objUpload = Nothing

  dim rsCategories
  strSQL = "SELECT CAT_ID, SG_FULL, SG_WRITE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & cat
  set rsA = my_Conn.execute(strSql)
  if rsA.eof then
    set rsA = nothing
    closeAndGo("dl.asp?bad=scatID")
  else
    parent = rsA("CAT_ID")
    s_full = hasAccess(rsA("SG_FULL"))
    s_write = hasAccess(rsA("SG_WRITE"))
  end if
  set rsA = nothing

  if not s_write then
    closeAndGo("error.asp?type=no_access")
  end if

  strSQL = "SELECT CG_FULL FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & parent
  set rsA = my_Conn.execute(strSql)
  if rsA.eof then
    set rsA = nothing
    closeAndGo("dl.asp?bad=catID")
  else
    c_full = hasAccess(rsA("CG_FULL"))
  end if
  set rsA = nothing

  strSQL = "select UP_ACTIVE, UP_ALLOWEDGROUPS, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'download'"
  dim rsUload
  set rsUload = server.CreateObject("adodb.recordset")
  rsUload.Open strSQL, my_Conn
  uActive = rsUload("UP_ACTIVE")
  uAllowed = hasAccess(rsUload("UP_ALLOWEDGROUPS"))
  extAllowed = rsUload("UP_ALLOWEDEXT")
  rsUload.Close
  set rsUload = nothing

  if trim(uLoad) <> "" and bFSO = true and strAllowUploads = 1 and uActive = 1 and uAllowed then
	banner = remotePath  & cat & "/"  & uLoad
	url = banner
  end if

  if len(trim(name)) = 0 then
	sString = sString & "<li>Please enter Program Title.</li>"
  end if

  isOK = false
  if len(trim(url)) <= 8 and trim(uLoad) = "" then 
	sString = sString & "<li>Please enter a Download URL.</li>"
  elseif len(trim(url)) > 8 and trim(uLoad) = "" then 
    tmpExt = split(extAllowed,",")
	for ex = 0 to ubound(tmpExt)
    if right(url,3) = tmpExt(ex) then
	  isOK = true
	end if
	next
	if not isOK then
	  sString = sString & "<li>Please enter a valid Download URL.</li>"
	end if
  else

	strSql="Select URL from DL where URL like '%" & URL & "%'"
	set rsA = my_Conn.execute(strSql)

	if not rsA.eof then
		sString = sString & "<li>This Download already exists in our database.</li>"
	end if
	set rsA = nothing
  end if

  if cat = "--Please select one--" then
	sString = sString & "<li>Please select category that match your program.</li>"
  end if

  if len(trim(sDesc)) = 0 then
	sString = sString & "<li>Please enter a Short Description.</li>"
  end if

  if len(trim(sDesc)) => 255 then
	sString = sString & "<li>" &len(trim(sDesc))&" characters. Your Short Description is too long. 255 characters max</li>"
  end if

  if len(trim(lDesc)) = 0 then
	sString = sString & "<li>Please enter a Long Description.</li>"
  end if

  if len(trim(Email)) = 0 then 
	sString = sString & "<li>You must give an email address.</li>"
  else
    if EmailField(Email) = 0 then 
	  sString = sString & "<li>You Must enter a valid email address.</li>"
    end if
  end if

  end if ':: if bSecCodeMatch then

end if

if sString = "" and bSecCodeMatch then
  ':: set default module permissions
  'setAppPerms "downloads","iName"

session.Contents("uploadType") = ""
strSql = "INSERT INTO DL"
strSql = strSql & "(NAME"
strSql = strSql & ", URL"
strSql = strSql & ", KEYWORD"
strSql = strSql & ", CATEGORY"
strSql = strSql & ", DESCRIPTION"
strSql = strSql & ", CONTENT"
strSql = strSql & ", EMAIL"
strSql = strSql & ", POST_DATE"
strSql = strSql & ", O_POST_DATE"
strSql = strSql & ", ACTIVE"
strSql = strSql & ", BADLINK"
strSql = strSql & ", FILESIZE "
'strSql = strSql & ", PARENT_ID"
'strSql = strSql & ", LICENSE "
'strSql = strSql & ", LANG "
'strSql = strSql & ", PLATFORM "
'strSql = strSql & ", PUBLISHER "
'strSql = strSql & ", PUBLISHER_URL "
strSql = strSql & ", UPLOADER "
strSql = strSql & sT1Sql
strSql = strSql & ")"
strSql = strSql & " VALUES ("
strSql = strSql & "'" & name & "'"
strSql = strSql & ", " & "'" & trim(url) & "'"
strSql = strSql & ", " & "'" & key & "'"
strSql = strSql & ", " & "'" & cat & "'"
strSql = strSql & ", " & "'" & sdesc & "'"
strSql = strSql & ", " & "'" & ldesc & "'"
strSql = strSql & ", " & "'" & email & "'"
strSql = strSql & ", " & "'" & today & "'"
strSql = strSql & ", " & "'" & today & "'"
if bAppFull or s_full or c_full then
strSql = strSql & ", " & "1"
else
strSql = strSql & ", " & "0"
end if
strSql = strSql & ", " & "0"
strSql = strSql & ", " & "'" & filesize & "'"
'strSql = strSql & ", " & "'" & parent & "'"
'strSql = strSql & ", " & "'" & license & "'"
'strSql = strSql & ", " & "'" & language & "'"
'strSql = strSql & ", " & "'" & platform & "'"
'strSql = strSql & ", " & "'" & publisher & "'"
'strSql = strSql & ", " & "'" & publisherurl & "'"
strSql = strSql & ", " & "'" & uploader & "'"
strSql = strSql & sT2Sql
strSql = strSql & ")"

	executeThis(strSQL)
	resetCoreConfig()
	
	if trim(uLoad) <> "" and bFSO = true and strAllowUploads = 1 and uActive = 1 and uAllowed then
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(remotePath) & "\"
		if fso.FolderExists(dirPath & cat) = false then
			fso.CreateFolder(dirPath & cat)
		end if
		if fso.FileExists(dirPath & uLoad) = true then
			fso.MoveFile dirPath & uLoad, dirPath & cat & "\" & uLoad
		end if
	set fso = nothing
	on error goto 0
	end if
	  
	  if bAppFull or s_full or c_full then
	    mod_increaseSubcatCount(cat)
	    if intSubscriptions = 1 and strEmail = 1 then
	      'send subscriptions emails
	      eSubject = strSiteTitle & " - New Download"
		  eMsg = "A new download has been submitted at " & strSiteTitle & vbCrLf
		  eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		  eMsg = eMsg & "You can view the new downloads by visiting " & strHomeUrl & "dl.asp?cmd=3" & vbCrLf
	      sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
		  'response.Write("<br />Email sent<br />" )
		end if
	  end if
	%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
<td class="leftPgCol">
<% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
</td>
<td class="mainPgCol">
<% 
  intSkin = getSkin(intSubSkin,2)
			spThemeBlock1_open(intSkin) %>
			<table cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="middle" width=100%>
			<center>
			<% if hasAccess(1) then%>Your file has been added to our database
			<%else%>Your File has been accepted for review.<br />
					Please wait 1-3 days for your File to be reviewed and added.
			<%end if%>
						<table border="0" cellpadding="4" cellspacing="0" width="60%" align="center">
       <tr>
        <td valign=top>
         <b>Keywords:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= KEYWORD %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b>Email:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= Email %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b>File Size:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= filesize %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b><%= txtDlLabel1 %>:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= STDATA1 %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b><%= txtDlLabel2 %>:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= STDATA2 %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b><%= txtDlLabel3 %>:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= STDATA3 %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b><%= txtDlLabel4 %>:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= STDATA4 %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b><%= txtDlLabel5 %>:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= STDATA5 %>
        </td>
       </tr>
       <tr>
        <td valign=top>
         <b>Uploaded by:</b>&nbsp;
        </td>
        <td valign=top align=left>
         <%= Uploader %>
        </td>
						</table>
					</td>
				</tr></table>
	<center><p><a href="dl.asp">Back to Download Categories </a></p>   
	</center>
			<%
spThemeBlock1_close(intSkin)%>
		</td>
	</tr>
</table>
<meta http-equiv="Refresh" content="5; URL=dl.asp">
<%
else
	'They have made an error, delete their upload, if there is one
	if bFSO = true and strAllowUploads = 1 and uLoad <> "" then
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(downloadDir) & "\" & uLoad
		if fso.FileExists(dirPath) = true then
			fso.DeleteFile dirPath
		end if
		dirPath = server.MapPath("files") & "\" & uLoad
		if fso.FileExists(dirPath) = true then
			fso.DeleteFile dirPath
		end if
	set fso = nothing
	end if %>
<br />
<table border="0" cellpadding="0" cellspacing="0" valign="top" width="100%">
	<tr>
<td class="leftPgCol">
<% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
</td>
<td class="mainPgCol">
<% 
  intSkin = getSkin(intSubSkin,2)
			spThemeBlock1_open(intSkin) %>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td valign=top align=center>
									<p align="center"><span class="fSubTitle">There Was A Problem.</span></p>
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
			<%
spThemeBlock1_close(intSkin)%>
		</td>
	</tr>
</table>
<%end if%>
<!--#INCLUDE FILE="inc_footer.asp" -->
