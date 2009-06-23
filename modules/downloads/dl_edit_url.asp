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
bDbug = true
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
  
CurPageInfoChk = ""
function CurPageInfo()
	CurPageInfo = "Editing Download"
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
  if bDbug then
    Response.Write "<br>objUpload is an object"
  end if
  if trim(objUpload.Fields("subcat").Value) = "" then
	closeAndGo("dl.asp")
  end if

  uLoad = filename
  if trim(uLoad) = "" then
	sString = ""
  end if
  itemID = clng(objUpload.Fields("itemID").Value)
  subcat = clng(objUpload.Fields("subcat").Value)
  orig_subcat = clng(objUpload.Fields("orig_subcat").Value)
  iname = ChkString(objUpload.Fields("name").Value,"sqlstring")
  orig_URL = trim(ChkString(objUpload.Fields("URL").Value,"url"))
  'filesize = formatSize(size)
  sdesc = ChkString(objUpload.Fields("sdes").Value,"sqlstring")
  ldesc = ChkString(objUpload.Fields("Message").Value,"message")
  email = ChkString(objUpload.Fields("email").Value,"sqlstring")
  key = ChkString(objUpload.Fields("key").Value,"sqlstring")
  
  if left(orig_URL,4) = "http" then
	filesize = ChkString(objUpload.Fields("size").Value,"sqlstring")
  else
    filesize = formatSize(size)
  end if
  
  sT1Sql = ""
  sT2Sql = ""

  if txtDlLabel1 <> "" then
    sTDATA1 = ChkString(objUpload.Fields("TDATA1").Value,"sqlstring")
	sT1Sql = ", TDATA1 = '" & sTDATA1 & "'"
  end if
  if txtDlLabel2 <> "" then
    sTDATA2 = ChkString(objUpload.Fields("TDATA2").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2 = '" & sTDATA2 & "'"
  end if
  if txtDlLabel3 <> "" then
    sTDATA3 = ChkString(objUpload.Fields("TDATA3").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3 = '" & sTDATA3 & "'"
  end if
  if txtDlLabel4 <> "" then
    sTDATA4 = ChkString(objUpload.Fields("TDATA4").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4 = '" & sTDATA4 & "'"
  end if
  if txtDlLabel5 <> "" then
    sTDATA5 = ChkString(objUpload.Fields("TDATA5").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5 = '" & sTDATA5 & "'"
  end if
  if txtDlLabel6 <> "" then
    sTDATA6 = ChkString(objUpload.Fields("TDATA6").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6 = '" & sTDATA6 & "'"
  end if
  if txtDlLabel7 <> "" then
    sTDATA7 = ChkString(objUpload.Fields("TDATA7").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7 = '" & sTDATA7 & "'"
  end if
  if txtDlLabel8 <> "" then
    sTDATA8 = ChkString(objUpload.Fields("TDATA8").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8 = '" & sTDATA8 & "'"
  end if
  if txtDlLabel9 <> "" then
    sTDATA9 = ChkString(objUpload.Fields("TDATA9").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9 = '" & sTDATA9 & "'"
  end if
  if txtDlLabel10 <> "" then
    sTDATA10 = ChkString(objUpload.Fields("TDATA10").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10 = '" & sTDATA10 & "'"
  end if
  
  unmarkbad = objUpload.Fields("unmarkbad").Value
  marknew = objUpload.Fields("marknew").Value
  markupd = objUpload.Fields("markupd").Value
  approve = objUpload.Fields("approve").Value
  today = strCurDateString

  Set objUpload = Nothing

  dim rsA
  strSQL = "SELECT CAT_ID, SG_FULL, SG_WRITE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & subcat
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
    closeAndGo("error.asp?type=noaccess")
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
  set rsB = server.CreateObject("adodb.recordset")
  rsB.Open strSQL, my_Conn
  uActive = rsB("UP_ACTIVE")
  uAllowed = hasAccess(rsB("UP_ALLOWEDGROUPS"))
  extAllowed = rsB("UP_ALLOWEDEXT")
  rsB.Close
  set rsB = nothing

  if trim(uLoad) <> "" and bFSO = true and strAllowUploads = 1 and uActive = 1 and uAllowed then
    if bDbug then
      Response.Write "<br>Is upload"
    end if
	banner = remotePath  & orig_subcat & "/"  & uLoad
	url = banner
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(remotePath) & "\"
		if fso.FolderExists(dirPath & orig_subcat) = false then
			fso.CreateFolder(dirPath & orig_subcat)
		end if
		if fso.FileExists(dirPath & uLoad) = true then
			fso.MoveFile dirPath & uLoad, dirPath & orig_subcat & "\" & uLoad
		end if
		if left(orig_URL,16) = "files/downloads/" then
		  if fso.FileExists(server.MapPath(orig_URL)) = true then
			deleteFile(server.MapPath(orig_URL))
		  end if
		end if
	set fso = nothing
	on error goto 0
  else
    url = orig_URL
  end if
  
  if approve = "1" then
    if bDbug then
      Response.Write "<br>mod_increaseSubcatCount"
    end if
	mod_increaseSubcatCount(subcat)
  end if
  
  ':: check for subcat change
  if orig_subcat <> subcat then
    if bDbug then
      Response.Write "<br>orig_subcat <> subcat"
    end if
    if approve <> "1" then
	  mod_increaseSubcatCount(subcat)
	  mod_decreaseSubcatCount(orig_subcat)
    end if
    if left(URL,16) = "files/downloads/" then
	  if instr(URL,"/" & orig_subcat & "/") > 0 then
    	if bDbug then
      	  Response.Write "<br>move file"
    	end if
	    ':: edit URL
  	    tUrl = replace(URL,"/" & orig_subcat & "/","/" & subcat & "/")
	    ':: move file to new subcat folder
		call moveFile(server.MapPath(URL),server.MapPath(tUrl))
	    URL = tUrl
	  end if
	end if
  end if

  if sString = "" then
    if bDbug then
      Response.Write "<br>Update DB"
    end if
  
    sSql = "UPDATE DL SET "
    sSql = sSql & "NAME='" & iname & "'"
    sSql = sSql & ",CATEGORY =" & subcat & ""
    sSql = sSql & ",EMAIL ='" & email & "'"
    sSql = sSql & ",DESCRIPTION ='" & sdesc & "'"
    sSql = sSql & ",CONTENT ='" & ldesc & "'"
    sSql = sSql & ",URL ='" & URL & "'"
    sSql = sSql & ",FILESIZE ='" & filesize & "'"
    sSql = sSql & ",KEYWORD ='" & key & " '"
    sSql = sSql & sT1Sql
    if marknew = "1" then
      sSql = sSql & ",UPDATED ='0'"
      sSql = sSql & ",O_POST_DATE ='" & today & "'"
      sSql = sSql & ",POST_DATE ='" & today & "'"
	elseif markupd = "1" then
      sSql = sSql & ",UPDATED ='" & today & "'"
      sSql = sSql & ",POST_DATE ='" & today & "'"
    end if
    if approve = "1" then
      sSql = sSql & ",ACTIVE = 1"
    end if
    if unmarkbad = "1" then
      sSql = sSql & ",BADLINK = 0"
    end if
    sSql = sSql & " WHERE DL_ID =" & itemID
    'response.Write(email)
    'response.End()
    executeThis(sSql)
    'Response.Write("<br>" & sSql)
    resetCoreConfig()
    'closeAndGo("stop")
    session.Contents("uploadType") = ""
	'if not bDbug then
      Call setSession("sMsg","Item successfully updated")
      'closeAndGo("dl.asp?cmd=6&cid=" & itemID)
      closeAndGo("dl.asp?cmd=23&item=" & itemID)
	'end if
  end if ':: if sString = ""
end if ':: if isObject(objUpload)

if sString <> "" then
    if bDbug then
      Response.Write "<br>Form error"
    end if
	'They have made an error, delete their upload, if there is one
	if bFSO = true and strAllowUploads = 1 and uLoad <> "" then
    if bDbug then
      Response.Write "<br>Delete upload"
    end if
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
	on error goto 0
	end if %>
<table border="0" cellpadding="0" cellspacing="0" valign="top" width="100%">
	<tr><td class="leftPgCol">
<% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
</td>
<td class="mainPgCol">
<% 
  intSkin = getSkin(intSubSkin,2)
  spThemeBlock1_open(intSkin) %>
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td valign="top" align="center">
		<p align="center"><span class="fSubTitle">There Was A Problem.</span></p>
		<table align="center" border="0"><tr><td>
		  <ul><% =sString %></ul>
		</td></tr></table>
		<p align="center"><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
	</td></tr>
  </table>
  <%
  spThemeBlock1_close(intSkin)%>
  </td></tr>
</table>
<%end if%>
<!--#INCLUDE FILE="inc_footer.asp" -->
