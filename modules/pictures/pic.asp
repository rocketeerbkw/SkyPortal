<!-- #include file="config.asp" --><%
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
dim CurPageType
CurPageType = "pictures"
CurPageInfoChk = "1"
'dim sAppRead, sAppWrite, sAppFull
'dim sCatRead, sCatWrite, sCatFull
'dim sSCatRead, sSCatWrite, sSCatFull
Dim iPgType, cat_id, sub_id, intDir, intShow, intLen, ShareCheck, lastid
Dim strPicCopyright, strPoster, strTURL, strURL, comments, sMode
Dim strDesc, strpicTitle, strPostDate, dateSince
Dim intHit, intpicID, sTxt, hp, search
iPgType = 0
sMode = 0
cat_id = 0
sub_id = 0
hp = 0
intPicID = 0
intHit = 0
intDir = 0

function CurPageInfo ()
	strOnlineQueryString = "?" & ChkActUsrUrl(Request.QueryString)
	if len(strOnlineQueryString) = 1 then
	  strOnlineQueryString = ""
	end if
	PageName = txtPics
	PageAction = txtViewing & "<br>" 
	PageLocation = "pic.asp" & strOnlineQueryString
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		sMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("cid") <> "" or  Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cat_id = cLng(Request("cid"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("sid") <> "" or  Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sub_id = cLng(Request("sid"))
	else
		closeAndGo("default.asp")
	end if
end if
%>
<!-- #include file="inc_functions.asp" -->
<!-- #include file="modules/pictures/pic_functions.asp" -->
<!-- #include file="modules/pictures/pic_custom.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<% 
arg1 = txtPics & "|pic.asp" 'this is for the page breadcrumb

':: set default module permissions
setAppPerms "pictures","iName"
'response.Write("intAppID: " & intAppID)
%>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse;" width="100%">
<tr>
<td class="leftPgCol">
<% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
</td>
<td class="mainPgCol">
<% 
intSkin = getSkin(intSubSkin,2)
  select case iPgType
	case 0
	  showall()
	case 1
	  showcat(cat_id)
	case 2
	  showsub()
	case 3
	  shownew()
	case 4
	  showpopular()
	case 5
	  showtoprated()
	case 6
	  lastid = 0
	  if Request.Cookies(strUniqueID & "picid") <> "" then
	  lastid = Request.Cookies(strUniqueID & "picid")
	  end if
	  if lastid <> "" and not isnumeric(lastid) then
		closeAndGo("pic.asp")
	  end if
	  if lastid <> cint(cat_id) then
		executeThis("UPDATE PIC SET HIT=HIT + 1 WHERE PIC_ID=" & cat_id)
	    Response.Cookies(strUniqueID & "picid") = cat_id
		Response.Cookies(strUniqueID & "picid").Expires = dateadd("d",7,now())
	  end If
	  showItem()
	case 7
	  doSearch()
	case 8
	  addPicture()
	case 10
	  closeAndGo("pic.asp?cmd=6&cid=" & sub_id)
	case else
	  showall()
  end select  %>
  <% app_MainColumn_bottom() %>
  <div class="clsSpacer"></div>
</td>
<td class="rightPgCol" width="190" valign="top">
<% intSkin = getSkin(intSubSkin,3)
app_RightColumn() %></td>
</tr>
</table>
<% app_Footer() %>
<!-- #INCLUDE FILE="inc_footer.asp" -->