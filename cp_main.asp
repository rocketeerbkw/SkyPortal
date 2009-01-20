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
CurPageType = "core"
%>
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="includes/inc_avatar_functions.asp" -->
<!--#include file="includes/inc_profile_functions.asp" -->
<!--#include file="includes/inc_subscriptions.asp" -->
<!--#include file="includes/inc_bookmarks.asp" -->
<%
'closeandgo("default.asp")
PageTitle = txtCtrlPnl

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("stop")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	end if
end if

if iPgType = 6 and intSubscriptions = 0 then
  closeAndGo("cp_main.asp")
end if
if iPgType = 7 and intBookmarks = 0 then
  closeAndGo("cp_main.asp")
end if

if Request.QueryString("mode") = "goEdit" or Request.QueryString("mode") = "goModify" then
  hasEditor = true
  strEditorElements = "Sig"
end if

CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = txtCtrlPnl
	PageAction = txtViewing & "<br />" 
	CurPageInfo = PageAction & PageName
end function
%>
<!--#include file="inc_top.asp" -->
<%
tmpVMsg = ""
	
'breadcrumb here
  arg1 = txtCtrlPnl & "|cp_main.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""

if Request.QueryString("mode") = "Moderator" and hasAccess(1) then
  select case Request.QueryString("action")
  case "del"
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LEVEL = 1 "
	strSql = strSql & " WHERE M_LEVEL = 2 AND MEMBER_ID = " & cLng(Request.QueryString("id"))
	call executeThis(strsql)
	
	strSql = "DELETE FROM " & strTablePrefix & "MODERATOR"
	strSql = strSql & " WHERE MEMBER_ID=" & cLng(Request.QueryString("id"))
    executeThis(strSql)
	
	strTxt = txtRevoked
  case "add"
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LEVEL = 2 "
	strSql = strSql & " WHERE M_LEVEL = 1 AND MEMBER_ID = " & cLng(Request.QueryString("id"))
	call executeThis(strsql)
	strTxt = txtAssign
  end select
  tmpVMsg = tmpVMsg & txtModStat & "&nbsp;" & strTxt
  tmpVMsg = tmpVMsg & "<meta http-equiv=""Refresh"" content=""2; url=members.asp"">"
end if

':::::::: email change routine :::::::::::::::::::::
if Request.QueryString("verkey") <> "" and len(Request.QueryString("verkey")) = 10 then
	verkey = chkString(Request.QueryString("verkey"),"sqlstring")

	strSql = "SELECT *"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_KEY = '" & verkey & "'"

	set rsKey = my_Conn.Execute(strSql)
	
	if rsKey.EOF or rsKey.BOF or strComp(verkey, rsKey("M_KEY")) <> 0 then
		'Error message to user
		tmpVMsg = tmpVMsg & ""
		tmpVMsg = tmpVMsg & "<p>&nbsp;</p><p align=""center"" class=""fTitle"">" & txtThereIsProb & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center"" class=""fSubTitle"">" & txtVerKeyNoMatch & "<br />"
		tmpVMsg = tmpVMsg & txtTryEmlChgAgain & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center""><a href=""cp_main.asp"">" & txtCtrlPnl & "</a><br />"
		tmpVMsg = tmpVMsg & "<a href=""default.asp"">" & txtHome & "</a></p><p>&nbsp;</p>"

	elseif rsKey("M_EMAIL") = rsKey("M_NEWEMAIL") then
		tmpVMsg = tmpVMsg & "<meta http-equiv=""Refresh"" content=""3; URL=cp_main.asp"">"
		tmpVMsg = tmpVMsg & "<p>&nbsp;</p><p align=""center"" class=""fTitle"">" & txtEmlAlrVer & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center"" class=""fTitle"">" & txtEmlAlrChgd & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center""><a href=""cp_main.asp"">" & txtCtrlPnl & "</a><br />"
		tmpVMsg = tmpVMsg & "<a href=""default.asp"">" & txtHome & "</a></p><p>&nbsp;</p>"
	else
		userID = rsKey("MEMBER_ID")

		'Update the user email
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_EMAIL = '" & chkString(rsKey("M_NEWEMAIL"),"SQLString") & "'"
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & userID

		call executeThis(strsql)
		Session(strUniqueID & "userID") = ""
		
		tmpVMsg = tmpVMsg & "<meta http-equiv=""Refresh"" content=""3; URL=cp_main.asp"">"
		tmpVMsg = tmpVMsg & "<p>&nbsp;</p><p align=""center"" class=""fTitle"">" & txtEmlUpdated & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center"" class=""fTitle"">" & txtEmlUpdInDB & "</p>"
		tmpVMsg = tmpVMsg & "<p align=""center""><a href=""cp_main.asp"">" & txtCtrlPnl & "</a><br />"
		tmpVMsg = tmpVMsg & "<a href=""default.asp"">" & txtHome & "</a></p><p>&nbsp;</p>"
    end if
	rsKey.close
   	set rsKey = nothing
end if

If not hasAccess("2") Then ' Not Logged in %>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
<td width="200" class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)
Menu_fp() 
affiliateBanners()
%>
</td>
<td class="mainPgCol" valign="top">
<%
intSkin = getSkin(intSubSkin,2)
  if tmpVMsg <> "" then
    call showMsgBlock(1,tmpVMsg)
  else
	spThemeBlock1_open(intSkin) %>
	<table border="0" cellpadding="0" cellspacing="0" width="60%" align="center">
	<tr align="center"><td><p>&nbsp;</p><p align="center"><span class="fSubTitle"><%= txtLgnToVwPg %></span>
	<br /><br /><%= txtNoRegis %>&nbsp;<a href="policy.asp"><u><%= txtRegNow %></u></a>.</p>
	<p>&nbsp;</p>
	</td></tr></table>
<%  spThemeBlock1_close(intSkin)
  end if
 %>
	</td></tr>
</table>
<%
Else  ':: they ARE logged in
	id = strUserMemberID
	flag_HasForums = false
	flag_showmytopics = 0
	flag_showrecenttopics = 0
	flag_maxtopics = 5
	flag_showpm = 1
	flag_showstatus = 1
	edit_flag = 0
	flag_PMstatus = 0
	flag_pm_layout = 0
	flag_isFirstTime = 0
	tmpResult = ""

	strSQL = "SELECT * FROM " & strTablePrefix & "CP_CONFIG WHERE MEMBER_ID=" & strUserMemberID
	Set objRS = my_Conn.Execute(strSQL)
	if objRS.EOF then
		strSQL = "INSERT INTO " & strTablePrefix & "CP_CONFIG (Member_ID) VALUES ('" & strUserMemberID & "')"
		executeThis(strSQL)
		flag_isFirstTime = 1
	end if
	set objRS = nothing

	strSQL = "SELECT * FROM " & strTablePrefix & "CP_CONFIG WHERE MEMBER_ID=" & strUserMemberID
	Set objRS = my_Conn.Execute(strSQL)
	if not objRS.EOF then
	  flag_showpm = objRS("SHOW_PM")
	  flag_showstatus = objRS("SHOW_STATUS")
	  flag_pm_layout = objRS("PM_OUTBOX")
  	  if chkApp("forums","USERS") then
		flag_showmytopics = objRS("SHOW_MY_TOPICS")
		flag_showrecenttopics = objRS("SHOW_RECENT_TOPICS")
		flag_maxtopics = objRS("MAX_MY_TOPICS")
		flag_HasForums = true
	  end if
	end if
	
	if flag_isFirstTime = 1 then
	  iPgType = 20
	  tmpVMsg = tmpVMsg & "<span style=""text-align:left;"">"
	  tmpVMsg = tmpVMsg & "<p align=""center"" class=""fSubTitle"">Welcome to your Control Panel</p>"
	  tmpVMsg = tmpVMsg & "<p>Here you can set your personal preferences<br />"
	  tmpVMsg = tmpVMsg & "You can always come back to your Control Panel<br />and change these later by "
	  tmpVMsg = tmpVMsg & "clicking on the<br />'Personal Settings' link in the menu at the left</p>"
	  tmpVMsg = tmpVMsg & "</span>"
	end if

select case iPgType
  case 1
    arg2 = txtEditAvatar & "|cp_main.asp?cmd=1"
  case 2
    arg2 = txtEditAvatar & "|cp_main.asp?cmd=1"
	editAvatar()
  case 3
    arg2 = txtEditAvatar & "|cp_main.asp?cmd=1"
  	arg3 = "Avatar Error|javascript:;"
  case 4
   if flag_HasForums then
	arg2 = txtMyRecTop & "|javascript:;"
   end if
  case 5 
	arg2 = txtPersSet & "|javascript:;"
    updateDateTime()
  case 6
  	arg2 = txtSubscMgr & "|cp_main.asp?cmd=6"
    modifySubscriptions()
  case 7
  	arg2 = txtBkmkMgr & "|cp_main.asp?cmd=7"
	modifyBookmarks()
  case 8 'view profile
    if IsNumeric(Request.QueryString("member")) = True then
  	  arg2 = txtViewProf & "|cp_main.asp?cmd=8&member=" & cLng(request.QueryString("member"))
	else
	  closeAndGo("default.asp")
	end if
  case 9 'edit profile
  	arg2 = txtEditProf & "|cp_main.asp?cmd=9"
  case 10 'edit member profile
  	arg2 = txtEditMProf & "|javascript:;"
  case 11
  case 12
  case 20
  	arg2 = txtPersSet & "|javascript:;"
  case else
end select
	
	'Get member info
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_GOLD, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_REP, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_RTOTAL, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PAGE_VIEWS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL "	
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strUserMemberID
	set rsMem = my_Conn.Execute (strSql)
	flag_PMstatus = rsMem("M_PMSTATUS")
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
<td width="200" class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)

showUserMenu() 
theme_changer()
%>
<script type="text/javascript">
function OpenSPreview(){
	var curCookie = "strSignaturePreview=" + escape(document.Form1.Sig.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_portal.asp?cmd=6', 'preview_page', 'scrollbars=yes,width=450,height=250')	
}
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
</script>
</td>
<td class="mainPgCol" valign="top">
<%
intSkin = getSkin(intSubSkin,2)
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
 select case iPgType
  case 1, 2
    if iPgType = 2 then
	  call showMsgBlock(1,tmpResult)
	end if
    showAVedit()
  case 3
    showAvatarError()
    showAVedit()
  case 4
   if flag_HasForums = true then
	showMyRecentTopics()
   end if
  case 5
    showPreferences()
  case 6
    if tmpResult <> "" then
	  call showMsgBlock(1,tmpResult)
	end if
    showMySubscriptions()
  case 7
    if tmpResult <> "" then
	  call showMsgBlock(1,tmpResult)
	end if
    showMyBookmarks()
  case 8
    if IsNumeric(Request.QueryString("member")) = True then
      displayProfile(cLng(request.QueryString("member")))
	else
	  closeAndGo("default.asp")
	end if
  case 9
    editProfile()
  case 10
  	if hasAccess(1) then
      spThemeTitle = txtEditMProf
      editMemberProfile()
	else
	  tmpResult = ""
	  tmpResult = tmpResult & "<p align=""center""><span class=""fAlert""><b>" & txtERROR & "</b></span></p>"
	  tmpResult = tmpResult & txtNoPermViewPg
	  'tmpResult = tmpResult & "<p>&nbsp;</p>"
	  call showMsgBlock(1,tmpResult)
	end if
  case 11
  case 12
  case 20
    call showMsgBlock(1,tmpVMsg)
	showPreferences()
  case else
	if tmpVMsg <> "" then
    	':: display any messages
    	call showMsgBlock(1,tmpVMsg)
	else
      spThemeTitle = txtMyProfile
      displayProfile(strUserMemberID)
	end if
end select

set rsMem=nothing
End if 
%>

</td></tr></table>
<!--#include file="inc_footer.asp" -->

<%

sub showUserMenu()
spThemeTitle = txtUsrOpts
spThemeBlock1_open(intSkin) %>
<table width="100%">
<tr><td valign="top" class="fSubTitle">
<%
sSQL = "select " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.M_DONATE from " & strMemberTablePrefix & "MEMBERS where " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strdbntusername & "'"
set rsP = my_Conn.execute(sSQL)
	if varBrowser = "ie" then
		if trim(rsP("M_GLOW")) <> "" or rsP("M_DONATE") > 0 then %> <a href="javascript:;"><img src="images/icons/icon_color.gif" onclick="openWindow('pop_glow.asp?cmd=2&id=<% =rsP("MEMBER_ID") %>')" alt="<%= txtEditGlo %>" title="<%= txtEditGlo %>" border="0"></a>&nbsp;&nbsp;&nbsp;<b><%= displayName(ChkString(strdbntusername,"display"),rsP("M_GLOW")) %></b>
  <% 	else %>
				&nbsp;<b><%= strdbntusername %></b>
	  <% 
		end if
	Else %>
		&nbsp;<b><%= strdbntusername %></b><% 
	End If  %>
	
<%	if rsP("M_DONATE") > 0 then
        Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & getDonor_Level(rsP("MEMBER_ID")) & vbcrlf
	end if
set rsP = nothing

response.Write("<hr />")
cp_userMenu()
response.Write("<hr />")
 
showMyStats()
showMemberGroups(strUserMemberID)
showAvatar() %>
</td></tr></table>
<%
spThemeBlock1_close(intSkin)
end sub

sub showMyRecentTopics()
  if flag_showmytopics > 0 then
    spThemeTitle = txtTopUStart
    spThemeBlock1_open(intSkin)
	' - Get all active topics from last visit
	if flag_maxtopics = 0 then
	  strSql = "SELECT "
	else
	  strSql = "SELECT TOP " & flag_maxtopics & " "
	end if

strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS, " 
strSql = strSql & strTablePrefix & "TOPICS.T_VIEW_COUNT, " & strTablePrefix & "TOPICS.FORUM_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.CAT_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_MAIL, " 
strSql = strSql & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_REPLIES, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, "  
strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "
strSql = strSql & "FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "FORUM, "
strSql = strSql & strTablePrefix & "TOPICS, " & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & "WHERE " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS_1.MEMBER_ID  "
strSql = strSql & "AND " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
strSql = strSql & "AND " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "TOPICS.CAT_ID "
strSql = strSql & "AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID=" & id
strSql = strSql & "AND " & strTablePrefix & "TOPICS.T_AUTHOR=" & id
strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

set rs = my_Conn.Execute (strSql)
%>
<table class="tPlain" width="100%">
      <tr>
        <td colspan="5" class="tTitle" align="center"><b><%= txtTopUStart %></b></td>
      </tr>
      <tr>
        <td align="center" width="10%" class="tSubTitle" valign="center">&nbsp;</td>
        <td align="center" class="tSubTitle"><b><%= txtTopic %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtReplies %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtRead %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtLstPost %></b></td>
      </tr>
<%If rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="5" class="fSubTitle"><b><%= txtNoTopFnd %></b></td>
      </tr>
<%else
	currForum = 0 
	fDisplayCount = 0 

	do until rs.EOF

	' - Find out if the Category is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
	strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & rs("CAT_ID")

	set rsCStatus = my_Conn.Execute (StrSql)

	' - Find out if the Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS " 
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & rs("FORUM_ID")

	set rsFStatus = my_Conn.Execute (StrSql)
			fDisplayCount = fDisplayCount + 1
			if currForum <> rs("FORUM_ID") then %>
				<tr>
					<td height="20" colspan="5" class="tAltSubTitle" valign="center"><a href="<% Response.Write("forum.asp?FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID") & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")) %>"><b><% =ChkString(rs("F_SUBJECT"),"display") %></b></a></td>

				</tr>
<%			end if %>
			<tr>
			<td align="center" class="fNorm" valign="center">
<%			if rsCStatus("CAT_STATUS") <> 0 and rsFstatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then
				if lcase(strHotTopic) = "1" then
					if rs("T_REPLIES") >= intHotTopicNum Then %>
						<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_hot.gif" height="15" width="15" border="0" hspace="0" title="<%= txtHotTop %>" alt="<%= txtHotTop %>"></a>
<%					else%>
						<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new.gif" title="<%= txtNewTop %>" alt="<%= txtNewTop %>" border="0" WIDTH="15" HEIGHT="15"></a>
<%					end if
				end if
			else %>
			<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<%=rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<%=ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_locked.gif" 
<% 			if rsCStatus("CAT_STATUS") = 0 then 
				Response.Write (" title=""" & txtCatLok & """ alt=""" & txtCatLok & """ ")
			elseif rsFStatus("F_STATUS") = 0 then 
				Response.Write (" title=""" & txtFrmLok & """ alt=""" & txtFrmLok & """ ")
			else
				Response.Write (" title=""" & txtTopLok & """ alt=""" & txtTopLok & """ ")
			end if %> border="0" WIDTH="15" HEIGHT="15"></a>
<%			end if %>
			</td>
			<td valign="center" class="fNorm"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a>&nbsp;</td>
			<td valign="center" align="center" class="fNorm"><% =rs("T_REPLIES") %></td>
			<td valign="center" align="center" class="fNorm"><% =rs("T_VIEW_COUNT") %></td>
			<%
			if IsNull(rs("T_LAST_POST_AUTHOR")) then
				strLastAuthor = ""
			else
				strLastAuthor = "<br />" & txtBy & ": " 
				strLastAuthor = strLastAuthor & "<a href=""cp_main.asp?cmd=8&member="& rs("T_LAST_POST_AUTHOR") & """><span class=""fSmall"">"
				strLastAuthor = strLastAuthor & rs("LAST_POST_AUTHOR_NAME") & "</span></a>"
			end if
			%>
			<td valign="center" align="center" nowrap><span class="fSmall"><b><%= ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></span></td>
			</tr>	
<%	
		currForum = rs("FORUM_ID") %>
<%		rs.MoveNext 
	loop 
	if fDisplayCount = 0 then %>
		  <tr><td colspan="6" class="fSubTitle"><b><%= txtNoTopFnd %></b></td></tr>
<%
	end if 
 end if
Response.Write("<tr><td colspan=""5"">&nbsp;</td></tr>")
myRecentTopics()
Response.Write("<tr><td colspan=""5"">&nbsp;</td></tr></table>")
spThemeBlock1_close(intSkin)

End If
end sub

':::::::::::::::::::::::::::::: My Recent Topics :::::::::::::::::::::::::::::::::::
sub myRecentTopics()
if flag_showrecenttopics > 0 then 
  if strRecentTopics = "1" then
	strStartDate = DateToStr(dateadd("d", -30, now()))
	' - Find all records for the member
	strsql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, " 
	strsql = strsql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS,  " & strTablePrefix & "TOPICS.T_LAST_POST, " & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_VIEW_COUNT "
	strsql = strsql & " FROM ((" & strTablePrefix & "FORUM LEFT JOIN " & strTablePrefix & "TOPICS "
	strsql = strsql & " ON " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY "
	strsql = strsql & " ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID) "
	strsql = strsql & " WHERE (T_DATE > '" & strStartDate & "') "
	strsql = strsql & " AND (" & strTablePrefix & "TOPICS.T_AUTHOR = " & id & " "
	strsql = strsql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & id & ")"
	strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC, " & strTablePrefix & "TOPICS.TOPIC_ID DESC"
	
	
	set rs2 = my_Conn.Execute(strsql) %>
 	<tr><td colspan="5" class="tSubTitle" align="center"><b><%= txtURecTopics %></b></td></tr>
      <tr>
        <td align="center" class="tSubTitle" valign="center">&nbsp;</td>
        <td align="center" class="tSubTitle"><b><%= txtTopic %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtReplies %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtRead %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtLstPost %></b></td>
      </tr><%	
	if rs2.EOF or rs2.BOF then  %>
 	  <tr><td colspan="5" class="fSubTitle">&nbsp;<br /><%= txtNoTopFnd %>...<br /></td></tr><%
	else 
	  currTopic = 0
	  TopicCount = 0				
	  do until rs2.EOF or (TopicCount = 10)
	    if chkDisplayForum(rs2("FORUM_ID")) then 
		  if currTopic <> rs2("TOPIC_ID") then %>
		    <tr><td width="10%" align="center"><a href="forum_topic.asp?TOPIC_ID=<% =rs2("TOPIC_ID") %>&FORUM_ID=<% =rs2("FORUM_ID") %>&CAT_ID=<% =rs2("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs2("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs2("F_SUBJECT"),"urlpath") %>">
<%			if rs2("T_STATUS") <> 0 then
	 		  if strHotTopic = "1" then
	  			if rs2("T_LAST_POST") > Session(strUniqueID & "last_here_date") then
	   			  if rs2("T_REPLIES") >= intHotTopicNum then%>
				    <img src="images/icons/icon_folder_new_hot.gif" height="15" width="15" title="<%= txtHotTop %>" alt="<%= txtHotTop %>" border="0"></a>
<%	   			  else%>
 					<img src="images/icons/icon_folder_new.gif" height="15" width="15" title="<%= txtNewTop %>" alt="<%= txtNewTop %>" border="0"></a>
<%	   			  end if
  	  			else
  	   			  if rs2("T_REPLIES") >= intHotTopicNum then%>
				    <img src="images/icons/icon_folder_hot.gif" height="15" width="15" title="<%= txtHotTop %>" alt="<%= txtHotTop %>" border="0"></a>
<%	   			  else%>
					<img src="images/icons/icon_folder.gif" height="15" width="15" border="0"></a>
<%	   			  end if
   	  			end if
     		  else
  	  			if rs2("T_LAST_POST") > Session(strUniqueID & "last_here_date") then %>
				  <img src="images/icons/icon_folder_new.gif" height="15" width="15" title="<%= txtNewTop %>" alt="<%= txtNewTop %>" border="0"></a>
<%	  			else%>
  				  <img src="images/icons/icon_folder.gif" height="15" width="15" border="0"></a> 
<%	  			end if
  	 		  end if
  			else 
  	 		  if rs2("T_LAST_POST") > Session(strUniqueID & "last_here_date") then %>
			    <img src="images/icons/icon_folder_new_locked.gif" title="<%= txtTopLok %>" alt="<%= txtTopLok %>" border="0"></a>
<%	 		  else %>
  				<img src="images/icons/icon_folder_locked.gif" title="<%= txtTopLok %>" alt="<%= txtTopLok %>" border="0"></a>
<%	 		  end if
  			end if %>
			&nbsp;</td>
  			<td valign="center" class="fNorm"><a href="forum_topic.asp?TOPIC_ID=<% =rs2("TOPIC_ID") %>&FORUM_ID=<% =rs2("FORUM_ID") %>&CAT_ID=<% =rs2("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs2("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs2("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs2("T_SUBJECT"), 50),"display") %></a></td>
			<td align="center" class="fNorm"><%= rs2("T_REPLIES") %></td>
			<td align="center" class="fNorm"><%= rs2("T_VIEW_COUNT") %></td>
			<%
			if IsNull(rs2("T_LAST_POST_AUTHOR")) then
				strLastAuthor = ""
			else
				strLastAuthor = "<br />" & txtBy & ": " 
				strLastAuthor = strLastAuthor & "<a href=""cp_main.asp?cmd=8&member="& rs2("T_LAST_POST_AUTHOR") & """><span class=""fSmall"">"
				strLastAuthor = strLastAuthor & getMemberName(rs2("T_LAST_POST_AUTHOR")) & "</span></a>"
			end if
			%>
			<td valign="center" align="center" nowrap><span class="fSmall"><b><% =ChkDate(rs2("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs2("T_LAST_POST")) %><%=strLastAuthor%></span></td>
			</tr>
			<% TopicCount = TopicCount + 1
   		  end if 
		  currTopic = rs2("TOPIC_ID")
  		end if
		rs2.MoveNext 
 	  loop %>				
<%	end if 'if rs2.eof check
  	rs2.close
  	set rs2 = nothing
  end if ' strRecentTopics
%>
 <tr><td colspan="5">&nbsp;</td></tr>
<%
End if
end sub

'::::::::::::::::::::::::::: show my stats ::::::::::::::::::::::::::::::::::::::::::::::::
sub showMemberGroups(id)
  sSql = "SELECT PORTAL_GROUPS.G_ID, PORTAL_GROUPS.G_NAME, PORTAL_GROUPS.G_DESC, PORTAL_GROUPS.G_ADDMEM, PORTAL_GROUP_MEMBERS.G_MEMBER_ID, PORTAL_GROUP_MEMBERS.G_GROUP_ID, PORTAL_GROUP_MEMBERS.G_GROUP_LEADER "
  sSql = sSql & "FROM PORTAL_GROUPS INNER JOIN PORTAL_GROUP_MEMBERS ON PORTAL_GROUPS.G_ID = PORTAL_GROUP_MEMBERS.G_GROUP_ID "
  sSql = sSql & "WHERE (((PORTAL_GROUP_MEMBERS.G_MEMBER_ID)=" & id & "));"
  set rsT = my_Conn.execute(sSql)
  if not rsT.eof then
    if strUserMemberID <> id then
	  sGrpTitle = txtGrpMbrshps
	else
	  sGrpTitle = txtMyGrps
	end if
	%>
	<script type="text/javascript">
	function popEditGrp(mID){
    //alert("Group : " + mID);
	if (mID != 0){
	if (mID != 2){
	if (mID != 3){
    var whereto = "pop_portal.asp?cmd=11&cid=" + mID;
	popUpWind(whereto,'egroups','430','580','yes','yes');
	}}}
	}
	</script>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr><td width="100%" class="tSubTitle"><b><%= sGrpTitle %></b></td></tr>
	<tr><td class="fNorm"><ul>
	<%
    do until rsT.eof
	  Response.Write("<li>")
	  Response.Write("<b>" & rsT("G_NAME") & "</b>")
	  if rsT("G_GROUP_LEADER") = 1 or hasAccess(1) then
	    Response.Write("<a href=""javascript:;"" onclick=""javascript:popEditGrp(" & rsT("G_ID") & ");"">")
	    Response.Write(icon(icnEdit,txtEdit,"","","hspace=""4""") & "</a>")
	  end if
	  Response.Write("</li>") & vbCRLF
      rsT.movenext
    loop
	%></ul></td></tr></table><hr align="center">
	<%  
  end if
  set rsT = nothing
end sub

'::::::::::::::::::::::::::: show my stats ::::::::::::::::::::::::::::::::::::::::::::::::
sub showMyStats()
if flag_showstatus = 1 then
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td width="100%" class="tSubTitle"><b><%= txtStatus %></b></td>
</tr>
<tr><td class="fNorm">
<b><%= txtRank %>:&nbsp;</b><%= getMember_Level(rsMem("M_TITLE"), rsMem("M_LEVEL"), rsMem("M_POSTS")) %>
<br /><b><%= txtRfls %>:&nbsp;</b><% =rsMem("M_RTOTAL")%>
<br /><b><%= txtProfViews %>:&nbsp;</b><% =rsMem("M_PAGE_VIEWS")%>
<br /><b><%= txtCurThm %>:&nbsp;</b><%= strTheme %>
<% if flag_HasForums then %>
<br /><b><%= txtPosts %>:&nbsp;</b><% =rsMem("M_POSTS")%>
<% End If %>
<% If showGold = 1 Then %><br /><b><%= txtGold %>:&nbsp;</b><% =rsMem("M_GOLD")%><% End If %>
<% If showRep = 1 Then %><br /><b><%= txtRepPts %>:&nbsp;</b><% =rsMem("M_REP")%><% End If %>
</td></tr>
<%
'## FLag_SQL - Get Flag from DB
		strSql = "SELECT " & strTablePrefix & "COUNTRIES.CO_FLAG,"& strMemberTablePrefix &"MEMBERS.M_COUNTRY"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS LEFT JOIN "
		strSql = strSql & strTablePrefix & "COUNTRIES ON "& strMemberTablePrefix & "MEMBERS.M_COUNTRY ="& strTablePrefix & "COUNTRIES.CO_NAME "
		strSql = strSql & "WHERE "& strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & strDBNTUserName & "'"
		set rsflag = my_Conn.Execute (strSql) %>
<% If trim(rsflag("M_COUNTRY")) <> "" and not IsNull(rsflag("M_COUNTRY")) Then %>
<% if Trim(rsflag("CO_FLAG")) <> "" and rsflag("CO_FLAG") <> " " and (IsNull(rsflag("CO_FLAG")) = false) then %>
<tr><td align="center" class="fNorm"><hr />
		<b><%= txtCntryFlg %></b><br />
        <img src="<% =rsflag ("CO_FLAG") %>" alt="<%=rsflag("M_COUNTRY") %>" title="<%=rsflag("M_COUNTRY") %>" align="absmiddle" border="0" hspace="0" /><br />
        <%=rsflag("M_COUNTRY") %>
</td></tr>
<% end If 
	end if
   set rsflag  = nothing%>
</table><hr align="center">
<%end if
end sub

sub showPreferences()
 Select Case Request("mode")
  Case "WriteOptions" ' Write the Topic Options in the database
	flag_maxtopics = clng(Request.QueryString("maxtopics"))
	flag_showmytopics = Request.QueryString("showmytopics")
	flag_showrecenttopics = Request.QueryString("showrecenttopics")
	
	strSQL = "UPDATE " & strTablePrefix & "CP_CONFIG SET "

	if flag_showrecenttopics = "1" then
		strSQL = strSQL & "SHOW_RECENT_TOPICS ='1', "
	else
		strSQL = strSQL & "SHOW_RECENT_TOPICS ='0', "
	end if

	if flag_showmytopics = "1" then
		strSQL = strSQL & "SHOW_MY_TOPICS ='1', "
	else
		strSQL = strSQL & "SHOW_MY_TOPICS ='0', "
	end if
	
	if Request.QueryString("showstatus") = "1" then
		strSQL = strSQL & "SHOW_STATUS ='1', "
	else
		strSQL = strSQL & "SHOW_STATUS ='0', "
	end if
	
	strSQL = strSQL & "MAX_MY_TOPICS ='" & flag_maxtopics & "' "
	strSQL = strSQL & "WHERE MEMBER_ID=" & strUserMemberID
	executeThis(strSQL)
	
	call showMsgBlock(1,txtChgsAppl)
			
  Case "PMoptions"
	'flag_showpm = clng(Request("showpm"))
    if chkApp("PM","USERS") then
	  flag_pm_layout = clng(Request("pm_layoutstorage"))
	  strSQL = "UPDATE " & strTablePrefix & "CP_CONFIG SET"
	  strSQL = strSQL & " PM_OUTBOX =" & flag_pm_layout
	  'strSQL = strSQL & ", SHOW_PM =" & flag_showpm
	  strSQL = strSQL & " WHERE MEMBER_ID=" & strUserMemberID
	  executeThis(strSQL)
	
	  strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	  strSql = strSql & "SET M_PMRECEIVE = '" & chkstring(Request.QueryString("statusstorage"), "sqlstring") & "'"
	  strSql = strSql & ", M_PMEMAIL = '" & chkstring(Request.QueryString("emailstorage"), "sqlstring") & "'"
	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	  executeThis(strSql)
	
	  'pm_layoutstorage
	
	  Response.Cookies(strCookieURL & "PmOutBox").Path = strCookieURL
	  Response.Cookies(strCookieURL & "PmOutBox") = flag_pm_layout
	  Response.Cookies(strCookieURL & "PmOutBox").Expires = dateAdd("d", 360, now())

	  call showMsgBlock(1,txtPMSetUpd)
	end if
 end select

  spThemeTitle= "<b>" & strDBNTUserName & "'s " & txtPersSet & "</b>"
  spThemeBlock1_open(intSkin)
	showDateTimeOpts()
	'flag_PMstatus = true
    if PMaccess = 1 and chkApp("PM","USERS") then
	  showPMOptions()
	end if
	if flag_HasForums then
	  showDisplayOpts()
	end if
  spThemeBlock1_close(intSkin)
end sub

sub updateDateTime()
  if request.Form("mode") = "DateOptions" then
	intTimeOffset = cLng(Request("strTimeAdjust"))
	'sDateFormat = Request("strDateType")
	sTimeType = Request("strTimeType")
	intMemberLCID = cLng(Request("intLCID"))
	if intMemberLCID = 0 or len(intMemberLCID) < 4 or len(intMemberLCID) > 5 then
	  intMemberLCID = intPortalLCID
	end if
	strSQL = "UPDATE " & strMemberTablePrefix & "MEMBERS SET "
	'strSQL = strSQL & "M_DATE_FORMAT='" & sDateFormat & "',"
	strSQL = strSQL & "M_TIME_OFFSET=" & intTimeOffset & ","
	strSQL = strSQL & "M_TIME_TYPE='" & sTimeType & "', "
	strSQL = strSQL & "M_LCID=" & intMemberLCID & " "
	strSQL = strSQL & "WHERE MEMBER_ID=" & strUserMemberID
	executeThis(strSQL)
	
	strMTimeAdjust = intTimeOffset
	strMTimeType = sTimeType
	strTimeType = strMTimeType
	session.LCID = intMemberLCID
	
	strMCurDateAdjust = DateAdd("h", (strTimeAdjust + strMTimeAdjust) , now())
	strMCurDateString = DateToStr(strMCurDateAdjust)
	strCurDate = ChkDate2(strMCurDateString)
	'strCurDateAdjust = strCurDate & chkTime2(strMCurDateString)
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
	Session(strUniqueID & "userID") = ""
	closeAndGo("cp_main.asp?cmd=5")
  end if
end sub

sub showDateTimeOpts() %>

<Form method="post" action="cp_main.asp?cmd=5" name="formEle" id="formEle">
<table border="0" cellspacing="1" cellpadding="1" width="100%">
<%
if lcase(strDBNTUserName) = "skydogg" then
 'showDateTimeTest() 
end if %>
  <tr valign="top">
    <td class="tSubTitle" colspan="2" align="center"><%= txtMDTConfig %></td>
  </tr>
  <!-- <tr>
    <td align="right" width="40%"><b>Server Time:</b>&nbsp;</td>
    <td>
    <b><%= now() %></b>
    </td>
  </tr> -->
  <tr>
    <td align="right" class="fNorm"><b><%= txtServLCID %>:</b>&nbsp;</td>
    <td>
    <b><%= session.LCID %></b>
    </td>
  </tr>
  <tr>
    <td align="right" class="fNorm"><b><%= txtPortTime %>:</b>&nbsp;</td>
    <td class="fNorm"><% portDate = chkDate2(strCurDateString) & chkTime2(strCurDateString) %>
	<b><%= portDate %></b>
	<%'= strToDate(strCurDateString) %>
    <%'= strCurDateAdjust %>
    </td>
  </tr>
  <tr>
    <td align="right" class="fNorm"><b><%= txtMemTime %>:</b>&nbsp;</td>
    <td class="fNorm">
	<%'= DateAdd("h", strMTimeAdjust , portDate) %>
    <%= strMCurDateAdjust %>
    </td>
  </tr>
  <tr>
    <td align="center" colspan="2"><hr /></td>
  </tr><%
	  %>
  <tr>
    <td align="right" class="fNorm"><b><%= txtMemLCID %>:</b>&nbsp;</td>
    <td>
	<% displayLCID() %>
		<!-- <input type="text" class="textbox" name="intLCID2" id="intLCID2" value="<%= intMemberLCID %>" maxlength="6" size="10"> -->
    </td>
  </tr>
  <tr>
    <td align="right" class="fNorm"><b><%= txtTimeDisp %>:</b>&nbsp;</td>
    <td class="fNorm">
    24hr <input type="radio" class="radio" name="strTimeType" value="24" <% if strMTimeType = "24" then Response.Write("checked") %>> 
    12hr <input type="radio" class="radio" name="strTimeType" value="12" <% if strMTimeType = "12" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#timetype')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr>
    <td align="right" class="fNorm"><b><%= txtTimeAdj %>:</b>&nbsp;</td>
    <td>
    <select name="strTimeAdjust">
	<% for xl = -24 to 24
		response.Write("<option Value=""" & xl & """" & chkSelect(strMTimeAdjust,xl) & ">" & xl & "</option>")
	   next %>
	</select>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#TimeAdjust')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
</table><br />
		<input type="hidden" name="mode" Value="DateOptions">
		<INPUT TYPE="submit" VALUE="<%= txtUpdSetngs %>" class="button">
</form><br />
<%
end sub

sub showDisplayOpts()
%><table width="100%">
<tr><td class="tSubTitle" align="center"><b><%= txtProfPref %></b></td></tr>
<tr><td align="center">
<Form method="get" action="cp_main.asp?cmd=5" name="formEle" id="formEle">
	<input type="hidden" name="cmd" Value="5">
	<input type="hidden" name="mode" Value="WriteOptions">
		<table align="center">
		  <tr>
		    <td class="fNorm">
      <input type="checkbox" name="showrecenttopics" value="1" <% if flag_showrecenttopics="1" then Response.Write("checked")%>><%= txtShoRecTop %><br />
     	<input type="checkbox" name="showmytopics" value="1" <% if flag_showmytopics="1" then Response.Write("checked")%>><%= txtShow %> 
	    <select name="maxtopics">
	      <option value="5" <% if flag_maxtopics="5" then Response.Write("selected")%>>5</option>
	     	<option value="10" <% if flag_maxtopics="10" then Response.Write("selected")%>>10</option>
		<option value="0" <% if flag_maxtopics="0" then Response.Write("selected")%>><%= txtAll %></option>
	    </select>&nbsp <%= txtLstTopStart %>.<br /><br />

	<input type="hidden" name="showpm" value="<%= flag_showpm %>">
	<input type="hidden" name="showstatus" value="<%= flag_showstatus %>"><%'= txtShoUsrStat %>
	<input type="submit" value="<%= txtUpdtViewProfSet %>" class="button">
    <br />&nbsp;
</td></tr></table>
</form>
</td>
</tr></table>
<%
end sub

sub showPMOptions()
%><table class="tPlain" width="100%">
<tr><td class="tSubTitle" align="center"><b><%= txtPvtMessg & " " & txtOptions %></b></td></tr>
<tr><td align="center">
<%
  	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_PMEMAIL "
  	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
  	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	if strDBType = "db" then
  	  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & chkString(Request.Cookies(strUniqueID & "User")("PWord"),"sqlstring") & "'"
	end if
  	set rsPmOpts = my_Conn.Execute(strSql)
%>
	<Form method="GET" action="cp_main.asp?cmd=5">
		<input type="hidden" name="cmd" Value="5">
		<table align="center">
		  <tr>
		    <td class="fNorm">
		      <b><%= txtUPMTurned %>&nbsp;<% if rsPmOpts("M_PMRECEIVE") = "1" then %> <%= ucase(txtOn) %><% else %> <%= ucase(txtOff) %><% end if %>.</b><br />
		      <INPUT TYPE="RADIO" NAME="statusstorage" VALUE="1" <% if rsPmOpts("M_PMRECEIVE") = "1" then Response.Write("checked")%>> <%= txtEnblPM %>.<br />
		      <INPUT TYPE="RADIO" NAME="statusstorage" VALUE="0" <% if rsPmOpts("M_PMRECEIVE") = "0" then Response.Write("checked")%>> <%= txtDisPM %>.<br />
		    </td>
		  </tr>
		  <tr><td class="fNorm">
		<B><%= txtEmlNoto %></B><br />
		<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="1" <% if rsPmOpts("M_PMEMAIL") = "1" then Response.Write("checked")%>>&nbsp;<%= txtRecEmlNotoPM %>.<br />
		<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="0" <% if rsPmOpts("M_PMEMAIL") = "0" then Response.Write("checked")%>>&nbsp;<%= txtNoPMNoto %>.</td>
		  </tr>
		  <tr>
		    <td class="fNorm"><% if flag_pm_layout= 3 then flag_pm_layout= 0 %>
		      <b><%= txtInOutPref %></b><br />
			 <input type="radio" name="pm_layoutstorage" value="1"<%= chkRadio(flag_pm_layout,1) %>>&nbsp;<%= txt2PgLayout %><br />
		      <INPUT TYPE="RADIO" NAME="pm_layoutstorage" VALUE="0"<%= chkRadio(flag_pm_layout,0) %>>&nbsp;<%= txtNoOutBx %><br /></td>
		  </tr>
		 <!--  <tr>
		    <td><b><%= txtPmAlerts %></b><br />
		      <INPUT TYPE="RADIO" NAME="pm_layoutstorage" VALUE="double" <% if request.cookies(strCookieURL & "PmOutBox") = "double" then Response.Write("checked")%>>&nbsp;<%= txtPopPMAlerts %><br />
			<input type="radio" name="showpm" value="1" <% if flag_showpm="1" then Response.Write("checked")%>>&nbsp;<%= txtBlinkPMAlerts %><br />
			<input type="radio" name="showpm" value="2" <% if flag_showpm="2" then Response.Write("checked")%>>&nbsp;<%= txtBothPMAlerts %><br /></td>
		  </tr> -->
		</table><br />
		<input type="hidden" name="mode" Value="PMoptions">
		<INPUT TYPE="submit" VALUE="<%= txtUpdSetngs %>" class="button">
</form><br />
</td></tr></table>
<%
	set rsPmOpts = nothing
end sub

sub showDateTimeTest() %>
  <tr valign="top">
    <td colspan="2" align="center">
	<table border="0" cellspacing="1" cellpadding="1" width="100%">
  	  <tr valign="top">
    	<td class="tSubTitle" width="33%" align="center">Member</td>
    	<td class="tSubTitle" width="33%" align="center" nowrap>Portal</td>
    	<td class="tSubTitle" align="center">Server</td>
	  </tr>
  	  <tr valign="top">
    	<td>
	<%
	response.Write("intPortalLCID: " & intPortalLCID & "<br />")
	response.Write("intMemberLCID: " & intMemberLCID & "<br /><br />")
	response.Write("strTimeType: " & strTimeType & "<br />")
	response.Write("strMTimeType: " & strMTimeType & "<br /><br />")
	response.Write("strTimeAdjust: " & strTimeAdjust & "<br />")
	response.Write("strMTimeAdjust: " & strMTimeAdjust & "<br /><br />")
	response.Write("session.LCID: " & session.LCID & "<br />")
	response.Write("now(): " & now() & "<br /><br />")
	'session.LCID = cLng(request.Cookies(strUniqueID & "User")("spLCID"))
	 %></td>
    	<td><% 
	response.Write("intMemberLCID: " & intMemberLCID & "<br />")
	response.Write("Portal offset from server time: " & strTimeAdjust & " - strTimeAdjust<br /><br />")
	response.Write("Member offset from Portal time: " & strMTimeAdjust & " - strMTimeAdjust<br /><br />")
	response.Write("Member offset from Server time: " & strTimeAdjust + strMTimeAdjust & " - strSTimeAdjust<br /><br />")
		 %></td>
    	<td align="center">&nbsp;</td>
	  </tr>
  	  <tr valign="top">
    	<td><% 
	response.Write("strMCurDateAdjust:<br /> " & strMCurDateAdjust & "<br /><br />")
	response.Write("strMCurDateString: " & strMCurDateString & "<br />")
	response.Write("strtodate(): " & strtodate(strMCurDateString) & "<br />")
	response.Write("ChkDate(): " & ChkDate(strMCurDateString) & "<br />")
	response.Write("ChkTime(): " & ChkTime(strMCurDateString) & "<br /><br />") 
	response.Write("ChkDate2(): " & ChkDate2(strMCurDateString) & "<br />")
	response.Write("ChkTime2(): " & ChkTime2(strMCurDateString) & "<br /><br />") 
	response.Write("strCurDateAdjust:<br />" & strCurDateAdjust & "<br />")
	response.Write("strCurDate: " & strCurDate & "<br />")
	response.Write("long date: " & FormatDateTime(strCurDateAdjust,1) & "<br />")
	
		%></td>
    	<td align="left"><% 
	tmpStrDateTime = DateToStr(now())
	'tmpStrDateTime = ""
	tmpTime = "12:50:36 PM"
	'FormatDateTime(time(),4)
	yr = left(tmpStrDateTime,4)
	mo = mid(tmpStrDateTime,5,2)
	dy = mid(tmpStrDateTime,7,2)
	'response.Write("DateSerial: " & yr & ":" & mo & ":" & dy & "<br />")
	'response.Write("DateSerial: " & DateSerial(yr,mo,dy) & "<br />")
	response.Write("strCurDateAdjust:<br /> " & strCurDateAdjust & "<br /><br />")
	response.Write("strCurDateString: " & strCurDateString & "<br />")
	response.Write("strtodate(): " & strtodate(strCurDateString) & "<br />")
	response.Write("ChkDate(): " & ChkDate(strCurDateString) & "<br />")
	response.Write("ChkTime(): " & ChkTime(strCurDateString) & "<br /><br />")
	response.Write("ChkDate2(): " & ChkDate2(strCurDateString) & "<br />")
	response.Write("ChkTime2(): " & ChkTime2(strCurDateString) & "<br /><br />")
	'response.Write("DateAdd(""h"", 24, now()): " & DateAdd("h", 24, now()) & "<br /><br />")
	
	'response.Write("FormatDateTime(ChkTime(strCurDateString),3): " & FormatDateTime(ChkTime(strCurDateString),3) & "<br />")
	'response.Write("FormatDateTime(ChkTime(strCurDateString),4): " & FormatDateTime(ChkTime(strCurDateString),4) & "<br />")
	'response.Write("FormatDateTime(strtodate(strCurDateString),3): " & FormatDateTime(strtodate(strCurDateString),3) & "<br />")
	'response.Write("FormatDateTime(strtodate(strCurDateString),4): " & FormatDateTime(strtodate(strCurDateString),4) & "<br /><br />")
	
	response.Write("<br />last_here_date: " & Session(strUniqueID & "last_here_date") & "<br />")
	
		%></td>
    	<td align="left"><% 
	timeAdj = 0
	CkxTime = DateAdd("h", timeAdj , now())
	tmpStrDateTime = DateToStr(CkxTime)
	response.Write("tmpStrDateTime:<br /> " & CkxTime & "<br /><br />")
	response.Write("tmpStrDateTime: " & tmpStrDateTime & "<br />")
	response.Write("strtodate(): " & strtodate(tmpStrDateTime) & "<br />")
	response.Write("ChkDate(): " & ChkDate(tmpStrDateTime) & "<br />")
	response.Write("ChkTime(): " & ChkTime(tmpStrDateTime) & "<br /><br />")
	response.Write("ChkDate2(): " & ChkDate2(tmpStrDateTime) & "<br />")
	response.Write("ChkTime2(): " & ChkTime2(tmpStrDateTime) & "<br /><br />")
		%>
		</td>
	  </tr>
	 </table>
	 </td>
  </tr>
  <%
end sub
%>