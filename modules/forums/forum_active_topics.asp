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
CurPageType = "forums"
CurPageInfoChk = "1"
cnter = 1
%>

<!-- #include file="lang/en/forum_core.asp" -->
<%
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtActvTopics
	PageAction = txtViewing & "<br />" 
	PageLocation = "forum_active_topics.asp"
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"

end function
%>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
<td class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1) %>
	<% menu_fp() %></td>
<td class="mainPgCol">
	<% intSkin = getSkin(intSubSkin,2) %>
<%
	' get module id
	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	end if
	
'## Do Cookie stuffs with reload
nRefreshTime = Request.Cookies(strCookieURL & "Reload")

if Request.form("cookie") = "1" then	
    Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	Response.Cookies(strCookieURL & "Reload") = Request.Form("RefreshTime")
	Response.Cookies(strCookieURL & "Reload").expires = DateAdd("d",365,now())
	nRefreshTime = Request.Form("RefreshTime")
end if

if nRefreshTime = "" then
	nRefreshTime = 0
end if

if not isNumeric(nRefreshTime) then
 nRefreshTime = 0
End if

ActiveSince = Request.Cookies(strCookieURL & "ActiveSince")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = Request.Form("ShowSinceDateTime")	
    Response.Cookies(strCookieURL & "ActiveSince").Path = strCookieURL
	Response.Cookies(strCookieURL & "ActiveSince") = ActiveSince
end if
Select Case ActiveSince
	Case "LastVisit" 
		lastDate = ""
	Case "LastHour" 
		lastDate = datetostr(DateAdd("h",-1,strCurDateAdjust))
	Case "Lastthree" 
		lastDate = datetostr(DateAdd("h",-3,strCurDateAdjust))
	Case "Lastsix" 
		lastDate = datetostr(DateAdd("h",-6,strCurDateAdjust))
	Case "Lasttwelve" 
		lastDate = datetostr(DateAdd("h",-12,strCurDateAdjust))
	Case "LastDay" 
		lastDate = datetostr(DateAdd("d",-1,strCurDateAdjust))
	Case "Last2Day" 
		lastDate = datetostr(DateAdd("d",-2,strCurDateAdjust))
	Case "Last3Day" 
		lastDate = datetostr(DateAdd("d",-3,strCurDateAdjust))
	Case "LastWeek" 
		lastDate = datetostr(DateAdd("d",-7,strCurDateAdjust))
	Case "LastMonth" 
		lastDate = datetostr(DateAdd("m",-1,strCurDateAdjust))
	Case Else
		lastDate = ""
End Select


%>
<script type="text/javascript">
<!--
function autoReload()
{
	document.ReloadFrm.submit()
}

function SetLastDate()
{
	document.LastDateFrm.submit()
}

function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<%
if IsEmpty(sLastHereDate) then
    sLastHereDate = ReadLastHereDate(strDBNTUserName)
end if
if Request.Form("AllRead") = "Y" then
	sLastHereDate = ReadLastHereDate(strDBNTUserName)
	sLastHereDate = ReadLastHereDate(strDBNTUserName)
	lastDate = sLastHereDate
	ActiveSince = ""
end if
if lastDate = "" then
	lastDate = sLastHereDate
end if

' - Get all active topics from last visit
strSql = "SELECT " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS, " 
strSql = strSql & strTablePrefix & "TOPICS.T_VIEW_COUNT, " & strTablePrefix & "TOPICS.FORUM_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "FORUM.CAT_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_MAIL, " 
strSql = strSql & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_POLL, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, "
strSql = strSql & strTablePrefix & "TOPICS.T_NEWS, " & strTablePrefix & "MEMBERS.M_LEVEL, " & strTablePrefix & "MEMBERS.M_GLOW, "
strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "
strSql = strSql & "FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "FORUM, "
strSql = strSql & strTablePrefix & "TOPICS, " & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & "WHERE " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS_1.MEMBER_ID  "
strSql = strSql & "AND " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
strSql = strSql & "AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & "AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & "AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

set rs = my_Conn.Execute (strSql)

  arg1 = txtForums & "|fhome.asp"
  arg2 = txtActvTopics & "|forum_active_topics.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

%>

<% 'shoUpcomingEventsWeek() %>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="fNorm"><form name="LastDateFrm" action="forum_active_topics.asp" method="post">
    &nbsp;&nbsp;Active Topics Since 
    <select name="ShowSinceDateTime" size="1" onchange="SetLastDate();">
        <option value="LastVisit" <% if ActiveSince = "LastVisit" or ActiveSince = "" then Response.Write(" selected=""selected""")%>>&nbsp;Last Visit on <%= ChkDate(Session(strUniqueID & "last_here_date")) %>&nbsp;<% =ChkTime(Session(strUniqueID & "last_here_date")) %>&nbsp;</option>
        <option value="LastHour"  <% if ActiveSince = "LastHour" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLstHr %></option>
        <option value="Lastthree"  <% if ActiveSince = "Lastthree" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLst3Hr %></option>
        <option value="Lastsix"  <% if ActiveSince = "Lastsix" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLst6Hr %></option>
        <option value="Lasttwelve"  <% if ActiveSince = "Lasttwelve" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLst12Hr %></option>
        <option value="LastDay"   <% if ActiveSince = "LastDay" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLstDy %></option>
        <option value="Last2Day"   <% if ActiveSince = "Last2Day" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLst2Dy %></option>
        <option value="Last3Day"   <% if ActiveSince = "Last3Day" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLst3Dy %></option>
        <option value="LastWeek"  <% if ActiveSince = "LastWeek" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLstWk %></option>
        <option value="LastMonth" <% if ActiveSince = "LastMonth" then Response.Write(" selected=""selected""")%>>&nbsp;<%= txtLstMn %></option>
     </select>
    <input type="hidden" name="Cookie" value="2" />
    
    </form>
    </td>
    <td align="center">&nbsp;</td>
    <td align="center">
    <form name="ReloadFrm" action="forum_active_topics.asp" method="post"> 
    <select name="RefreshTime" size="1" onchange="autoReload();">
        <option value="0"  <% if nRefreshTime = "0" then Response.Write(" selected=""selected""")%>>
		<%= txtNoRld %></option>
        <option value="1"  <% if nRefreshTime = "1" then Response.Write(" selected=""selected""")%>>
		<%= txtRld1m %></option>
        <option value="5"  <% if nRefreshTime = "5" then Response.Write(" selected=""selected""")%>>
		<%= txtRld5m %></option>
        <option value="10" <% if nRefreshTime = "10" then Response.Write(" selected=""selected""")%>>
		<%= txtRld10m %></option>
        <option value="15" <% if nRefreshTime = "15" then Response.Write(" selected=""selected""")%>>
		<%= txtRld15m %></option>
        <option value="30" <% if nRefreshTime = "30" then Response.Write(" selected=""selected""")%>>
		<%= txtRld30m %></option>
    </select>
    <input type="hidden" name="Cookie" value="1" />
    
    </form>
    </td>
  </tr>
</table>
<% colsp = 7 %>
<%
spThemeBlock1_open(intSkin) %>
  <table width="100%" cellpadding="0" cellspacing="0">
  <tr><td align="center">
	<table border="0" class="tCellHover" cellspacing="1" cellpadding="2" style="border-collapse: collapse;" width="100%">
      <tr>
        <td align="center" width="50" class="tSubTitle" valign="middle">
        <%If not(rs.EOF or rs.BOF) and (hasAccess(2)) then %>
			<form name="MarkRead" action="forum_active_topics.asp" method="post">
			<input type="hidden" name="AllRead" value="Y" />
			<input type="image" src="images/icons/icon_topic_all_read.gif" value="Mark all read" id="submit1" name="submit1" alt="<%= txtMkAllRead %>" height="20" width="20" border="0" hspace="0" title="<%= txtMkAllRead %>" onclick /></form>
			
        <% else %>
			&nbsp;
        <% end if %>
        </td>
        <td align="center" class="tSubTitle">Topic</td>
        <!-- <td align="center" width="100" class="tSubTitle"><span class="fSubTitle"><b><%= txtAuthor %></b></span></td> -->
        <td align="center" width="50" class="tSubTitle"><%= txtReplies %></td>
        <td align="center" width="50" class="tSubTitle"><%= txtRead %></td>
        <td align="center" width="180" class="tSubTitle"><%= txtLstPost %></td>
        <td align="center" width="50" class="tSubTitle"><%= txtOptions %></td>
      </tr>
<%If rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="6"><span class="fSubTitle"><b><%= txtNoTopFnd %></b></span></td>
      </tr>
<%else
	currForum = 0 
	fDisplayCount = 0 
	bgClr = "tCellAlt1"

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
		if chkForumAccess(strUserMemberID,rs("FORUM_ID")) then
			if (hasAccess(1)) or (chkForumModerator(rs("FORUM_ID"), Session(strCookieURL & "username"))= "1") or (chkForumModerator(rs("FORUM_ID"), chkString(Request.Cookies(strCookieURL & "User")("Name"),"sqlstring")) = "1") then
 				AdminAllowed = 1
 			else   
 				AdminAllowed = 0
 			end if
			fDisplayCount = fDisplayCount + 1
			if currForum <> rs("FORUM_ID") then
'			     %>
				<tr>
					<td height="20" colspan="5" class="tSubTitle" valign="middle" ><a href="<% Response.Write("FORUM.asp?FORUM_ID=" & rs("FORUM_ID") & "&amp;CAT_ID=" & rs("CAT_ID") & "&amp;Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")) %>">&nbsp;<% =ChkString(rs("F_SUBJECT"),"display") %></a></td>
					<td align="center" class="tSubTitle" valign="middle" nowrap="nowrap">
<%
			if (strForumSubscription = 1 or strForumSubscription = 3) and hasAccess(2) then 
			  subscription_id = chkIsSubscribed(intAppID,"0",rs("FORUM_ID"),"0",strUserMemberID)
			  if subscription_id <> 0 then
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=9&amp;cid=" & subscription_id &"');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""" & txtUnSubScrFm & """ alt=""" & txtUnSubScr & """ border=""0"" /></a>&nbsp;" 
			  else
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=7&amp;cmd=2&amp;cid="&rs("FORUM_ID")&"');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""" & txtSubScrFm & """ alt=""" & txtSubScr & """ border=""0"" /></a>&nbsp;" 
			  end if
			end if
			  if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
				<% ForumAdminOptions() %>
			<%else %>
				&nbsp;
<%		      end if %>
				  </td>
				</tr>
<%			end if
	  if bgClr = "tCellAlt2" then
	  	bgClr = "tCellAlt0"
	  else
	  	bgClr = "tCellAlt2"
	  end if
				response.Write("<tr class=""" & bgClr & """ onmouseover=""this.className='tCellHover';"" onmouseout=""this.className='" & bgClr & "';"">")  %>
			<td align="center" valign="middle">
<%			if rsCStatus("CAT_STATUS") <> 0 and rsFstatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then
				if lcase(strHotTopic) = "1" then
					if rs("T_REPLIES") >= intHotTopicNum Then %>
						<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_hot.gif" height="15" width="15" border="0" hspace="0" title="<%= txtHotTop %>" alt="<%= txtHotTop %>" /></a>
<%					else%>
						<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new.gif" title="<%= txtNewTop %>" alt="<%= txtNewTop %>" border="0" /></a>
<%					end if
				end if
			else %>
			<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<%=rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<%=ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_locked.gif" alt=""
<% 			if rsCStatus("CAT_STATUS") = 0 then 
				Response.Write (" title=""" & txtCatLck & """ ")
			elseif rsFStatus("F_STATUS") = 0 then 
				Response.Write (" title=""" & txtFrmLck & """ ")
			else
				Response.Write (" title=""" & txtTopLck & """ ")
			end if %>
			border="0" /></a>
<%			end if
':: sdfnsldijsildfjlisjdisuchclaoicejhoa;cjdckd %>
			</td>
			<td valign="middle" class="fNorm"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;FORUM_ID=<% =rs("FORUM_ID") %>&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a><% if rs("T_NEWS") = 1 then%>&nbsp;<img src="images/icons/icon_topic_news.gif" alt="" /><% end if %><% if rs("T_POLL") <> 0 then %>&nbsp;<img src="images/icons/icon_topic_poll.gif" alt="" /><% end if %>
				<% if strShowPaging = "1" then TopicPaging() end if%></td>
			<!-- <td valign="middle" align="center">
			<%
				strIMmsg = txtView & " " & ChkString(rs("M_NAME"),"display") & "'s " & txtProfile %>
				<a href="cp_main.asp?cmd=8&member=<% =rs("T_AUTHOR") %>" title="<%= strIMmsg %>"><b><%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %></b></a></td> -->
			<td valign="middle" align="center" class="fNorm"><% =rs("T_REPLIES") %></td>
			<td valign="middle" align="center" class="fNorm"><% =rs("T_VIEW_COUNT") %></td>
			<%
			if IsNull(rs("T_LAST_POST_AUTHOR")) then
				strLastAuthor = ""
			else
				strLastAuthor = "<br />" & txtBy & ": <a href=""cp_main.asp?cmd=8&amp;member="& rs("T_LAST_POST_AUTHOR") & """>"
				strLastAuthor = strLastAuthor & rs("LAST_POST_AUTHOR_NAME") & "</a>"
			end if
			'::dkufjdlkfjelidfj.kdzfj.zkdjf.dkjfdzfzdfd
			%>
			<td valign="middle" align="center" nowrap="nowrap"><span class="fSmall"><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></span><a href="link.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&amp;view=lasttopic"><img src="Themes/<%= strTheme %>/icons/arrow1.gif" alt="<%= txtRdLstPst %>" title="<%= txtRdLstPst %>" border="0" hspace="0" /></a></td>
				<td valign="middle" align="center" nowrap="nowrap">
		<%	if intSubscriptions = 1 and hasAccess(2) then 
			  subscription_id = chkIsSubscribed(intAppID,"0","0",rs("TOPIC_ID"),strUserMemberID)
			  if subscription_id <> 0 then
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=9&amp;cid=" & subscription_id &"');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""" & txtUnSubScrTp & """ alt=""" & txtUnSubScr & """ border=""0"" /></a>&nbsp;" 
			  else
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=7&amp;cmd=3&amp;cid="&rs("TOPIC_ID")&"');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""" & txtSubScrTp & """ alt=""" & txtSubScr & """ border=""0"" /></a>&nbsp;" 
			  end if
			end if %>
<%			if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
				<% showTopicOptions() %>
<%			else %>
				&nbsp;
<%			end if %>
				</td>
			</tr>	
<%		end if	
		currForum = rs("FORUM_ID") %>
<%		rs.MoveNext 
	loop 
	if fDisplayCount = 0 then %>
		  <tr>
		 <td colspan="6" class="tCellAlt1"><span class="fTitle"><b><%= txtNoTopFnd %></b></span></td>
		</tr>
<%
	end if 
 end if %>
 </table></td></tr></table>
<%
spThemeBlock1_close(intSkin)
%>

<table width="95%" border="0" align="center">
  <tr>
    <td>&nbsp;
    
    </td>
    <td align="right">
    <!--#INCLUDE file="modules/forums/inc_jump_to.asp" -->
    </td>
  </tr>
</table>
<script type="text/javascript">
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</script>
<% 
set rsCStatus = nothing
set rsFStatus = nothing
rs.close
set rs = nothing %>
</td></tr></table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
sub ForumAdminOptions() 
  cnter = cnter + 1 %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');mwpHSs('formJmpTo','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onmouseover="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onmouseout="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtFrmOpts %>" alt="<%= txtFrmOpts %>" border="0" hspace="0" align="absmiddle" /></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:120px; z-index:100; display:none; position:absolute; right:50px;">
<fieldset><legend><b><%= txtFrmOpts %> </b></legend>
<table width="100%" align="center"><tr><td align="center" nowrap="nowrap">
<%	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then 
		if rsCStatus("CAT_STATUS") = 0 then 
			if hasAccess(1) then %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="images/icons/icon_folder_unlocked.gif" alt="<%= txtUnlokCat %>" title="<%= txtUnlokCat %>" height="15" width="15" border="0" /></a>
<%			else %>
    <img src="images/icons/icon_folder_locked.gif" alt="<%= txtCatLok %>" title="<%= txtCatLok %>" height="15" width="15" border="0" />
<%			end if 
		else 
			if rsFStatus("F_STATUS") <> 0 then %>
    <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_folder_locked.gif" alt="<%= txtLkFrm %>" title="<%= txtLkFrm %>" height="15" width="15" border="0" /></a>
<%			else %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_folder_unlocked.gif" alt="<%= txtUnLkFrm %>" title="<%= txtUnLkFrm %>" height="15" width="15" border="0" /></a>
<%			end if 
		end if 
		if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0) or (AdminAllowed = 1) then %>
          <a href="forum_post.asp?method=EditForum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&type=0"><img src="images/icons/icon_folder_pencil.gif" title="<%= txtEdFrmProp %>" alt="<%= txtEdFrmProp %>" border="0" hspace="0" /></a>
<%		end if %>
    <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_folder_delete.gif" title="<%= txtDelFrm %>" alt="<%= txtDelFrm %>" height="15" width="15" border="0" /></a>
    <a href="forum_post.asp?method=Topic&FORUM_ID=<% =rs("FORUM_ID")%>&CAT_ID=<% =rs("CAT_ID")%>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="<%= txtNewTop %>" alt="<%= txtNewTop %>" height="15" width="15" border="0" /></a>
<%	end if %>
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); mwpHSs('formJmpTo','1');"><span class="fSmall"><%= txtClose %></span></a></center>
</td></tr></table>
</fieldset></div> <%
end sub

sub showTopicOptions()
  				cnter = cnter + 1 %>
          		<a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');mwpHSs('formJmpTo','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onmouseover="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onmouseout="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtTopOpts %>" alt="<%= txtTopOpts %>" border="0" hspace="0" align="absmiddle" /></a>
				<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:112px; z-index:100; display:none; position:absolute; right:50px;">
<fieldset><legend><b><%= txtTopOpts %> </b></legend>
<table width="100%" align="center"><tr><td align="center" nowrap="nowrap">
				<b>
<%				if rsCStatus("CAT_STATUS") = 0 then %>
					<a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnlokCat %>" alt="<%= txtUnlokCat %>" border="0" hspace="0" /></a>
<%				else 
					if rsFStatus("F_STATUS") = 0 then %>
						<a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLkFrm %>" alt="<%= txtUnLkFrm %>" border="0" hspace="0" /></a>
<%					else 
						if rs("T_STATUS") <> 0 then %>
							<a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_lock.gif" title="<%= txtLkTop %>" alt="<%= txtLkTop %>" border="0" hspace="0" /></a>
<%						else %>
							<a href="JavaScript:openWindow('forum_pop_open.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLkTop %>" alt="<%= txtUnLkTop %>" border="0" hspace="0" /></a>
<%						end if 
					end if 
				end if 
				if (AdminAllowed = 1) or (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0) then %>
					<a href="forum_post.asp?method=EditTopic&amp;TOPIC_ID=<% =rs("TOPIC_ID") %>&&amp;FORUM_ID=<% =rs("FORUM_ID") %>&&amp;CAT_ID=<% =rs("CAT_ID") %>&amp;auth=<% =rs("T_AUTHOR") %>&amp;Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&amp;Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_pencil.gif" title="<%= txtEdMsg %>" alt="<%= txtEdMsg %>" border="0" hspace="0" /></a>
<%				end if %>
				<a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_trashcan.gif" title="<%= txtDelTop %>" alt="<%= txtDelTop %>" border="0" hspace="0" /></a>
				<a href="forum_post.asp?method=Reply&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" title="<%= txtRplyTop %>" alt="<%= txtRplyTop %>" height="15" width="15" border="0" /></a>
				</b><br />
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); mwpHSs('formJmpTo','1');"><span class="fSmall"><%= txtClose %></span></a></center>
</td></tr></table>
<%'spThemeBlock3_close()%>
  </fieldset>
</div> <%
end sub


sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td valign=""middle""><img src=""images/icons/icon_posticon.gif"" border=""0"" alt="""" /></td>")
		for counter = 1 to mxpages
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr><td>&nbsp;</td>")
			end if
			ref = "<td align=""center"" valign=""middle"" class=""tCellAlt1"">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<a href='forum_topic.asp?"
            		ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
		        ref = ref & "&amp;FORUM_ID=" & rs("FORUM_ID")
		        ref = ref & "&amp;CAT_ID=" & rs("CAT_ID")
		        ref = ref & "&amp;Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
		        ref = ref & "&amp;Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")
			ref = ref & "&amp;whichpage=" & counter
			ref = ref & "'><span class=""fSmall"">" & counter & "</span></a></td>"
			Response.Write ref 
		next			
        Response.Write("</tr></table>")
	end if
end sub
%>