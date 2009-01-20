<% 

':: PORTAL DEFAULT MENU ::::
sub defaultMenu()
 if bFso then
  mnu.menuName = "def_main"
  mnu.template = 5
  mnu.thmBlk = 0
  mnu.title = ""
  mnu.shoExpanded = 0
  mnu.canMinMax = 1
  mnu.keepOpen = 0
  mnu.GetMenu()
 else
  d_menu()
 end if
end sub

':: PORTAL NAVBAR MENU ::::
sub navbarMenu()
 if bFso then
  mnu.menuName = "nav_main"
  mnu.template = 3
  mnu.GetMenu()
 else
 %><!--#include file="menu_com_rc2.asp"--><%  
 end if
end sub

':: PORTAL USER CONTROL PANEL MENU ::::
sub cp_userMenu()
 if bFso then
  mnu.menuName = "cp_main"
  mnu.template = 4
  mnu.thmBlk = 0
  mnu.title = ""
  mnu.shoExpanded = 0
  mnu.canMinMax = 1
  mnu.keepOpen = 0
  mnu.GetMenu()
 else
  cpMenu()
 end if
end sub

sub cpMenu() %>
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<A href="default.asp"><%= txtHome %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<A href="cp_main.asp?cmd=5"><%= txtPersSet %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=9"><%= txtEditProf %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=8&member=<%= id%>"><%= txtViewProf %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=1"><%= txtEditAvatar %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=7"><%= txtMyBkmks %></a>
	&nbsp;(<%= getCount("BOOKMARK_ID","" & strTablePrefix & "BOOKMARKS","M_ID=" & strUserMemberID) %>)<br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=6"><%= txtMySubsc %></a>
	&nbsp;(<%= getCount("SUBSCRIPTION_ID","" & strTablePrefix & "SUBSCRIPTIONS","M_ID=" & strUserMemberID) %>)<br />
<% if flag_HasForums then %>
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="cp_main.asp?cmd=4"><%= txtShoRecTopics %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="forum_active_topics.asp"><%= txtActvTopics %></a><br />
<% End If %>
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="members.asp"><%= txtMbrLst %></a><br />
	<img src="Themes/<%= strTheme %>/icons/arrow1.gif">&nbsp;<a href="pm.asp"><%= txtPvtMsgs %></a>
<%
end sub

sub d_menu() %>
<table><tr><td width="100%">
<div class="menu">
<% if not hasAccess(2) and strNewReg = 1 then %>
<a href="policy.asp">- <%= txtRegister %><br /></a>
<% End If %>
<% if chkApp("forums","USERS") then %>
<a href="forum_active_topics.asp">- <%= txtActvTopics %>&nbsp;<%= cntActiveTopics() %><br /></a>
<% End If %>
<% if chkApp("events","USERS") then %>
<a href="<%= eUrl %>">- <%= txtEvents %> <% If eCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% End If %>
<% if chkApp("article","USERS") then %>
<a href="<%= aUrl %>">- <%= txtArticles %> <% If aCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% End If %>
<% if chkApp("downloads","USERS") then %>
<a href="<%= dlUrl %>">- <%= txtDownloads %> <% If dlCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% End If %>
<% if chkApp("pictures","USERS") then %>
<a href="<%= pUrl %>">- <%= txtPics %> <% If pCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% End If %>
<% if chkApp("classifieds","USERS") then %>
<a href="<%= clUrl %>">- <%= txtClassifieds %> <% If clCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% End If %>
<% if chkApp("links","USERS") then %>
<a href="<%= linkUrl %>">- <%= txtLinks %> <% If linkCnt <> 0 Then %><img src="themes/<%= strTheme %>/new.gif" border="0" alt="" title="" /><% End If %><br /></a>
<% end if

if not hasAccess(2) Then 
%>		
<% if chkApp("PM","USERS") then %>
	  <a href="javascript:;" title="">- <%= txtMsg %><br /></a>
<% end if
   if intBookmarks then %>
	  <a href="javascript:;" title="<%= txtBkmkLnkHov %>">- <%= txtMyBkmks %><br /></a>
<% end if
   if intSubscriptions then %>
	  <a href="javascript:;" title="<%= txtSubscLnkHov %>">- <%= txtSubsc %><br /></a>
<% end if %>
	  <a href="javascript:;" title="<%= txtStatsLnkHov %>">- <%= txtSStats %><br /></a>
	<% 
else
   if chkApp("PM","USERS") then %>
	<a href="pm.asp"><% if not pmCount = 0 then%><b>- <%= txtMsg %></b> <img src="images/icons/icon_pm2.gif" border="0"><% Else %>- <%= txtMsgs %><%end if%><br /></a>
<% end if
   if intBookmarks then %>
	<a href="cp_main.asp?cmd=7" title="<%= txtView & " " & txtMyBkmks %>">- <%= txtMyBkmks %><br /></a>
<% end if
   if intSubscriptions then %>
	<a href="cp_main.asp?cmd=6" title="<%= txtView & " " & txtMySubsc %>">- <%= txtMySubsc %><br /></a>
<% end if %>
	<a href="statistics.asp" title="<%= txtSStats %>">- <%= txtSStats %><br /></a>
<% if hasAccess(1) then %>
<a href="site_monitor.asp" title="<%= txtMonSite %>" target="_search">- <%= txtMxMon %><br /></a>
<% End If %>
<%if hasAccess(1) and chkApp("forums","USERS") Then 
rptCnt = getCount("R_STATUS",strTablePrefix & "REPORTED_POST","R_STATUS=0") %>
<% if rptCnt <> 0 then%>
	<a href="forum_report_post_moderate.asp" title="<%= txtRptdPst %>!">- <span class="fAlert"><b><%= txtRptdPst %></b></span></a>
<% Else %>
	<a href="forum_report_post_moderate.asp" title="<%= txtRptdPst %>">- <%= txtRptdPst %><br /></a>
<% end if%>
<%end if %>
<% if hasAccess(1) then %>
<a href="admin_home.asp">- <%= txtAdminOpts %><br /></a>
<% 
   If PTcnt <> 0 Then %>
<a href="admin_home.asp"><b>- <%= txtPndTsks %> <%= cntPendTsks() %></b><br /></a>
<% end if
end if %>
<% End If %>
</div></td></tr></table>
 <%
end sub

%>
