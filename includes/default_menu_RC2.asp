<%
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


%>
<%

function menu_fp()

' Count the number of MEMBERS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Members] "
else
	strSqL = "SELECT count(UserID) Members  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) <> '" & txtGuest & "' "

Set rsMembers = my_Conn.Execute(strSql)
Members = rsMembers("Members")
strOnlineMembersCount = rsMembers("Members")
Set rsMembers = nothing

' Count the number of GUESTS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) = '" & txtGuest & "' "

Set rsGuests = my_Conn.Execute(strSql)
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")
Set rsGuests = nothing

':::::::::::::::::::::::::: get item counts ::::::::::::::::::::::::::::::::::::::
PTcnt = 0

if chkApp("events","USERS") then
eCnt = getCount("EVENT_ID","PORTAL_EVENTS","PENDING = 0 AND DATE_ADDED >= '" & Session(strUniqueID & "last_here_date") & "' AND PRIVATE = 0")
If eCnt = 0 Then eUrl = "events.asp" else eUrl = "events.asp?mode=newEvents" end if
'Pending Events count
PTcnt = PTcnt + getCount("EVENT_ID",strTablePrefix & "EVENTS","PENDING=1") 
end if

if chkApp("article","USERS") then
aCnt = getCount("ARTICLE_ID","ARTICLE","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND show = 1")
If aCnt = 0 Then aUrl = "article.asp" else aUrl = "article.asp?cmd=3" end if
' Pending Articles count
PTcnt = PTcnt + getCount("ARTICLE_ID","ARTICLE","SHOW=0")
end if

if chkApp("downloads","USERS") then
dlCnt = getCount("DL_ID","DL","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND show = 1")
If dlCnt = 0 Then dlUrl = "dl.asp" else dlUrl = "dl.asp?cmd=3" end if
' Pending Downloads count
PTcnt = PTcnt + getCount("DL_ID","DL","SHOW=0 OR BADLINK<>0")
end if

if chkApp("pictures","USERS") then
pCnt = getCount("PIC_ID","PIC","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND show = 1")
If pCnt = 0 Then pUrl = "pic.asp" else pUrl = "pic.asp?cmd=3" end if
' Pending Pictures count
PTcnt = PTcnt + getCount("PIC_ID","pic","SHOW=0 OR BADLINK<>0") 
end if

if chkApp("classifieds","USERS") then
clCnt = getCount("CLASSIFIED_ID","CLASSIFIED","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND show = 1")
If clCnt = 0 Then clUrl = "classified.asp" else clUrl = "classified.asp?cmd=3" end if
' Pending Classifieds count
PTcnt = PTcnt + getCount("CLASSIFIED_ID","CLASSIFIED","SHOW=0 OR BADLINK<>0")
end if

if chkApp("links","USERS") then
linkCnt = getCount("LINK_ID","LINKS","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND show = 1")
If linkCnt = 0 Then linkUrl = "links.asp" else linkUrl = "links.asp?cmd=3" end if
' Pending Links count
PTcnt = PTcnt + getCount("LINK_ID","LINKS","SHOW=0 OR BADLINK<>0")
end if

totalCnt = aCnt + dlCnt + pCnt + clCnt + linkCnt 


'::::::::::::::::::::::: Start the menu HTML ::::::::::::::::::::::::::::::
spThemeTitle= txtMenu
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)
%>
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
<a href="admin_home.asp"><b>- <%= txtPndTsks %> (<%= PTcnt %>)</b><br /></a>
<% end if
end if %>
<% End If %>
</div></td></tr>
<tr><td width="100%"><hr /></td></tr>
<% if hasAccess(2) then
strSql = "SELECT " & strTablePrefix & "TOTALS.U_COUNT "
strSql = strSql & " FROM " & strTablePrefix & "TOTALS"
set rs1 = my_Conn.Execute(strSql)
Users = rs1("U_COUNT")
rs1.Close
set rs1 = nothing
%>
<tr><td width="100%"><span class="fSmall"><a href="members.asp"><%= txtMembers %>: <% =Users%></a></span></td></tr>
<% End If %>
<tr><td width="100%"><a href="active_users.asp"><span class="fSmall"><%= txtActvUsrs %>: <br /><%=strOnlineMembersCount & " " & txtMembers & " " & txtAnd & " " & strOnlineGuestsCount & " " & txtGuests %></span></a></td></tr></table>

<% 
spThemeBlock1_close(intSkin)
end function
%>