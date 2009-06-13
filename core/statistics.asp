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

CurPageInfoChk = "1" %>
<!--#include file="lang/en/forum_core.asp" -->
<!--#include file="inc_functions.asp" -->
<%
function CurPageInfo ()
	PageName = txtSStats
	PageAction = txtViewing & "<br />" 
	PageLocation = "statistics.asp"
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"

end function
%>
<!--#include file="inc_top.asp" -->
<%if strDBNTUserName = "" then
	doNotLoggedInForm
else


dim intTopPosters
dim intPageViews
dim intLastSeen
dim intReadTopics
dim intRepliedTopics
dim intTopReferrers

intTopPosters = 25	' number of top posters to display
intLastSeen	= 25	' number of members last seen to display
intReadTopics = 25	' number of top read topics to display
intRepliedTopics = 25	' number of top replied to topics to display
intTopReferrers	= 25	' number of top referrers to display
intPageViews = 10	' number of top viewed profiles to display

dim boolPM, boolBookmarks
boolPM = 0
boolBookmarks = 0
boolPM = 1
boolBookmarks = 1
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td class="leftPgCol" valign="top">
	<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtJmpTo
spThemeBlock1_open(intSkin)
%><table cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td align="left" class="fNorm">
	<ul style="margin-left:3px;">
    <li><a href="#stats"><%= strSiteTitle & "&nbsp;" & txtStats %></a></li>
<% if hasAccess(1) then%>
    <li><a href="#lastseen"><%= intLastSeen & "&nbsp;" & txtMLstSeen %> </a></li>
<%end if%>
    <li><a href="#referrers"><%= txtTop & "&nbsp;" & intTopReferrers & "&nbsp;" & txtRefers %></a></li>
<% If chkApp("forums","USERS") Then %>
    <li><a href="#posters"><%= txtTop & "&nbsp;" & intTopPosters & "&nbsp;" & txtPosters %></a></li>
  <% if hasAccess(1) then %>
    <li><a href="#replied"><%= txtTop & "&nbsp;" & intTopReferrers & "&nbsp;" & txtRplyToTops %></a></li>
    <li><a href="#read"><%= txtTop & "&nbsp;" & intReadTopics & "&nbsp;" & txtRdTops %></a></li>
  <%end if%>
<% End If %>
    <li><a href="#pageview"><%= txtCelebs %></a></li>
<% if hasAccess(1) then%>
    <li><a href="statisticsx.asp"><%= txtExStats %></a></li>
<%end if%>    
    </ul>
    </td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)
	
	menu_fp() %>
	</td>
	<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = strSiteTitle & " " & txtStats & "|statistics.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<center><div>
<%
	dim SQLMemberCount, RSMEmberCount
	SQLMemberCount = "SELECT COUNT(Member_id) AS MemberCount FROM " & strMemberTablePrefix & "MEMBERS WHERE M_NAME not like '%skyiamdogg%'"
	Set RSMemberCount = my_Conn.execute(SQLMemberCount)
	intMemberCount = RSMemberCOUNT("MemberCount")
	set RSMEmberCount = nothing

spThemeTitle= strSiteTitle & "&nbsp;" & txtStats
spThemeBlock1_open(intSkin)
Response.Write("<table cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""tCellAlt1"">")
%>
  <tr>
    <td align="left" class="fNorm"><a name="stats"></a>
            
      <b><u><%= txtMembers %></u></b><br />
      <%= txtMembers %>:  <b><%= intMemberCount %></b><br />
      <%= txtF & "&nbsp;" & txtMembers %>:  <b><%= getFemalesCount%></b><br />
      <%= txtM & "&nbsp;" & txtMembers %>:  <b><%= getMalesCount%></b><br />
      <br />
	<% If chkApp("forums","USERS") Then %>
      <b><u><%= txtForum %></u></b><br />
      <% displayStats%><br /><br />
    <% End If %>
	
	<% If chkApp("PM","USERS") and boolPM = 1 Then %>
      <b><u><%= txtPvtMessgs %></u></b><br />
	  <% pCt = GetPMCount 
	 	 pmToday = GetPMToday %>
     	  <b><%= pCt %></b>&nbsp;<%= txtPMinSys %>.<br />
      	  <b><%= GetPMToday%></b>&nbsp;<%= txtPM24Hrs %>.<br /><br />
    <% End If %>
	
    <% if intBookmarks = 1 then %>
      <b><u><%= txtBookmks %></u></b><br />
      	<b><%= GetBookmarkCount %></b>&nbsp;<%= txtBkmkInSys %>.<br /><br />
    <% end if %>
	
    <% if intSubscriptions = 1 then %>
      <b><u><%= txtSubsc %></u></b><br />
      	<b><%= GetSubscriptionCount %></b>&nbsp;<%= txtSubscInSys %>.<br /><br />
    <% end if %>
      <br />
    </td>
  </tr>
<% if hasAccess(1) then%>
  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="lastseen"></a>
      <b>
      <%= intLastSeen & " " & txtMLstSeen %>
      </b>
    </td>
  </tr>
  <tr>
    <td align="center" class="fNorm">
    <%=displayLastSeen%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>
<%end if%>

  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="referrers"></a>
      <b><%= txtTop & "&nbsp;" & intTopReferrers & "&nbsp;" & txtRefers %></b>
    </td>
  </tr>
  <tr>
    <td align="center">
    <%=displayTopReferrers%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="pageview"></a>
      <b><%= txtCelebs %></b>
    </td>
  </tr>
  <tr>
    <td align="center">
    <%=displayPageViews%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>

<% If chkApp("forums","USERS") Then %>
  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="posters"></a>
      <b><%= txtTop & "&nbsp;" & intTopPosters & "&nbsp;" & txtPosters %></b>
    </td>
  </tr>
  <tr>
    <td align="center">
    <%=displayTopPosters%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>
<% if hasAccess(1) then%>

  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="replied"></a>
      <b><%= txtTop & "&nbsp;" & intTopReferrers & "&nbsp;" & txtRplyToTops %></b>
    </td>
  </tr>
  <tr>
    <td align="center">
    <%=displayRepliedTopics%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>

  <tr>
    <td align="left" class="tSubTitle" width="100%"><a name="read"></a>
      <b><%= txtTop & "&nbsp;" & intReadTopics & "&nbsp;" & txtRdTops %></b>
    </td>
  </tr>
  <tr>
    <td align="center">
    <%=displayReadTopics%>
    <p align="right"><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" width="15" height="15" alt="" /></a></p>
    </td>
  </tr>
<%
 end if
 End If
Response.Write("</table>")
spThemeBlock1_close(intSkin)%></div></center>
    </td>
  </tr>
</table>
<% 
end if %>
<!-- #include file="inc_footer.asp" -->
<%
function getFemalesCount
	dim rs, intFemales, intMales 
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(*) as females FROM " & strMemberTablePrefix & "MEMBERS WHERE m_sex = 'female' and M_NAME not like '%iamviet %'"
	rs.Open strSql, my_Conn
	getFemalesCount = rs("females")
	rs.Close
end function

function getMalesCOUNT()
	dim rs, intFemales, intMales 
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(*) as males FROM " & strMemberTablePrefix & "MEMBERS WHERE m_sex = 'male' and M_NAME not like '%iamviet %'"
	rs.Open strSql, my_Conn
	getMalesCount = rs("males")
	rs.Close
end function

function DisplayLastSeen()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	
strSql = "SELECT TOP " & intLastSeen & " member_id, m_name, m_lastheredate FROM " & strMemberTablePrefix & "MEMBERS ORDER BY M_LASTHEREDATE DESC"
rs.Open strSql, my_Conn
%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b><%= txtMemName %></b></td>
		<td valign="middle" align="center"><b><%= txtLstSeen %></b></td>
	</tr>
	<% 
	intCounter = 1
	do while not rs.EOF%>
	<tr>
		<td class="fNorm"><%= intCounter%>.</td>
		<td align="left" class="fNorm"><a href="cp_main.asp?cmd=8&amp;member=<%=rs("member_id")%>"><%= rs("m_name")%></a></td>
		<td align="center" class="fNorm">
		  <% 
		  minutesTotal = datediff("n",strtodate(rs("m_lastheredate")),strMCurDateAdjust)
		  'minutesTotal = datediff("n",strMCurDateAdjust,strtodate(rs("m_lastheredate")))
		  hoursTotal = int(minutesTotal / 60)
		  daysSince = int(hoursTotal / 24)
		  hoursSince = hoursTotal mod 24
		  minutesSince = minutestotal mod 60
		  if daysSince = 1 then Response.Write daysSince & " " & txtDay & ", "
		  if daysSince > 1 then Response.Write daysSince & " " & txtDays & ", "			
		  if hoursSince = 1 then Response.Write hoursSince & " " & txtHour & ", "		  
		  if hoursSince > 1 then Response.Write hoursSince & " " & txtHours & ", "
		  Response.Write minutesSince & "&nbsp;" & txtMinsAgo
		  'Response.Write minutesTotal & "&nbsp;X" & txtMinsAgo
		  %>.
		</td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop
	rs.Close
	%>
</table>
<%
end function

function displayStats()
	set rs1 = Server.CreateObject("ADODB.Recordset")

	'## " & strTablePrefix & "SQL
	strSql = "SELECT " & strTablePrefix & "TOTALS.P_COUNT, " & strTablePrefix & "TOTALS.T_COUNT, " & strTablePrefix & "TOTALS.U_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "TOTALS"

	rs1.open strSql, my_Conn

	Users = rs1("U_COUNT")
	Topics = rs1("T_COUNT")
	Posts = rs1("P_COUNT")

	rs1.Close
	set rs1 = nothing

	ShowLastHere = hasAccess(2)
		Response.Write replace(replace(txtTotCnt,"[%postcnt%]","<b>" & Posts & "</b>"),"[%topiccnt%]","<b>" & Topics & "</b>")
end function

function GetPMCOUNT()
	dim rs, intFemales, intMales 
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(M_ID) AS PMCount FROM " & strTablePrefix & "PM"
	rs.Open strSql, my_Conn
	GetPMCount = rs("PMCOUNT")
	rs.Close
end function

function GetPMToday()
	dim rs, intFemales, intMales 
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(M_ID) AS PMCount FROM " & strTablePrefix & "PM WHERE m_sent > '" & DateToStr(DateAdd("h",-24,DateAdd("h", strTimeAdjust , Now()))) & "'"
	rs.Open strSql, my_Conn
	GetPMToday = rs("PMCOUNT")
	rs.Close
end function

function displayTopPosters()
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	strSql = "SELECT TOP " & intTopPosters & " member_id, m_name, m_posts FROM " & strMemberTablePrefix & "MEMBERS ORDER BY M_posts DESC"
	rs.Open strSql, my_Conn
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b><%= txtMemName %></b></td>
		<td valign="middle" align="center"><b><%= txtPosts %></b></td>
	</tr>
	<% 
	dim intCounter
	intCounter = 1
	do while not rs.EOF and intCounter < intTopPosters + 1%>
	<tr>
		<td class="fNorm"><%= intCounter%>.</td>
		<td align="left" class="fNorm"><a href="cp_main.asp?cmd=8&amp;member=<%=rs("member_id")%>"><%= rs("m_name")%></a></td>
		<td align="center" class="fNorm"><%= rs("m_posts")%></td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop%>
</table>
<%

	rs.Close
	set rs = nothing
end function

function displayReadTopics()
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	strSql = "SELECT TOP " & intReadTopics & " TOPIC_ID, T_SUBJECT, T_VIEW_COUNT FROM " & strTablePrefix & "TOPICS ORDER BY T_VIEW_COUNT DESC"
	rs.Open strSql, my_Conn
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b>Topic Subject</b></td>
		<td valign="middle" align="center"><b>View Count</b></td>
	</tr>
	<% 
	dim intCounter
	intCounter = 1
	do while not rs.EOF%>
	<tr>
		<td class="fNorm"><%= intCounter%>.</td>
		<td align="left" class="fNorm"><a href="link.asp?TOPIC_ID=<%=rs("TOPIC_ID")%>"><%= rs("T_SUBJECT")%></a></td>
		<td align="center" class="fNorm"><%= rs("T_VIEW_COUNT")%></td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop%>
</table>
<%

	rs.Close
	set rs = nothing
end function

function displayRepliedTopics()
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	strSql = "SELECT TOP " & intRepliedTopics & " TOPIC_ID, T_SUBJECT, T_REPLIES FROM " & strTablePrefix & "TOPICS ORDER BY T_REPLIES DESC"
	rs.Open strSql, my_Conn
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b><%= txtSubject %></b></td>
		<td valign="middle" align="center"><b><%= txtRplyCnt %></b></td>
	</tr>
	<% 
	dim intCounter
	intCounter = 1
	do while not rs.EOF and intCounter < intRepliedTopics + 1%>
	<tr>
		<td class="fNorm"><%= intCounter%>.</td>
		<td align="left" class="fNorm"><a href="link.asp?TOPIC_ID=<%=rs("TOPIC_ID")%>"><%= rs("T_SUBJECT")%></a></td>
		<td align="center" class="fNorm"><%= rs("T_REPLIES")%></td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop%>
</table>

<%

	rs.Close
	set rs = nothing
end function

function displayTopReferrers()
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	strSql = "SELECT TOP " & intTopReferrers & " member_id, m_name, M_RTOTAL FROM " & strMemberTablePrefix & "MEMBERS WHERE M_RTOTAL <> 0 ORDER BY M_RTOTAL DESC"
	rs.Open strSql, my_Conn
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b><%= txtMemName %></b></td>
		<td valign="middle" align="center"><b><%= txtRfls %></b></td>
	</tr>
	<% 
	dim intCounter
	intCounter = 1
	do while not rs.EOF and intCounter < intTopReferrers + 1%>
	<tr class="fNorm">
		<td><%= intCounter%>.</td>
		<td align="left"><a href="cp_main.asp?cmd=8&amp;member=<%=rs("member_id")%>"><%= rs("m_name")%></a></td>
		<td align="center"><%= rs("M_RTOTAL")%></td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop%>
</table>
<%

	rs.Close
	set rs = nothing
end function

function displayPageViews()
	dim rs
	set rs = server.CreateObject("adodb.recordset")

	strSql = "SELECT TOP " & intPageViews & " member_id, m_name, M_PAGE_VIEWS FROM " & strMemberTablePrefix & "MEMBERS ORDER BY M_PAGE_VIEWS DESC"
	rs.Open strSql, my_Conn
	%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr class="fSubTitle">
		<td valign="middle" width="10%">&nbsp;</td>
		<td valign="middle" align="left"><b><b><%= txtMemName %></b></b></td>
		<td valign="middle" align="center"><b><%= txtProfViews %></b></td>
	</tr>
	<% 
	dim intCounter
	intCounter = 1
	do while not rs.EOF and intCounter < intPageViews + 1%>
	<tr class="fNorm">
		<td><%= intCounter%>.</td>
		<td align="left"><a href="cp_main.asp?cmd=8&amp;member=<%=rs("member_id")%>"><%= rs("m_name")%></a></td>
		<td align="center"><%= rs("M_PAGE_VIEWS")%></td>
	</tr>
	<%
	intCounter = intCounter + 1
	rs.movenext
	loop%>
</table>
<%

	rs.Close
	set rs = nothing
end function

function GetUploadCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(UPLOAD_ID) AS UPLOADS FROM UPLOADS WHERE WEBLINK = 0"
	rs.Open strSql, my_Conn
	GetUploadCount = rs("UPLOADS")
	rs.Close
	set rs = nothing
end function

function GetWeblinkCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(UPLOAD_ID) AS WEBLINKS FROM UPLOADS WHERE WEBLINK = 1"
	rs.Open strSql, my_Conn
	GetWeblinkCount = rs("WEBLINKS")
	rs.Close
	set rs = nothing
end function

function GetUploadMembersCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(distinct UPLOAD_BY) AS UPLOADS FROM UPLOADS"
	rs.Open strSql, my_Conn
	GetUploadMembersCount = rs("UPLOADS")
	rs.Close
	set rs = nothing
end function

function GetBookmarkCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(*) AS BOOKMARKS FROM " & strTablePrefix & "BOOKMARKS"
	rs.Open strSql, my_Conn
	GetBookmarkCount = rs("BOOKMARKS")
	rs.Close
	set rs = nothing
end function

function GetSubscriptionCount()
	dim rs
	set rs = server.CreateObject("adodb.recordset")
	strSql = "SELECT COUNT(*) AS SUBSCRIPTIONS FROM " & strTablePrefix & "SUBSCRIPTIONS"
	rs.Open strSql, my_Conn
	GetSubscriptionCount = rs("SUBSCRIPTIONS")
	rs.Close
	set rs = nothing
end function
%>