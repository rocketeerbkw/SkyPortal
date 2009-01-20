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
curPageType = "core"
%>
<!--#include file="inc_functions.asp" -->
<% thispage = "monitor" 
intSkin = 1
ActiveSince = chkString(Request.Cookies(strCookieURL & "ActiveSince"),"sqlstring")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = chkString(Request.Form("ShowSinceDateTime"),"sqlstring")	
    Response.Cookies(strCookieURL & "ActiveSince").Path = strCookieURL
	Response.Cookies(strCookieURL & "ActiveSince") = ActiveSince
end if
Select Case ActiveSince
	Case "LastVisit" 
		lastDate = ""
	Case "LastHour" 
		lastDate = DateToStr(DateAdd("h",-1,DateAdd("h", strTimeAdjust , Now())))
	Case "Lastthree" 
		lastDate = DateToStr(DateAdd("h",-3,DateAdd("h", strTimeAdjust , Now())))
	Case "Lastsix" 
		lastDate = DateToStr(DateAdd("h",-6,DateAdd("h", strTimeAdjust , Now())))
	Case "Lasttwelve" 
		lastDate = DateToStr(DateAdd("h",-12,DateAdd("h", strTimeAdjust , Now())))
	Case "LastDay" 
		lastDate = DateToStr(DateAdd("d",-1,DateAdd("h", strTimeAdjust , Now())))
	Case "Last2Day" 
		lastDate = DateToStr(DateAdd("d",-2,DateAdd("h", strTimeAdjust , Now())))
	Case "Last3Day" 
		lastDate = DateToStr(DateAdd("d",-3,DateAdd("h", strTimeAdjust , Now())))
	Case "LastWeek" 
		lastDate = DateToStr(DateAdd("d",-7,DateAdd("h", strTimeAdjust , Now())))
	Case "LastMonth" 
		lastDate = DateToStr(DateAdd("m",-1,DateAdd("h", strTimeAdjust , Now())))
	Case Else
		lastDate = ""
End Select
%>
<!--#include file="inc_top_short.asp" -->
<%
if intSubSkin > 0 then
  intSkin = 2
end if
	if not hasAccess(1) Then
		'closeAndGo("stop")
	end if

'## Do Cookie stuffs with reload
nRefreshTime = chkString(Request.Cookies(strCookieURL & "Reload"),"sqlstring")

if Request.form("cookie") = "1" then	
    Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	Response.Cookies(strCookieURL & "Reload") = chkString(Request.Form("RefreshTime"),"sqlstring")
	Response.Cookies(strCookieURL & "Reload").expires = cdate(strCurDateAdjust) + 365
	nRefreshTime = chkString(Request.Form("RefreshTime"),"sqlstring")
end if

if nRefreshTime = "" then
	nRefreshTime = 6
end if
%>

<%
' Get Guest count for display on Default.asp
set rsGuests = Server.CreateObject("ADODB.Recordset")

if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID," & len(txtGuest) & ") = '" & txtGuest &"' "

Set rsGuests = my_Conn.Execute(strSql)
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")

spThemeTitle= "<a href=""active_users.asp"" target=""_main"">" & txtWhoOnl & "</a>"
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
	<tr>
		<td valign="middle" align="left" nowrap>
<%
        mypage = trim(chkString(request("whichpage"),"sqlstring"))

	If  mypage = "" then
	   mypage = 1
	end if

	mypagesize = 15

	If  mypagesize = "" then
	   mypagesize = 15
	end if

	set rs = Server.CreateObject("ADODB.Recordset")
	'
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated "
	strSql = strSql & " FROM " & strMemberTablePrefix & "ONLINE "
	strSql = strSql & " WHERE Right(" & strTablePrefix & "ONLINE.UserID, 5) NOT IN('" & txtGuest & "') " 
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated DESC"

	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3

	i = 0 

	If rs.EOF or rs.BOF then
		Response.Write ""
	Else

		rs.movefirst
		num = 0
		rs.pagesize = mypagesize
		maxpages = cint(rs.pagecount)
		maxrecs = cint(rs.pagesize)
		rs.absolutepage = mypage
		howmanyrecs = 0
		rec = 1
		do until rs.EOF or (rec = mypagesize+1)
			if Right(rs("UserID"), len(txtGuest)) <> "" & txtGuest & "" then 
				strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE,  " & strTablePrefix & "ONLINE.UserID "
				strSql = strSql & " FROM " & strTablePrefix & "MEMBERS, " & strTablePrefix & "ONLINE "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & rs("UserID") & "' "
				set rsMember =  my_Conn.Execute(strSql)
			end if

			if Right(rs("UserID"), len(txtGuest)) <> txtGuest then
				Response.Write("&nbsp;<a href=""cp_main.asp?cmd=8&member=" & rsMember("MEMBER_ID") & """ target=""_main"">")
				Response.Write(rs("UserID") & "</a> ")
			  if chkApp("PM","USERS") and rsMember("M_PMSTATUS") = 1 and rsMember("M_PMRECEIVE") = 1 then
				  Response.Write("<a href=""Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=" & rsMember("MEMBER_ID") & "');"">")
				  Response.Write("<img src=""images/icons/pm.gif"" border=""0"" width=""11"" height=""17"" align=""middle"" hspace=""6""></a><br />")
			  else
			    Response.Write("<br />")
			  end if
			end if
			rs.MoveNext
			rec = rec + 1
		loop
 %>
		&nbsp;<% =Guests %><% if Guests=1 then %>&nbsp;<%= txtGuest %><% else %>&nbsp;<%= txtGuests %><br /><% end if %>
<%	end if %>
		</td>
	</tr>
	<tr>
		<td valign="middle" align="left" nowrap>
		
<%      ' Get Private Message count
	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSqL = "SELECT count(M_TO) as pmcount " 
    end if
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
		strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 

	Set rsPM = my_Conn.Execute(strSql)
		pmcount = rsPM("pmcount")
%>


<%
if chkApp("PM","USERS") then
	if strDBNTUserName = "" Then %>
		
<%	else
		if pmcount = 0 then %>
		<br /><A HREF="pm.asp" target="_main"><IMG SRC="images/icons/icon_pm.gif" align=absmiddle border=0 hspace=6></a>(<% =pmcount %>)&nbsp;<%= txtNew %>&nbsp;<A HREF="pm.asp" target="_main"><%= txtMsgs %></a>.
<%		end if
        if pmcount >= 1 then %>
        <EMBED SRC="themes/<%= strTheme %>/newpm.wav" WIDTH="1" HEIGHT="1" HIDDEN="true" AUTOSTART="true" LOOP="false" volume="100"></EMBED>
<%
      	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
      	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
      	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
      	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
      	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

      	Set rsMessage = my_Conn.Execute(strSql)


      	i = 0
      		do Until rsMessage.EOF or (i = 3)

	  	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID  "
	  	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	  	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsMessage("M_FROM") & ""

          	Set rsFrom = my_Conn.Execute(strSql)	
          	if rsMessage("M_READ") = "0" then %>
			<script type="text/javascript">
			//openWindowPM('<%= strHomeURL %>pm_pop.asp?mode=1&cid=<% =rsMessage("M_ID") %>');
			</script>
<%        		i = i + 1
		end if
	    	rsMessage.MoveNext
	    	
       		Loop

%>
        <A HREF="pm.asp" target="_main"><IMG SRC="images/icons/icon_pm_new.gif" align="absmiddle" border="0" hspace="6"></a>(<b><% =pmcount %></b>)&nbsp;<%= txtNew %>&nbsp;<A HREF="pm.asp" target="_main"><% if pmcount = 1 then %><%= txtMsg %><% else %><%= txtMsgs %><% end if %></a>.
<%		end if 
	end if
end if 'PM app check %>	
		</td>
	</tr>
	<tr>
	<form name="ReloadFrm" action="site_monitor.asp" method="post"> 	
		<td height="10" valign="middle" align="right" nowrap>
    			<select name="RefreshTime" size="1" onchange="autoReload();" style="font-size:10px;">
        			<option value="0"  <% if nRefreshTime = "0" then Response.Write(" SELECTED")%>><%= txtNoRef %></option>
        			<option value="3"  <% if nRefreshTime = "3" then Response.Write(" SELECTED")%>><%= txtRef30s %></option>
        			<option value="4.5"  <% if nRefreshTime = "4.5" then Response.Write(" SELECTED")%>><%= txtRef45s %></option>
        			<option value="6" <% if nRefreshTime = "6" then Response.Write(" SELECTED")%>><%= txtRef1m %></option>
        			<option value="12" <% if nRefreshTime = "12" then Response.Write(" SELECTED")%>><%= txtRef2m %></option>
        			<option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>><%= txtRef5m %></option>
    			</select>
    		</td>	
	<input type="hidden" name="Cookie" value="1">
	</form>
    	</tr>
<% if maxpages > 1 then %>    	
    	<tr>
		<td valign="middle" align="left"><b><% Call Paging() %></b></td>
  	</tr>
<% end if
Response.Write("</table>")
spThemeBlock1_close(intSkin)%>
<%
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = chkString(Request.QueryString("whichpage"),"sqlstring")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 width=95% cellspacing=0 cellpadding=1 align=top><tr><td align=left>" & txtPages & ": </td>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then
				ref = "<td align=right>" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname
				ref = ref & "?whichpage=" & counter
				ref = ref & "&pagesize=" & mypagesize
				if top = "1" then
					ref = ref & "'><b>" & counter & "</b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=right>" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></td>")
			end if
			if counter mod 15 = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub 
%>

<%
'##############################################################################
 '----------------- start active topics -------------------
'##############################################################################
if chkApp("forums","USERS") then
if IsEmpty(Session(strUniqueID & "last_here_date")) then
	Session(strUniqueID & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if
if lastDate = "" then
	lastDate = Session(strUniqueID & "last_here_date")
end if
if Request.Form("AllRead") = "Y" then
	Session(strUniqueID & "last_here_date") = ReadLastHereDate(strDBNTUserName)
	Session(strUniqueID & "last_here_date") = ReadLastHereDate(strDBNTUserName)
	lastDate = Session(strUniqueID & "last_here_date")
	ActiveSince = ""
end if

' - Get all active topics from last visit
strSql = "SELECT " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS, " 
strSql = strSql & strTablePrefix & "TOPICS.T_VIEW_COUNT, " & strTablePrefix & "TOPICS.FORUM_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.CAT_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_MAIL, " 
strSql = strSql & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_POLL, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, "
strSql = strSql & strTablePrefix & "TOPICS.T_NEWS, "
strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "
strSql = strSql & "FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "FORUM, "
strSql = strSql & strTablePrefix & "TOPICS, " & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & "WHERE " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS_1.MEMBER_ID  "
strSql = strSql & "AND " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
strSql = strSql & "AND " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "TOPICS.CAT_ID "
strSql = strSql & "AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & "AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

set rs = my_Conn.Execute (strSql)
%>
<table width="180" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="LastDateFrm" action="site_monitor.asp" method="post">
  <tr>
    <td align="left">
	&nbsp;Last visit:<br />&nbsp;<%= ChkDate(Session(strUniqueID & "last_here_date")) %> @ <% =ChkTime(Session(strUniqueID & "last_here_date"))%><br />
              &nbsp;<a href="forum_active_topics.asp" target="_main"><%= txtActvTopics %></a>&nbsp;<%= txtSince %>: 
              <select name="ShowSinceDateTime" size="1" onchange="SetLastDate();">
                <option value="LastVisit" <% if ActiveSince = "LastVisit" or ActiveSince = "" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLstVst %>&nbsp;</option>
                <option value="LastHour"  <% if ActiveSince = "LastHour" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLstHr %></option>
                <option value="Lastthree"  <% if ActiveSince = "Lastthree" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLst3Hr %></option>
                <option value="Lastsix"  <% if ActiveSince = "Lastsix" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLst6Hr %></option>
                <option value="Lasttwelve"  <% if ActiveSince = "Lasttwelve" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLst12Hr %></option>
                <option value="LastDay"   <% if ActiveSince = "LastDay" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLstDy %></option>
				<option value="Last2Day"   <% if ActiveSince = "Last2Day" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLst2Dy %></option>
				<option value="Last3Day"   <% if ActiveSince = "Last3Day" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLst3Dy %></option>
                <option value="LastWeek"  <% if ActiveSince = "LastWeek" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLstWk %></option>
                <option value="LastMonth" <% if ActiveSince = "LastMonth" then Response.Write(" SELECTED")%>>&nbsp;<%= txtLstMn %></option>
              </select>
    <input type="hidden" name="Cookie" value="2"><br />
    
    </td>
  </tr></form>
</table>
<%
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")%>
      <tr>
        <td valign="top" align="center" width="20" class="fSubTitle">
        <%If not(rs.EOF or rs.BOF) then %>
			<form name="MarkRead" action="site_monitor.asp" method="post">
			<input type="hidden" name="AllRead" value="Y">
			<input type="image" src="images/icons/icon_topic_all_read.gif" value="<%= txtMkAllRead %>" id="submit1" name="submit1" alt="<%= txtMkAllRead %>" height="20" width="20" border="0" hspace="0" onclick></form>
        <% else %>
			&nbsp;
        <% end if %>
		</td>
        <td valign="top" align="center" class="fSubTitle"><b><a href="forum_active_topics.asp" target="_main"><%= txtActvTopics %></a></b></td>
      </tr>
<%If rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="2"><b><%= txtNoTopFnd %></b></td>
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
		if chkForumAccess(strUserMemberID,rs("FORUM_ID")) then
			fDisplayCount = fDisplayCount + 1
			if currForum <> rs("FORUM_ID") then %>
				<tr>
					<td height="20" colspan="2" class="tAltSubTitle" valign="top" ><a href="<% Response.Write("FORUM.asp?FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID") & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")) %>" target="_main"><b><% =ChkString(rs("F_SUBJECT"),"display") %></b></a></td>
				</tr>
<%			end if %>
			<tr>
			<%
			if IsNull(rs("T_LAST_POST_AUTHOR")) then
				strLastAuthor = ""
			else
				strLastAuthor = txtLstRplyBy & ": " 
				strLastAuthor = strLastAuthor & "<a href=""cp_main.asp?cmd=8&member="& rs("T_LAST_POST_AUTHOR") & """  target=""_main"">"
				strLastAuthor = strLastAuthor & rs("LAST_POST_AUTHOR_NAME") & "</a>"
			end if
			%>
			<td colspan="2" valign="center"><b><a title="<%= txtReadAll %>" href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>" target="_main"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a></b><% if rs("T_NEWS") = 1 then%>&nbsp;<img src="images/icons/icon_topic_news.gif"><% end if %><% if rs("T_POLL") <> 0 then %>&nbsp;<img src="images/icons/icon_topic_poll.gif"><% end if %><br />
			<%=strLastAuthor%><br />
			<%= txtOn %>: <% =ChkDate(rs("T_LAST_POST")) %>&nbsp;<% =ChkTime(rs("T_LAST_POST"))  %>
			<% if rs("T_REPLIES") <> 0 then %>
			<br />[<a href="link.asp?TOPIC_ID=<%= rs("TOPIC_ID") %>&view=lasttopic" target="_main"><%= txtViewRply %>&nbsp;(<% =rs("T_REPLIES") %>)</a>]
			<% End If %>
			</td>
			</tr>
			<tr>
			<td colspan="2" height="3"></td>
			</tr>
			<tr>
			<td colspan="2" height="1"></td>
			</tr>
			<tr>
			<td colspan="2" height="4"></td>
			</tr>	
<%		end if	
		currForum = rs("FORUM_ID") %>
<%		rs.MoveNext 
	loop 
	if fDisplayCount = 0 then %>
		  <tr>
		 <td colspan="2"><b><%= txtNoTopFnd %></b></td></tr>
<%
	end if 
 end if
Response.Write("</table>")
spThemeBlock1_close(intSkin)

end if
'----------------- end active topics ------------------- %>

<SCRIPT>
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 5000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 10000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</SCRIPT>
    </td>
  </tr>
</table>
</BODY>
</HTML>
<% 
my_Conn.Close
set my_Conn = nothing

sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=0 cellspacing=0 cellpadding=0><tr><td valign=""center""><img src=""images/icons/icon_posticon.gif"" border=""0""></td>")
		for counter = 1 to mxpages
			ref = "<td align=right valign=""center"">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<a href='forum_topic.asp?"
            		ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
		        ref = ref & "&FORUM_ID=" & rs("FORUM_ID")
		        ref = ref & "&CAT_ID=" & rs("CAT_ID")
		        ref = ref & "&Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
		        ref = ref & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")
			ref = ref & "&whichpage=" & counter
			ref = ref & "'>" & counter & "</a></td>"
			Response.Write ref 
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr><td>&nbsp;</td>")
			end if
		next				
        Response.Write("</tr></table>")
	end if
end sub
function chkForumModerator(fForum_ID, fMember_Name)
	strSql = "SELECT * FROM " & strTablePrefix & "MODERATOR"
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID
	strSql = strSql & " AND MEMBER_ID = " & strUserMemberID
	set rsChk = my_Conn.Execute (strSql)
	if rsChk.eof then
		chkForumModerator = "0"
	else
		chkForumModerator = "1"
	end if 
	set rsChk = nothing
end function %>

