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

if Request.QueryString("TOPIC_ID") = "" and Request.QueryString("mode") <> "getIP" and Request.Form("Method_Type") <> "login" and Request.Form("Method_Type") <> "logout" then
	Response.Redirect "fhome.asp"
	Response.End
end if

if Request.QueryString("FORUM_ID") <> "" or  Request.QueryString("FORUM_ID") <> " " then
	if IsNumeric(Request.QueryString("FORUM_ID")) = True then
		strRqForumID = cLng(Request.QueryString("FORUM_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.QueryString("CAT_ID") <> "" or Request.QueryString("CAT_ID") <> " " then
	if IsNumeric(Request.QueryString("CAT_ID")) = True then
		strRqCatID = cLng(Request.QueryString("CAT_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.QueryString("TOPIC_ID") <> "" or Request.QueryString("TOPIC_ID") <> " " then
	if IsNumeric(Request.QueryString("TOPIC_ID")) = True then
		strRqTopicID = cLng(Request.QueryString("TOPIC_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.QueryString("REPLY_ID") <> "" or Request.QueryString("REPLY_ID") <> " " then
	if IsNumeric(Request.QueryString("REPLY_ID")) = True then
		strRqReplyID = cLng(Request.QueryString("REPLY_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if 
%>
<!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<%

if (strAuthType = "nt") then
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

%>
<!--#INCLUDE FILE="inc_top.asp" -->
<% 
Member_ID = strUserMemberID
	intSkin = getSkin(intSubSkin,2)

' - Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID

set rsCStatus = my_Conn.Execute (StrSql)

' - Find out if the Forum is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS "
strSql = strSql & " FROM " & strTablePrefix & "FORUM "
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumID

set rsFStatus = my_Conn.Execute (StrSql)
strPrivateForums =rsFStatus("F_PRIVATEFORUMS")

' - Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "TOPICS.T_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & strRqTopicID

set rsTStatus = my_Conn.Execute (StrSql)

if rsCStatus.EOF or rsCStatus.BOF or rsFStatus.EOF or rsFStatus.BOF or rsTStatus.EOF or rsTStatus.BOF then
	TopicOpen = false
else
	TopicOpen = true
	mypage = request("whichpage")
end if

if mypage = "" then
	mypage = 1
end if

if strPrivateForums = "1" then
	if Request("Method_Type") = "" then
		chkUser4()
	end if
end if

if (hasAccess(1)) or (chkForumModerator(strRqForumID, ChkString(STRdbntUserName, "decode"))= "1") or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if

'
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strTablePrefix & "ARCHIVE_REPLY.REPLY_ID, " & strTablePrefix & "ARCHIVE_REPLY.R_AUTHOR, " & strTablePrefix & "ARCHIVE_REPLY.TOPIC_ID, " & strTablePrefix & "ARCHIVE_REPLY.R_MESSAGE, " & strTablePrefix & "ARCHIVE_REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "ARCHIVE_REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "ARCHIVE_REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & strRqTopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ARCHIVE_REPLY.R_DATE"

	if strDBType = "mysql" then 'MySql specific code

		' - Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "ARCHIVE_REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "ARCHIVE_REPLY "
		strSql2 = strSql2 & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "ARCHIVE_REPLY.R_AUTHOR "
		strSql2 = strSql2 & " AND   TOPIC_ID = " & strRqTopicID & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 
	
		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize= strPageSize

		rs.open  strSql,  my_Conn, 3
		if not(rs.EOF) then
			rs.movefirst
		end if
		
	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")

		rs.cachesize = strPageSize
		rs.open  strSql,  my_Conn, 3

		If not (rs.EOF or rs.BOF) then  '## No replies found in DB
			rs.movefirst
			rs.pagesize = strPageSize
			rs.absolutepage = mypage '**
			maxpages = cint(rs.pagecount)
		end if
	end if
	i = 0 
 %>
<table border="0" width="95%">
  <tr class="breadcrumb">
	<td width="50%" align="left" nowrap>
	<img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
	<img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="forum_archive.asp?FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">Forum Archives</a><br />
	<img src="images/icons/icon_blank.gif" height=15 width=15 border="0"><img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("Topic_Title"),"display") %>
    </td>
    <td align="center" width="50%">&nbsp;</td>
  </tr>
  <tr>
<td align="right" colspan=2>
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><b>Pages: </b></td>
        <td valign="top"><% Call Paging2() %></td>
      </tr>
    </table>
<% else %>
	<td align=right>&nbsp;</td>
    &nbsp;
<% end if %>

</td>
  </tr>
</table>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" class="tTitle" width="120" nowrap><b>Author</b></td>
        <td align="center" class="tTitle" width="100%"><b><% Call Topic_nav() %></b></td>
<%	if (AdminAllowed = 1) then %>
        <td align=right class="tTitle" colspan=2 nowrap><% call AdminOptions() %></td>
<%	else %>
        <td align=right class="tTitle" nowrap>&nbsp;</td>
<%	end if %>
</tr>
<% 
	if mypage = 1 then 
		Call GetFirst() 
	end if
 %>
<% 
	' - Get all topicsFrom DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_ICQ, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_Homepage, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_REPLY.REPLY_ID, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_REPLY.R_AUTHOR, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_REPLY.TOPIC_ID, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_REPLY.R_MESSAGE, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "ARCHIVE_REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "ARCHIVE_REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & strRqTopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ARCHIVE_REPLY.R_DATE"
	
	if strDBType = "mysql" then 'MySql specific code
		if mypage > 1 then 
			intOffSet = CInt((mypage - 1) * strPageSize) - 1
			strSql = strSql & " LIMIT " & intOffSet & ", " & CInt(strPageSize) & " "
		end if

		' - Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "ARCHIVE_REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strTablePrefix & "ARCHIVE_REPLY "
		strSql2 = strSql2 & " WHERE  TOPIC_ID = " & strRqTopicID & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 

		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize = strPageSize

		rs.open  strSql,  my_Conn, 3

	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.cachesize = 20
		'response.write strSQL
		rs.open  strSql,  my_Conn, 3
	
		if not(rs.EOF or rs.BOF) then  '## Replies found in DB
			rs.movefirst
			rs.pagesize = strPageSize
			maxpages = cint(rs.pagecount)
			rs.absolutepage = mypage
		end if
	end if		
	if rs.EOF or rs.BOF then  '## No replies found in DB
		Response.Write ""
	else
		'rs.movefirst			
		intI = 0 
		howmanyrecs = 0
		rec = 1
	
		do until rs.EOF or (mypage = 1 and rec > strPageSize) or (mypage > 1 and rec > strPageSize) '**		
			if intI = 0 then 
				CColor = "tCellAlt1"
			else
				CColor = "tCellAlt2"
			end if
 %>
 <tr>
        <td class="<% =CColor %>" valign="top">
        <% if strUseExtendedProfile then %>
		<a href="cp_main.asp?cmd=8&member=<% =rs("R_AUTHOR") %>">
        <% else %>
		<a href="JavaScript:openWindow3('cp_main.asp?cmd=8&member=<% =rs("R_AUTHOR") %>')">
		<% end if %>	
		<b><% =ChkString(rs("M_NAME"),"display") %></a>
        </b>
<%				if strShowRank = 1 or strShowRank = 3 then %>
        <br /><span class="fSmall"><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></span>
<%				end if %>
<%				if strShowRank = 2 or strShowRank = 3 then %>
        <br /><% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %>
<%				end if %>
        <br />
        <br /><span class="fSmall"><% =rs("M_COUNTRY") %></span>
        <br /><span class="fSmall"><% =rs("M_POSTS") %> Posts</span></td>
        <td class="<% =CColor %>" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top"><img src="images/icons/icon_posticon.gif" border="0" hspace="3"><span class="fSmall">Posted&nbsp;-&nbsp;<% =ChkDate(rs("R_DATE")) %>&nbsp;:&nbsp;<% =ChkTime(rs("R_DATE")) %></span>
        <% if strUseExtendedProfile then %>
		&nbsp;<a href="cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>"><img src="images/icons/icon_profile.gif" height=15 width=15 alt="Show Profile" border="0" align="absmiddle" hspace="6"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_profile.gif" height=15 width=15 alt="Show Profile" border="0" align="absmiddle" hspace="6"></a>
		<% end if %>	
   
<%		if (lcase(strEmail) = "1") then 
			if (hasAccess(2)) or (not hasAccess(2) and  strLogonForMail <> "1") then 
%>
				&nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_email.gif" height=15 width=15 alt="Email Poster" border="0" align="absmiddle" hspace="6"></a>
<%			end if
		else
%>
			&nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_email.gif" height=15 width=15 alt="Email Poster" border="0" align="absmiddle" hspace="6"></a>
<%		end if %>  
<%			if strHomepage = "1" then %>
<%				if rs("M_Homepage") <> " " then %>
        &nbsp;<a href="<% =rs("M_Homepage") %>"><img src="images/icons/icon_homepage.gif" height=15 width=15 alt="Visit <% = ChkString(rs("M_NAME"),"display") %>'s Homepage" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<%			if strICQ = "1" then %>
<%			  if Trim(rs("M_ICQ")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=1&ICQ=<% =ChkString(rs("M_ICQ"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"JSurlpath") %>')"><img src="images/icons/icon_icq.gif" height=15 width=15 alt="Send <% = ChkString(rs("M_NAME"),"display")  %> an ICQ Message" border="0" align="absmiddle" hspace="6"></a>
<%			  end if %>
<%			end if %>
<%			if strYAHOO = "1" then %>
<%			  if Trim(rs("M_YAHOO")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rs("M_YAHOO"), "JSurlpath") %>&.src=pg')"><img src="images/icons/icon_yahoo.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display")  %> a Yahoo! Message" border="0" align="absmiddle" hspace="6"></a>
<%			  end if %>
<%			end if %>
<%			if (strAIM = "1") then %>
<%				if Trim(rs("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rs("M_AIM"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="images/icons/icon_aim.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display")  %> an instant message" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<% If TopicOpen then %>
        &nbsp;<a href="forum_post.asp?method=ReplyQuote&REPLY_ID=<% =rs("REPLY_ID") %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>&M=<% =Request.QueryString("M") %>"><img src="images/icons/icon_reply_topic.gif" height=15 width=15 alt="Reply with Quote" border="0" align="absmiddle" hspace="6"></a>
<% End If %>

        <hr noshade>
        
        <% =formatStr(rs("R_MESSAGE")) %><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right" alt="Go to Top of Page"></a></td>
      </tr>
<%			
		    rs.MoveNext
		    intI  = intI + 1
		    if intI = 2 then 
				intI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>
 
 
 
 </table></td>
  </tr>
  <tr>
    <td colspan="2">
    <table border="0" width="100%">
      <tr>
        <td>
<% if maxpages > 1 then %>
        <table border=0>
          <tr>
            <td valign="top"><b>Topic is <% =maxpages %> Pages Long: </td>
            <td valign="top"><% Call Paging() %></td>
          </tr>
        </table>
<% else %>
	<td valign="top">&nbsp;</td>
<% end if %>
        </td>
        <td align="right" nowrap>&nbsp;

        </td> 
      </tr>
    </table></td>
  </tr>
</table>
</div>

<table width="100%">
  <tr>
    <td align="center" valign="top" width="50%">&nbsp;</td>
    <td align="center" valign="top" width="50%"><!--#INCLUDE file="modules/forums/inc_jump_to.asp" --></td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%


sub GetFirst()

	' - Get Origional Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " 	& strMemberTablePrefix & "MEMBERS.M_ICQ, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.T_DATE, " & strTablePrefix & "ARCHIVE_TOPICS.T_SUBJECT, " & strTablePrefix & "ARCHIVE_TOPICS.T_AUTHOR, " 
	strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID, " & strTablePrefix & "ARCHIVE_TOPICS.T_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "ARCHIVE_TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "ARCHIVE_TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID = " &  strRqTopicID 

	set rs = my_Conn.Execute (strSql)

	if rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write "  <tr>" & vbCrLf
		Response.Write "    <td colspan=5>No Topics Found</td>" & vbCrLf
		Response.Write "  </tr>" & vbCrLf
	else
 %>
      <tr>
        <td class="tCellAlt0" valign="top">
        
        <% if strUseExtendedProfile then %>
		<a href="cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>">
        <% else %>
		<a href="JavaScript:openWindow3('cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>')">
		<% end if %>	
		<b><% =ChkString(rs("M_NAME"),"display") %></a>
        </b>
<%		if strShowRank = 1 or strShowRank = 3 then %>
        <br /><span class="fSmall"><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></span>
<%		end if %>
<%		if strShowRank = 2 or strShowRank = 3 then %>
        <br /><% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %>
<%		end if %>
        <br />
        <br /><span class="fSmall"><% =rs("M_COUNTRY") %></span>
        <br /><span class="fSmall"><% =rs("M_POSTS") %> Posts</span></td>
        <td class="tCellAlt0" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top"><img src="images/icons/icon_posticon.gif" border="0" hspace="3"><span class="fSmall">Posted&nbsp;-&nbsp;<% =ChkDate(rs("T_DATE")) %>&nbsp;:&nbsp;<% =ChkTime(rs("T_DATE")) %></span>

        <% if strUseExtendedProfile then %>
		&nbsp;<a href="cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>"><img src="images/icons/icon_profile.gif" height=15 width=15 alt="Show Profile" border="0" align="absmiddle" hspace="6"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_profile.gif" height=15 width=15 alt="Show Profile" border="0" align="absmiddle" hspace="6"></a>
		<% end if %>	
        
<%		if (lcase(strEmail) = "1") then 
			if (hasAccess(2)) or (not hasAccess(2) and  strLogonForMail <> "1") then 
%>
				&nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_email.gif" height=15 width=15 alt="Email Poster" border="0" align="absmiddle" hspace="6"></a>
<%			end if
		else
%>
			&nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="images/icons/icon_email.gif" height=15 width=15 alt="Email Poster" border="0" align="absmiddle" hspace="6"></a>
<%		end if %>        
<%		if (strHomepage = "1") then %>
<%				if rs("M_Homepage") <> " " then %>
        &nbsp;<a href="<% =rs("M_Homepage") %>"><img src="images/icons/icon_homepage.gif" height=15 width=15 alt="Visit <% =ChkString(rs("M_NAME"),"display")  %>'s Homepage" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%		end if %>
<%		if (strICQ = "1") then %>
<%			if Trim(rs("M_ICQ")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=1&ICQ=<% =ChkString(rs("M_ICQ"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "JSurlpath") %>')"><img src="images/icons/icon_icq.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display")  %> an ICQ Message" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%		end if %>
<%		if (strYAHOO = "1") then %>
<%		  if Trim(rs("M_YAHOO")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rs("M_YAHOO"), "JSurlpath") %>&.src=pg')"><img src="images/icons/icon_yahoo.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display") %> a Yahoo! Message" border="0" align="absmiddle" hspace="6"></a>
<%		  end if %>
<%		end if %>
<%		if (strAIM = "1") then %>
<%			if Trim(rs("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rs("M_AIM"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "JSurlpath") %>')"><img src="images/icons/icon_aim.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display") %> an instant message" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%		end if %>
<%			if TopicOpen then

			if (cint(strPrivateForums) <> 10 and cint(strPrivateForums) <> 12) or AdminAllowed = 1 then %>
        &nbsp;<a href="forum_post.asp?method=TopicQuote&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" height=15 width=15 alt="Reply with Quote" border="0" align="absmiddle" hspace="6"></a>
<%			end if 
			end if%>

        <hr noshade size="1">
        
        <% =formatStr(rs("T_MESSAGE")) %></td>
      </tr>
<%
	end if

	'
	strSql = "UPDATE " & strTablePrefix & "ARCHIVE_TOPICS "
	strSql = strSql & " SET " & strTablePrefix & "ARCHIVE_TOPICS.T_VIEW_COUNT = (" & strTablePrefix & "ARCHIVE_TOPICS.T_VIEW_COUNT + 1) "
	strSql = strSql & " WHERE (" & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID = " & strRqTopicID & ");"

	my_conn.Execute (strSql)

	set rs = nothing

End Sub
sub DisplayIP()
	usr = (chkForumModerator(strRqForumID, STRdbntUserName))
	if hasAccess(1) then 
		usr = 1
	end if
	if usr then
		if strRqTopicID <> "" then

			'
			strSql = "SELECT " & strTablePrefix & "ARCHIVE_TOPICS.T_IP, " & strTablePrefix & "ARCHIVE_TOPICS.T_SUBJECT "
			strSql = strSql & " FROM " & strTablePrefix & "ARCHIVE_TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & strRqTopicID

			rsIP = my_Conn.Execute(strSql)

			IP = rsIP("T_IP")
			Title = rsIP("T_Subject")
		else
			if strRqReplyID <> "" then
				'
				strSql = "SELECT " & strTablePrefix & "ARCHIVE_REPLY.R_IP "
				strSql = strSql & " FROM " & strTablePrefix & "ARCHIVE_REPLY "
				strSql = strSql & " WHERE REPLY_ID = " & strRqReplyID

				rsIP = my_Conn.Execute(strSql)

				IP = rsIP("R_IP")
			end if
		end if
		set rsIP = nothing
 %>
<P align=center><b>User's IP address:</b><br />
<% =ip %></P>
<%	else %>
<p align=center><b>Only moderators and administrators can perform this action.</B></p>
<%
	end If
end sub
sub PostingOptions() 
 %>
    

    <img src="images/icons/icon_closed_topic.gif" height=15 width=15 border=0>&nbsp;Topic Locked
<%
		if (lcase(strEmail) = "1") then 
			if (hasAccess(2)) or (not hasAccess(2) and  strLogonForMail <> "1") then %>
				<br />	
				<a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<%= strRqTopicID  %>')"><img border="0" src="images/icons/icon_send_topic.gif" height=15 width=15></a>&nbsp;<a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<%= strRqTopicID  %>')">Send Topic to a Friend</a>
<%			end if
		end if %>

    
<% 
end sub 
sub AdminOptions() 
 %>
    
<%	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
    <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Topic&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="images/icons/icon_folder_delete.gif" alt="Delete Topic" height=15 width=15></a>

<%	end if %>
    

<% 
end sub 
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = Request.QueryString("whichpage")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 width=100% cellspacing=0 cellpadding=1 align=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then   
				ref = "<td align=right>" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname 
				ref = ref & "?whichpage=" & counter
				'ref = ref & "&pagesize=" & mypagesize 
				ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_Title"),"urlpath") 
				ref = ref & "&Topic_Title=" & ChkString(Request.QueryString("Topic_Title"),"urlpath")
				ref = ref & "&CAT_ID=" & strRqCatID
				ref = ref & "&FORUM_ID=" & strRqForumID 
				ref = ref & "&TOPIC_ID=" & strRqTopicID & "'"
				if top = "1" then
					ref = ref & ">"
					ref = ref & "<b>" & counter & "</b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=right>" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></td>")
			end if
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub 
sub Paging2()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = Request.QueryString("whichpage")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = Request.QueryString("method")
		end if


		Response.Write("<form name=""PageNum"" action=""forum_archive_display.asp"">")
		Response.Write("<input type=""hidden"" name=""CAT_ID"" value=""" & strRqCatID & """>")
		Response.Write("<input type=""hidden"" name=""FORUM_ID"" value=""" & strRqForumID & """>")
		Response.Write("<input type=""hidden"" name=""TOPIC_ID"" value=""" & strRqTopicID & """>")
		Response.Write("<input type=""hidden"" name=""Topic_Title"" value=""" & ChkString(Request.QueryString("Topic_Title"),"urlpath") & """>")
		Response.Write("<input type=""hidden"" name=""Forum_Title"" value=""" & ChkString(Request.QueryString("FORUM_Title"),"urlpath") & """>")

		Response.Write("<select name=""whichpage"" size=""1"" onchange=""javascript:PageNum.submit()"">")
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				Response.Write "<OPTION VALUE=""" & counter &  """>" & counter
			else
				Response.Write "<OPTION SELECTED VALUE=""" & counter &  """>" & counter
			end if
		next
		Response.Write("</select>")

	end if
end sub 

Sub Topic_nav()    
    set rsLastPost = Server.CreateObject("ADODB.Recordset")
    
    strSql = "SELECT T_LAST_POST FROM " & strTablePrefix & "ARCHIVE_TOPICS " 
    strSql = strSql & "WHERE TOPIC_ID = " & strRqTopicID
                
    set rsLastPost = my_Conn.Execute (StrSql)
    
    T_LAST_POST = rsLastPost("T_LAST_POST")
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "ARCHIVE_TOPICS "
    strSql = strSql & "WHERE T_LAST_POST > '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & strRqForumID
    strSql = strSql & " ORDER BY T_LAST_POST;"
                
    set rsPrevTopic = my_conn.Execute (strSQL)
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "ARCHIVE_TOPICS "
    strSql = strSql & "WHERE T_LAST_POST < '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & strRqForumID
    strSql = strSql & " ORDER BY T_LAST_POST DESC;"
                
    set rsNextTopic = my_conn.Execute (strSQL)
    
    if rsPrevTopic.EOF then
        prevTopic = "<img src=""images/icons/icon_blank.gif"" height=15 width=15 alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        prevTopic = "<a href=forum_archive_display.asp?cat_id=" & strRqCatID & _
                    "&FORUM_ID=" & strRqForumID & _
                    "&TOPIC_ID=" & rsPrevTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsPrevTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""images/icons/icon_topic_prev.gif"" alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
                    
    if rsNextTopic.EOF then
        NextTopic = "<img src=""images/icons/icon_blank.gif"" height=15 width=15 alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        NextTopic = "<a href=forum_archive_display.asp?cat_id=" & strRqCatID & _
                    "&FORUM_ID=" & strRqForumID & _
                    "&TOPIC_ID=" & rsNextTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsNextTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""images/icons/icon_topic_next.gif"" alt=""Next Topic"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
    
    Response.Write (prevTopic & "<b>&nbsp;Topic&nbsp;</b>" & nextTopic)
    
    rsLastPost.close
    rsPrevTopic.close
    rsNextTopic.close
    set rsLastPost = nothing
    set rsPrevTopic = nothing
    set rsNextTopic = nothing
    
end sub

 %>