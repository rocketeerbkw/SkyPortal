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
CurPageType = "forums"
%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<%
if Request.QueryString("FORUM_ID") = "" then
	Response.Redirect "default.asp"
end if
if Request.QueryString("FORUM_ID") <> "" or  Request.QueryString("FORUM_ID") <> " " then
	if IsNumeric(Request.QueryString("FORUM_ID")) = True then
		strRqForumID = cLng(Request.QueryString("FORUM_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.QueryString("CAT_ID") <> "" or Request.QueryString("CAT_ID") <> " " then
	if IsNumeric(Request.QueryString("CAT_ID")) = True then
		strRqCatID = cLng(Request.QueryString("CAT_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if
if Request.QueryString("TOPIC_ID") <> "" or Request.QueryString("TOPIC_ID") <> " " then
	if IsNumeric(Request.QueryString("TOPIC_ID")) = True then
		strRqTopicID = cLng(Request.QueryString("TOPIC_ID"))
	else
		Response.Redirect("default.asp")
	end if
end if 

mypage = chkString(request("whichpage"),"numeric")

if mypage = "" then
	mypage = 1
end if

nDays = chkString(Request.Cookies(strCookieURL & "NumDays"),"numeric")

if Request.form("cookie") = 1 then
	Response.Cookies(strCookieURL & "NumDays").Path = strCookieURL
	Response.Cookies(strCookieURL & "NumDays") = chkString(Request.Form("days"),"numeric")
	Response.Cookies(strCookieURL & "NumDays").expires = dateAdd("d", 360, now())
	nDays = chkString(Request.Form("Days"),"numeric")
	mypage = 1
end if

if nDays = "" then
	nDays = 30
end if

defDate = datetostr(dateadd("d", -(nDays), now()))
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<%
if strPrivateForums = "1" then
	if Request("Method_Type") = "" and (not hasAccess(1)) then
		chkUser4()
	end if
end if


if (hasAccess(1)) or (chkForumModerator(strRqForumID, STRdbntUserName)= "1") or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if


' - Get all topics from DB
strSql ="SELECT " & strTablePrefix & "ARCHIVE_TOPICS.T_STATUS, " & strTablePrefix & "ARCHIVE_TOPICS.CAT_ID, " 
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.FORUM_ID, " & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID, " 
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.T_VIEW_COUNT, " & strTablePrefix & "ARCHIVE_TOPICS.T_SUBJECT, " 
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.T_MAIL, " & strTablePrefix & "ARCHIVE_TOPICS.T_AUTHOR, " 
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.T_REPLIES, " & strTablePrefix & "ARCHIVE_TOPICS.T_LAST_POST, "
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS.T_LAST_POST_AUTHOR, "  
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, "
strSql = strSql & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, "
strSql = strSql & strTablePrefix & "ARCHIVE_TOPICS, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "ARCHIVE_TOPICS.T_AUTHOR "
strSql = strSql & " AND " & strTablePrefix & "ARCHIVE_TOPICS.T_LAST_POST_AUTHOR = "& strMemberTablePrefix & "MEMBERS_1.MEMBER_ID "
strSql = strSql & " AND " & strTablePrefix & "ARCHIVE_TOPICS.FORUM_ID = " & strRqForumID & " "
if nDays = "-1" then
	strSql = strSql & " AND   " & strTablePrefix & "ARCHIVE_TOPICS.T_STATUS <> 0 "
end if
if nDays > "0" then
	strSql = strSql & " AND " & strTablePrefix & "ARCHIVE_TOPICS.T_LAST_POST > '" & defDate & "'"
end if
strSql = strSql & " ORDER BY " & strTablePrefix & "ARCHIVE_TOPICS.T_LAST_POST DESC "

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		intOffset = CInt((mypage-1) * strPageSize)
		strSql = strSql & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	' - Get the total pagecount 
	strSql2 = "SELECT COUNT(" & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strTablePrefix & "ARCHIVE_TOPICS " 
	strSql2 = strSql2 & " WHERE   " & strTablePrefix & "ARCHIVE_TOPICS.TOPIC_ID > 0 " 
	if nDays = "-1" then
		strSql2 = strSql2 & " AND   " & strTablePrefix & "TOPICS.T_STATUS <> 0 "
	end if
	if nDays > "0" then
		strSql2 = strSql2 & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & defDate & "'"
	end if

	set rsCount = my_Conn.Execute(strSql2)
	if not rsCount.eof then
		maxpages = (rsCount("PAGECOUNT") \ strPageSize )
			if rsCount("PAGECOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
	else
		maxpages = 0
	end if 

	rsCount.close
	
	set rs = Server.CreateObject("ADODB.Recordset")
'	rs.cachesize=20

	rs.open  strSql, my_Conn, 3
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
	end if
 
else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize=20

	rs.open  strSql, my_Conn, 3
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
		rs.pagesize = strPageSize
		maxpages = cint(rs.pagecount)
		rs.absolutepage = mypage
	end if
	
end if

' - Get all Forum Categories From DB
strSql = "SELECT CAT_ID FROM " & strTablePrefix & "CATEGORY"

set rsCat = my_Conn.Execute (StrSql)

%>
<script type="text/javascript">
<!----- 
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}

function setDays() {document.DaysFilter.submit(); return 0;}
// -->
</script>
<table border="0" width="95%" cellpadding="0" cellspacing="0">
  <tr class="breadcrumb">
    <td width="33%" align="left" nowrap>
    <a href="fhome.asp"><img src="images/icons/icon_folder_open.gif" alt="All Forums" height=15 width=15 border="0"></a>&nbsp;<a href="fhome.asp">All Forums</a><br />
    <img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_closed_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("FORUM_Title"),"sqlstring") %>
    </td>
    <td align="center" width="33%">&nbsp;
</td>
    <td align="center" width="33%">
    <form action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & ChkString(Request.Querystring,"SQLString")  %>" method="post" name="DaysFilter">
    <select name="Days" onchange="javascript:setDays();">
      <option value="0" <% if ndays = "0" then Response.Write(" SELECTED")%>>Show all topics</option>
      <option value="-1" <% if ndays = "-1" then Response.Write(" SELECTED")%>>Show all open topics</option>
      <option value="1" <% if ndays = "1" then Response.Write(" SELECTED")%>>Show topics from last day</option>
      <option value="2" <% if ndays = "2" then Response.Write(" SELECTED")%>>Show topics from last 2 days</option>
      <option value="5" <% if ndays = "5" then Response.Write(" SELECTED")%>>Show topics from last 5 days</option>
      <option value="7" <% if ndays = "7" then Response.Write(" SELECTED")%>>Show topics from last 7 days</option>
      <option value="14" <% if ndays = "14" then Response.Write(" SELECTED")%>>Show topics from last 14 days</option>
      <option value="30" <% if ndays = "30" then Response.Write(" SELECTED")%>>Show topics from last 30 days</option>
      <option value="60" <% if ndays = "60" then Response.Write(" SELECTED")%>>Show topics from last 60 days</option>
      <option value="120" <% if ndays = "120" then Response.Write(" SELECTED")%>>Show topics from last 120 days</option>
      <option value="365" <% if ndays = "365" then Response.Write(" SELECTED")%>>Show topics from the last year</option>
    </select>
    <input type="hidden" name="Cookie" value="1">
   </form>
    </td>
  </tr>
  <tr>
	<td colspan=2>
	</td>
     <td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><b>Pages:</b> &nbsp;</td>
        <td valign="top"><% Call Paging2() %></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
    </td>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" class="tTitle"><b><span class="fSubTitle">&nbsp;</span></b></td>
        <td align="center" class="tTitle"><b><span class="fSubTitle">Topic</span></b></td>
        <td align="center" class="tTitle"><b><span class="fSubTitle">Author</span></b></td>
        <td align="center" class="tTitle"><b><span class="fSubTitle">Replies</span></b></td>
        <td align="center" class="tTitle"><b><span class="fSubTitle">Read</span></b></td>
        <td align="center" class="tTitle"><b><span class="fSubTitle">Last Post</span></b></td>
<% if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
        <td align="center" class="tTitle"><b><span class="fSubTitle"><%	call ForumAdminOptions() %></span></b></td>
<% end if %>
      </tr>
<% if rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="7" class="tCellAlt1"><span class="fTitle"><b>No Topics Found</b></span></td>
      </tr>
<% else
	rec = 1
	do until rs.EOF or (rec = strPageSize + 1) %>
      <tr>
        <td class="tCellAlt1" align=center valign="center">
			<% if rs("T_LAST_POST") > Session(strUniqueID & "last_here_date") then %>
					<a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"sqlstring") %>">
<%					Response.Write "<img src=""images/icons/icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""Topic Locked"">"
				else %>
					<a href="forum_archive_display.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"sqlstring") %>">
<%					Response.Write "<img src=""images/icons/icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""Topic Locked"">"
				end if%></a></td>
        <td class="tCellAlt1" valign="center" align="left"><a href="forum_archive_display.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a>&nbsp;<% if strShowPaging = "1" then TopicPaging() end if%></td>
        <td class="tCellAlt1" valign="center" align="center"><% =ChkString(rs("M_NAME"),"display") %></td>
        <td class="tCellAlt1" valign="center" align="center"><% =rs("T_REPLIES") %></td>
        <td class="tCellAlt1" valign="center" align="center"><% =rs("T_VIEW_COUNT") %></td>
        <%
        if IsNull(rs("T_LAST_POST_AUTHOR")) then
            strLastAuthor = ""
        else
            strLastAuthor = "<br />by: " 
            if strUseExtendedProfile then
				strLastAuthor = strLastAuthor & "<a href=""cp_main.asp?cmd=8&member="& ChkString(rs("T_LAST_POST_AUTHOR"), "JSurlpath") & """>"
			else
				 strLastAuthor = strLastAuthor & "<a href=""JavaScript:openWindow2('cp_main.asp?cmd=8&member=" & ChkString(rs("T_LAST_POST_AUTHOR"), "JSurlpath") & "')"">"
			end if
            strLastAuthor = strLastAuthor & ChkString(rs("LAST_POST_AUTHOR_NAME"), "display") & "</a>"
        end if
        %>
        <td class="tCellAlt1" valign="center" align="center" nowrap><span class="fSmall"><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></span></td>
<% 		if AdminAllowed = 1 or strNoCookies = "1" then %>
        <td class="tCellAlt1" valign="center" align="center" nowrap>&nbsp;

        </td>
<%		end if %>
      </tr>
<%		rec = rec + 1 
		rs.MoveNext 
	loop 
 end if %>
    </table>
    </td>
  </tr>
  <tr>
  <td colspan=5>
  <% if maxpages > 1 then %>
    <table border=0 align="left">
      <tr>
        <td valign="top"><b>There are <% =maxpages %> Pages of Topics: &nbsp;&nbsp; </b></td>
        <td valign="top"><% Call Paging() %></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
	</td>
	</tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top" width="33%">
    <table>
      <tr>
        <td>
		<p> 
		<img alt="New Posts" src="images/icons/icon_folder_new.gif" width="8" height="9"> New posts since last logon.<br />
		<img alt="Old Posts" src="images/icons/icon_folder.gif" width="8" height="9"> Active topic. <% if lcase(strHotTopic) = "1" then %>(<img alt="Hot Topic" src="images/icons/icon_folder_hot.gif" width="8" height="9"> <% =intHotTopicNum %> replies or more.)<% end if %><br />
		<img alt="Locked Topic" src="images/icons/icon_folder_locked.gif" width="8" height="9"> Locked topic.<br />
		</p>
	    </td>
	  </tr>
	</table>
    </td>
    <td align="center" valign="top" width="33%">&nbsp;

    </td>
    <td align="center" valign="top" width="33%">
<!--#INCLUDE file="modules/forums/inc_jump_to.asp" -->
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
Function ChkIsNew(dt)
	if lcase(strHotTopic) = "1" then
		if dt > Session(strUniqueID & "last_here_date") then
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src='images/icons/icon_folder_new_hot.gif' height=15 width=15 border=0 hspace=0 alt='Hot Topic'>"
			else
			        ChkIsNew =  "<img src='images/icons/icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Topic'>"
			end if
		Else
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src='images/icons/icon_folder_hot.gif' height=15 width=15 border=0 hspace=0 alt='Hot Topic'>"
			else
			        ChkIsNew = "<img src='images/icons/icon_folder.gif' height=15 width=15 border=0 hspace=0>" 
			end if
		end if
	else
		if dt > Session(strUniqueID & "last_here_date") then
			ChkIsNew =  "<img src='images/icons/icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Topic'>" 
		Else
			ChkIsNew = "<img src='images/icons/icon_folder.gif' height=15 width=15 border=0 hspace=0>" 
		end if
	end if
End Function

sub PostNewTopic() 
%>
&nbsp;
<%
end sub

sub ForumAdminOptions() 
%>
&nbsp;
<%
end sub

sub Paging()
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 cellspacing=0 cellpadding=1 valign=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then   
				ref = "<td align=right>" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname 
				ref = ref & "?FORUM_ID=" & strRqForumID 
				ref = ref & "&CAT_ID=" & strRqCatID
				ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_TITLE"),"urlpath") 
				ref = ref & "&whichpage=" & counter
				ref = ref & "'>" & counter & "</a></td>"
				Response.Write ref 
			else
				Response.Write("<td align=right>" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></td>")
			end if
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
end sub

sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=0 cellspacing=0 cellpadding=0><tr><td valign=""center""><img src=""images/icons/icon_posticon.gif"" border=""0""></td>")
		for counter = 1 to mxpages
			ref = "<td align=right class=""tCellAlt1"">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<a href='forum_archive_display.asp?"
            ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
            ref = ref & "&FORUM_ID=" & rs("FORUM_ID")
            ref = ref & "&CAT_ID=" & rs("CAT_ID")
            ref = ref & "&Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
            ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_Title"),"urlpath")
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
sub Paging2()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = chkString(Request.QueryString("whichpage"),"numeric")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = chkString(Request.QueryString("method"),"sqlstring")
		end if

		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<form name=""PageNum"" action=""topic_stats.asp"">")
		Response.Write("<input type=""hidden"" name=""Topic_Title"" value=""" & ChkString(Request.QueryString("Topic_Title"),"sqlstring") & """>")
		Response.Write("<input type=""hidden"" name=""Forum_Title"" value=""" & ChkString(Request.QueryString("FORUM_Title"),"sqlstring") & """>")
		Response.Write("<input type=""hidden"" name=""CAT_ID"" value=""" & strRqCatID & """>")
		Response.Write("<input type=""hidden"" name=""FORUM_ID"" value=""" & strRqForumID & """>")
		Response.Write("<input type=""hidden"" name=""TOPIC_ID"" value=""" & strRqTopicID & """>")
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


%>