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
CurPageType = "forums" %>
<!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<%
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	if Request.QueryString("mode") = "news" then
	PageName = "News Search"
	else
	PageName = "Forum Search"
	end if
	PageAction = "Viewing<br />" 
	PageLocation = "forum_search.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function

%>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<% set rs = Server.CreateObject("ADODB.Recordset")
	intSkin = getSkin(intSubSkin,2) %>
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
<% If request.QueryString("mode")="" Then %>
function focuspass() { document.forms.SearchForm.Search.focus(); }
<% End If %>
function RefreshS() {
if (document.SearchForm.news.checked) {
	window.location ="forum_search.asp?mode=news";
} else {
	window.location ="forum_search.asp";
}
}
function checklength() {
	var isOK = true;
	var tmpA = ""
	if (document.SearchForm.SearchMember.value == 0) {
	  if (document.SearchForm.Search.value.length < 3) {
	  tmpA = tmpA + 'Your search must be at least 3 characters long.\n';
	  isOK = false;
	  }
	}

 	if (!CheckName(document.forms.SearchForm.Search.value)) {
 	tmpA = tmpA + '\nSearch word(s) can not contain any of the\nfollowing characters: \\ / : ; *  \" < > |';
	document.forms.SearchForm.Search.value = "";
	isOK = false;
 	}
	
	if (!isOK){
	alert(tmpA);
	document.forms.SearchForm.Search.focus();
 	return false;
	}
}
function CheckName(str) {
	var re;
	re = /[\\\/:;*'?"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
function memberlist() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=search", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }
// done hiding -->
</script>
  <table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
<td class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1) %>
	<% menu_fp() %></td>
<td class="mainPgCol">
	<% intSkin = getSkin(intSubSkin,2)
	
  arg1 = txtForums & "|fhome.asp"
  arg2 = "Search Forum|forum_active_topics.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

If Request.QueryString("mode") = "DoIt" then
    srchWords = chkstring(Request("Search"),"sqlstring")
	if srchWords <> "" and len(srchWords) >= 3 then
		keywords = split(srchWords, " ")
		keycnt = ubound(keywords)

		' - Find all records with the search criteria in them
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_STATUS, " & strTablePrefix & "TOPICS.T_LAST_POST, " & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_NEWS, " & strTablePrefix & "TOPICS.T_POLL, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
		strSql = strSql & " FROM ((" & strTablePrefix & "FORUM LEFT JOIN " & strTablePrefix & "TOPICS "
		strSql = strSql & " ON " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY "
		strSql = strSql & " ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " ON " & strTablePrefix & "TOPICS.T_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " WHERE ("
	'################# New Search Code ######################
			if request("SearchMessage") = 1 then
				if request("andor") = "phrase" then
					'strSql = strSql & "     " & strTablePrefix & "FORUM.F_SUBJECT LIKE '%" & ChkString(request("Search"), "SQLString") & "%'"
					strSql = strSql & "(" & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & ChkString(request("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql = strSql & "(" & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & SearchWord & "%') "
						if cnt < keycnt then strSql = strSql &  chkString(request("andor"),"sqlstring")
						cnt = cnt + 1
					next
				end if
			else
				if request("andor") = "phrase" then
						strSql = strSql & "     (" & strTablePrefix & "REPLY.R_MESSAGE LIKE '%" & ChkString(request("Search"), "SQLString") & "%'"
						'strSql = strSql & " OR   " & strTablePrefix & "FORUM.F_DESCRIPTION LIKE '%" & ChkString(request("Search"), "SQLString") & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & ChkString(request("Search"), "SQLString") & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_MESSAGE LIKE '%" & ChkString(request("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql = strSql & "     (" & strTablePrefix & "REPLY.R_MESSAGE LIKE '%" & SearchWord & "%'"
						'strSql = strSql & " OR   " & strTablePrefix & "FORUM.F_DESCRIPTION LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_MESSAGE LIKE '%" & SearchWord & "%') "
						if cnt < keycnt then strSql = strSql &  chkString(request("andor"),"sqlstring")
						cnt = cnt + 1
					next
				end if
			    if request("news") = "news" then
				  strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_NEWS = " & 1
			    end if
'################# New Search Code #################################################
			end if
			strSql = strSql & " ) "
		cnt = 0
		if Request("Forum") <> 0 then
			ft = cLng(Request("Forum"))
			strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & ft & " "
		end if
		if request("SearchDate") <> 0 then
			dt = cLng(request("SearchDate"))
			strSql = strSql & " AND (T_DATE > '" & datetostr(dateadd("d", -dt, now())) & "')"
		end if
		if Request("SearchMember") <> 0 then
			strSql = strSql & " AND (" & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & chkString(Request("SearchMember"), "SQLString") & " "
			strSql = strSql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & chkString(Request("SearchMember"), "SQLString") & ") "
		end if
		strSql = strSql & " AND " & strTablePrefix & "FORUM.F_TYPE = " & 0 
		strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.FORUM_ID DESC, "
		strSql = strSql & "          " & strTablePrefix & "TOPICS.T_LAST_POST DESC"

		mypage = trim(chkString(request("whichpage"), "SQLString"))

		If mypage = "" then
			mypage = 1
		end if
		rs.Open strSql, my_Conn, 3,1

spThemeBlock1_open(intSkin)
%><table cellpadding="4" cellspacing="1" width="100%">
      <tr>
        <td align="center" class="tSubTitle">&nbsp;</td>
        <td align="center" class="tSubTitle"><b>Topic</b></td>
        <td align="center" class="tSubTitle"><b>Author</b></td>
        <td align="center" class="tSubTitle"><b>Replies</b></td>
        <td align="center" class="tSubTitle"><b>Last Post</b></td>
      </tr>
<%		if rs.EOF or rs.BOF then  '## No topic %>
      <tr>
        <td colspan="5" class="fSubTitle"><b>No Matches Found</b></td>
      </tr>
<%
		else 
			rs.MoveFirst
			currForum = 0 
			currTopic = 0
			do until rs.EOF
				if chkForumAccess(strUserMemberID,rs("FORUM_ID")) then 

					' - Find out if the Category is Locked or Un-Locked and if it Exists
					strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS "
					strSql = strSql & ", " & strTablePrefix & "FORUM.F_STATUS "
					strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
					strSql = strSql & " , " & strTablePrefix & "FORUM "
					strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strTablePrefix & "FORUM.CAT_ID "
					strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & rs("FORUM_ID")
		
					set rsCFStatus = my_Conn.Execute (StrSql)

					if (currForum <> rs("FORUM_ID")) and (currTopic <> rs("TOPIC_ID")) then 
%>
      <tr>
        <td height="20" colspan="5" class="tAltSubTitle"><b>&nbsp;<% =ChkString(rs("F_SUBJECT"),"display") %></b></td>
      </tr>
						
<%					currForum = rs("FORUM_ID")
					end if %>
<%					if currTopic <> rs("TOPIC_ID") then
 					  if sCColor = "tCellAlt0" then
					    sCColor = "tCellAlt2"
					  else
					    sCColor = "tCellAlt0"
					  end if %>
      					<tr class="<%= sCColor %>">
<%						if rsCFStatus("CAT_STATUS") <> 0 and rsCFstatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then %>
        						<td class="fNorm" align="center"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkIsNew(rs("T_LAST_POST")) %></a></td>
<%						else %>
						        <td class="fNorm" align="center"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<%=rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<%=ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_locked.gif"
<% 							if rsCFStatus("CAT_STATUS") = 0 then 
								Response.Write ("alt='Category Locked'")
							elseif rsCFStatus("F_STATUS") = 0 then 
								Response.Write ("alt='Forum Locked'")
							else
								Response.Write ("alt='Topic Locked'")
							end if %>
								border="0"></a>
							</td>
<%						end if %>
        <td class="fNorm"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a>&nbsp;
        <% if rs("T_POLL") <> 0 then %>&nbsp;<img src="images/icons/icon_topic_poll.gif"><% end if %><% if rs("T_NEWS") = 1 then%>&nbsp;<img src="images/icons/icon_topic_news.gif"><% end if %></td>
        <td class="fNorm" valign="top" align="center"><% =ChkString(rs("M_NAME"),"display") %></td>
        <td class="fNorm" valign="top" align="center"><% =rs("T_REPLIES") %></td>
        <td valign="top" align="center" nowrap><span class="fSmall"><b><% =ChkDate(left(rs("T_LAST_POST"),8) & "000000") %></b></span></td>
      </tr>
<%					currTopic = rs("TOPIC_ID")
					end if 
				end if
				rs.MoveNext 
			loop 
		end if 
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
<p align="center">
<a href="JavaScript:history.go(-1)">Back To Search Page</a>
</p>

<%	Else %>

<p align="center"><span class="fSubTitle">You must enter searchwords</span></p>

<p align="center"><a href="JavaScript:history.go(-1)">Back To Search Page</a></p>
<meta http-equiv="Refresh" content="2; URL=JavaScript:history.go(-1)">
<%
	end if
Else
%>

<form name="SearchForm" action="forum_search.asp?mode=DoIt" method="post" id="formEle" onsubmit="return checklength()">
<%
spThemeTableCustomCode = "align=""center"" width=""65%"""
spThemeBlock1_open(intSkin)
%><table cellpadding="3" cellspacing="0" width="100%" class="tCellAlt1">
      <tr>
        <td align="right" valign="top"><span class="fNorm"><b>Search For:</b></span><br /><br /><br />
        <input type="checkbox" value="news" name="news" <% if Request.QueryString("mode") = "news" then%>checked<%end if%> onClick="RefreshS();"><span class="fNorm">News only</span></td>
        <td align="left" valign="top"><input type="text" name="Search" size="40" value="<%=chkString(Request.QueryString("Search"), "sqlstring")%>"><br />
        
		<input type="radio" class="radio" name="andor" value="phrase"><span class="fNorm">Match exact phrase</span><br />
		<input type="radio" class="radio" name="andor" value=" and " checked><span class="fNorm">Search for all Words</span><br />
        <input type="radio" class="radio" name="andor" value=" or "><span class="fNorm">Match any of the words</span></td>
      </tr>
      <tr>
        <td align="right" valign="top" class="fNorm"><b>Search Forum:</b></td>
        <td align="left" valign="top">
        <select name="Forum" size="1">
          <option value="0">All Forums</option>
<%
	'
	strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM"

	set rs = my_Conn.execute (strSql)

	On Error Resume Next

	do until rs.EOF
		if chkForumAccess(strUserMemberID,rs("FORUM_ID")) then
			Response.Write "          <option value=""" & rs("FORUM_ID") & """>" & ChkString(left(rs("F_SUBJECT"), 50),"display") & "</option>" & vbCrLf
		end if
		rs.movenext
	loop
%>
        </select>
        </td>
      </tr>
	  <tr>
        <td class="fNorm" align="right" valign="top"><b>Search In:</b></td>
        <td align="left" valign="top">
        <select NAME="SearchMessage">
          <option value="0">Entire Message</option>
          <option value="1">Subject Only</option>		  
        </select>
        </td>
      </tr>
      <tr>
        <td class="fNorm" align="right" valign="top"><b>Search By Date:</b></td>
        <td align="left" valign="top">
        <select NAME="SearchDate">
          <option value="0">Any Date</option>
          <option VALUE="1">Since Yesterday</option>
          <option VALUE="2">Since 2 Days Ago</option>
          <option VALUE="5">Since 5 Days Ago</option>
          <option VALUE="7">Since 7 Days Ago</option>
          <option VALUE="14">Since 14 Days Ago</option>
          <option VALUE="30" selected>Since 30 Days Ago</option>
          <option VALUE="60">Since 60 Days Ago</option>
          <option VALUE="120">Since 120 Days Ago</option>
          <option VALUE="365">In the Last Year</option>
        </select>
        </td>
      </tr>
      <tr>
        <td class="fNorm" align="right" valign="top"><b>Search By Member:</b></td>
        <td align="left" valign="top">
        <select name="SearchMember">
          <option value="0"><%if Request.QueryString("mode") = "news" then%>All News Editors<%else%>All Members<%end if%>
        </select><%if not Request.QueryString("mode") = "news" then%>
		<a href="JavaScript:memberlist();"><span class="fNorm">Select Member</span></a>
	<%end if%>
        </td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2"><input type="submit" value="Search" class="button"></td>
      </tr></table>
<%
spThemeBlock1_close(intSkin)%>
</form>
<% end if %>
<% set rs = nothing %>
    </td>
  </tr>
</table>
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
<% If request.QueryString("mode")="" Then %>
focuspass();
<% End If %>
-->
</script>
<!--#INCLUDE FILE="inc_footer.asp" -->