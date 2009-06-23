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
cnter = 0
%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" --><%
CurPageInfoChk = "1"
function CurPageInfo () 
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString) 
	PageName = txtForums & ": " & chkString(Request.QueryString("Forum_Title"),"sqlstring")
	PageAction = txtViewing & "<br />" 
	PageLocation = "forum.asp?" & strOnlineQueryString & "" 
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function 

if Request.QueryString("FORUM_ID") = "" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") then
	Response.Redirect "fhome.asp"
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
%>

<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<%
dim mypage : mypage = trim(chkString(request("whichpage"),"sqlstring"))
if ((Trim(mypage) = "") Or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = CInt(mypage)

' Topic Sorting Variables
dim strtopicsortord :strtopicsortord = chkString(request("sortorder"),"sqlstring")
dim strtopicsortfld :strtopicsortfld = chkString(request("sortfield"),"sqlstring")
dim strtopicsortday :strtopicsortday = chkString(request("days"),"sqlstring")
dim inttotaltopics : inttotaltopics = 0
dim strSortCol, strSortOrd

Select Case strtopicsortord
	Case "asc"
		strSortOrd = " ASC"
	Case Else
		strSortOrd = " DESC"
End Select

Select Case strtopicsortfld
	Case "topic"
		strSortCol = "T_SUBJECT" & strSortOrd
	Case "author"
		strSortCol = "T_AUTHOR" & strSortOrd
	Case "replies"
		strSortCol = "T_REPLIES" & strSortOrd
	Case "views"
		strSortCol = "T_VIEW_COUNT" & strSortOrd
	Case "lastpost"
		strSortCol = "T_LAST_POST" & strSortOrd
	Case Else
		strSortCol = "T_LAST_POST" & strSortOrd
End Select
strQStopicsort = "&FORUM_ID=" & chkstring(Request("FORUM_ID"), "sqlstring") &_
	"&CAT_ID=" & chkstring(Request("CAT_ID"), "sqlstring") &_
	"&Forum_Title=" & ChkString(Request("FORUM_Title"),"sqlstring")

' Paging Variables
dim scriptname, intPagingLinks, strQS
scriptname = request.servervariables("script_name")
intPagingLinks = 5
strQS = "&sortorder=" & strtopicsortord &_
	"&sortfield=" & strtopicsortfld &_
	"&days=" & strtopicsortday &_
	"&FORUM_ID=" & chkstring(Request("FORUM_ID"), "sqlstring") &_
	"&CAT_ID=" & chkstring(Request("CAT_ID"), "sqlstring") &_
	"&Forum_Title=" & ChkString(Request("FORUM_Title"),"sqlstring")

nDays = trim(chkString(Request.Cookies(strCookieURL & "NumDays"),"sqlstring"))

if Request.form("cookie") = 1 then
	Response.Cookies(strCookieURL & "NumDays").Path = strCookieURL
	Response.Cookies(strCookieURL & "NumDays") = chkString(Request.Form("days"),"sqlstring")
	Response.Cookies(strCookieURL & "NumDays").expires = dateAdd("d", 360, now())
	nDays = chkString(Request.Form("Days"),"sqlstring")
	mypage = 1
end if

if trim(nDays) = "" then
	nDays = 0
end if

defDate = datetostr(dateadd("d", -(nDays), now()))

%>
<!--#INCLUDE FILE="inc_top.asp" -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1) %>
	<% menu_fp() %></td>
<td class="mainPgCol">
	<% intSkin = getSkin(intSubSkin,2) %>
<%
strPageSize = 15

	' get module id
	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	end if
	
if strPrivateForums = "1" then
	if Request("Method_Type") = "" and (not hasAccess(1)) then
		chkUser4()
	end if
end if


if (hasAccess(1)) or (chkForumModerator(strRqForumID, STRdbntUserName)= "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if

' - Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID

set rsCStatus = my_Conn.Execute (StrSql)

' - Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.FORUM_ID" 
strSql = strSql & " FROM " & strTablePrefix & "FORUM"
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumID

set rsFStatus = my_Conn.Execute (StrSql)

' - Get all topics from DB
strSql ="SELECT " & strTablePrefix & "TOPICS.T_STATUS, " 
strSql = strSql & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.T_VIEW_COUNT, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_POLL, "
strSql = strSql & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_AUTHOR, " 
strSql = strSql & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_LAST_POST, "
strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, " & strTablePrefix & "TOPICS.T_MSGICON, " & strTablePrefix & "TOPICS.T_INPLACE, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_GLOW, "
strSql = strSql & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME, "
strSql = strSql & strTablePrefix & "TOPICS.T_NEWS "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, "
strSql = strSql & strTablePrefix & "TOPICS, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = "& strMemberTablePrefix & "MEMBERS_1.MEMBER_ID "
strSql = strSql & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & strRqForumID & " "
if nDays = "-1" then
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.T_STATUS <> 0 "
end if
if nDays > "0" then
	strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & defDate & "'"
end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.T_INPLACE DESC "
	strSql = strSql & " , " & strTablePrefix & "TOPICS." & strSortCol & " "

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		intOffset = CInt((mypage-1) * strPageSize)
		strSql = strSql & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	' - Get the total pagecount 
	strSql2 = "SELECT COUNT(" & strTablePrefix & "TOPICS.TOPIC_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strTablePrefix & "TOPICS " 
	strSql2 = strSql2 & " WHERE   " & strTablePrefix & "TOPICS.TOPIC_ID > 0 " 
	strSql2 = strSql2 & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & strRqForumID & " "
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
		inttotaltopics = rs.Recordcount
	end if
	
end if

' - Get all Forum Categories From DB
strSql = "SELECT CAT_ID FROM " & strTablePrefix & "CATEGORY"

set rsCat = my_Conn.Execute (StrSql)

if rsCStatus.EOF = true OR rsFStatus.EOF = true then
	closeAndGo("fhome.asp")
end if
%>
<script type="text/javascript">
<!----- 
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}

function setDays() {document.DaysFilter.submit(); return 0;}
// -->
</script>
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtForums & "|fhome.asp"
  arg2 = rsFStatus("F_SUBJECT") & "|link.asp?forum_id=" &  rsFStatus("FORUM_ID")
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td>
<% 'shoUpcomingEventsWeek() %>
</td></tr><tr><td align="center">
<% if (hasAccess(2)) then %>
		<center><% call PostNewTopic() %></center>       
<% else %>
        &nbsp;
<% end if %>
</td></tr><tr><td>
<%
'spThemeTitle = rsFStatus("F_SUBJECT")
spThemeBlock1_open(intSkin)
%>
<table width="100%" cellpadding="0" cellspacing="0"><tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
      <tr>
        <td colspan="8" height="30" class="tTitle" align="center"><%= rsFStatus("F_SUBJECT") %></td>
      </tr>
      <tr>
        <td align="center" class="tSubTitle"><b>&nbsp;</b></td>
        <td align="center" class="tSubTitle"><b>&nbsp;</b></td>
        <td align="center" class="tSubTitle"><b><%= txtTopic %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtAuthor %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtReplies %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtRead %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtLstPost %></b></td>
        <td align="center" class="tSubTitle">
<% if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
       <% call ForumAdminOptions() %>
<% end if %>
        </td>
      </tr>
<% if rs.EOF or rs.BOF then %>
<% else
	rec = 1
	bgClr = "tCellAlt2"
	do until rs.EOF or (rec = strPageSize + 1)
	  if bgClr = "tCellAlt2" then
	  	bgClr = "tCellAlt0"
	  else
	  	bgClr = "tCellAlt2"
	  end if
      response.Write("<tr class=""" & bgClr & """ onMouseOver=""this.className='tCellHover';"" onMouseOut=""this.className='" & bgClr & "';"">") %>
        <td align="center" valign="center" class="fNorm"><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><%
		if rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then %><% =ChkIsNew(rs("T_LAST_POST")) %><%
		else
			if rs("T_LAST_POST") > Session(strUniqueID & "last_here_date") then
				Response.Write "<img src=""images/icons/icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""Topic Locked"">"
			else
				Response.Write "<img src=""images/icons/icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""Topic Locked"">"
			end if
		end if %></a></td>
        <td  valign="center" align="center" class="fNorm"><img src="images/icons/icon_mi_<% =rs("T_MSGICON") %>.gif" height="15" width="15" border="0" hspace="0"></td>
		<td valign="center" align="left" class="fNorm"><% if rs("T_INPLACE") = 1 then%>Sticky: <%end if%><a href="forum_topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><%= left(rs("T_SUBJECT"), 100) %></a>&nbsp;<% if rs("T_NEWS") = 1 then%><img src="images/icons/icon_topic_news.gif"><% end if %><% if rs("T_POLL") <> "0" then %>&nbsp;<img src="images/icons/icon_topic_poll.gif"><% end if %><% if strShowPaging = "1" then TopicPaging() end if%></td>
        <td  valign="center" align="center" class="fNorm">
			<% strIMmsg = txtView & " " & ChkString(rs("M_NAME"),"display") & "'s " & txtProfile %>
				<a href="cp_main.asp?cmd=8&member=<% =rs("T_AUTHOR") %>" title="<%= strIMmsg %>">	
	  <b><%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %></b></a></td>
        <td  valign="center" align="center" class="fNorm"><% =rs("T_REPLIES") %></td>
        <td  valign="center" align="center" class="fNorm"><% =rs("T_VIEW_COUNT") %></td>
        <%
        if IsNull(rs("T_LAST_POST_AUTHOR")) then
            strLastAuthor = ""
        else
            strLastAuthor = "<br />" & txtBy & ": " 
            if strUseExtendedProfile then
				strLastAuthor = strLastAuthor & "<a href=""cp_main.asp?cmd=8&member="& ChkString(rs("T_LAST_POST_AUTHOR"), "JSurlpath") & """>"
			else
				strLastAuthor = strLastAuthor & "<a href=""JavaScript:openWindow2('cp_main.asp?cmd=8&member=" & ChkString(rs("T_LAST_POST_AUTHOR"), "JSurlpath") & "')"">"
			end if
            strLastAuthor = strLastAuthor &  ChkString(rs("LAST_POST_AUTHOR_NAME"), "display") & "</a>"
        end if
        %>
        <td  valign="center" align="center" nowrap><span class="fSmall"><a href="link.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&view=lasttopic"><img src="Themes/<%= strTheme %>/icons/arrow1.gif" title="<%= txtRdLstPst %>" alt="<%= txtRdLstPst %>" border="0" hspace="0"></a><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></span></td>
        <td  valign="center" align="center" nowrap>
		<%	if intSubscriptions = 1 and hasAccess(2) and (strForumSubscription = 2 or strForumSubscription = 3) then 
			  subscription_id = chkIsSubscribed(intAppID,"0","0",rs("TOPIC_ID"),strUserMemberID)
			  if subscription_id <> 0 then
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=9&amp;cid=" & subscription_id &"');""><img src=""themes/" &  strTheme & "/icon_pmread.gif"" title=""" & txtUnSubScrTp & """ alt=""" & txtUnSubScr & """ border=""0""></a>&nbsp;" 
			  else
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=7&amp;cmd=3&amp;cid="&rs("TOPIC_ID")&"');""><img src=""themes/" &  strTheme & "/icon_pmold.gif"" title=""" & txtSubScrTp & """ alt=""" & txtSubScr & """ border=""0""></a>&nbsp;" 
			  end if
			end if %>
<% 		if AdminAllowed = 1 or strNoCookies = "1" then %>
<%	
  cnter = cnter + 1 %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onMouseOver="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onMouseOut="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtTopOpts %>" alt="<%= txtTopOpts %>" border="0" hspace="0" align="absmiddle"></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:105px; z-index:100; display:none; position:absolute; right:50px;">
<%  'cnter = 1
'spThemeTitle= "Topic Options "
'spThemeBlock3_open()
Response.Write("<table width=""90"" cellpadding=""0"" cellspacing=""0""><tr><td align=""center"" nowrap=""nowrap"">")
Response.Write("<b>Topic Options:</b><br />")
				if rs("T_INPLACE") <> 0 then %>
          <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=UTopic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_next_find.gif" title="<%= txtUnStkTop %>" alt="<%= txtUnStkTop %>" border="0" hspace="0"></a>
<%					else %>
          <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=STopic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_prev_find.gif" title="<%= txtStkTop %>" alt="<%= txtStkTop %>" border="0" hspace="0"></a>
<%					end if %>
<%			if rsCStatus("CAT_STATUS") = 0 then %>
          <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =strRqCatID %>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLokCat %>" alt="<%= txtUnLokCat %>" border="0" hspace="0"></a>
<%			else
				if rsFStatus("F_STATUS") = 0 then %>
          <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLkFrm %>" alt="<%= txtUnLkFrm %>" border="0" hspace="0"></a>
<%				else 
					if rs("T_STATUS") <> 0 then %>
          <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_lock.gif" title="<%= txtLkTop %>" alt="<%= txtLkTop %>" border="0" hspace="0"></a>
<%					else %>
          <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLkTop %>" alt="<%= txtUnLkTop %>" border="0" hspace="0"></a>
<%					end if 
				end if
			end if 
			if (AdminAllowed = 1) or (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0) then %>
          <a href="forum_post.asp?method=EditTopic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =strRqCatID %>&auth=<% =ChkString(rs("T_AUTHOR"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_pencil.gif" title="<%= txtEdMsg %>" alt="<%= txtEdMsg %>" border="0" hspace="0"></a>
<%			end if %>
          <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_trashcan.gif" title="<%= txtDelTop %>" alt="<%= txtDelTop %>" border="0" hspace="0"></a>
          <a href="forum_post.asp?method=Reply&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" title="<%= txtRplyTop %>" alt="<%= txtRplyTop %>" height="15" width="15" border="0"></a><br />
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); shwFm('formEle');"><span class="fSmall"><%= txtClose %></span></a></center>
		  <% Response.Write("</td></tr></table>")
'spThemeBlock3_close() %>
</div>
<%		end if %>
        </td>
      </tr>
<%		rec = rec + 1 
		rs.MoveNext 
	loop 
 end if %>
 <%
 	response.write("<tr>" &_
    	"<td align=""center"" class=""tAltSubTitle"" colspan=""7"">")

	dim topicreclow, topicrechigh, topicpage

	topicpage = mypage

	if (topicpage <= 1) then
		topicreclow = 1
	else
		topicreclow = ((topicpage - 1) * strPageSize) + 1
	end if

	topicrechigh = topicreclow + (rec - 2)

	Response.Write("<form method=""post"" name=""topicsort"" id=""pagelist"" action=""" & scriptname & "?nothing=0"& strQStopicsort & """>")
	Response.Write("<table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""right"" width=""100%""><tr><td align=""center"" class=""tAltSubTitle""><b>" & txtShoTops & " " & topicreclow & " " & txtTo & " " & topicrechigh & " " & txtOf & " " & inttotaltopics & ",<br />" & txtSortBy & "</b>&#160;")
	Response.Write("<select name=""sortfield"" style=""font-size:10px;"">" & vbCrLf)
	Response.Write("<option value=""topic""" & CheckSelected(strtopicsortfld,"topic") & ">topic title" & vbCrLf)
	Response.Write("<option value=""lastpost""" & CheckSelected(strtopicsortfld,"lastpost") & ">last post time" & vbCrLf)
	Response.Write("<option value=""replies""" & CheckSelected(strtopicsortfld,"replies") & ">number of replies" & vbCrLf)
	Response.Write("<option value=""views""" & CheckSelected(strtopicsortfld,"views") & ">number of views" & vbCrLf)
	Response.Write("<option value=""author""" & CheckSelected(strtopicsortfld,"author") & ">topic author" & vbCrLf)
	Response.Write("</select>")
	Response.Write("&#160;<b>" & txtIn & "</b>&#160;")
	Response.Write("<select name=""sortorder"" style=""font-size:10px;"">" & vbCrLf)
	Response.Write("<option value=""desc""" & CheckSelected(strtopicsortord,"desc") & ">descending" & vbCrLf)
	Response.Write("<option value=""asc""" & CheckSelected(strtopicsortord,"asc") & ">ascending" & vbCrLf)
	Response.Write("</select>")
	Response.Write("&#160;<b>" & txtOrder & ",<br> " & txtFrom & "</b><nobr>&#160;")

	' Select box for show topic choice
	response.write ("<select name=""Days"" style=""font-size:10px;"">" & vbCrLf &_
	  "<option value=""0""" & CheckSelected(ndays,0) & ">" & txtAllTops & "</option>" & vbCrLf &_
	  "<option value=""-1""" & CheckSelected(ndays,-1) & ">" & txtAllOpnTops & "</option>" & vbCrLf &_
	  "<option value=""1""" & CheckSelected(ndays,1) & ">" & txtLstDay & "</option>" & vbCrLf &_
	  "<option value=""2""" & CheckSelected(ndays,2) & ">" & txtLst2Day & "</option>" & vbCrLf &_
	  "<option value=""5""" & CheckSelected(ndays,5) & ">" & txtLst5Day & "</option>" & vbCrLf &_
	  "<option value=""7""" & CheckSelected(ndays,7) & ">" & txtLst7Day & "</option>" & vbCrLf &_
	  "<option value=""14""" & CheckSelected(ndays,14) & ">" & txtLst14Day & "</option>" & vbCrLf &_
	  "<option value=""30""" & CheckSelected(ndays,30) & ">" & txtLst30Day & "</option>" & vbCrLf &_
	  "<option value=""60""" & CheckSelected(ndays,60) & ">" & txtLst60Day & "</option>" & vbCrLf &_
	  "<option value=""90""" & CheckSelected(ndays,90) & ">" & txtLst90Day & "</option>" & vbCrLf &_
	  "<option value=""120""" & CheckSelected(ndays,120) & ">" & txtLst120Day & "</option>" & vbCrLf &_
	  "<option value=""365""" & CheckSelected(ndays,365) & ">" & txtLstYr & "</option>" & vbCrLf &_
	  "</select>" & vbCrLf)
	Response.Write("<input type=""hidden"" name=""Cookie"" value=""1"">")
	Response.Write("<input type=""submit"" name=""" & txtGo & """ value=""" & txtGo & """ class=""button"">")
	Response.Write("</td></tr></table>" & vbCrLf)
	Response.Write("</form>")

	Response.Write("</td>")

 	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then
    	response.write("<td align=""center"" class=""tAltSubTitle""><b>")
    	call ForumAdminOptions()
    	response.write("</b></td>")
  	end if

    response.write ("</tr>" & vbCrLf)

%>
    </table>
    </td>
  </tr>
  <tr>
  <td colspan=5>
  <% if maxpages > 1 then %>
    <table border=0 align="left">
      <tr>
        <td valign="top" class="fNorm"><b><%= replace(txtMaxPgsTops,"[%maxpages%]",maxpages) %>: &nbsp;&nbsp; </b></td>
        <td valign="top"><% Call Paging() %></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
	</td>
	</tr></table>
<%spThemeBlock1_close(intSkin)%>
	</td>
	</tr></table>

<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top" width="33%">
    <table cellpadding="0" cellspacing="0">
      <tr>
        <td>
		<p>
		<img title="<%= txtNewPosts %>" alt="<%= txtNewPosts %>" src="images/icons/icon_folder_new.gif" width="15" height="15">&nbsp;<%= txtNewLstVst %>.<br />
		<img title="<%= txtOldPosts %>" alt="<%= txtOldPosts %>" src="images/icons/icon_folder.gif" width="15" height="15">&nbsp;<%= txtNoNewLstVst %>.<br />
		</p>
	    </td>
	  </tr>
	</table>
    </td>
    <td align="center" valign="top" width="33%">
<% if hasAccess(2) then %>
		<p align="center"><center><% call PostNewTopic() %></center></p>       
<% else %>
        &nbsp;
<% end if %>
    </td>
    <td align="center" valign="top" width="33%"><span id="formEle">
<!--#INCLUDE file="modules/forums/inc_jump_to.asp" -->
    </span></td>
  </tr>
</table>
	</td>
	</tr></table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
Function ChkIsNew(dt)
	if lcase(strHotTopic) = "1" then
		if dt > Session(strUniqueID & "last_here_date") then
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src=""images/icons/icon_folder_new_hot.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtHotTop & """ alt=""" & txtHotTop & """>"
			else
			        ChkIsNew =  "<img src=""images/icons/icon_folder_new.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtNewTop & """ alt=""" & txtNewTop & """>"
			end if
		Else
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src=""images/icons/icon_folder_hot.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtHotTop & """ alt=""" & txtHotTop & """>"
			else
			        ChkIsNew = "<img src=""images/icons/icon_folder.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"">" 
			end if
		end if
	else
		if dt > Session(strUniqueID & "last_here_date") then
			ChkIsNew =  "<img src=""icon_folder_new.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtNewTop & """ alt=""" & txtNewTop & """>" 
		Else
			ChkIsNew = "<img src=""icon_folder.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"">" 
		end if
	end if
End Function

sub PostNewTopic() 
%>

<% if rsCStatus("CAT_STATUS") = 0 or rsFStatus("F_STATUS") = 0 then 
	if (AdminAllowed = 1) then %>
		<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_locked.gif" title="<%= txtCatLok %>" alt="<%= txtCatLok %>" height="15" width="15" border="0"></a>&nbsp;<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><%= txtNewTop %></a>
<%	else %>
		<img src="images/icons/icon_folder_locked.gif" title="<%= txtCatLok %>" alt="<%= txtCatLok %>" height="15" width="15" border="0">&nbsp;<%= txtCatLok %>
<%	end if 
	else 
		if rsFStatus("F_STATUS") <> 0 then %>
			<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="<%= txtCreNewTop %>" alt="<%= txtCreNewTop %>" height="15" width="15" border="0"></a>&nbsp;<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><%= txtCreNewTop %></a>
	<%	else %>
			<img src="images/icons/icon_folder_locked.gif" title="<%= txtFrmLok %>" alt="<%= txtFrmLok %>" height="15" width="15" border="0">&nbsp;<%= txtFrmLok %>
	<%	end if 
	end if
end sub

sub ForumAdminOptions() 
  cnter = cnter + 1 %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');mwpHSs('formEle','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onMouseOver="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onMouseOut="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtFrmOpts %>" alt="<%= txtFrmOpts %>" border="0" hspace="0" align="absmiddle"></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:110px; z-index:100; display:none; position:absolute; right:50px;">
<%  'cnter = 1
'spThemeTitle= "Forum Options "
'spThemeBlock3_open()
Response.Write("<table width=""100""><tr><td align=""center"" nowrap=""nowrap"">")
Response.Write("<b>" & txtFrmOpts & ":</b><br />")

	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then 
		if rsCStatus("CAT_STATUS") = 0 then 
			if hasAccess(1) then %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =strRqCatID %>')"><img src="images/icons/icon_folder_unlocked.gif" title="<%= txtUnlokCat %>" alt="<%= txtUnlokCat %>" height="15" width="15" border="0"></a>
<%			else %>
    <img src="images/icons/icon_folder_locked.gif" title="<%= txtCatLok %>" alt="<%= txtCatLok %>" height="15" width="15" border="0">
<%			end if 
		else 
			if rsFStatus("F_STATUS") <> 0 then %>
    <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Forum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>')"><img src="images/icons/icon_folder_locked.gif" title="<%= txtLkFrm %>" alt="<%= txtLkFrm %>" height="15" width="15" border="0"></a>
<%			else %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>')"><img src="images/icons/icon_folder_unlocked.gif" title="<%= txtUnLkFrm %>" alt="<%= txtUnLkFrm %>" height="15" width="15" border="0"></a>
<%			end if 
		end if 
		if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0) or (AdminAllowed = 1) then %>
          <a href="forum_post.asp?method=EditForum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&type=0"><img src="images/icons/icon_folder_pencil.gif" title="<%= txtEdFrm %>" alt="<%= txtEdFrm %>" border="0" hspace="0"></a>
<%		end if %>
    <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Forum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>')"><img src="images/icons/icon_folder_delete.gif" title="<%= txtDelFrm %>" alt="<%= txtDelFrm %>" height="15" width="15" border="0"></a>
    <a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="<%= txtCreNewTop %>" alt="<%= txtCreNewTop %>" height="15" width="15" border="0"></a>
<%	end if %><br />
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); mwpHSs('formEle','1');"><span class="fSmall"><%= txtClose %></span></a></center>
<% Response.Write("</td></tr></table>")
  'spThemeBlock3_close()%>
</div>
<%
end sub

sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""1""><tr><td valign=""center""><img src=""images/icons/icon_posticon.gif"" border=""0"">&nbsp;</td>")
		for counter = 1 to mxpages
			ref = "<td align=""center"" class=""tCellAlt1"">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if	
			if counter > 0 then	
			ref = ref & widenum(counter) & "<a href='forum_topic.asp?"
            ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
            ref = ref & "&FORUM_ID=" & rs("FORUM_ID")
            ref = ref & "&CAT_ID=" & strRqCatID
            ref = ref & "&Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
            ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_Title"),"urlpath")
			ref = ref & "&whichpage=" & counter
			ref = ref & "'><span class=""fSmall"">" & counter & "</span></a></td>"
			else
			ref = ref & "<span class=""fSmall"">" & counter & "</span></td>"
			end if
			Response.Write ref 
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr><td>&nbsp;</td>")
			end if
		next				
        Response.Write("</tr></table>")
	end if
end sub

sub Paging()

	if (IsNumeric(intPagingLinks) = 0) AND (Trim(intPagingLinks) = "") then intPagingLinks = 10
	if (maxpages > 1) and (Trim(strQS) <> "") then
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0"" valign=""top"" align=""center"">" & vbCrLf & "<tr align=""center"">" & vbCrLf)
		if maxpages > 10 then
			Response.Write("<td class=""fNorm"">")
			Response.Write("<form method=""post"" name=""pagelist"" id=""pagelist"" action=""" & scriptname & "?n=0"& strQS & """>")
			Response.Write("<table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""right""><tr><td><b>" & txtGoToPg & "</b>:&#160;</td><td>")
			Response.Write("<select name=""whichpage"" onchange=""jumpToPage(this)"" style=""font-size:10px;"">" & vbCrLf)
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """>&#160;-" & vbCrLf)
			pgeselect = ""
			if pgenumber = mypage then pgeselect = " selected"
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """" & pgeselect & ">1" & vbCrLf)
			for counter = 1 to (maxpages/5)
				pgenumber = (counter*5)
				pgeselect = ""
				if pgenumber = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & pgenumber & strQS & """" & pgeselect & ">" & pgenumber & vbCrLf)
			next
			if (maxpages mod 5) > 0 then
				pgeselect = ""
				if maxpages = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & maxpages & strQS & """" & pgeselect & ">" & maxpages & vbCrLf)
			end if
			Response.Write("</select>")
			Response.Write("</td></tr></table>" & vbCrLf)
			Response.Write("</form>")
			Response.Write("</td><td nowrap>&#160;&#160;</td>")
		end if
		
		dim pgelow, pgehigh, pgediv
		if maxpages > intPagingLinks then
			pgediv = Int(Abs(intPagingLinks/2))
			pgelow = mypage - pgediv
			pgehigh = mypage + (intPagingLinks - (pgediv + 1))
			if pgelow < 1 then
				pgelow = 1
				pgehigh = pgelow + (intPagingLinks - 1)
			end if
			if pgehigh > maxpages then
				pgehigh = maxpages
				pgelow = pgehigh - (intPagingLinks - 1)
			end if
		else
			pgelow = 1
			pgehigh = maxpages
		end if

		Response.Write("<td class=""fNorm"" nowrap>&#160;")
		if pgelow > 1 then
			response.write("<a href=""" & scriptname & "?whichpage=1" & strQS & """>&lt;&lt;</a>&#160;")
		else
			response.write("&#160;&#160;&#160;&#160;")
		end if
		Response.Write("</td><td class=""fNorm"">&#160;")
		for counter = pgelow to pgehigh
			if counter <> mypage then
				response.write("&#160;<a href=""" & scriptname & "?whichpage=" & counter & strQS & """>" & counter & "</a>")
			else
				response.write("&#160;" & counter)
			end if
			if counter < pgehigh then response.write("&#160;&#160;|&#160;")
		next
		Response.Write("</td><td class=""fNorm"" nowrap>&#160;")
		if pgehigh < maxpages then
			response.write("&#160;<a href=""" & scriptname & "?whichpage=" & maxpages & strQS & """>&gt;&gt;</a>&#160;")
		else
			response.write("&#160;&#160;&#160;&#160;")
		end if
		Response.Write("</td><td class=""fNorm"" nowrap>&#160;")
		
		' Previous Page Link
		if mypage = 1 then
			response.write(txtPrevious)
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage - 1) & strQS & """>" & txtPrevious & "</a>")
		end if
		response.write("&#160;|&#160;")
		
		' Next Page Link
		if mypage = maxpages then
			response.write(txtNext)
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage + 1) & strQS & """>" & txtNext & "</a>")
		end if
		response.write("&#160;|&#160;")
		
		' Reload Page Link
		response.write("<a href=""" & scriptname & "?whichpage=" & mypage & strQS & """>" & txtReload & "</a>")
		Response.Write("</td></tr></table>")


	else
		response.write("<div class=""fNorm"">&#160;</div>")
	end if

end sub
%>