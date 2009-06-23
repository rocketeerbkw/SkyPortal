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
PgType = "adminForums"
%>
<!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/core_admin.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE file="includes/inc_adminvar.asp" -->
<%
if Session(strCookieURL & "Approval") = "256697926329" then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<!--#INCLUDE file="includes/inc_admin_functions.asp" -->
<!--#INCLUDE file="modules/forums/forum_functions.asp" -->
<%
if Request.Form("Method_Type") = "forumConfig" then 
		Err_Msg = ""
			if Request.Form("intHotTopicNum") = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Hot Topic Number</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "-" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "+" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number without the <b>+</b></li>"
			end if

		if Err_Msg = "" then

			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRIPLOGGING = " & Request.Form("strIPLogging") & ""
			strSql = strSql & ", C_STRSHOWMODERATORS = " & Request.Form("strShowModerators") & ""
			strSql = strSql & ", C_INTHOTTOPICNUM  = " & Request.Form("intHotTopicNum") & ""
			strSql = strSql & ", C_STREDITEDBYDATE = " & Request.Form("strEditedByDate") & ""
			strSql = strSql & ", C_STRPAGESIZE = " & Request.Form("strPageSize") & ""
			strSql = strSql & ", C_STRPAGENUMBERSIZE = " & Request.Form("strPageNumberSize") & ""
			strSql = strSql & ", C_STRQUICKREPLY = " & Request.Form("strQuickReply") & ""
			strSql = strSql & " WHERE CONFIG_ID = 1"

			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""

			Session.Contents("forumHome") = "<li><span class=""fSubTitle"">Feature Configuration Updated!</span></li>"
		else 
			Err_Msg1 = "<li><span class=""fSubTitle"">There Was A Problem With Your Details</span></li>"
			Session.Contents("forumHome") = Err_Msg1 & Err_Msg
		end if	
    closeAndGo("admin_forums.asp")
end if

if Request.Form("Method_Type") = "forumOrder" then 

	Err_Msg = ""
	if Err_Msg = "" then

	i = 1
	do until i > cint(Request.Form("NumberCategories"))
		SelectName = Request.Form("SortCategory" & i)
		SelectID   = Request.Form("SortCatID" & i)

		strSql = "UPDATE " & strTablePrefix & "CATEGORY "
		strSql = strSql & " SET CAT_ORDER = " & SelectName
		strSql = strSql & " WHERE CAT_ID = " & SelectId
		executeThis(strSql)

		j = 1
		do until j > cint(Request.Form("NumberForums" & SelectID))
			SelectNamec = Request.Form("SortCat" & i & "SortForum" & j)
			SelectIDc   = Request.Form("SortCatID" & i & "SortForumID" & j)
			
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET FORUM_ORDER = " & SelectNamec
		strSql = strSql & " WHERE FORUM_ID = " & SelectIDc
		strSql = strSql & " AND CAT_ID = " & SelectID
		executeThis(strSql)

			j = j + 1
		loop
		i = i + 1
	loop
			Session.Contents("forumHome") = "<li><span class=""fSubTitle"">Category/Forum Order Updated!</span></li>"
		else 
			Err_Msg1 = "<li><span class=""fSubTitle"">There Was A Problem With Your Details</span></li>"
			Session.Contents("forumHome") = Err_Msg1 & Err_Msg
		end if	
    closeAndGo("admin_forums.asp?cmd=4")
end if

if trim(request.form("Fvalue")) <> "" then
	DownMSG = request.form("downmsg")
	if strForumStatus = "down" then
	  Fstatus = "up"
	else
	  Fstatus = "down"
	end if
	strSql = "UPDATE " & strTablePrefix & "CONFIG "
	strSql = strSql & " SET C_DOWNMSG = '" & DownMSG & "'"
	strSql = strSql & ",    C_FORUMSTATUS = '" & Fstatus & "'"

	executeThis(strSql)
	Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
	if Fstatus = "down" then 
	 tMsg = "Forum is down."
    else
     tMsg = "Forum Restarted."
    end if
	
	Session.Contents("forumHome") = "<li><span class=""fSubTitle"">" & tMsg & "</span></li>"
    closeAndGo("admin_forums.asp?cmd=5")
end if

if request.querystring("actionLT") = "updateLT" then
	'slImages	=	cint(request.form("strIMGInPosts"))
	slEncode	=	cint(request.form("slEncode"))

	'if slImages <> 1 then
	'	slImages = 0
	'end if

	if slEncode <> 1 then
		slEncode = 0
	end if

  strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & request.form("slPosts") & _
                            "' WHERE m_name='slash' AND m_code = 'slPosts';"
  executeThis(strSQL) 

  strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & request.form("slLength") & _
                            "' WHERE m_name='slash' AND m_code = 'slLength';"
  executeThis(strSQL) 

  strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & request.form("slSort") & _
                            "' WHERE m_name='slash' AND m_code = 'slSort';"
  executeThis(strSQL) 

  'strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & slImages & "' WHERE m_name='slash' AND m_code = 'slImages';"
  'executeThis(strSQL) 

  strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & slEncode & _
                            "' WHERE m_name='slash' AND m_code = 'slEncode';"
  executeThis(strSQL) 

  slMessage = "Settings updated."
	Session.Contents("forumHome") = "<li><span class=""fSubTitle"">" & slMessage & "</span></li>"
    closeAndGo("admin_forums.asp?cmd=6")
end if

if request.querystring("actionN") = "updateN" then
			slImages	=	cint(request.form("strIMGInPosts"))
			slEncode	=	cint(request.form("slEncode"))

			if slImages <> 1 then
				slImages = 0
			end if

			if slEncode <> 1 then
				slEncode = 0
			end if

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & cint(request.form("slPosts")) & _
                            "' WHERE m_name='news' AND m_code = 'slPosts';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & cint(request.form("slLength")) & _
                            "' WHERE m_name='news' AND m_code = 'slLength';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & cint(request.form("slSort")) & _
                            "' WHERE m_name='news' AND m_code = 'slSort';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & slImages & _
                            "' WHERE m_name='news' AND m_code = 'slImages';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & slEncode & _
                            "' WHERE m_name='news' AND m_code = 'slEncode';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & cint(request.form("slColumns")) & _
                            "' WHERE m_name='news' AND m_code = 'slColumns';"
            executeThis(strSQL) 

            strSQL = "UPDATE " & strTablePrefix & "mods SET m_value = '" & chkstring(request.form("slDefimg"),"displayimage") & _
                            "' WHERE m_name='news' AND m_code = 'slDefimg';"
            executeThis(strSQL) 

            slMessage = "Settings updated."
	Session.Contents("forumHome") = "<li><span class=""fSubTitle"">" & slMessage & "</span></li>"
    closeAndGo("admin_forums.asp?cmd=7")
end if

if Request.Form("Method_Type") = "pollConfig" then 
		Err_Msg = ""
		if Request.Form("strPollCreate") = "" then 
			Err_Msg = Err_Msg & "<li>You must select an option</li>"
		end if
		if Request.Form("strFeaturedPoll") = "" then 
			Err_Msg = Err_Msg & "<li>You must select an option</li>"
		end if

		if Err_Msg = "" then

			'
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_POLLCREATE = " & Request.Form("strPollCreate") & ""
if Request.Form("strFeaturedPoll") = "0" or strFeaturedPoll = "0" then
			strSql = strSql & ",    C_FEATUREDPOLL = " & Request.Form("strFeaturedPoll")
end if
			strSql = strSql & " WHERE CONFIG_ID = 1"
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
			Session.Contents("forumHome") = "<li><span class=""fSubTitle"">Poll Configuration Updated!</span></li>"
		else 
			Err_Msg1 = "<li><span class=""fSubTitle"">There Was A Problem With Your Details</span></li>"
			Session.Contents("forumHome") = Err_Msg1 & Err_Msg
		end if	
    closeAndGo("admin_forums.asp?cmd=8")
end if
 %>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	forumConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
	<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = "Admin Area|admin_home.asp"
  arg2 = "Forum Configuration|admin_forums.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
   spThemeBlock1_open(intSkin)
	  response.Write("<div id=""zz"" style=""display:block;"">")
	if Session.Contents("forumHome") <> "" then
	  response.Write("<p align=""center""><ul>")
	  response.Write(Session.Contents("forumHome"))
	  response.Write("</ul></p>")
	  Session.Contents("forumHome") = ""
	end if
	  response.Write("</div>")
	forumConfig()
	forumModerators()
	mergeForums()
	forumArchive()
	forumOrder()
	forumDown()
	lastTopics() 
	forumNews() 
	forumPolls()
	%>
	<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF

sub forumConfig() %>
	<div id="aa" style="display:<%= aa %>;">
<form action="admin_forums.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="forumConfig">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt1">
<table border="0" cellspacing="1" cellpadding="1" class="grid">
  <tr>
    <td class="tTitle" colspan="2"><b>Forum Configuration</b></td>
  </tr><!--
  <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Allow Forum Subscriptions:</b> </td>
    <td class="tCellAlt1">
    <select name="strForumSubscription">
      <option value="0"<% if strForumSubscription = "0" then Response.Write(" checked") %>>None</option>
      <option value="1"<% if strForumSubscription = "1" then Response.Write(" checked") %>>Forums Only</option>
      <option value="2"<% if strForumSubscription = "2" then Response.Write(" checked") %>>Topics Only</option>
      <option value="3"<% if strForumSubscription = "3" then Response.Write(" checked") %>>Forums & Topics</option>
    </select>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#forumSubscriptions')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Restrict Moderators to  <br /> moving their own topics:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strMoveTopicMode" value="1"<% if strMoveTopicMode <> "0" then Response.Write(" checked") %>> 
    Off: <input type="radio" class="radio" name="strMoveTopicMode" value="0"<% if strMoveTopicMode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#MoveTopicMode')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr> -->
  <tr>
    <td class="tCellAlt1" align="right"><b>IP Logging:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strIPLogging" value="1"<% if strIPLogging <> "0" then Response.Write(" checked") %>> 
    Off: <input type="radio" class="radio" name="strIPLogging" value="0"<% if strIPLogging = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#IPLogging')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <!-- <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Private Forums:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strPrivateForums" value="1"<% if strPrivateForums <> "0" then Response.Write(" checked") %>> 
    Off: <input type="radio" class="radio" name="strPrivateForums" value="0"<% if strPrivateForums = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#privateforums')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr> -->
  <tr>
    <td class="tCellAlt1" align="right"><b>Show Moderators:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strShowModerators" value="1"<% if strShowModerators <> "0" then Response.Write(" checked") %>> 
    Off: <input type="radio" class="radio" name="strShowModerators" value="0"<% if strShowModerators = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#ShowModerator')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr><!-- 
  <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Images in Posts:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strIMGInPosts" value="1" <% if (lcase(strIMGInPosts) <> "0") then Response.Write("checked")%>> 
    Off: <input type="radio" class="radio" name="strIMGInPosts" value="0" <% if (lcase(strIMGInPosts) = "0") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#imginposts')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
   </td>
  </tr> -->
  <tr>
    <td class="tCellAlt1" align="right"><b>Edited By on Date:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strEditedByDate" value="1" <% if lcase(strEditedByDate) <> "0" then Response.Write("checked") %>> 
    Off: <input type="radio" class="radio" name="strEditedByDate" value="0" <% if lcase(strEditedByDate) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#editedbydate')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" align="right"><b>Hot Topics:</b> </td>
    <td class="tCellAlt1">
    <!-- On: <input type="radio" class="radio" name="strHotTopic" value="1" <% if (strHotTopic <> "0" or lcase(HotTopic) <> "0") then Response.Write("checked") %>> 
    Off: <input type="radio" class="radio" name="strHotTopic" value="0" <% if (strHotTopic = "0" or lcase(HotTopic) = "0") then Response.Write("checked") %>> -->
    <input type="text" name="intHotTopicNum" size="5" value="<% if intHotTopicNum <> "" then Response.Write(intHotTopicNum) else Response.Write("20") %>">
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#hottopics')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <!--  <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Detailed Statistics:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strShowStatistics" value="1" <% if strShowStatistics <> "0" then Response.Write("checked") %>> 
    Off: <input type="radio" class="radio" name="strShowStatistics" value="0" <% if strShowStatistics = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#stats')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr> --><!-- 
  <tr valign="top">
    <td class="tCellAlt1" align="right"><b>Show Quick Paging:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strShowPaging" value="1" <% if strShowPaging <> "0" then Response.Write("checked") %>> 
    Off: <input type="radio" class="radio" name="strShowPaging" value="0" <% if strShowPaging = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#ShowPaging')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr> -->
  <tr>
    <td class="tCellAlt1" align="right"><b>Show Quick Reply Box:</b> </td>
    <td class="tCellAlt1">
    On: <input type="radio" class="radio" name="strQuickReply" value="1" <% if strQuickReply <> "0" then Response.Write("checked") %>> 
    Off: <input type="radio" class="radio" name="strQuickReply" value="0" <% if strQuickReply = "0" then Response.Write("checked") %>>
   <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#QuickReply')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" align="right"><b>Replies per page:</b> </td>
    <td class="tCellAlt1">
    <input type="text" name="strPageSize" size="5" value="<% if strPageSize <> "" then Response.Write(strPageSize) else Response.Write("15") %>">
   <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#ItemsPerPage')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" align="right"><b>Pagenumbers per row:</b> </td>
    <td class="tCellAlt1">
    <input type="text" name="strPageNumberSize" size="5" value="<% if strPageNumberSize <> "" then Response.Write(strPageNumberSize) else Response.Write("10") %>">
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&this=1#PageNumbersPerRow')"><img src="images/icons/icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" colspan="2" align="center"><input type="submit" value="Submit New Config" id="submit1" name="submit1" class="button"> <input type="reset" value="Reset Old Values" id="reset1" name="reset1" class="button"></td>
  </tr>
</table>
    </td>
  </tr>
</table>
</form>
	</div>
<%
end sub

sub forumModerators() %>
	<div id="ab" style="display:<%= ab %>;">
<%
	if request("Forum") = "" then
		txMessage = "Select a forum to edit the moderators for that forum"
	else
		if request("userid") = "" then
			txMessage = "Select a user to grant/revoke moderator powers for that user.<br />Users witn the <img src=""images/icons/icon_mod.gif"" /> icon are currently moderators of this forum."
		else
			if Request("action") = "" then
				txMessage = "Select an action for this user"
			else
				txMessage = "Action Successful"
			end if
		end if
	end if
%>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
<table border="0" width="100%" cellspacing="1" cellpadding="4" class="grid">
  <tr>
    <td class="tSubTitle"><div class="tSubTitle"><b>Moderator Configuration</b><%if txMessage <> "" Then%><br /><%=txMessage%><%End If%></div></td>
  </tr>
  <tr>
    <td class="tCellAlt1">
<%	if Request("Forum") = "" then %>
<UL>
<%

		'
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM "
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.CAT_ID ASC, " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rs = my_Conn.Execute(strSql)

		do until rs.EOF
%>
<br />
<LI><a href="admin_forums.asp?cmd=1&forum=<%=rs("FORUM_ID")%>"><%=rs("F_SUBJECT")%></a></LI>
<%
rs.MoveNext
		loop
%>
</UL>
<%
	else
		if Request("action") = "" then
			if Request("UserID") = "" then

				'
				strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
				strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"

				strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"

				set rs = my_Conn.Execute(strSql)
%>
<br />
<UL>
<%
				do until rs.EOF
%>
  <LI><% if chkForumModerator(Request("Forum"), rs("M_NAME")) then %><img src="images/icons/icon_mod.gif" /><% end if %><a href="admin_forums.asp?cmd=1&forum=<%=Request("Forum")%>&UserID=<%=rs("MEMBER_ID")%>"><%=rs("M_NAME")%></a></LI>
<%
				rs.MoveNext
				loop
%>
</UL><br />
<br />
<center><a href="admin_forums.asp?cmd=1">Back to Moderator Options</a></center>
<%
			else

				'
				strSql = "SELECT " & strTablePrefix & "MODERATOR.FORUM_ID, " & strTablePrefix & "MODERATOR.MEMBER_ID, " & strTablePrefix & "MODERATOR.MOD_TYPE "
				strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
				strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & Request("UserID") & " "
				strSql = strSql & " AND " & strTablePrefix & "MODERATOR.FORUM_ID = " & Request("Forum") & " "

				set rs = my_Conn.Execute(strSql)
				
				  sSQL = "SELECT M_NAME FROM " & strTablePrefix & "MEMBERS WHERE MEMBER_ID = " & Request("UserID")
				  set rsNam = my_Conn.execute(sSQL)
				  nam = rsNam(0)
				  set rsNam = nothing

				if rs.EOF then
%>
<center>
<br />
<b><%= nam %></b> is not a moderator of the selected forum<br />
<br />
If you would like to make <b><%= nam %></b> a moderator of this forum, <a href="admin_forums.asp?cmd=1&forum=<%=Request("Forum")%>&UserID=<%=Request("UserID")%>&action=1">click here</a>.<br />
<br />
<a href="admin_forums.asp?cmd=1">Back to Moderator Options</a>
</center>
<br />
<%				else %>
<center>
<br />
<b><%= nam %></b> is currently a moderator of the selected forum<br />
<br />
If you would like to remove <b><%= nam %>'s</b> moderator status in this forum, <a href="admin_forums.asp?cmd=1&forum=<%=Request("Forum")%>&UserID=<%=Request("UserID")%>&action=2">click here</a>.<br />
<br />
<a href="admin_forums.asp?cmd=1">Back to Moderator Options</a>
</center>
<br />
<%
				end if
			end if
		else
			sSQL = "SELECT M_NAME FROM " & strTablePrefix & "MEMBERS WHERE MEMBER_ID = " & Request("UserID")
			set rsNam = my_Conn.execute(sSQL)
			nam = rsNam(0)
			set rsNam = nothing
			
			select case Request("action")
				case 1

					'

					strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
					strSql = strSql & "(FORUM_ID"
					strSql = strSql & ", MEMBER_ID"
					strSql = strSql & ") VALUES (" 
					strSql = strSql & Request("Forum")
					strSql = strSql & ", " & Request("UserID")
					strSql = strSql & ")"

					executeThis(strSql)
%>
<br />
<center>
<b><%= nam %></b> is now a moderator of the selected forum<br />
<br />
<a href="admin_forums.asp?cmd=1">Back to Moderator Options</a>
</center><br />
<%
				case 2

					'
					strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
					strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & Request("Forum") & " "
					strSql = strSql & " AND   " & strTablePrefix & "MODERATOR.MEMBER_ID = " & Request("UserID")

					executeThis(strSql)
%>
<br />
<center>
<b><%= nam %>'s</b> moderator status in the selected forum has been removed<br />
<br />
<a href="admin_forums.asp?cmd=1">Back to Moderator Options</a>
</center>
<br />
<%
			end select
		end if
	end if
%>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>
	</div>
<%
end sub

sub getMergeIds
  strSql = "SELECT FORUM_ID, F_SUBJECT" & _
           " FROM " & strTablePrefix & "FORUM"
  dim rs
  set rs = my_conn.execute(strSql)
  Response.write "<center><b>Note:</b> Items merged will be deleted from source forum.<br />If source forum deleted than any unmerged information will be lost.<br /></center>" & _
                 "<form action=""admin_forums.asp?cmd=2"" method=""post"" id=""formEle"" name=""Form1"">" & _
                 "  <input type=""hidden"" name=""Method_Type"" value=""process"">" & _
                 "  <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & _
                 "    <tr>" & _
                 "      <td class=""tCellAlt2"">" & _
                 "        <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & _
                 "          <tr valign=""top"">" & _
                 "            <td class=""tTitle"" colspan=""2""><b>Merge Forums</b></td>" & _
                 "          </tr>" & _
                 "          <tr valign=""top"">" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Source Forum</b>&nbsp;</td>" & _
                 "            <td class=""tCellAlt0"">&nbsp;" & _
                 "              <select name=""srcForum"">"
  do until rs.EOF
    response.write "                <option value=""" & rs("Forum_ID") & """>" & rs("F_SUBJECT") & "</option>"
    rs.movenext
  loop
  response.write "              </select>"
  rs.movefirst
  Response.write "            </td>" & _
                 "          </tr>" & _
                 "          <tr valign=""top"">" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Destination Forum</b>&nbsp;</td>" & _
                 "            <td class=""tCellAlt0"">&nbsp;" & _
                 "              <select name=""destForum"">"
  do until rs.EOF
    response.write "                <option value=""" & rs("Forum_ID") & """>" & rs("F_SUBJECT") & "</option>"
    rs.movenext
  loop
  response.write "              </select>"
  rs.close
  set rs = nothing
  Response.write "            </td>" & _
                 "          </tr>" & _
				 "          <tr>" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Merge Topics</b>&nbsp;" & _
                 "            </td>" & _
                 "            <td class=""tCellAlt0"" align=""left"">" & _
                 "              &nbsp;<input type=""checkbox"" name=""doTopicsMerge"" value=""true"">" & _
                 "            </td>" & _
                 "          </tr>"
  Response.write "           <tr>" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Merge Moderators</b>&nbsp;" & _
                 "            </td>" & _
                 "            <td class=""tCellAlt0"" align=""left"">" & _
                 "              &nbsp;<input type=""checkbox"" name=""doModsMerge"" value=""true"">" & _
                 "            </td>" & _
                 "          </tr>" & _
                 "          <tr>" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Merge Members Lists</b>&nbsp;" & _
                 "            </td>" & _
                 "            <td class=""tCellAlt0"" align=""left"">" & _
                 "              &nbsp;<input type=""checkbox"" name=""doMembersMerge"" value=""true"">" & _
                 "            </td>" & _
                 "          </tr>" & _
                 "          <tr>" & _
                 "            <td class=""tCellAlt0"" align=""right""><b>Delete Source Forum</b>&nbsp;" & _
                 "            </td>" & _
                 "            <td class=""tCellAlt0"" align=""left"">" & _
                 "              &nbsp;<input type=""checkbox"" name=""doDelete"" value=""true"">" & _
                 "            </td>" & _
                 "          </tr>" & _
                 "          <tr valign=""top"">" & _
                 "            <td class=""tCellAlt0"" colspan=""2"" align=""center""><input type=""submit"" value=""Merge"" id=""submit1"" name=""submit1"" class=""button""></td>" & _
                 "          </tr>" & _
                 "        </table>" & _
                 "      </td>" & _
                 "    </tr>" & _
                 "  </table>" & _
                 "</form>"
end sub

sub process

  if Request.Form("destForum") = Request.Form("srcForum") then
    response.write "<p align=""center""><span class=""fSubTitle"">There has been a problem!</span></p>" &_
                   "<p align=""center""><span class=""fSubTitle"">You must choose two different forums!</span></p>"

	getMergeIds

  else
    response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & _
                 "  <tr>" & _
                 "    <td class=""tCellAlt2"">" & _
                 "      <table align=""center"" border=""0"" cellspacing=""1"">" & _
                 "        <tr>" & _
                 "          <td align=""center"" class=""tTitle""><p><b>Forums Merged</b><br />" & _
                 "          </td>" & _
                 "        </tr>"
  strSql = "SELECT CAT_ID FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & int(Request.Form("destForum"))
  set rs = my_conn.execute(strSql)
  strDestCat = rs("CAT_ID")
  intDestForum = int(Request.Form("destForum"))
  intSrcForum = int(Request.Form("srcForum"))
  if Request.Form("doTopicsMerge") = "true" then
    strSql = "SELECT F_TOPICS, F_COUNT FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strMovedCount = rs("F_TOPICS")
    strMovedPosts = rs("F_COUNT")
    'Move those topics!
    strSql = "UPDATE " & strTablePrefix & "TOPICS SET CAT_ID='" & strDestCat & "', FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    'Lets get the replies too!
    strSql = "UPDATE " & strTablePrefix & "REPLY SET CAT_ID='" & strDestCat & "', FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    strSql = "UPDATE " & strTablePrefix & "FORUM SET F_TOPICS=0, F_COUNT=0 WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    'Confirm topics, replies, and subs moved.
    response.write "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Topics Moved: " & strMovedCount & "</td>" & _
                   "        </tr>" & _
                   "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Posts Moved: " & strMovedPosts & "</td>" & _
                   "        </tr>"
    'Update the destination forum count
    strSql = "SELECT count(TOPIC_ID) AS cnt FROM " & strTablePrefix & "TOPICS WHERE FORUM_ID=" & intDestForum
    set rs = my_conn.execute(strSql)
    strDestinTopicCount = rs("cnt")
    strSql = "SELECT count(REPLY_ID) AS cnt FROM " & strTablePrefix & "REPLY WHERE FORUM_ID=" & intDestForum
    set rs = my_conn.execute(strSql)
    strDestinReplyCount = rs("cnt")
    strDestinCount = strDestinReplyCount + strDestinTopicCount
    strSql = "UPDATE " & strTablePrefix & "FORUM SET F_TOPICS=" & strDestinTopicCount & ", F_COUNT=" & strDestinCount & " WHERE FORUM_ID=" & intDestForum
    executeThis(strSql) 
  end if
  if Request.Form("doModsMerge") = "true" then
    strSql = "SELECT count(MOD_ID) AS cnt FROM " & strTablePrefix & "MODERATOR WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strModsCount = rs("cnt")
    strSql = "SELECT MEMBER_ID FROM " & strTablePrefix & "MODERATOR WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strSql = "UPDATE " & strTablePrefix & "MODERATOR SET FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum & "AND ("
    do until rs.EOF
      strSql = strSql & "MEMBER_ID <> " & rs("MEMBER_ID") & " OR "
	  rs.movenext
	loop
	strSql = strSql & "MEMBER_ID <> 0)"
    executeThis(strSql)
	strSql = "DELETE FROM " & strTablePrefix & "MODERATOR WHERE FORUM_ID = " & intSrcForum
	executeThis(strSql)
    response.write "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Moderators Moved: " & strModsCount & "</td>" & _
                   "        </tr>"
  end if
  'If merge Members, lets count em and do so.
  if Request.Form("doMembersMerge") = "true" then
    strSql = "SELECT count(MEMBER_ID) AS cnt FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strMembersCount = rs("cnt")
    strSql = "SELECT MEMBER_ID FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strSql = "UPDATE " & strTablePrefix & "ALLOWED_MEMBERS SET FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum & "AND ("
    do until rs.EOF
      strSql = strSql & "MEMBER_ID <> " & rs("MEMBER_ID") & " OR "
	  rs.movenext
	loop
	strSql = strSql & "MEMBER_ID <> 0)"
	executeThis(strSql)
	strSql = "DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID = " & intSrcForum
	executeThis(strSql)
    response.write "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Allowed Members Moved: " & strMembersCount & "</td>" & _
                   "        </tr>"
  end if
  'If requested, move archives
  if Request.Form("doArchives") = "true" then
    'Lets count how many we are going to move
    strSql = "SELECT F_A_TOPICS, F_A_COUNT FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & intSrcForum
    set rs = my_conn.execute(strSql)
    strMovedACount = rs("F_A_TOPICS")
    strMovedAPosts = rs("F_A_COUNT")
	'Ok, we can move them now.
    strSql = "UPDATE " & strTablePrefix & "A_TOPICS SET CAT_ID='" & strDestCat & "', FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    strSql = "UPDATE " & strTablePrefix & "A_REPLY SET CAT_ID='" & strDestCat & "', FORUM_ID=" & intDestForum & " WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    'We must update the archived counts for both source and destination forums
    strSql = "UPDATE " & strTablePrefix & "FORUM SET F_A_TOPICS=0, F_A_COUNT=0 WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    strSql = "SELECT count(TOPIC_ID) AS cnt FROM " & strTablePrefix & "A_TOPICS WHERE FORUM_ID=" & intDestForum
    set rs = my_conn.execute(strSql)
    strDestinATopicCount = rs("cnt")
    strSql = "SELECT count(REPLY_ID) AS cnt FROM " & strTablePrefix & "A_REPLY WHERE FORUM_ID=" & intDestForum
    set rs = my_conn.execute(strSql)
    strDestinAReplyCount = rs("cnt")
    strDestinACount = strDestinAReplyCount + strDestinATopicCount
    strSql = "UPDATE " & strTablePrefix & "FORUM SET F_A_TOPICS=" & strDestinATopicCount & ", F_A_COUNT=" & strDestinACount & " WHERE FORUM_ID=" & intDestForum
    executeThis(strSql)
    'The last archived date issue would not be updateable, becase one forum might have older posts not archived, so after being merged that is irrelivant.
    response.write "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Archived Topics Moved: " & strMovedACount & "</td>" & _
                   "        </tr>" & _
                   "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Archived Posts Moved: " & strMovedAPosts & "</td>" & _
                   "        </tr>"
  end if
  'If user requested, delete the source forum
  if Request.Form("doDelete") = "true" then
    strSql = "DELETE FROM " & strTablePrefix & "FORUM WHERE FORUM_ID=" & intSrcForum
    executeThis(strSql)
    response.write "        <tr>" & _
                   "          <td align=""center"" valign=""top"" class=""tCellAlt0"">Forums Deleted: 1</td>" & _
                   "        </tr>"
  end if
  'Thats a wrap folks, lets close er down
  response.write "       <tr>" & _
                 "         <td align=""center"" valign=""top"" class=""tCellAlt0"">Forum Counts Updated</td>" & _
                 "       </tr>" & _
                 "       <tr>" & _
                 "         <td align=""center"" valign=""top"" class=""tCellAlt0"">&nbsp</td>" & _
                 "       </tr>" & _
                 "       <tr>" & _
                 "         <td align=""center"" valign=""top"" class=""tCellAlt0""><a href=""admin_forums.asp?cmd=2"">Merge Again</a></td>" & _
                 "       </tr>" & _
                 "      </table>" & _
                 "    </td>" & _
                 "  </tr>" & _
                 "</table>" & _
                 "<br />"
  	rs.close
  	set rs = nothing
  end if
end sub

sub mergeForums() %>
	<div id="ac" style="display:<%= ac %>;"><%
  select case Request.Form("Method_Type")
    case "process"
      process
    case else
      getMergeIds
  end select %>
	</div>
<%
end sub

sub forumArchive() %>
	<div id="ad" style="display:<%= ad %>;">
<% 
strWhatToDo = request.querystring("action")
if strWhatToDo = "" then
	strWhatToDo = "default"
End if

Select Case strWhatToDo

	Case "default" %>

	<table border="0" width="50%" cellspacing="0" cellpadding="0" class="tBorder" align="center">
		<tr>
			<td class="tCellAlt2">
				
				<table border="0" width="100%" cellspacing="1" cellpadding="4">
					<tr>
						<td class="tSubTitle" colspan=2>
							<span class="tAltSubTitle">
								<b>Administrative Forum Archive Functions</b>
							</span>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<b>Forum Options:</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
							
								<br />
								<ul>
									<li><a href="admin_forums.asp?cmd=3&action=archive">Archive topics from a forum</a>
									<li><a href="admin_forums.asp?cmd=3&action=deletearchive">Delete selected topics from an archive</a>
									<li><a href="admin_forums.asp?cmd=3&action=delete">Delete <b>all</b> topics from a forum</a>
								</ul>
							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

<% Case "delete" %>

	<table border="0" width="50%" cellspacing="0" cellpadding="0" class="tBorder" align="center">
		<tr>
			<td class="tCellAlt2">
				
				<table border="0" width="100%" cellspacing="1" cellpadding="4">
					<tr>
						<td class="tSubTitle" colspan=2>
							<b>Administrative Forum Delete Functions</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<b>Delete all topics:</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<br />
								<ul>
<% strForumIDN = request.querystring("id")
	If request.querystring("confirm") = "" AND request.querystring("id") = "" then
		strsql = "Select CAT_ID,FORUM_ID, F_SUBJECT from " & strTablePrefix & "FORUM ORDER BY CAT_ID"
		set drs = my_conn.execute(strsql) %>
			
<%
		if drs.eof then
			response.write("No Forums Found!")
		else
			Do until drs.eof
				response.write("<li><a href=""admin_forums.asp?cmd=3&action=delete&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a></li>" & vbcrlf)
				drs.movenext
			Loop
		End if

		set drs = nothing

	Elseif request.querystring("confirm") = "true" AND request.querystring("id") <> "" then
		Call subdeletestuff(strForumIDN)
			response.write("<p>Deletion Completed, <a href=""admin_forums.asp"">Click Here</a> To return to Forum Administration<br /><br />")

	Elseif request.querystring("confirm") = "" AND request.querystring("id") <> "" then
			response.write("<center>Are you sure you want to delete <b>all</b> topics in this forum? This is <B><STRONG>NOT</STRONG></B> reversable</center>" & vbcrlf & "<center><a href=""admin_forums.asp?cmd=3&action=delete&id=" & request.querystring("id") & "&confirm=true"">Yes</a> | <a href=""admin_forums.asp?cmd=3&action=delete&id=" & request.querystring("id") & "&confirm=false"">No</a><br /><br /></center>")

	Elseif request.querystring("confirm") = "false" then
    	response.write("Topics in Forum have NOT been deleted<br />" & vbcrlf & "<a href=""admin_forums.asp"">Back to Forums Administration</a><br /><br />")

    End if %>
								</ul><br />
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<% Case "archive" %>

	<table border="0" width="50%" cellspacing="0" cellpadding="0" class="tBorder" align="center">
		<tr>
			<td class="tCellAlt2">
				
				<table border="0" width="100%" cellspacing="1" cellpadding="4">
					<tr>
						<td class="tSubTitle" colspan=2>
								<b>Administrative Forum Archive Functions</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<b>Archive all topics:</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<br />
								<ul>

<%
	strForumIDN = request.querystring("id")
		If strForumIDN = "" then
			strsql = "Select CAT_ID, FORUM_ID, L_ARCHIVE, F_SUBJECT from " & strTablePrefix & "FORUM ORDER BY CAT_ID, F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)    
			thisCat = 0
			if drs.eof then
	           	response.write("No Forums Found!")
	        else
	           	response.write("<li><a href=""admin_forums.asp?cmd=3&action=archive&id=-1"">All Forums</a>" & vbcrlf)

				Do until drs.eof

					if (IsNull(drs("L_ARCHIVE"))) or (drs("L_ARCHIVE") = "") then 
						archive_date = "Has not been archived" 
					else 
						archive_date = StrToDate(drs("L_ARCHIVE"))
					end if

					if thisCat <> drs("CAT_ID") then response.write "</ul><ul>" 
						response.write("<li><a href=""admin_forums.asp?cmd=3&action=archive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a> Last archive date: " & archive_date & "</li>" & vbcrlf)
						thisCat = drs("Cat_ID")
						drs.movenext
	            Loop
					End if
				set drs = nothing %>

							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<%		Elseif strForumIDN <> "" then
			If request.querystring("confirm") = "" then %>
	        	<form method="post" action="admin_forums.asp?cmd=3&action=archive&id=<%=strForumIDN%>&confirm=no">
					<br />
						Archive Topics which are older than:
					<br />
					
						<select name="archiveolderthan" size=1>
							<% for counter = 1 to 6  %>
								<option value="<%=datetostr(DateAdd("m", -counter, now()))%>">
									<%= counter %> Month<% if counter > 1 then response.write "s" %>
								</option>
							<% next %>
								<option value="<%=datetostr(DateAdd("m", -12, now()))%>">
									One Year
								</option>
	                    </select>
							&nbsp;&nbsp;
						<input type="submit" value="  Archive  " class="button">

	        	</form>

<%        	elseif request.querystring("confirm") = "no" then
		        response.write("<center>Are you sure you want to archive these topics?</center><br /><br />" & vbcrlf & _
        	    "<center><a href=""admin_forums.asp?cmd=3&action=archive&id=" & strForumIDN & "&confirm=yes&adate=" & request.form("archiveolderthan") & """>Yes</a> | <a href=""admin_forums.asp?cmd=3&action=archive&id=" & strForumIDN & "&confirm=cancel"">No</a><br /><br /><br /></center>")
            elseif request.querystring("confirm") = "yes" then
            	Call subarchivestuff(request.querystring("adate"))
            elseif request.querystring("confirm") = "cancel" then
            	response.write("<center>Archiving Cancelled. <a href=""admin_forums.asp"">Click here</a> to return to Forum Administration<br /><br /></center>")
            end if %>
								<br />
							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

<%        End if %>


<%        Case "deletearchive" %>

	<table border="0" width="50%" cellspacing="0" cellpadding="0" class="tBorder" align="center">
		<tr>
			<td class="tCellAlt2">
				
				<table border="0" width="100%" cellspacing="1" cellpadding="4">
					<tr>
						<td class="tSubTitle" colspan=2>
								<b>Administrative Forum Archive Functions</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<b>Archive all topics:</b>
						</td>
					</tr>
					<tr>
						<td class="tCellAlt1" valign=top>
								<br />
								<ul>

<%	strForumIDN = request.querystring("id")
			if request.querystring("id") = "" and request.querystring("confirm") = "" then   
        		response.write("<center>Select a forum from which to delete archived topics</center><br /><ul>" & vbcrlf)
        		strsql = "Select CAT_ID, FORUM_ID, F_SUBJECT from " & strTablePrefix & "FORUM "
        		strSQL = strSQL & "WHERE FORUM_ID IN (SELECT FORUM_ID FROM " & strTablePrefix & "ARCHIVE_TOPICS) ORDER BY CAT_ID, F_SUBJECT DESC"
				set drs =  my_conn.execute(strsql)
            	thisCat = 0
				if drs.eof then
            	''# do nothing
            	else
                	Do until drs.eof
                		if thisCat <> drs("CAT_ID") then response.write "</ul><ul>" 
						response.write("<li><a href=""admin_forums.asp?cmd=3&action=deletearchive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a>")
                		thisCat = drs("Cat_ID")
                		drs.movenext
                	Loop
           		End if %>
								</ul><br />
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

<%     		elseif request.querystring("id") <> "" and request.querystring("confirm") = "" then
     
     			response.write("<center>Select how many months old the Topics should be that you wish to delete?</center>")
            %>
            	<form method="post" action="admin_forums.asp?cmd=3&action=deletearchive&id=<%=strForumIDN%>&confirm=no">
	        	<br /><br />Delete archived Topics which are older than:<br /><br /> <select name="archiveolderthan" size=1>
	        	
				<% for counter = 1 to 6  %>
						<option value="<%=datetostr(DateAdd("m", -counter, now()))%>"><%= counter %> Month<% if counter > 1 then response.write "s" %></option>
				<% next %>
	            <option value="<%=datetostr(DateAdd("m", -12, now()))%>">One Year</option>
	            </select>&nbsp;&nbsp;<input type="submit" value="  Delete  " class="button">
	        	</form>
								<br />
							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
				
            <%
     		elseif request.querystring("id") <> "" and request.querystring("confirm") = "no" then
     			response.write("<center>Are you sure you want to delete these topics from the archive?<br />" & vbcrlf & _
            	"<a href=""admin_forums.asp?cmd=3&action=deletearchive&id=" & strForumIDN & "&confirm=yes&date="&request.form("archiveolderthan")&_
            	""">Yes</a> | <a href=""admin_forums.asp?cmd=3&action=delete&confirm=false&id=" & strForumIDN & """>No</a><br /></center>") %>
								<br />
							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
            
<%     		elseif strForumIDN <> "" and request.querystring("confirm") = "yes" then
     			call subdeletearchivetopics(strForumIDN, request.querystring("date"))
            	response.write("<center>Topics older than " & StrToDate(request.querystring("date")) & " have been deleted from the selected archive forum" & vbcrlf & "<br /><a href=""admin_forums.asp"">Back to Forum Admin</a></center>")%>
								<br />
							
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

<%     		End if   
End Select %>
	</div>
<%
end sub

Sub subdeletearchivetopics(strForum_id, strDateOlderThan)
        	strsql = "Delete from " & strTablePrefix & "ARCHIVE_TOPICS WHERE FORUM_ID=" & strForum_id & " AND T_LAST_POST < '" & strDateOlderThan & "'"
            executeThis(strsql)
            strsql = "DELETE FROM " & strTablePrefix & "ARCHIVE_REPLY WHERE FORUM_ID=" & strForum_id & " AND R_DATE < '" & strDateOlderThan & "'"
            executeThis(strsql)
				Call subdoupdates()
End Sub

Sub subarchivestuff(fdateolderthan)
 Server.ScriptTimeout = 10000 
	 rqID = request("id")
    strsql = "SELECT * FROM " & strTablePrefix & "REPLY where R_DATE < '" & fdateolderthan & "'"
    
	if rqID <> "-1" then strSQL = strSQL & " AND FORUM_ID=" & rqID
	set drs = my_conn.execute(strsql)
			 strsql = "UPDATE " & strTablePrefix & "FORUM SET L_ARCHIVE= '" & datetostr(now) & "'"
			 if rqID <> "-1" then strSQL = strSQL & " WHERE FORUM_ID=" & rqID
			 executeThis(strsql)
	    if drs.eof then
        	response.write("<center>No replies were archived, none found</center>")
        else
        	do until drs.eof
				
                strsqlvalues = "'" & drs("CAT_ID") & "', '" & drs("FORUM_ID") & "', '" & drs("TOPIC_ID") & "', '" & drs("REPLY_ID")
                strsqlvalues = strsqlvalues & "', '" & drs("R_MAIL") & "', '" & drs("R_AUTHOR") & "', '" & chkstring(drs("R_MESSAGE"),"message")
                strsqlvalues = strsqlvalues & "', '" & drs("R_DATE") & "', '" & drs("R_IP") & "'"
            	strsql = "insert into " & strTablePrefix & "ARCHIVE_REPLY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID, R_MAIL, R_AUTHOR, R_MESSAGE, R_DATE, R_IP)"
                strsql = strsql & " Values (" & strsqlvalues & ")"

            	on error resume next
				
				executeThis(strsql)
            	drs.movenext
            Loop
        End if

	strsql = "Select * from " & strTablePrefix & "TOPICS where T_DATE < '" & fdateolderthan & "'"
	if rqID <> "-1" then strSQL = strSQL & " AND FORUM_ID=" & rqID
    
	set drs = my_conn.execute(strsql)
   	if drs.eof then
       	response.write("<center>No Topics were Archived</center>")
    else
       	do until drs.eof
       		strSQL = "SELECT TOPIC_ID FROM " & strTablePrefix & "ARCHIVE_TOPICS WHERE TOPIC_ID=" & drs("TOPIC_ID")
			set rsTcheck = my_conn.execute(strSQL)
			if rsTcheck.eof then
				strsqlvalues = "'" & drs("CAT_ID") & "', '" & drs("FORUM_ID") & "', '" & drs("TOPIC_ID") & "', '" & drs("T_STATUS")
	           	strsqlvalues = strsqlvalues & "', '" & drs("T_MAIL") & "', '" & chkstring(drs("T_SUBJECT"),"message") & "', '" & chkstring(drs("T_MESSAGE"),"message")
	           	strsqlvalues = strsqlvalues & "', '" & drs("T_AUTHOR") & "', '" & drs("T_REPLIES") & "', '" & drs("T_VIEW_COUNT")
	           	strsqlvalues = strsqlvalues & "', '" & drs("T_LAST_POST") & "', '" & drs("T_DATE") & "', '" & drs("T_LAST_POSTER")
	           	strsqlvalues = strsqlvalues & "', '" & drs("T_IP") & "', '" & drs("T_LAST_POST_AUTHOR") & "'"
	           	on error resume next
	       		strsql = "insert into " & strTablePrefix & "ARCHIVE_TOPICS (CAT_ID, FORUM_ID, TOPIC_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, T_REPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR)"
				strsql = strsql & " Values (" & strsqlvalues & ")"
	            executeThis(strsql)
				msg = "<center>All topics older than " & strToDate(fdateolderthan) & " were archived</center>"
			else
	       		strsql = "UPDATE " & strTablePrefix & "ARCHIVE_TOPICS SET " &_
					"T_STATUS = " & drs("T_STATUS") &_
					", T_MAIL = " & drs("T_MAIL") &_
					", T_SUBJECT = '" & chkstring(drs("T_SUBJECT"),"message") & "'" &_
					", T_MESSAGE = '" & chkstring(drs("T_MESSAGE"),"message") & "'" &_
					", T_REPLIES = T_REPLIES + " & drs("T_REPLIES") &_
					", T_VIEW_COUNT = T_VIEW_COUNT + " & drs("T_VIEW_COUNT") &_
					", T_LAST_POST = '" & drs("T_LAST_POST") & "'" &_ 
					",T_LAST_POST_AUTHOR = " & drs("T_LAST_POST_AUTHOR") &_
					"WHERE TOPIC_ID = " & drs("TOPIC_ID")

				newstrsql = "UPDATE " & strTablePrefix & "FORUM SET L_ARCHIVE= '" & datetostr(now) & "' WHERE FORUM_ID = " & drs("FORUM_ID")
				executeThis(newstrsql)

	            executeThis(strsql)
				msg = "Topic exists, Stats Updated......"
			end if
            if err.number = 0 then
               	strsql = "delete from " & strTablePrefix & "TOPICS where TOPIC_ID=" & drs("TOPIC_ID") & " AND T_LAST_POST < '" & fdateolderthan & "'"
               	executeThis(strsql)
                response.write msg
            	if err.number = 0 then
            		strsql = "delete from " & strTablePrefix & "REPLY where TOPIC_ID=" & drs("TOPIC_ID") & "AND R_DATE < '" & fdateolderthan & "'"
            		executeThis(strsql)
            		response.write("<center>All replies for the topic  were archived</center>")
				else
					response.write err.description
				end if

			end if
           drs.movenext
	    Loop
		drs.close
    	Call subdoupdates()
    End if

	response.write("<br /><center><a href=""admin_forums.asp"">Click Here</a> to return to Forums Admin</center>")
	response.write "</tr></table>"
End Sub

Sub subdeletestuff(fstrid)
        	strsql = "Delete from " & strTablePrefix & "TOPICS WHERE FORUM_ID=" & fstrid
            executeThis(strsql)
            strsql = "DELETE FROM " & strTablePrefix & "REPLY WHERE FORUM_ID=" & fstrid
            executeThis(strsql)
				Call subdoupdates()
End Sub

Sub subdoupdates()
	response.write("<table align=""center"" border=""0"">" &_
  "<tr>" & _
   " <td align=""center"" colspan=""2""><p><span class=""fTitle""><b>Updating Counts</b></span><br />" & _
    "&nbsp;</p></td>" & _
 " </tr>")
set rs = Server.CreateObject("ADODB.Recordset")
set rs1 = Server.CreateObject("ADODB.Recordset")

Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=""right"" valign=""top"">Topics:</td>" & vbCrLf
Response.Write "    <td valign=""top"">"

' - Get contents of the Forum table related to counting
strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn
rs.MoveFirst
i = 0 

do until rs.EOF
i = i + 1

	' - count total number of topics in each forum in Topics table
	strSql = "SELECT count(FORUM_ID) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")

	rs1.Open strSql, my_Conn

	if rs1.EOF or rs1.BOF then
		intF_TOPICS = 0
	Else
		intF_TOPICS = rs1("cnt")
	End if
	
	strSql = "UPDATE " & strTablePrefix & "FORUM "
	strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
	
	executeThis(strSql)
	
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br />" & vbCrLf
		i = 0
	End if
loop
rs.Close

Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=""right"" valign=""top"">Topic Replies:</td>" & vbCrLf
Response.Write "    <td valign=""top"">"

'
strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS"

rs.Open strSql, my_Conn
i = 0 

do until rs.EOF
i = i + 1

	' - count total number of replies in Topics table
	strSql = "SELECT count(REPLY_ID) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")

	rs1.Open strSql, my_Conn
	if rs1.EOF or rs1.BOF or (rs1("cnt") = 0) then
		intT_REPLIES = 0
		
		set rs2 = Server.CreateObject("ADODB.Recordset")

		' - Get post_date and author from Topic
		strSql = "SELECT T_AUTHOR, T_DATE "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID") & " "
				
		set rs2 = my_Conn.Execute (strSql)
			
		if not(rs2.eof or rs2.bof) then
			strLast_Post = rs2("T_DATE")
			strLast_Post_Author = rs2("T_AUTHOR")
		else
			strLast_Post = ""
			strLast_Post_Author = ""
		end if
				
		rs2.Close
		set rs2 = nothing
		
	Else
		intT_REPLIES = rs1("cnt")
		
		' - Get last_post and last_post_author for Topic
		strSql = "SELECT R_DATE, R_AUTHOR "
		strSql = strSql & " FROM " & strTablePrefix & "REPLY "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID") & " "
		strSql = strSql & " ORDER BY R_DATE DESC"
				
		set rs3 = my_Conn.Execute (strSql)
			
		if not(rs3.eof or rs3.bof) then
			rs3.movefirst
			strLast_Post = rs3("R_DATE")
			strLast_Post_Author = rs3("R_AUTHOR")
		else
			strLast_Post = ""
			strLast_Post_Author = ""
		end if
	
		rs3.close
		set rs3 = nothing
		
	End if
	
	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
	if strLast_Post <> "" then 
		strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
		if strLast_Post_Author <> "" then 

			strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author 

		end if
	end if
	strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")
	
	executeThis(strSql)
		
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br />" & vbCrLf
		i = 0
	End if
loop
rs.Close

Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=""right"" valign=""top"">Forum Replies:</td>" & vbCrLf
Response.Write "    <td valign=""top"">"

' - Get values from Forum table needed to count replies
strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn, 2, 3

do until rs.EOF

	' - Count total number of Replies
	strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & rs("FORUM_ID")
	
	rs1.Open strSql, my_Conn
	
	if rs1.EOF or rs1.BOF then
		intF_COUNT = 0
		intF_TOPICS = 0
	Else
		intF_COUNT = rs1("cnt") + rs1("SumOfT_REPLIES")
		intF_TOPICS = rs1("cnt") 
	End if
	If IsNull(intF_COUNT) then intF_COUNT = 0 
	if IsNull(intF_TOPICS) then intF_TOPICS = 0 
	
	' - Get last_post and last_post_author for Forum
	strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID") & " "
	strSql = strSql & " ORDER BY T_LAST_POST DESC"

	set rs2 = my_Conn.Execute (strSql)
			
	if not (rs2.eof or rs2.bof) then
		strLast_Post = rs2("T_LAST_POST")
		strLast_Post_Author = rs2("T_LAST_POST_AUTHOR")
	else
		strLast_Post = ""
		strLast_Post_Author = ""
	end if
			
	rs2.Close
	set rs2 = nothing
	
	strSql = "UPDATE " & strTablePrefix & "FORUM "
	strSql = strSql & " SET F_COUNT = " & intF_COUNT
	strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
	if strLast_Post <> "" then 
		strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
		if strLast_Post_Author <> "" then 
			strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
		end if
	end if
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
	
	'Response.Write strSql
	'Response.End
	
	executeThis(strSql)
		
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br />" & vbCrLf
		i = 0
	End if	
loop
rs.Close

Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=""right"" valign=""top"">Totals:</td>" & vbCrLf
Response.Write "    <td valign=""top"">"

' - Total of Topics
strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
strSql = strSql & " AS SumOfF_TOPICS "
strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn

Response.Write "Total Topics: " & RS("SumOfF_TOPICS") & "<br />" & vbCrLf

' - Write total Topics to Totals table
strSql = "UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET T_COUNT = " & rs("SumOfF_TOPICS")

rs.Close

executeThis(strSql)

' - Total all the replies for each topic
strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
strSql = strSql & " AS SumOfF_COUNT "
strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

set rs = my_Conn.Execute (strSql)
'rs.Open strSql, my_Conn

if rs("SumOfF_COUNT") <> "" then
	Response.Write "Total Posts: " & RS("SumOfF_COUNT") & "<br />" & vbCrLf
	strSumOfF_COUNT = rs("SumOfF_COUNT")
else
	Response.Write "Total Posts: 0<br />" & vbCrLf
	strSumOfF_COUNT = "0"
end if

' - Write total replies to the Totals table
strSql = "UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET P_COUNT = " & strSumOfF_COUNT

rs.Close

executeThis(strSql)

' - Total number of users
strSql = "SELECT Count(MEMBER_ID) "
strSql = strSql & " AS CountOf "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"

rs.Open strSql, my_Conn

Response.Write "Registered Users: " & RS("Countof") & "<br />" & vbCrLf

' - Write total number of users to Totals table
strSql = " UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET U_COUNT = " & cint(RS("Countof"))

rs.Close

executeThis(strSql)

Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=""center"" colspan=""2"">&nbsp;<br />" & vbCrLf
Response.Write "    <span class=""fTitle""><b>Count Update Complete</b></span></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf

'on error resume next

set rs = nothing
set rs1 = nothing
response.write("</table>")
End Sub

sub forumOrder() %>
	<div id="ae" style="display:<%= ae %>;">
<form action="admin_forums.asp?cmd=4" method="post" id="formEle" name="Form1">
<input type="hidden" name="Method_Type" value="forumOrder">
<%
' - Get all Forums From DB
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_STATUS, " 
strSql = strSql & strTablePrefix & "CATEGORY.CAT_NAME, " & strTablePrefix & "CATEGORY.CAT_ORDER "
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_ORDER "
strSql = strSql & ", " & strTablePrefix & "CATEGORY.CAT_NAME "

set rs = Server.CreateObject("ADODB.Recordset")
rs.cachesize = 20
rs.open  strSql, my_Conn, 3

rs.movefirst
rs.pagesize = 1

if strDBType = "mysql" then
	'
	strSql2 = "SELECT COUNT(" & strTablePrefix & "CATEGORY.CAT_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strTablePrefix & "CATEGORY " 
				
	set rsCount = my_Conn.Execute(strSql2)

	categorycount = rsCount("PAGECOUNT")
	rsCount.close
	set rsCount = nothing
else
	categorycount = cint(rs.pagecount)
end if
%>
	<input name="NumberCategories" type="hidden" value="<% =categorycount %>"> 
<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
      <table border="0" cellspacing="1" class="grid" cellpadding="4">
        <tr>
          <td align="center" class="tTitle" valign="top" nowrap="nowrap"><b>Category/Forum</b></td>
          <td align="center" class="tTitle" valign="top" nowrap="nowrap"><b>Order</b></td>
        </tr>
<%
        if rs.EOF or rs.BOF then
%>
        <tr>
          <td class="tSubTitle" colspan="2"><div class="fAltSubTitle"><b>No Categories/Forums Found</b></div></td>
        </tr>
<%      else

	catordercount = 1
	do until rs.EOF 
		' - Build SQL to get forums via category
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " 
		strSql = strSql & strTablePrefix & "FORUM.CAT_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_TYPE, " 
		strSql = strSql & strTablePrefix & "FORUM.FORUM_ORDER " 
		strSql = strSql & "FROM " & strTablePrefix & "FORUM " 
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rs("CAT_ID") & " "
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.FORUM_ORDER ASC"
		strSql = strSql &  ", " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rsForum =  Server.CreateObject("ADODB.Recordset")
		rsForum.cachesize = 20
                rsForum.open  strSql, my_Conn, 3

		'if strDBType = "mysql" then
			'
			strSql2 = "SELECT COUNT(" & strTablePrefix & "FORUM.FORUM_ID) AS PAGECOUNT "
			strSql2 = strSql2 & " FROM " & strTablePrefix & "FORUM " 
			strSql2 = strSql2 & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rs("CAT_ID") & " "

			set rsCount = my_Conn.Execute(strSql2)

			forumcount = rsCount("PAGECOUNT")
			rsCount.close
			set rsCount = nothing
		'else
		'	forumcount = cint(rsForum.pagecount)
		'end if

%>
		<input name="NumberForums<% =rs("CAT_ID") %>" type="hidden" value="<% =forumcount %>"> 
<%
		chkDisplayHeader = true

		if rsForum.eof or rsForum.bof then
		    SelectName = "SortCategory" & catordercount
	            SelectID   = "SortCatID" & catordercount
%>
      	          <tr>
        	    <td colspan="2" class="tSubTitle"><div class="fAltSubTitle"><b><% =ChkString(rs("CAT_NAME"),"display") %></b></div></td>
                  </tr>
	          <tr>
	            <td colspan="2" class="tCellAlt1"><input name="<% =SelectID %>" type="hidden" value="<% =rs("CAT_ID") %>"> <input name="<% =SelectName %>" type="hidden" value="<% =rs("CAT_ID") %>">
	              <b>No Forums Found</b></td>
	          </tr>
<%		else
		rsForum.movefirst
		rsForum.pagesize = 1
		  forumordercount = 1
		  do until rsForum.Eof
		  	if ChkDisplayForum(rsForum("FORUM_ID")) then
		  		if rsForum("F_TYPE") <> "1" then 
		  			intForumCount = intForumCount + 1
		  		end if
		  		if chkDisplayHeader then
%>
                                  <tr>
				    <td class="tSubTitle" align="left">
				     <div class="fAltSubTitle"><b><% =ChkString(rs("CAT_NAME"),"display") %></b></div></td>
				    <td class="tSubTitle" align="center">
<%				    SelectName = "SortCategory" & catordercount
			            SelectID   = "SortCatID" & catordercount
%>
			            <input name="<% =SelectID %>" type="hidden" value="<% =rs("CAT_ID") %>" /> 
			            <SELECT name="<% =SelectName %>">
<%			            i = 1
			            do while i <= categorycount%>
					    <option value="<% =i %>"<% if (i = rs("CAT_ORDER")) then Response.Write(" selected=""selected""") %>><% =i %></option>
<%					    i = i + 1
			            loop 
%>
        				</select></td>
				  </tr>
<%
				  chkDisplayHeader = false
				end if
%>
				<tr>
				  <td class="tCellAlt1" align="left">
                                    <b><% =ChkString(rsForum("F_SUBJECT"),"display") %></b></td>
				    <td class="tCellAlt1" align="center">
<%				    SelectName = "SortCat" & catordercount & "SortForum" & forumordercount
			            SelectID   = "SortCatID" & catordercount & "SortForumID" & forumordercount
%>
			            <input name="<% =SelectID %>" type="hidden" value="<% =rsForum("FORUM_ID") %>" /> 
			            <SELECT name="<% =SelectName %>">
<%			            i = 1
			            do while i <= ForumCount  %>
					    <option value=<% =i %> <% if (i = rsForum("FORUM_ORDER")) then Response.Write(" selected=""selected""") %>><% =i %></option>
<%					    i = i + 1
			            loop 
%>
        		</select></td>
				</tr>
<%                              
			end if ' ChkDisplayForum() 
		 	forumordercount = forumordercount + 1
			rsForum.MoveNext
		  loop
		end if
		catordercount = catordercount + 1	
		rs.MoveNext
	loop
end if 
%>
  	<tr valign="top">
    	  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="Submit Order" id="submit1" name="submit1" class="button" /> <input type="reset" value="Reset Old Values" id="reset1" name="reset1" class="button" /></td>
  	</tr>
    </table>           
    </td>
  </tr>
</table>
</form>
<%
rs.close
rsForum.close
set rs = nothing 
set rsForum = nothing %>
	</div>
<%
end sub

sub forumDown() %>
	<div id="af" style="display:<%= af %>;">
<%
	strSql = "SELECT C_DOWNMSG"
	strSql = strSql & " FROM " & strTablePrefix & "CONFIG "
	set rs = my_Conn.Execute(strSql)
%>
<form name="editTask" method="post" action="admin_forums.asp?cmd=5">
<table border="0" cellpadding="0" cellspacing="0"  width="95%" style="border-collapse: collapse" align="center"><tr>
<tr>
<td class="tCellAlt1" width="100%"><b>Current Status:</b> <span class="fAlert"><% =strForumStatus%></span></td>
</tr>
<tr>
<td class="tCellAlt1" width="100%"><b>Down Message:</b> (This message will be displayed only when the forum is down)<br /><textarea name="downmsg" rows="10" cols="42"><% =Trim(CleanCode(rs("C_DOWNMSG")))%></textarea></td>
</tr>
<tr>
<td class="tCellAlt1" width="100%"><input type="submit" value="<%if strForumStatus = "down" then%>Restart Forum<%else%>Close Down Forum<%end if%>" name="update" class="button" />&nbsp;<input type="reset" class="button" /><input type="hidden" value="<%if strForumStatus = "down" then%>up<%else%>down<%end if%>" name="Fvalue" /></td>
</tr>
</table></form>
<br />
	</div>
<%
end sub

sub lastTopics() %>
	<div id="ag" style="display:<%= ag %>;">
	<% 
		'Set objRec  =   Server.CreateObject("ADODB.RecordSet")
		Set objDict =   CreateObject("Scripting.Dictionary")   
        strSQL = "SELECT m_code, m_value FROM " & strTablePrefix & "mods WHERE m_name = 'slash';"
        set objRec = my_conn.Execute(strSQL)

        while not objRec.EOF    
            objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
            objRec.moveNext
        wend     

        slPosts     	=   cint(objDict.Item("slPosts"))
        slLength    	=   cint(objDict.Item("slLength"))
        slSort      	=   cint(objDict.Item("slSort"))
        slEncode      	=   cint(objDict.Item("slEncode"))
		'strIMGInPosts	=	cint(objDict.Item("slImages"))

        set objDict 	=   nothing        

		'if strIMGInPosts = 1 then
		'	slImages = "checked"
		'end if

		if slEncode = 1 then
			slEncode = "checked"
		end if

        Select Case slSort
        Case "2"    '   last post
            select2 = "selected"
    	Case "3"    '   last replied
            select3 = "selected"
    	Case Else   '   last created
            select1 = "selected"
    	End Select
%>
<form name="lasttopics" method="post" action="admin_forums.asp?cmd=6&actionLT=updateLT" id="formEle">
<table border="0" cellspacing="0" cellpadding="4" class="grid" align="center">
    <TR> 
      <TD colspan="2" class="tSubTitle">
        <P><B>Last Topics</B></P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>How many posts to show?</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT type="text" name="slPosts" size="3" maxlength="3" value="<%= slPosts %>" />
          posts</P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>How many characters display?</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT type="text" name="slLength" size="3" maxlength="3" value="<%= slLength %>" />
          characters</P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>Order by?</B></P>
      </TD>
      <TD> 
        <P> 
          <SELECT name="slSort">
            <OPTION value="1" <%= select1 %>>last created</OPTION>
            <OPTION value="2" <%= select2 %>>last post</OPTION>
            <OPTION value="3" <%= select3 %>>hot topics</OPTION>
          </SELECT>
        </P>
      </TD>
    </TR>
    <!-- <TR class="tCellAlt1"> 
      <TD> 
        <P><B>Don't allow images?</B></P>
      </TD>
      <TD> 
        <P> 
            <INPUT type="checkbox" name="strIMGInPosts" value="1" <%= slImages %>>
          </P>
      </TD>
    </TR> -->
    <TR align="center" class="tCellAlt1"> 
      <TD colspan="2"> 
        <INPUT type="submit" value="Update" class="button" />
      </TD>
    </TR>
  </TABLE>
</FORM>
	</div>
<%
end sub

sub forumNews() %>
	<div id="ah" style="display:<%= ah %>;">
	<%
        Set objDict =   CreateObject("Scripting.Dictionary") 
        strSQL      =   "SELECT m_code, m_value FROM " & strTablePrefix & "mods WHERE m_name = 'news';"
        set objRec  =   my_conn.Execute(strSQL)

        while not objRec.EOF    
            objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
            objRec.moveNext
        wend     

        slPosts     	=   cint(objDict.Item("slPosts"))
        slLength    	=   cint(objDict.Item("slLength"))
        slSort      	=   cint(objDict.Item("slSort"))
        slEncode      	=   cint(objDict.Item("slEncode"))
		strIMGInPosts	=	cint(objDict.Item("slImages"))
		strColumns	=	cint(objDict.Item("slColumns"))
		strDefimg	=	chkString(objDict.Item("slDefimg"),"display")

        set objDict 	=   nothing        

		if strIMGInPosts = 1 then
			slImages = "checked"
		end if

		if slEncode = 1 then
			slEncode = "checked"
		end if
		
        Select Case slSort
        Case "2"    '   last post
            select2 = "selected"
    	Case "3"    '   last replied
            select3 = "selected"
    	Case Else   '   last created
            select1 = "selected"
    	End Select
		
		if strColumns = 2 then
			slColumn1 = ""
			slColumn2 = "selected"
		else
			slColumn1 = "selected"
			slColumn2 = ""
		end if
%>
<FORM name="news" method="post" action="admin_forums.asp?cmd=7&actionN=updateN" id="formEle">
  <TABLE border="0" cellspacing="0" cellpadding="4" class="grid" align="center">
    <TR> 
      <TD colspan="2" class="tSubTitle"> 
        <P><B>Front Page News</B></P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>How many posts to 
          show?</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT class="textbox" type="text" name="slPosts" size="3" maxlength="3" value="<%= slPosts %>" />
          posts</P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>How many characters 
          display?</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT class="textbox" type="text" name="slLength" size="4" maxlength="5" value="<%= slLength %>" />
          characters</P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>Order by?</B></P>
      </TD>
      <TD> 
        <P> 
          <SELECT name="slSort">
            <OPTION value="1" <%= select1 %>>last 
            created</OPTION>
            <OPTION value="2" <%= select2 %>>last 
            post</OPTION>
            <OPTION value="3" <%= select3 %>>hot 
            topics</OPTION>
          </SELECT>
        </P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>Use default image</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT type="checkbox" name="strIMGInPosts" value="1" <%= slImages %> />
        </P>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>How many columns?</B></P>
      </TD>
      <TD align="left" valign="middle">
        <select name="slColumns">
          <option value="1" <%= slColumn1 %>>1</option>
          <option value="2" <%= slColumn2 %>>2</option>
        </select>
      </TD>
    </TR>
    <TR class="tCellAlt1"> 
      <TD> 
        <P><B>Allow Forum code?</B></P>
      </TD>
      <TD> 
        <P> 
          <INPUT type="checkbox" name="slEncode" value="1" <%= slEncode %> />
        </P>
      </TD>
    </TR>
    <TR align="center" class="tCellAlt1"> 
      <TD colspan="2"> 
        <INPUT type="submit" value="Update" class="button" />
      </TD>
    </TR>
  </TABLE>
</FORM>
	</div>
<%
end sub

sub forumPolls() %>
	<div id="ai" style="display:<%= ai %>;">
<form action="admin_forums.asp" method="post">
<input type="hidden" name="Method_Type" value="pollConfig" />
<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b>Poll Options</b></td>
  </tr>
  <tr valign="top">
        <td align="center" class="tCellAlt0"><b>Featured Poll</b></td>
        <td align="center" class="tCellAlt0"><b>On <input type="radio" class="radio" name="strFeaturedPoll" value="1" <% if strFeaturedPoll <> "0" then Response.Write("checked")%> /> Off</b><input type="radio" class="radio" name="strFeaturedPoll" value="0" <% if strFeaturedPoll = "0" then Response.Write("checked")%> /></td>
  </tr>
  <tr valign="top">
    <td class="tTitle" colspan="2"><b>Who can create polls</b></td>
  </tr>
  <tr valign="top">
        <td align="center" class="tCellAlt0"><b>All Members</b></td>
        <td align="center" class="tCellAlt0"><input type="radio" class="radio" name="strPollCreate" value="1" <% if strPollCreate = "1" then Response.Write("checked")%> /> </td>
  </tr>
  <tr valign="top">
        <td align="center" class="tCellAlt0"><b>Adminstrators and Moderators</b></td>
        <td align="center" class="tCellAlt0"><input type="radio" class="radio" name="strPollCreate" value="2" <% if strPollCreate = "2" then Response.Write("checked")%> /> </td>
  </tr>
  <tr valign="top">
        <td align="center" class="tCellAlt0"><b>Adminstrators Only</b></td>
        <td align="center" class="tCellAlt0"><input type="radio" class="radio" name="strPollCreate" value="3" <% if strPollCreate = "3" then Response.Write("checked")%> /> </td>
  </tr>
  <tr valign="top">
        <td align="center" class="tCellAlt0"><b>Disable Poll</b></td>
        <td align="center" class="tCellAlt0"><input type="radio" class="radio" name="strPollCreate" value="0" <% if strPollCreate = "0" then Response.Write("checked")%> /> </td>
  </tr>
  <tr valign="top">
    <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="Submit New Config" class="button" /> <input type="reset" value="Reset Old Values" class="button" /></td>
  </tr>
</table>
    </td>
  </tr>
</table></form>
	</div>
<%
end sub %>
