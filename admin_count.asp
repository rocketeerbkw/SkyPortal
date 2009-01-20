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

Server.ScriptTimeout = 6000
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!-- #include file="lang/en/core_admin.asp" -->
<!-- #include file="includes/inc_DBfunctions.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<%
function clearOnlineUsers()
	droptable("PORTAL_ONLINE")
	sSQL = "CREATE TABLE [PORTAL_ONLINE]([ID] int IDENTITY (1, 1) PRIMARY KEY NOT NULL, [CheckedIn] TEXT(100), [DateCreated] TEXT(100), [LastChecked] TEXT(100), [LastDateChecked] TEXT(100), [M_BROWSE] MEMO, [UserID] TEXT(100), [UserIP] TEXT(255), [UserAgent] TEXT(100));"

	createTable(checkIt(sSQL))

	createIndex("CREATE INDEX [UserID] ON [PORTAL_ONLINE]([UserID]);")
end function

intStep = Request.QueryString("Step")
if intStep = "" or IsNull(intStep) then
	intStep = 1
else
	intStep = cLng(intStep)
end if

if intStep < 5 then 
	Response.write "<meta http-equiv=""Refresh"" content=""3; URL=admin_count.asp?Step=" & intStep + 1 & """>"
	shoMnu = false
else
	shoMnu = true
end if

%>
<table border="0" width="98%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	if shoMnu then
	forumConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	end if
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
	<% 
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtACUpd & "|javascript:;"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if strACntResult <> "" then
    call showMsgBlock(1,strACntResult)
  end if
	
	spThemeBlock1_open(intSkin) %>
<table align=center border=0>
  <tr>
    <td align=center colspan=2><p><span class="fTitle"><b><%= txtACUpdating %></b></span><br />
    &nbsp;</p></td>
  </tr>
<%

if intStep = 1 then 

	Response.Write "<tr>" & vbNewline
	Response.Write "<td align=""right"" valign=""top"">" & txtTopics & ":</td>" & vbNewline
	Response.Write "<td valign=""top"">"

	' - Get contents of the Forum table related to counting
	strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recForumCount = ""
	else
		allForumData = rs.GetRows(adGetRowsRest)
		recForumCount = UBound(allForumData,2)
	end if

	rs.close
	set rs = nothing

	if recForumCount <> "" then
		fFORUM_ID = 0
		fF_TOPICS = 1
		i = 0 

		for iForum = 0 to recForumCount
			ForumID = allForumData(fFORUM_ID,iForum)
			ForumTopics = allForumData(fF_TOPICS,iForum)

			i = i + 1

			' - count total number of topics in each forum in Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & ForumID
			strSql = strSql & " AND T_STATUS <= 1 "

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_TOPICS = 0
			else
				intF_TOPICS = rs1("cnt")
			end if

			set rs1 = nothing

			Response.Write "."
			if i = 80 then 
				Response.Write "<br />" & vbNewline
				i = 0
			end if
		next
	end if

	Response.Write "</td></tr>" & vbNewline

elseif intStep = 2 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top"">" & txtACTopReplies & ":</td>" & vbNewline
	Response.Write "          <td valign=""top"">"

	'
	strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS"
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_STATUS <= 1"

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recTopicCount = ""
	else
		allTopicData = rs.GetRows(adGetRowsRest)
		recTopicCount = UBound(allTopicData,2)
	end if

	rs.close
	set rs = nothing

	if recTopicCount <> "" then
		fTOPIC_ID = 0
		fT_REPLIES = 1
		i = 0 

		for iTopic = 0 to recTopicCount
			TopicID = allTopicData(fTOPIC_ID,iTopic)
			TopicReplies = allTopicData(fT_REPLIES,iTopic)

			i = i + 1

			' - count total number of replies in Topics table
			strSql = "SELECT count(REPLY_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rs.EOF then
				recReplyCntCount = ""
			else
				allReplyCntData = rs.GetRows(adGetRowsRest)
				recReplyCntCount = UBound(allReplyCntData,2)
			end if

			rs.close
			set rs = nothing

			if recReplyCntCount <> "" then
				fReplyCnt = 0

				for iCnt = 0 to recReplyCntCount
					ReplyCnt = allReplyCntData(fReplyCnt,iCnt)

					intT_REPLIES = ReplyCnt

					' - Get last_post and last_post_author for Topic
					strSql = "SELECT R_DATE, R_AUTHOR "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
					strSql = strSql & " ORDER BY R_DATE DESC"

					set rs2 = my_Conn.Execute (strSql)

					if not(rs2.eof or rs2.bof) then
						rs2.movefirst
						strLast_Post = rs2("R_DATE")
						strLast_Post_Author = rs2("R_AUTHOR")
					else
						strLast_Post = ""
						strLast_Post_Author = ""
					end if

					set rs2 = nothing
				next
                        else
				intT_REPLIES = 0

				set rs2 = Server.CreateObject("ADODB.Recordset")

				' - Get post_date and author from Topic
				strSql = "SELECT T_AUTHOR, T_DATE "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
				strSql = strSql & " AND T_STATUS <= 1"

				set rs2 = my_Conn.Execute(strSql)

				if not(rs2.eof or rs2.bof) then
					strLast_Post = rs2("T_DATE")
					strLast_Post_Author = rs2("T_AUTHOR")
				else
					strLast_Post = ""
					strLast_Post_Author = ""
				end if

				set rs2 = nothing

			end if

			strSql = "UPDATE " & strTablePrefix & "TOPICS "
			strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
			if strLast_Post <> "" then 
				strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author 
				end if
			end if
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "<br />" & vbNewline
				i = 0
			end if
		next
	end if

	Response.Write "</td></tr>" & vbNewline

elseif intStep = 3 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top"">" & txtACFrmRplys & ":</td>" & vbNewline
	Response.Write "          <td valign=top>"

	' - Get values from Forum table needed to count replies
	strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recForumCount = ""
	else
		allForumData = rs.GetRows(adGetRowsRest)
		recForumCount = UBound(allForumData,2)
	end if

	rs.close
	set rs = nothing

	if recForumCount <> "" then
		fFORUM_ID = 0
		fF_COUNT = 1
		i = 0

		for iForum = 0 to recForumCount
			ForumID = allForumData(fFORUM_ID,iForum)
			ForumCount = allForumData(fF_COUNT,iForum)

			i = i + 1

			' - Count total number of Replies
			strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & ForumID
			strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_COUNT = 0
				intF_TOPICS = 0
			else
				intF_COUNT = rs1("cnt") + rs1("SumOfT_REPLIES")
				intF_TOPICS = rs1("cnt") 
			end if
			if IsNull(intF_COUNT) then intF_COUNT = 0 
			if IsNull(intF_TOPICS) then intF_TOPICS = 0 

			set rs1 = nothing

			' - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & ForumID & " "
			strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"
			strSql = strSql & " ORDER BY T_LAST_POST DESC"

			set rs2 = my_Conn.Execute (strSql)

			if not (rs2.eof or rs2.bof) then
				strLast_Post = rs2("T_LAST_POST")
				strLast_Post_Author = rs2("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if

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
			strSql = strSql & " WHERE FORUM_ID = " & ForumID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "<br />" & vbNewline
				i = 0
			end if	
		next
	end if
	Response.Write "</td></tr>" & vbNewline

elseif intStep = 4 then 
	clearOnlineUsers()
	Response.Write "<tr><td align=""right"" valign=""top"">" & txtACAUDel & "</td>" & vbNewline
	Response.Write "<td valign=top></td></tr>" & vbNewline

	' - Get values from Forum table needed to count replies
	'strSql = "DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP <> """""
	'my_Conn.execute(strSql)

elseif intStep = 5 then 

	Response.Write "<tr><td align=""right"" valign=""top"">" & txtACTotals & ":</td>" & vbNewline
	Response.Write "<td valign=""top"">"

	
	' - Total of Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
	strSql = strSql & " AS SumOfF_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = my_Conn.Execute(strSql)

	if rs("SumOfF_TOPICS") <> "" then
		Response.Write txtACTotTops & ": " & rs("SumOfF_TOPICS") & "<br />" & vbNewline
		intSumOfF_TOPICS = rs("SumOfF_TOPICS")
	else
		Response.Write txtACTotTops & ": 0<br />" & vbNewLine
		intSumOfF_TOPICS = 0
	end if

	' - Write total Topics to Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET T_COUNT = " & intSumOfF_TOPICS

	set rs = nothing

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	' - Total all the replies for each topic
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
	strSql = strSql & " AS SumOfF_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = my_Conn.Execute (strSql)

	if rs("SumOfF_COUNT") <> "" then
		Response.Write txtACTotPosts & ": " & RS("SumOfF_COUNT") & "<br />" & vbNewline
		intSumOfF_COUNT = rs("SumOfF_COUNT")
	else
		Response.Write txtACTotPosts & ": 0<br />" & vbNewline
		intSumOfF_COUNT = 0
	end if

	' - Write total replies to the Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET P_COUNT = " & intSumOfF_COUNT

	set rs = nothing

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	' - Total number of users
	strSql = "SELECT Count(MEMBER_ID)"
	strSql = strSql & " AS CountOf"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"

	set rs = my_Conn.Execute(strSql)

	Response.Write txtACReg & ": " & rs("Countof") & "<br />" & vbNewline

	' - Write total number of users to Totals table
	strSql = " UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET U_COUNT = " & cLng(RS("Countof"))

	set rs = nothing

	my_Conn.Execute(strSql)
	

	Response.Write txtACAUDel & "<br />" & vbNewline
	Response.Write "</td></tr>" & vbNewline
	Response.Write "<tr><td align=""center"" colspan=""2"">&nbsp;<br />" & vbNewline
	Response.Write "<span class=""fSubTitle""><b>" & txtACCntComp & "</b></span>" & vbNewline
	Response.Write "</td></tr>" & vbNewline
	Response.Write "<tr><td align=""center"" colspan=""2"">&nbsp;<br />" & vbNewline
	Response.Write "<a href=""admin_home.asp"">" & txtAdminHome & "</a></td></tr>" & vbNewline
	'strSql = "delete from " & strTablePrefix & "ONLINE where UserID <> ""123dogg123"""
	'my_Conn.Execute(strSql)
end if
%>
	</table>
	<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
