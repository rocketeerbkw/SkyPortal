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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
 strRqTopicID = 0
 strRqForumID = 0

if Request.QueryString("TOPIC_ID") <> "" or Request.QueryString("TOPIC_ID") <> " " or Request.QueryString("TOPIC_ID") <> 0 then
	if IsNumeric(Request.QueryString("TOPIC_ID")) = True then
		strRqTopicID = cLng(Request.QueryString("TOPIC_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if
if Request.QueryString("FORUM_ID") <> "" or Request.QueryString("FORUM_ID") <> " " then
	if IsNumeric(Request.QueryString("FORUM_ID")) = True then
		strRqForumID = cLng(Request.QueryString("FORUM_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if
if strRqTopicID <> 0 and strRqForumID = 0 then

	' - Find out if the Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "TOPICS.T_REPLIES "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS, " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & strRqTopicID
	strSql = strSql & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & strTablePrefix & "FORUM.FORUM_ID"

	set rsTopicInfo = my_Conn.Execute (StrSql)

	if (rsTopicInfo.EOF and rsTopicInfo.BOF) then
		closeAndGo("fhome.asp")
	else
		ReTopicId = RsTopicInfo("TOPIC_ID")
		ReForumId = RsTopicInfo("FORUM_ID")
		ReCatId   = RsTopicInfo("CAT_ID")
		ReForumTitle = ChkString(RsTopicInfo("F_SUBJECT"),"urlpath")
		ReTopicTitle = ChkString(RsTopicInfo("T_SUBJECT"),"urlpath")
        RsTReplies = rsTopicInfo("T_REPLIES")
        RsPvtForum = rsTopicInfo("F_PRIVATEFORUMS")
		'RsTopicInfo.Close
		set RsTopicInfo = nothing
		
		if not hasAccess(2) and RsPvtForum > 0 then
		  closeAndGo("fhome.asp")
		end if
		if not chkForumAccess(strUserMemberID,ReForumId) then
		  closeAndGo("fhome.asp")
		end if
		
		if request.querystring("view") = "lasttopic" and not RsTReplies = 0 then

			strSql = "SELECT " & strTablePrefix & "REPLY.REPLY_ID "
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			'if not trim(Session(strUniqueID & "last_here_date")) = "" then
			'strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & strRqTopicID & " AND " & strTablePrefix & "REPLY.R_DATE > '" & Session(strUniqueID & "last_here_date") & "' ORDER BY R_DATE ASC"
			'else
			strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & strRqTopicID & " ORDER BY R_DATE DESC"
			'end if
			set rsReplyInfo = my_Conn.Execute (StrSql)
			
			if (rsReplyInfo.EOF and rsReplyInfo.BOF) then
				closeAndGo("forum_topic.asp?TOPIC_ID=" & ReTopicId & "&FORUM_ID=" & ReForumId & "&CAT_ID=" & ReCatId & "&Forum_Title=" & ReForumTitle & "&Topic_Title=" & ReTopicTitle)
			else
				replyIde = RsReplyInfo("REPLY_ID")
				totalReplies = RsTReplies
				pageNum = 0
					do while not totalReplies =< 0
				        totalReplies = totalReplies - strPageSize
			    		pageNum = pageNum + 1
					loop
				RsReplyInfo.Close
				set RsReplyInfo = nothing

				closeAndGo("forum_topic.asp?TOPIC_ID=" & ReTopicId & "&FORUM_ID=" & ReForumId & "&CAT_ID=" & ReCatId & "&Forum_Title=" & ReForumTitle & "&Topic_Title=" & ReTopicTitle & "&whichpage=" & pageNum & "&tmp=1#pid" & replyIde)
			end if		
		else
		closeAndGo("forum_topic.asp?TOPIC_ID=" & ReTopicId & "&FORUM_ID=" & ReForumId & "&CAT_ID=" & ReCatId & "&Forum_Title=" & ReForumTitle & "&Topic_Title=" & ReTopicTitle)
		end if
	end if
'download this portal at SkyPortal.net
elseif strRqForumID <> 0 and strRqTopicID = 0 then

	' - Find out if the Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS " 
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumID

	set rsForumInfo = my_Conn.Execute (StrSql)

	if (rsForumInfo.EOF and rsForumInfo.BOF) then
		closeAndGo("fhome.asp")
	else
        RsPvtForum = rsForumInfo("F_PRIVATEFORUMS")
		if not hasAccess(2) and RsPvtForum > 0 then
		  closeAndGo("fhome.asp")
		end if
		if not chkForumAccess(strUserMemberID,rsForumInfo("FORUM_ID")) then
		  closeAndGo("fhome.asp")
		end if
		
		closeAndGo("forum.asp?FORUM_ID=" & rsForumInfo("FORUM_ID") & "&CAT_ID=" & rsForumInfo("CAT_ID") & "&Forum_Title=" & ChkString(rsForumInfo("F_SUBJECT"),"urlpath"))
	end if
else
	closeAndGo("default.asp")
	Response.End()
end if
%>