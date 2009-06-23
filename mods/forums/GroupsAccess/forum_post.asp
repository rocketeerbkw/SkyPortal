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
%>
<!--#INCLUDE FILE="config.asp" --> 
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<%

  hasEditor = true
  strEditorElements = "Message"
  editorFull = true

CurPageInfoChk = "1"
function CurPageInfo () 
	strOnlineQueryString = ChkActUsrUrl(chkString(Request.QueryString,"sqlstring"))
	If Request.QueryString("method") = "Reply" or Request.QueryString("method") = "ReplyQuote" or Request.QueryString("method") = "TopicQuote" Then
		PageName = chkString(Request.QueryString("Topic_Title"), "sqlstring")
		PageAction = "Replying To Message<br />" 
		PageLocation = "forum_topic.asp?" & strOnlineQueryString & ""
	ElseIf Request.QueryString("method") = "Topic" Then
		PageName = chkString(Request.QueryString("Forum_Title"), "sqlstring")
	        PageAction = "Posting New Topic in<br />"
		PageLocation = "forum.asp?" & strOnlineQueryString & "" 
	else
		PageName = "Unknown"
		PageAction = "Unknown"
		PageLocation = "fhome.asp"
	end if

	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function
%>
<%
'#################################################################################
'## Variable declaration 
'#################################################################################
dim strSelecSize
dim intCols, intRows
'#################################################################################
'## Initialise variables 
'#################################################################################
strSelectSize = chkString(Request.Form("SelectSize"),"sqlstring")
strRqMethod = chkString(Request.QueryString("method"), "SQLString")
if Request.QueryString("TOPIC_ID") <> "" or Request.QueryString("TOPIC_ID") <> " " then
	if IsNumeric(Request.QueryString("TOPIC_ID")) = True then
		strRqTopicID = cLng(Request.QueryString("TOPIC_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.QueryString("FORUM_ID") <> "" or Request.QueryString("FORUM_ID") <> " " then
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
if Request.QueryString("REPLY_ID") <> "" or Request.QueryString("REPLY_ID") <> " " then
	if IsNumeric(Request.QueryString("REPLY_ID")) = True then
		strRqReplyID = cLng(Request.QueryString("REPLY_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
'response.Write(strRqReplyID & "<br />")
strCkPassWord = chkString(Request.Cookies(strUniqueID & "User")("Pword"), "SQLString")

if strSelectSize = "" or IsNull(strSelectSize) then 
	strSelectSize = chkString(Request.Cookies(strCookieURL & "strSelectSize"),"sqlstring")
end if
if not(IsNull(strSelectSize)) then 
	Response.Cookies(strCookieURL & "strSelectSize") = strSelectSize
	Response.Cookies(strCookieURL & "strSelectSize").expires = Now() + 365
end if
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<%
intSkin = getSkin(intSubSkin,2)
	
if strRqMethod = "Edit" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "Topic" or _
strRqMethod = "TopicQuote" then
 
	' - Find out if the Category/Forum/Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS, " & strTablePrefix & "FORUM.F_STATUS "
	if strRqMethod <> "Topic" then
		strSql = strSql & ", " & strTablePrefix & "TOPICS.T_STATUS, " & strTablePrefix & "TOPICS.T_POLL "
    end if
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY, " & strTablePrefix & "FORUM"
	if strRqMethod <> "Topic" then
		strSql = strSql & ", " & strTablePrefix & "TOPICS "
    end if
	strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID
	strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumID
	strSql = strSql & " AND " & strTablePrefix & "FORUM.CAT_ID = " & strRqCatID
	if strRqMethod <> "Topic" then
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.TOPIC_ID = " & strRqTopicID
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & strRqForumID
		'strSql = strSql & " AND " & strTablePrefix & "TOPICS.CAT_ID = " & strRqCatID
    end if

	set rsStatus = my_Conn.Execute(strSql)
  
	if rsStatus.EOF or rsStatus.BOF then
 
		Go_Result "Please don't attempt to edit the URL<br />to gain access to locked Forums."
	else
 
		blnCStatus = rsStatus("CAT_STATUS")
		blnFStatus = rsStatus("F_STATUS")
		if strRqMethod <> "Topic" then
			blnTStatus = rsStatus("T_STATUS")
        end if
	if strRqMethod <> "Topic" then
                poll_id = rsStatus("T_POLL")
 	end if
		rsStatus.close
		set rsStatus = nothing
	end if
 
	if (hasAccess(1)) or (chkForumModerator(strRqForumID, ChkString(strDBNTUserName, "decode"))= "1") or (lcase(strNoCookies) = "1") then
		AdminAllowed = 1
	else
		AdminAllowed = 0
	end if 
 
	select case strRqMethod
		case "Topic"
			if (blnCStatus = 0) and (AdminAllowed = 0) then
				Go_Result "You have attempted to post a New Topic to a Locked Category"
			end if
			if (blnFStatus = 0) and (AdminAllowed = 0) then
				Go_Result "You have attempted to post a New Topic to a Locked Forum"
			end if
		case "EditTopic"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to edit a Locked Topic"
			end if
		case "Reply", "ReplyQuote", "TopicQuote"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to Reply to a Locked Topic"
			end if
		case "Edit"
			if ((blnCStatus = 0) or (blnFStatus = 0) or (blnTStatus = 0)) and (AdminAllowed = 0) then
				Go_Result "You have attempted to Edit a Reply to a Locked Topic"
			end if
	end select
end if
%>
<%
select case strSelectSize
	case "1"
		intCols = 45
		intRows = 6
	case "2"
		intCols = 70
		intRows = 12
	case "3"
		intCols = 85
		intRows = 12
	case "4"
		intCols = 125
		intRows = 15
	case else
		intCols = 70
		intRows = 12
end select
%>
<script type="text/javascript">
<!--

function selectUsers()
{
	if (document.PostTopic.AuthUsers.length == 1)
	{
		document.PostTopic.AuthUsers.options[0].value = "";
		return;
	}
	if (document.PostTopic.AuthUsers.length == 2)
		document.PostTopic.AuthUsers.options[0].selected = true
	else
	for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
		document.PostTopic.AuthUsers.options[x].selected = true;
}

function MoveWholeList(strAction)
{
	if (strAction == "Add")
	{
		if (document.PostTopic.AuthUsersCombo.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsersCombo.length - 1 ;x++)
			document.PostTopic.AuthUsersCombo.options[x].selected = true;
			InsertSelection("Add");
		}
	}
	else
	{
		if (document.PostTopic.AuthUsers.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
			document.PostTopic.AuthUsers.options[x].selected = true;
			InsertSelection("Del");
		}
	}
}

function InsertSelection(strAction)
{
	var pos,user,mText;
	var count,finished;

	if (strAction == "Add")
	{
		pos = document.PostTopic.AuthUsers.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsersCombo.options[count].selected)
			{
				document.PostTopic.AuthUsers.length +=1;
				document.PostTopic.AuthUsers.options[pos].value = document.PostTopic.AuthUsers.options[pos-1].value;	
				document.PostTopic.AuthUsers.options[pos].text = document.PostTopic.AuthUsers.options[pos-1].text;
				document.PostTopic.AuthUsers.options[pos-1].value = document.PostTopic.AuthUsersCombo.options[count].value;	
				document.PostTopic.AuthUsers.options[pos-1].text = document.PostTopic.AuthUsersCombo.options[count].text;
				document.PostTopic.AuthUsers.options[pos-1].selected = true;
			}
			pos = document.PostTopic.AuthUsers.length;
			count += 1;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsersCombo.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsersCombo.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsersCombo.length-1;z++)
				{	
					document.PostTopic.AuthUsersCombo.options[z].value = document.PostTopic.AuthUsersCombo.options[z+1].value;	
					document.PostTopic.AuthUsersCombo.options[z].text = document.PostTopic.AuthUsersCombo.options[z+1].text;
				}
				document.PostTopic.AuthUsersCombo.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}	

	if (strAction == "Del")
	{
		pos = document.PostTopic.AuthUsersCombo.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected)
			{
				document.PostTopic.AuthUsersCombo.length +=1;
				document.PostTopic.AuthUsersCombo.options[pos].value = document.PostTopic.AuthUsersCombo.options[pos-1].value;	
				document.PostTopic.AuthUsersCombo.options[pos].text = document.PostTopic.AuthUsersCombo.options[pos-1].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].value = document.PostTopic.AuthUsers.options[count].value;	
				document.PostTopic.AuthUsersCombo.options[pos-1].text = document.PostTopic.AuthUsers.options[count].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].selected = true;
			}
			count += 1;
			pos = document.PostTopic.AuthUsersCombo.length;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsers.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}	
}

function DeleteSelection()
{
	var user,mText;
	var count,finished;

		finished = false;
		count = 0;	
		finished = false;
		count = document.PostTopic.AuthUsers.length - 1;
		if (count<1) {
			return;
		}
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
}
function autoReload(objform)
{
	var tmpCookieURL = '<%=strCookieURL%>';
	if (objform.SelectSize.value == 1)
	{
		document.PostTopic.Message.cols = 45;
		document.PostTopic.Message.rows = 6;
	}
	if (objform.SelectSize.value == 2)
	{
		document.PostTopic.Message.cols = 70;
		document.PostTopic.Message.rows = 12;
	}
	if (objform.SelectSize.value == 3)
	{
		document.PostTopic.Message.cols = 85;
		document.PostTopic.Message.rows = 12;
	}
	if (objform.SelectSize.value == 4)
	{
		document.PostTopic.Message.cols = 125;
		document.PostTopic.Message.rows = 15;
	}
	document.cookie = tmpCookieURL + "strSelectSize=" + objform.SelectSize.value
}

function allowmembers() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=allowmember", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }

function authChange(obj) {
     if(obj.options[obj.selectedIndex].value == 13 || obj.options[obj.selectedIndex].value == 14) {
          document.getElementById('AuthPassword').disabled = true;
          document.getElementById('authPassGroup').disabled = false;
     } else {
          document.getElementById('AuthPassword').disabled = false;
          document.getElementById('authPassGroup').disabled = true;
     }
}
//-->
</script>
<% 
if strRqMethod = "EditForum" then
	if (hasAccess(1)) or (chkForumModerator(strRqForumId, strDBNTUserName) = "1") then
		'## Do Nothing
	else
		Response.Write "<p>ERROR: Only moderators and administrators can edit forums</p>" & vbcrlf
%>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
		Response.End
	end if
end if

Msg = ""

select case strRqMethod 
	case "Reply", "ReplyQuote", "TopicQuote", "Edit"
			if (strNoCookies = 1) or (strDBNTUserName = "") then
				Msg = Msg & "<b>Note:</b> You must be registered in order to post a reply.<br />"
				Msg = Msg & "To register, <a href=""policy.asp"">click here</a>. Registration is FREE!<br />"
			end if
	case "Topic", "EditTopic"
			if not hasAccess(2) then
				Msg = Msg & "<b>Note:</b> You must be registered in order to post a Topic.<br />"
				Msg = Msg & "To register, <a href=""policy.asp"">click here</a>. Registration is FREE!<br />"
			end if
	case "Forum"
		Msg = Msg & "<b>Note:</b> You must be an administrator to create a new forum.<br />"
	case "URL"
		Msg = Msg & "<b>Note:</b> You must be an administrator to create a new web link.<br />"
	case "Edit", "EditTopic"
		Msg = Msg & "<b>Note:</b> Only the poster of this message, and the Moderator can edit the message."
	case "EditForum"
		Msg = Msg & "<b>Note:</b> Only the Moderator can edit the message."
	case "EditCategory"
		Msg = Msg & "Note: Only an administrator can edit the subject."
end select

if strRqMethod = "Edit" or _
strRqMethod = "ReplyQuote" then
	'
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & strRqReplyID

	set rs = my_Conn.Execute (strSql)
	
	strAuthor = rs("R_AUTHOR")
	strAuthorName = getMemberName(strAuthor)

	if strRqMethod = "Edit" then
		TxMsg = rs("R_MESSAGE")
		  if strAllowHtml <> 1 then
			TxMsg = CleanCode(TxMsg)
		  else
 			If InStr(lcase(TxMsg),"[code]")>0 and InStr(lcase(TxMsg),"[/code]")>0 Then
			  TxMsg = server.HTMLEncode(TxMsg)
			end if
			'TxMsg = chkString(TxMsg,"message")
		  end if
	else
		if strRqMethod = "ReplyQuote" then
'			TxMsg = "[quote="&strAuthorName&"]" & vbCrLf
		  if strAllowHtml <> 1 then
			TxMsg = "[quote]" & vbCrLf
			TxMsg = TxMsg & rs("R_MESSAGE") & vbCrLf
			TxMsg = TxMsg & "[/quote]"
			TxMsg = CleanCode(TxMsg)
		  else
 			If InStr(lcase(rs("R_MESSAGE")),"[code]")>0 and InStr(lcase(rs("R_MESSAGE")),"[/code]")>0 Then
			  TxMsg = doMsgCode(rs("R_MESSAGE"))
			else
			  TxMsg = rs("R_MESSAGE")
			end if
			  TxMsgg = "<span class=quote><I>" & getMemberName(strAuthor) & " wrote:</i><BR>" & TxMsg & "</span><BR>" & vbCrLf
			  TxMsg = TxMsgg
			  'TxMsg = chkString(TxMsgg,"message")
			  'TxMsg = chkString(rs("R_MESSAGE"),"message")
		  end if
		end if
	end if
	if strDBNTUserName = strAuthorName then 
		boolReply =rs("R_MAIL") 
	end if
end if

if strRqMethod = "EditTopic" or _
strRqMethod = "TopicQuote" then
	'
	strSql = "SELECT " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_NEWS, " & strTablePrefix & "TOPICS.T_MESSAGE"
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & strRqTopicID

	set rs = my_Conn.Execute (strSql)

	TxSub = rs("T_SUBJECT")
	strAuthor = rs("T_AUTHOR")
	boolNews = rs("T_NEWS")

	if strRqMethod = "EditTopic" then
		TxMsg = rs("T_MESSAGE")
		  if strAllowHtml <> 1 then
			TxMsg = CleanCode(TxMsg)
		  else
			'TxMsg = chkString(TxMsg,"message")
		  end if
	else
		if strRqMethod = "TopicQuote" then
		  if strAllowHtml <> 1 then
			TxMsg = "[quote]" & vbCrLf
			TxMsg = TxMsg & rs("T_MESSAGE") & vbCrLf
			TxMsg = TxMsg & "[/quote]"
			TxMsg = CleanCode(TxMsg)
		  else
			TxMsgg = "<span class=quote><I>" & getMemberName(strAuthor) & " wrote:</i><BR>" & rs("T_MESSAGE") & "</span><BR>" & vbCrLf
			TxMsg = TxMsgg
			'TxMsg = chkString(TxMsgg,"message")
			'TxMsg = chkString(rs("T_MESSAGE"),"message")
		  end if
		end if
	end if
	if strDBNTUserName = getMemberName(strAuthor) then boolTopic = rs("T_MAIL")
end if



if strRqMethod = "EditForum" or _
strRqMethod = "EditURL" then
	'
	strSql = "SELECT " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_URL, " & strTablePrefix & "FORUM.F_DESCRIPTION, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumId

	set rs = my_Conn.Execute (strSql)
	
	TxUrl = rs("F_URL")
	ForumAuthType = rs("F_PRIVATEFORUMS")
	pass_new = rs("F_PASSWORD_NEW")
	TxSub = rs("F_SUBJECT")
	TxMsg = rs("F_DESCRIPTION")
end if

if strRqMethod = "EditCategory" then
	'
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_NAME "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
	strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID

	set rs = my_Conn.Execute (strSql)

	if strRqMethod = "EditCategory" then
		TxSub = rs("CAT_NAME")
	end if
end if

select case strRqMethod 
	case "Category"
		btn = "Post New Category"
	case "Forum"
		btn = "Post New Forum"
	case "Topic"
		btn = "Post New Topic"
	case "URL"
		btn = "Post New URL"
	case "Edit", "EditCategory", "EditForum", "EditTopic", "EditURL"
		btn = "Post Changes"
	case "Reply", "ReplyQuote", "TopicQuote"
		btn = "Post New Reply"
	case else
		btn = "Post"
end select
%>
<table class="tPlain" width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td class="leftPgCol" valign="top">
	<% 
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %></td>
	<td class="mainPgCol" valign="top">
<div id="formEles" class="breadcrumb">
<table border="0" width="100%" align=center cellpadding="0" cellspacing="0">
  <tr>
    <td width="33%" align="left">
    <img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
<%
if strRqMethod = "EditCategory" then
%>
    <img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<% =ChkString(Request.QueryString("Cat_Title"),"display") %><br />
<% 
else 
	if strRqMethod = "Edit" or _
	strRqMethod = "EditTopic" or _
	strRqMethod = "Reply" or _
	strRqMethod = "ReplyQuote" or _
	strRqMethod = "Topic" or _
	strRqMethod = "TopicQuote" then 
%>
    <img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="FORUM.asp?CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(Request.QueryString("FORUM_Title"),"display") %></a><br />
<%
	end if 
end if 

if strRqMethod = "Edit" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "TopicQuote" then 
%>
    <img src="images/icons/icon_blank.gif" border="0"><img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open_topic.gif" border="0">&nbsp;<a href="forum_topic.asp?TOPIC_ID=<% =strRqTopicID %>&CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(left(Request.QueryString("Topic_title"), 50),"urlpath") %>"><% =ChkString(Request.QueryString("Topic_Title"),"display") %></a>
<%
end if 
%>
    </td>
  </tr>
</table></div>
<%
intSkin = getSkin(intSubSkin,2)
spThemeBlock1_open(intSkin)
%>
<p align="center">
<% =Msg %>
</p>
<table class="tBorder" width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<form action="forum_post_info.asp" method="post" name="PostTopic" onSubmit="return validate();">
    <table border="0" class="tBorder" cellspacing="1" cellpadding="1" align="center">
	<tr><td width="150">&nbsp;</td><td>&nbsp;</td></tr>

<input name="Method_Type" type="hidden" value="<% =strRqMethod %>">
<input name="Type" type="hidden" value="<% =chkString(Request.QueryString("type"),"sqlstring") %>">
<input name="REPLY_ID" type="hidden" value="<% =strRqReplyID %>">
<input name="TOPIC_ID" type="hidden" value="<% =strRqTopicID %>">
<input name="FORUM_ID" type="hidden" value="<% =strRqForumId %>"> 
<input name="CAT_ID" type="hidden" value="<% =strRqCatID %>">
<input name="Author" type="hidden" value="<% =strAuthor %>">
<input name="Mod_ID" type="hidden" value="<% =chkString(Request.QueryString("mod"),"sqlstring") %>">
<input name="Cat_Title" type="hidden" value="<% =ChkString(Request.QueryString("Cat_Title"), "sqlstring") %>">
<input name="FORUM_Title" type="hidden" value="<% =ChkString(Request.QueryString("FORUM_Title"), "sqlstring") %>">
<input name="Topic_Title" type="hidden" value="<% =ChkString(Request.QueryString("TOPIC_Title"), "sqlstring") %>">
<input name="M" type="hidden" value="<% 'Request.QueryString("M") %>">
<input name="cookies" type="hidden" value="yes">
<%if strRqMethod = "Reply" or strRqMethod = "ReplyQuote" then%>
<input name="Refer" type="hidden" value="<% =strHomeURL %>link.asp?TOPIC_ID=<% =strRqTopicID %>&view=lasttopic">
<%else%>
<input name="Refer" type="hidden" value="<% =strHomeURL %>forum.asp?FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"), "urlpath") %>">
<%end if
if hasAccess(2) then 
%>
    <input name="UserName" type="hidden" Value="<% =strDBNTUserName%>">
    <input name="Password" type="hidden" value="<% =strCkPassWord%>">
<%
else
	if (lcase(strNoCookies) = "1") or _
	(strDBNTUserName = "" or strDBNTUserName = " " or _
	strCkPassWord = "") then 
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>UserName:</b></td>
        <td><input name="UserName" maxLength="25" size="25" type="text" value="<%=chkString(Request.Form("UserName"),"sqlstring")%>"></td>
      </tr>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Password:</b></td>
        <td valign="top"><input name="Password" maxLength="13" size="13" type="password" value="<%=chkString(Request.Form("password"),"sqlstring")%>"></td>
      </tr>
<%
	end if 
end if

if strRqMethod = "Forum" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "EditForum" then 
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Category:</b></td>
        <td>
        <select name="Category" size="1">
<%
'
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_NAME "
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
if mlev = 3 then 
	strSql = strSql & " WHERE CAT_ID = " & strRqCatID
end if 
strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_NAME ASC;"

set rsCat = my_Conn.execute (strSql)

'On Error Resume Next
do until rsCat.eof
	Response.Write "          <option value=""" & rsCat("CAT_ID") & """"
	if cint(strRqCatID) = rsCat("CAT_ID") then
		Response.Write " selected"
	end if
	Response.Write ">" & ChkString(rsCat("CAT_NAME"),"display") & "</option>" & vbCrLf
	rsCat.movenext
loop
set rsCat = nothing
%>
        </select>
        </td>
      </tr>
<%
end if

if (strRqMethod = "EditTopic") then
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Forum:</b></td>
        <td>
<%
	'
	strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_TYPE "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE F_TYPE = 0 "
	if (mLev < 3) then
		'strSql = strSql & " AND   FORUM_ID = " & rs("FORUM_ID")
		strSql = strSql & " AND   FORUM_ID = " & strRqForumId
	end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

	set rsForum = my_Conn.execute (strSql)

	if mlev = 3 or hasAccess(1) then
        response.Write("<select name=""Forum"" size=""1"">")
		on error resume next
		do until rsForum.eof
		  if chkForumAccess(strUserMemberID,rsForum("FORUM_ID")) then
			Response.Write "          <option value='" & rsForum("CAT_ID") & "|" & rsForum("FORUM_ID") & "'"
			if cint(strRqForumId) = rsForum("FORUM_ID") then
				Response.Write " selected"
			end if
			Response.Write ">" & ChkString(rsForum("F_SUBJECT"),"display") & "</option>" & vbCrLf
		  end if
		  rsForum.movenext
		loop
	else
		Response.Write ChkString(rsForum("F_SUBJECT"),"display")
		Response.Write "<input type='hidden' name='Forum' value='" & rsForum("CAT_ID") & "|" & rsForum("FORUM_ID") & "'>"
        Response.Write "</select>" & vbcrlf
	end if
	set rsForum = nothing
	%></td></tr><%
end if 

if strRqMethod = "Category" or _
strRqMethod = "EditCategory" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "Forum" or _
strRqMethod = "EditForum" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Topic" then
 %>
      <tr>
        <td noWrap vAlign="top" align="right"><b>
Subject:</b></td>
	<td><input maxLength="99" name="Subject" value="<% =Trim(ChkString(TxSub,"edit")) %>" size="40"></td>
      </tr>
<script type="text/javascript">document.PostTopic.Subject.focus();</script>
<% 
end if

if strRqMethod = "URL" or strRqMethod = "EditURL" then 
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Address:</b></td>
        <td><input maxLength="150" name="Address" value="<% if (TxUrl <> "") then Response.Write(TxUrl) else Response.Write("http://") %>" size="40"></td>
      </tr>
<%
end if
if strRqMethod = "Topic" then 
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Message Icon:</b></td>
	<td vAlign="top"><input type="radio" class="radio" name="strMessageIcon" value="1" <% Response.Write(" checked") %>>&nbsp;<img src="images/icons/icon_mi_1.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="2">&nbsp;<img src="images/icons/icon_mi_2.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="3">&nbsp;<img src="images/icons/icon_mi_3.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="4">&nbsp;<img src="images/icons/icon_mi_4.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="5">&nbsp;<img src="images/icons/icon_mi_5.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="6">&nbsp;<img src="images/icons/icon_mi_6.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="7">&nbsp;<img src="images/icons/icon_mi_7.gif" border="0"><br />
	<input type="radio" class="radio" name="strMessageIcon" value="8">&nbsp;<img src="images/icons/icon_mi_8.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="9">&nbsp;<img src="images/icons/icon_mi_9.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="10">&nbsp;<img src="images/icons/icon_mi_10.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="11">&nbsp;<img src="images/icons/icon_mi_11.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="12">&nbsp;<img src="images/icons/icon_mi_12.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="13">&nbsp;<img src="images/icons/icon_mi_13.gif" border="0">&nbsp;<input type="radio" class="radio" name="strMessageIcon" value="14">&nbsp;<img src="images/icons/icon_mi_14.gif" border="0"></td>
      </tr>
<% end if
  
  strST = ""

if strRqMethod = "Topic" or strRqMethod = "EditTopic" then 
  strST = strST & "<b>Message: </b><br /><br />"
end if 

if strRqMethod = "Edit" or strRqMethod = "Reply" or strRqMethod = "ReplyQuote" or strRqMethod = "TopicQuote" then 
  strST = strST & "<b>Reply: </b><br /><br />"
end if 

if strAllowHTML = 1 then
	strMsgText = replace(replace(chkString(TxMsg,"message"),"''","'"),"''","'")
else
	strMsgText = replace(replace(Trim(CleanCode(TxMsg)),"''","'"),"''","'")
end if

if strAllowForumCode = 1 then
 strST = strST & "<a href=""JavaScript:openWindow3('pop_portal.asp?cmd=10')"">Forum Code</a><br />"
end if
	
  If strAllowHtml = 1 Then  %>
<tr>
<td align="right" valign="top"><br /><%= strST %>&nbsp;</td>
<td align="left"><div style="width:200px;"><br /><textarea id="Message" name="Message" rows="25" cols="85" class="textbox"><%= strMsgText %></textarea></div><br />
<%	if strAllowHTML = 1 and trim(editorJS) <> "" then
  		response.Write(editorJS)
	End If %>
  </td>
</tr><% 
  	'displayHTMLeditor "Message", strST, strMsgText
  else
  	displayPLAINeditor 1,strMsgText
  end if %>

      <tr>
        <td>&nbsp;</td>
        <td valign="top">
<% 
	if strRqMethod = "Edit" or _
	strRqMethod = "EditTopic" or _
	strRqMethod = "Reply" or _
	strRqMethod = "ReplyQuote" or _
	strRqMethod = "Topic" or _
	strRqMethod = "TopicQuote" then 
%>
        <input name="Sig" type="checkbox" tabindex="3" value="yes" checked>Check here to include your profile signature.<br />
<%
		if strEmail = 1 then 
			if strRqMethod = "Topic" or _ 
			strRqMethod = "EditTopic" then %>
        <input name="rmail" tabindex="4" type="checkbox" value="1" <%=Chked(boolTopic)%>>Check here to be notified by email whenever someone replies to this topic.<br />
<%			else %>
        		  <input name="rmail" tabindex="4" type="checkbox" value="1" <%=Chked(boolReply)%>>Check here to be notified by email whenever anyone replies to this topic.<br /> <%
			end if
		end if

		if 	((hasAccess(1)) or (chkForumModerator(strRqForumId, strDBNTUserName) = "1")) _
		and (strRqMethod = "Topic" or strRqMethod = "Reply" or _
		strRqMethod = "ReplyQuote" or strRqMethod = "TopicQuote") then %>
          <input name="lock" tabindex="5" type="checkbox" value="1">Check here to lock the topic after this post.<br /><% 
		end if%>
		
<% 		if (mlev = 3) then
		  if strRqMethod = "Topic" then %>
        	<input name="news" tabindex="6" type="checkbox" value="1" <%=Chked(boolNews)%>>Check here to post the topic on the front page (Public Forums Only). <b>News</b><br /> <%
		  elseif strRqMethod = "EditTopic" and boolNews = 1 then%>
			<input name="news" type="hidden" value="1" <%=Chked(boolNews)%>>
<%		  end if 
   		end if %>

<% 		if (hasAccess(1)) then
		  if strRqMethod = "Topic" or strRqMethod = "EditTopic" then %>
        	<input name="news" tabindex="6" type="checkbox" value="1" <%=Chked(boolNews)%>>Check here to post the topic on the front page (Public Forums Only). <b>News</b><br /> <%
		  end if 
   		end if

   		if (strPollCreate <> 0) and ((strPollCreate = 1 and hasAccess(2)) or (strPollCreate = 2 and mLev >=3) or (strPollCreate = 3 and hasAccess(1))) then
		
	  	  if strRqMethod = "Topic" then %>
        	<input name="poll" tabindex="7" type="checkbox" value="1">Check here to create a poll.<br />	
<%	  	  end if
	  	  if (strRqMethod = "EditTopic" and poll_id = "0") then %>
        	<input name="poll" tabindex="7" type="checkbox" value="1">Check here to create a poll.<br />	
			<input name="pollTopic_ID" type="hidden" value="<%= strRqTopicID %>">
<%	  	  end if
		end if		
	end if
%>
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" value="<% =btn %>" tabindex="8" accesskey="s" title="Shortcut Key: Alt+S" class="button">
<%
		if (strRqMethod = "Reply" or _
		strRqMethod = "Edit" or _ 
		strRqMethod = "EditTopic" or _ 
		strRqMethod = "ReplyQuote" or _
		strRqMethod = "Topic" or _
		strRqMethod = "TopicQuote") AND strAllowHtml <> 1 then 
%>      
		&nbsp;<input name="Preview" tabindex="9" type="button" class="Button" value=" Preview " onclick="OpenPreview()">
<%
		end if %>
		</td>
      </tr>
<% 
if strPrivateForums <> "0" then 
	if strRqMethod = "Forum" or _
	strRqMethod = "URL" or _
	strRqMethod = "EditURL" or _
	strRqMethod = "EditForum" then 
		if strRqMethod = "EditForum" or _
		strRqMethod = "EditURL" then
			':: ForumAuthType = rs("F_PRIVATEFORUMS")
		else
			ForumAuthType = 0
		end if
%>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Auth Type:</b>
		</td>
        <td><SELECT name="AuthType" onChange="authChange(this)">
        <option value="0"<% if ForumAuthType = 0 then Response.Write(" selected") %>>All Visitors</option>
        <option value="4"<% if ForumAuthType = 4 then Response.Write(" selected") %>>Members Only</option>
        <option value="5"<% if ForumAuthType = 5 then Response.Write(" selected") %>>Members Only (Hidden)</option>
        <option value="1"<% if ForumAuthType = 1 then Response.Write(" selected") %>>Allowed Member List</option>
        <option value="6"<% if ForumAuthType = 6 then Response.Write(" selected") %>>Allowed Member List (Hidden)</option>
<%
		if strRqMethod = "Forum" or _
		strRqMethod = "EditForum" then 
%>
        <!-- <option value="2" <% if ForumAuthType = 2 then Response.Write(" selected") %>>Password Protected</option> -->
        <option value="7"<% if ForumAuthType = 7 then Response.Write(" selected") %>>Members Only & Password Protected</option>
        <option value="3"<% if ForumAuthType = 3 then Response.Write(" selected") %>>Allowed Member List & Password Protected</option>
<%
		end if 
%>
<%
	if strNTGroups = "1" then 
%>
        <option value="9" <% if ForumAuthType = 9 then Response.Write(" selected") %>>NT Global Group</option>
        <option value="8" <% if ForumAuthType = 8 then Response.Write(" selected") %>>NT Global Group (Hidden)</option>
<%
	end if
%>
        <option value="13" <% if ForumAuthType = 13 then Response.Write("selected") %>>Groups</option>
        <option value="14" <% if ForumAuthType = 14 then Response.Write("selected") %>>Groups (Hidden)</option>
        </select>
<%
		if strRqMethod = "Forum" or _
		strRqMethod = "EditForum" then 
			if strRqMethod = "EditForum" then 
				If pass_new <> " " Then
					strPassword = ChkString(pass_new,"password")
				else 
					strPassword = " "
				end if
			else
				strPassword = " "
			end if
%>
        <br /><b>Password <% if strNTGroups = "1" then Response.Write("or Global Groups") %>:</b><br /> 
        <input maxLength="255" type="text" name="AuthPassword" id="AuthPassword" size="60" value="<%=strPassword%>" <% if ForumAuthType >= 13 then Response.Write(" disabled=""disabled""") %>>
<%
		end if 
%>
          <br /><span id="groups"><b>Groups: </b><br />
          <select name="AuthPassword" id="authPassGroup" multiple size="4" <% 'if ForumAuthType < 13 then Response.Write("disabled=""disabled""") %>>
          <% 
          gSql = "select G_ID, G_NAME from " & strTablePrefix & "GROUPS order by g_name;"
          set gSqlRs = my_Conn.execute(gSql)
          
          do While NOT gSqlRs.EOF
               response.Write "<option value=""" & gSqlRs("G_ID") & """"
               if instr(strPassword, ",") > 0 then
                    strPass = Split(strPassword, ",")
                    for i=0 to ubound(strPass)
                         strPassClean = cLng(trim(strPass(i)))
                         if strPassClean = gSqlRs("G_ID") then
                              response.write " selected"
                         end if
                    next
               elseif len(trim(strPassword&" ")) > 0 then
                    if cLng(trim(strPassword)) = gSqlRs("G_ID") then
                         response.write " selected"
                    end if
               end if
               response.write ">" & gSqlRs("G_NAME") & "</option>" & vbCrLf
               gSqlRs.MoveNext
          Loop
          %>
          </select></span>
        </td>
      </tr>
      <tr>
        <td noWrap vAlign="top" align="right"><b>Member List:</b></td>
<%
		'#################################################################################
		'## Allowed User - listbox Code
		'#################################################################################
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"
		'on error resume next
		'set rsMember = my_Conn.execute (strSql)

		strSql = "SELECT " & strMemberTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID "
		strSql = strSql & " FROM " & strMemberTablePrefix & "ALLOWED_MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "ALLOWED_MEMBERS.FORUM_ID = " & strRqForumID

		set rsAllowedMember = my_Conn.execute (strSql)
%>
	<td>
		<table><tr>
			<td>
				<b>Forum Members:</b><br />
				<a href="JavaScript:allowmembers();">Memberlist</a>
		</td>
		<td width="15" align="center" valign="middle">
			<a href="javascript:DeleteSelection()"><img src="images/icons/icon_private_remove.gif" width="23" height="22" border="0" alt="Remove Selected Member(s)"></a>	
		</td>
	<td>
		<b>Select Members:</b><br />
		<select name="AuthUsers" size="<%=SelectSize %>" multiple >
<%
	'## Selected List
	'rsAllowedMember.movefirst
	if strRqMethod = "EditForum" or strRqMethod = "EditURL" then	
	do until rsAllowedMember.EOF
		Response.Write "          <option value=""" & rsAllowedMember("MEMBER_ID") & """>" & ChkString(getMemberName(rsAllowedMember("MEMBER_ID")),"display") & "</option>" & vbCrLf
		rsAllowedMember.movenext
	loop
	end if
	set rsAllowedMember = nothing
%>
        	<option value="<% if tmpStrUserList <> "" then Response.Write tmpStrUserList end if %>"></option>
		</select>
    </td>
  </tr>
</table>
	</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" value="<% =btn %>" onclick="selectUsers()" class="button">&nbsp;<input name="Reset" type="reset" value="Clear Fields" class="button"></td>
      </tr>
<%
		'#################################################################################
		'## Allowed User - End of listbox code
		'#################################################################################
	end if
end if 
%>
    </table></form>
    </td>
  </tr></table>
  
<%
spThemeBlock1_close(intSkin)%>

<%
if strRqMethod = "Reply" or _
strRqMethod = "TopicQuote" or _
strRqMethod = "ReplyQuote" then

spThemeSmallBlock_open()%>
      <tr>
        <td class="tSubTitle" colspan="2" align="center"><b>Topic Review (Newest First)</b></td>
      </tr>
<%
	' - Get all replies to Topic from the DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.R_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & strRqTopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE DESC"

	set rs = Server.CreateObject("ADODB.Recordset")

'	rs.cachesize=15
	rs.open  strSql, my_Conn, 3

	strI = 0 
	if rs.EOF or rs.BOF then
		Response.Write ""
	else
		rs.movefirst
		do until rs.EOF
			if strI = 0 then
	 			CColor = "tCellAlt1"
			else
				CColor = "tCellAlt2"
			end if
			Response.Write "      <TR>" & vbCrLf   & _
						   "        <TD class='" & CColor & "' valign='top' nowrap>"
			Response.Write "<b>" &  ChkString(rs("M_NAME"),"display") & "</b></td>" & vbCrLf & _
						   "        <TD class='" & CColor & "' valign='top'"
			Response.Write ">" & formatStr(rs("R_MESSAGE")) & "</td>" & vbCrLf & _
						   "      </tr>" & vbCrLf
			rs.MoveNext
			strI = strI + 1
			if strI = 2 then 
				strI = 0
			end if
		loop
	end if
	
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "TOPICS.T_MESSAGE " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR AND "
	strSql = strSql & "       " & strTablePrefix & "TOPICS.TOPIC_ID = " &  strRqTopicID

	set rs = my_Conn.Execute (strSql) 

	Response.Write "      <tr>" & vbCrLf
	Response.Write "        <td class=""tCellAlt0"" valign=top width='120' nowrap>"
	Response.Write "<b>" & ChkString(rs("M_NAME"),"display") & "</b></td>" & vbCrLf
	Response.Write "        <td class=""tCellAlt1"" valign='top' width='100%'>" 
	Response.Write formatStr(rs("T_MESSAGE")) & "</td>" & vbCrLf
	Response.Write "      </tr>" & vbCrLf
	spThemeSmallBlock_close()
	Response.Write	""  
end if
%>
    </td></tr></table><!--#INCLUDE FILE="inc_footer.asp" -->

<%
sub Go_Result(str_err_Msg)
%>
<table border="0" width="100%">
  <tr>
	<td width="33%" align="left">
	<img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">All Forums</a><br />
<% 
	if strRqMethod = "Edit" or _
	        strRqMethod = "EditTopic" or _
		strRqMethod = "Reply" or _
		strRqMethod = "ReplyQuote" or _
		strRqMethod = "TopicQuote" then
%>
	<img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="FORUM.asp?CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(Request.QueryString("FORUM_Title"),"sqlstring") %></a><br />
	<img src="images/icons/icon_blank.gif" height=15 width=15 border="0"><img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<a href="forum_topic.asp?TOPIC_ID=<% =strRqTopicID %>&CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(left(Request.QueryString("Topic_title"), 50),"urlpath") %>"><% =ChkString(Request.QueryString("Topic_Title"),"sqlstring") %></a>
<% 
	end if 
%>
    </td>
  </tr>
</table>

<p align="center"><span class="fTitle">There has been a problem!</span></p>
<p align="center"><span class="fAlert"><% =str_err_Msg %></span></p>
<p align="center"><a href="JavaScript:history.go(-1)">Go back to correct the problem.</a></p>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
Response.End
end sub
%>