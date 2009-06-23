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
<!--#INCLUDE FILE="inc_top.asp" -->
<%
intSkin = getSkin(intSubSkin,2)

if strDBNTUserName = "" then
	strTmpPassword = pEncrypt(pEnPrefix & chkString(Request.Form("Password"),"sqlstring"))
else
	strTmpPassword = ChkString(Request.Form("Password"), "sqlstring")
end if

if strAuthType = "db" then
	strDBNTUserName = ChkString(Request.Form("UserName"), "sqlstring")
end if

set rs = Server.CreateObject("ADODB.RecordSet")

err_Msg = ""
ok = "" 

MethodType = chkString(Request.Form("Method_Type"),"SQLString")
MethodTypeP = chkString(Request.Querystring("Method_Type"),"display")
'response.Write(MethodTypeP & "<br />")
if Request.Form("CAT_ID") <> "" then
	if IsNumeric(Request.Form("CAT_ID")) = True then
		Cat_ID = cLng(Request.Form("CAT_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.Form("FORUM_ID") <> "" then
	if IsNumeric(Request.Form("FORUM_ID")) = True then
		Forum_ID = cLng(Request.Form("FORUM_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.form("TOPIC_ID") <> "" then
	if IsNumeric(Request.form("TOPIC_ID")) = True then
		Topic_ID = cLng(Request.form("TOPIC_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.Querystring("TOPIC_IDP") <> "" then
	if IsNumeric(Request.Querystring("TOPIC_IDP")) = True then
		Topic_IDP = cLng(Request.Querystring("TOPIC_IDP"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.Querystring("POLL_ID") <> "" then
	if IsNumeric(Request.Querystring("POLL_ID")) = True then
		POLL_ID = cLng(Request.Querystring("POLL_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if
if Request.Form("REPLY_ID") <> "" then
	if IsNumeric(Request.Form("REPLY_ID")) = True then
		Reply_ID = cLng(Request.Form("REPLY_ID"))
	else
		Response.Redirect("fhome.asp")
	end if
end if

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::            makePoll
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

if MethodTypeP = "makePoll" and Topic_ID <> "" and hasAccess(2) then
	tmpUserId = getMemberId(ChkString(Request.Cookies(strUniqueID & "User")("Name"), "title"))
	if strAuthType = "db" then

		strSql = "SELECT TOPIC_ID "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID
		if not hasAccess(1) then
			strSql = strSql & " AND T_AUTHOR = " & tmpUserId 
		end if
		Set rsCheck = my_Conn.Execute(strSql)

		if rsCheck.BOF or rsCheck.EOF then
			Go_Result "Wrong user!", 0
		end if
		set rsCheck = nothing
	end if

	pollQuestion = ChkString(Request.Form("pollQuestion"),"SQLString")
	pollAns1 = ChkString(Request.Form("pollAns1"),"SQLString")
	pollAns2 = ChkString(Request.Form("pollAns2"),"SQLString")
	pollAns3 = ChkString(Request.Form("pollAns3"),"SQLString")
	pollAns4 = ChkString(Request.Form("pollAns4"),"SQLString")
	pollAns5 = ChkString(Request.Form("pollAns5"),"SQLString")
	pollAns6 = ChkString(Request.Form("pollAns6"),"SQLString")
	pollAns7 = ChkString(Request.Form("pollAns7"),"SQLString")
	pollAns8 = ChkString(Request.Form("pollAns8"),"SQLString")
	pollAns9 = ChkString(Request.Form("pollAns9"),"SQLString")
	pollAns10 = ChkString(Request.Form("pollAns10"),"SQLString")
	pollAns11 = ChkString(Request.Form("pollAns11"),"SQLString")
	pollAns12 = ChkString(Request.Form("pollAns12"),"SQLString")
	if Request.Form("pollMultiple") = "1" then 'multiple
		pollMultiple = "1"
	else
		pollMultiple = "0"
	end if
	if Request.Form("pollGuest") = "1" then 'allow guest
		pollGuest = "1"
	else
		pollGuest = "0"
	end if
	if Request.Form("pollExpire") <> "" then
		if IsNumeric(Request.Form("pollExpire")) = True then
			pollExpire = cLng(Request.Form("pollExpire"))
		else
			pollExpire = 0
		end if
	else
		pollExpire = 0
	end if

	if pollQuestion = " " then
		Go_Result "You must post a Question", 0 
		%><!--INCLUDE FILE="inc_footer.asp" --><%
		'Response.End
	end if
			
	if pollAns1 = " " then
		Go_Result "You must post at least one answer", 0
		%><!--INCLUDE FILE="inc_footer.asp" --><%
		'Response.End
	end if

	' - Add new post to Topics Table
	strSql = "INSERT INTO " & strTablePrefix & "POLLS (POLL_TYPE"
	strSql = strSql & ", POLL_ALLOW"
	strSql = strSql & ", POLL_QUESTION"
	strSql = strSql & ", ANSWER1"
	if trim(pollAns2) <> "" then
		strSql = strSql & ", ANSWER2"
	end if
	if trim(pollAns3) <> "" then
		strSql = strSql & ", ANSWER3"
	end if
	if trim(pollAns4) <> "" then
		strSql = strSql & ", ANSWER4"
	end if
	if trim(pollAns5) <> "" then
		strSql = strSql & ", ANSWER5"
	end if
	if trim(pollAns6) <> "" then
		strSql = strSql & ", ANSWER6"
	end if
	if trim(pollAns7) <> "" then
		strSql = strSql & ", ANSWER7"
	end if
	if trim(pollAns8) <> "" then
		strSql = strSql & ", ANSWER8"
	end if
	if trim(pollAns9) <> "" then
		strSql = strSql & ", ANSWER9"
	end if
	if trim(pollAns10) <> "" then
		strSql = strSql & ", ANSWER10"
	end if
	if trim(pollAns11) <> "" then
		strSql = strSql & ", ANSWER11"
	end if
	if trim(pollAns12) <> "" then
		strSql = strSql & ", ANSWER12"
	end if
	strSql = strSql & ", POST_DATE"
	strSql = strSql & ", END_DATE"
	strSql = strSql & ", POLL_AUTHOR"
	strSql = strSql & ") VALUES ("
	strSql = strSql & pollMultiple
	strSql = strSql & ", " & pollGuest
	strSql = strSql & ", '" & pollQuestion & "'"
	strSql = strSql & ", '" & pollAns1 & "'"
	if trim(pollAns2) <> "" then
		strSql = strSql & ", '" & pollAns2 & "'"
	end if
	if trim(pollAns3) <> "" then
		strSql = strSql & ", '" & pollAns3 & "'"
	end if
	if trim(pollAns4) <> "" then
		strSql = strSql & ", '" & pollAns4 & "'"
	end if
	if trim(pollAns5) <> "" then
		strSql = strSql & ", '" & pollAns5 & "'"
	end if
	if trim(pollAns6) <> "" then
		strSql = strSql & ", '" & pollAns6 & "'"
	end if
	if trim(pollAns7) <> "" then
		strSql = strSql & ", '" & pollAns7 & "'"
	end if
	if trim(pollAns8) <> "" then
		strSql = strSql & ", '" & pollAns8 & "'"
	end if
	if trim(pollAns9) <> "" then
		strSql = strSql & ", '" & pollAns9 & "'"
	end if
	if trim(pollAns10) <> "" then
		strSql = strSql & ", '" & pollAns10 & "'"
	end if
	if trim(pollAns11) <> "" then
		strSql = strSql & ", '" & pollAns11 & "'"
	end if
	if trim(pollAns12) <> "" then
		strSql = strSql & ", '" & pollAns12 & "'"
	end if
	strSql = strSql & ", '" & strCurDateString & "'"
	strSql = strSql & ", '" & datetostr(DateAdd("d","+"&pollExpire,strCurDateAdjust)) & "'"
	strSql = strSql & ", " & tmpUserId
	strSql = strSql & " )"
	executeThis(strSql)
		
	strSql = "SELECT POLL_ID "
	strSql = strSql & " FROM " & strTablePrefix & "POLLS "
	strSql = strSql & " WHERE POLL_QUESTION = '" & pollQuestion & "'"
	strSql = strSql & " AND POLL_AUTHOR = " & tmpUserId 
	strSql = strSql & " ORDER BY POST_DATE DESC"
	Set rsCheck = my_Conn.Execute(strSql)
'response.Write(strSql & "<br />")
	if rsCheck.BOF or rsCheck.EOF then
		Go_Result "Error! in Poll", 0
	else
		newPollId = rsCheck("POLL_ID")
	end if
	set rsCheck = nothing
		
	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & " SET T_POLL         = " & newPollId & ""
	strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID
	my_Conn.Execute (strSql)

	if hasAccess(1) and request.form("featuredPoll") = "1" then
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_FEATUREDPOLL = " & newPollId & ""
		strSql = strSql & " WHERE CONFIG_ID = " & 1
		my_Conn.Execute (strSql)
		Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
	end if
	if Err.description <> "" then 
		err_Msg = "There was an error = " & Err.description
	else
		err_Msg = "Updated OK"
	end if

	Go_Result err_Msg, 1
	'Response.End
end if

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::            Create Poll; makePoll; editPoll; updatePoll;
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

if MethodTypeP = "CreatePoll" or MethodTypeP = "makePoll" or MethodTypeP = "editPoll" or MethodTypeP = "updatePoll" then
'response.Write("ok here 2:" & MethodTypeP & ": <br />")
	if (strPollCreate <> 0) and ((strPollCreate = 1 and hasAccess(2)) or (strPollCreate = 2 and mLev >=3) or (strPollCreate = 3 and hasAccess(1))) then

		if MethodTypeP = "CreatePoll" and Topic_IDP <> "" and hasAccess(2) then
			if trim(Request.querystring("pollNum")) <> "" then
				if IsNumeric(Request.querystring("pollNum")) = True then
					pollNum = cLng(Request.querystring("pollNum"))
				else
					Response.Redirect("fhome.asp")
				end if
			end if %>
			<script type="text/javascript">
			function ChangePage(pollNum){
			document.location.href="forum_post_info.asp?Method_Type=CreatePoll&TOPIC_IDP=<%=Topic_IDP%>&pollNum="+pollNum;
			}
			</script>
			<table border="0" width="100%">
  			  <tr>
				<td width="33%" align="left">
				  <img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
				  <img src="images/icons/icon_blank.gif" border="0"><img src="images/icons/icon_bar.gif" border="0">
				  <img src="images/icons/icon_folder_open_topic.gif" border="0">&nbsp;Create Poll</td>
			  </tr>
			</table>
<%
spThemeTableCustomCode = "align=""center"" width=""33%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
  <tr>
    <td>
    <table border="0" cellspacing="1" cellpadding="1" align="center">
<form action="forum_post_info.asp?Method_Type=makePoll" method="post" name="createPoll">
<input name="TOPIC_ID" type="hidden" value="<%=Topic_IDP%>">
<input name="Method_Type" type="hidden" value="makePoll">
<input name="refer" type="hidden" value="link.asp?TOPIC_ID=<%=Topic_IDP%>">
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Poll Question:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollQuestion" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right" colspan="2">
Number of Answers:
<select name="ansNum" size="1" onchange="ChangePage(this.value)">
<OPTION VALUE="4" <%if pollNum = "4" or pollNum = "" then response.write "SELECTED" end if%>>4
<OPTION VALUE="8" <%if pollNum = "8" then response.write "SELECTED" end if%>>8
<OPTION VALUE="12" <%if pollNum = "12" then response.write "SELECTED" end if%>>12
</select>
        </td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 1:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns1" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 2:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns2" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 3:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns3" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 4:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns4" size="40"></td>
      </tr>
<%if pollNum = "8" or pollNum = "12" then %>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 5:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns5" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 6:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns6" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 7:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns7" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 8:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns8" size="40"></td>
      </tr>
<%end if
if pollNum = "12" then %>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 9:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns9" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 10:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns10" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 11:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns11" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 12:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns12" size="40"></td>
      </tr>
<%end if%>
      <tr>
        <td class="tCellAlt0" align="right" colspan="2"><b>Allow multiple choice</b>&nbsp;<input name="pollMultiple" type="checkbox" value="1"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="2"><b>Allow guests to vote</b>&nbsp;<input name="pollGuest" type="checkbox" value="1"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="2"><b>Poll expires after &nbsp;<input name="pollExpire" type="text" maxLength="4" value="0" size="3">&nbsp;days</b>&nbsp;(0 never expires)</td>
      </tr>
<%if hasAccess(1) then%>
      <tr>
        <td class="tCellAlt0" align="right" colspan="2"><b>Make this a featured poll</b>&nbsp;<input name="featuredPoll" type="checkbox" value="1"></td>
      </tr>
<%end if%>
      <tr>
        <td colspan="2" align="center" class="tCellAlt0"><input name="Submit" type="submit" value="Create Poll" accesskey="s" title="Shortcut Key: Alt+S" class="button">&nbsp;<input name="Reset" type="reset" value="Reset Fields" class="button" class="button"></td>
      </tr></form>
      </table>
      </td></tr></table>
	  <%
spThemeBlock1_close(intSkin)%>
<script language="JavaScript" type="text/javascript">document.createPoll.pollQuestion.focus();</script>
	<p align=center><a href="link.asp?TOPIC_ID=<%=Topic_IDP%>">Cancel</a></p>
<%
end if

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

		if MethodTypeP = "editPoll" and Topic_IDP <> "" and hasAccess(1) then
			strSql = "SELECT POLL_TYPE, POLL_ID, POLL_ALLOW, POLL_QUESTION," 
        	strSql = strSql & " ANSWER1, ANSWER2, ANSWER3, ANSWER4, ANSWER5, ANSWER6, ANSWER7, ANSWER8, ANSWER9, ANSWER10, ANSWER11, ANSWER12,"
	        strSql = strSql & " RESULT1, RESULT2, RESULT3, RESULT4, RESULT5, RESULT6, RESULT7, RESULT8, RESULT9, RESULT10, RESULT11, RESULT12,"
	        strSql = strSql & " POST_DATE, END_DATE, POLL_AUTHOR "
		strSql = strSql & " FROM " & strTablePrefix & "POLLS "
		strSql = strSql & " WHERE POLL_ID = " & POLL_ID
		Set rsEditPoll = my_Conn.Execute(strSql)

		if not (rsEditPoll.BOF or rsEditPoll.EOF) then
			if rsEditPoll("POLL_TYPE") = "0" then
			strPollType = 0
			else
			strPollType = 1
			end if
			strPollAllow = rsEditPoll("POLL_ALLOW")
			strPollQuestion = rsEditPoll("POLL_QUESTION")
		
			strPollAns1 = rsEditPoll("ANSWER1")
			strPollAns2 = rsEditPoll("ANSWER2")
			strPollAns3 = rsEditPoll("ANSWER3")
			strPollAns4 = rsEditPoll("ANSWER4")
			strPollAns5 = rsEditPoll("ANSWER5")
			strPollAns6 = rsEditPoll("ANSWER6")
			strPollAns7 = rsEditPoll("ANSWER7")
			strPollAns8 = rsEditPoll("ANSWER8")
			strPollAns9 = rsEditPoll("ANSWER9")
			strPollAns10 = rsEditPoll("ANSWER10")
			strPollAns11 = rsEditPoll("ANSWER11")
			strPollAns12 = rsEditPoll("ANSWER12")
			
			strPostDate = rsEditPoll("POST_DATE")
			strEndDate = rsEditPoll("END_DATE")


	if rsEditPoll("RESULT1") <> "" then
		strPollRes1 = cInt(rsEditPoll("RESULT1"))
	else
		strPollRes1 = 0
        end if
	if rsEditPoll("RESULT2") <> "" then
		strPollRes2 = cInt(rsEditPoll("RESULT2"))
	else
		strPollRes2 = 0
        end if
	if rsEditPoll("RESULT3") <> "" then
		strPollRes3 = cInt(rsEditPoll("RESULT3"))
	else  
		strPollRes3 = 0
        end if		
	if rsEditPoll("RESULT4") <> "" then
		strPollRes4 = cInt(rsEditPoll("RESULT4"))
	else
		strPollRes4 = 0
        end if
	if rsEditPoll("RESULT5") <> "" then
		strPollRes5 = cInt(rsEditPoll("RESULT5"))
	else  
		strPollRes5 = 0
        end if		
	if rsEditPoll("RESULT6") <> "" then
		strPollRes6 = cInt(rsEditPoll("RESULT6"))
	else  
		strPollRes6 = 0
        end if		
	if rsEditPoll("RESULT7") <> "" then
		strPollRes7 = cInt(rsEditPoll("RESULT7"))
	else  
		strPollRes7 = 0
        end if		
	if rsEditPoll("RESULT8") <> "" then
		strPollRes8 = cInt(rsEditPoll("RESULT8"))
	else  
		strPollRes8 = 0
        end if		
	if rsEditPoll("RESULT9") <> "" then
		strPollRes9 = cInt(rsEditPoll("RESULT9"))
	else		
		strPollRes9 = 0
        end if
	if rsEditPoll("RESULT10") <> "" then
		strPollRes10 = cInt(rsEditPoll("RESULT10"))
	else  
		strPollRes10 = 0
        end if		
	if rsEditPoll("RESULT11") <> "" then
		strPollRes11 = cInt(rsEditPoll("RESULT11"))
	else  
		strPollRes11 = 0
        end if		
	if rsEditPoll("RESULT12") <> "" then
		strPollRes12 = cInt(rsEditPoll("RESULT12"))
	else  
		strPollRes12 = 0
        end if
		else
			Go_Result "Wrong poll!", 0
		end if
		'rsEditPoll.close
		set rsEditPoll = nothing
%>
<table border="0" width="100%">
  <tr>
	<td width="33%" align="left">
	<img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
	<img src="images/icons/icon_blank.gif" border="0"><img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open_topic.gif" border="0">&nbsp;Edit <a href="link.asp?TOPIC_ID=<%=Topic_IDP%>">Poll</a>
    </td>
  </tr>
</table>
<%
spThemeTableCustomCode = "align=""center"" width=""40%"""
spThemeBlock1_open(intSkin)
%>
<table cellpadding="0" cellspacing="0">
  <tr>
    <td>
    <table border="0" cellspacing="1" cellpadding="1">
<form action="forum_post_info.asp?Method_Type=updatePoll" method="post" name="pollform">
<input name="TOPIC_ID" type="hidden" value="<%=Topic_IDP%>">
<input name="POLL_ID" type="hidden" value="<%=POLL_ID%>">
<input name="Method_Type" type="hidden" value="updatePoll">
<input name="refer" type="hidden" value="link.asp?TOPIC_ID=<%=Topic_IDP%>">
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Poll Question:</b></td>
	<td colspan="2" class="tCellAlt0"><input maxLength="50" name="pollQuestion" value="<%=strPollQuestion%>" size="40"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 1:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns1" value="<%=strPollAns1%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes1" value="<%=strPollRes1%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 2:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns2" value="<%=strPollAns2%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes2" value="<%=strPollRes2%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 3:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns3" value="<%=strPollAns3%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes3" value="<%=strPollRes3%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 4:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns4" value="<%=strPollAns4%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes4" value="<%=strPollRes4%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 5:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns5" value="<%=strPollAns5%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes5" value="<%=strPollRes5%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 6:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns6" value="<%=strPollAns6%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes6" value="<%=strPollRes6%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 7:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns7" value="<%=strPollAns7%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes7" value="<%=strPollRes7%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 8:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns8" value="<%=strPollAns8%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes8" value="<%=strPollRes8%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 9:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns9" value="<%=strPollAns9%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes9" value="<%=strPollRes9%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 10:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns10" value="<%=strPollAns10%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes10" value="<%=strPollRes10%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 11:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns11" value="<%=strPollAns11%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes11" value="<%=strPollRes11%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" noWrap vAlign="top" align="right"><b>
Answer 12:</b></td>
	<td class="tCellAlt0"><input maxLength="50" name="pollAns12" value="<%=strPollAns12%>" size="40"></td>
	<td class="tCellAlt0"><input maxLength="5" name="pollRes12" value="<%=strPollRes12%>" size="5"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="3"><b>Allow multiple choice</b>&nbsp;<input name="pollMultiple" type="checkbox" value="1" <% if strPollType = "1" then Response.Write("checked")%>></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="3"><b>Allow guests to vote</b>&nbsp;<input name="pollGuest" type="checkbox" value="1" <% if strPollAllow = "1" then Response.Write("checked")%>></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="3"><span class="fSmall"><b>Starting from now, poll will expire in &nbsp;<input name="pollExpire" type="text" maxLength="4" value="<% =DateDiff("d", strtodate(strPostDate), strtodate(strEndDate))%>" size="3">&nbsp;days </b>&nbsp;(0 never expires)</span></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" colspan="3"><b>Make this a featured poll</b>&nbsp;<input name="featuredPoll" type="checkbox" value="1"<%= chkRadio(cint(strFeaturedPoll),cint(POLL_ID)) %>></td>
      </tr>
      <tr>
        <td colspan="3" align="center" class="tCellAlt0"><input name="Submit" type="submit" value="Edit Poll" accesskey="s" title="Shortcut Key: Alt+S" class="button">&nbsp;<input name="Reset" type="reset" value="Reset Fields" class="button"></td>
      </tr></form>
      </table>
      </td></tr></table>
	  <%
spThemeBlock1_close(intSkin)%>
<%
end if

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
		if MethodTypeP = "updatePoll" and request.form("POLL_ID") <> "" and hasAccess(1) then

			if request.form("pollMultiple") = "1" then
			strPollType = 1
			else
			strPollType = 0
			end if
			if request.form("pollGuest") = "1" then
			strPollAllow = 1
			else
			strPollAllow = 0
			end if
			strPollQuestion = chkString(request.form("pollQuestion"),"sqlstring")
		
		strSql = "UPDATE " & strTablePrefix & "POLLS "
		strSql = strSql & " SET POLL_QUESTION = '" & strPollQuestion & "'"
		strSql = strSql & ", POLL_TYPE         = " & strPollType & ""
		strSql = strSql & ", POLL_ALLOW         = " & strPollAllow & ""

		strSql = strSql & ", ANSWER1 = '" & chkString(request.form("pollAns1"), "SQLString") & "'"
		strSql = strSql & ", ANSWER2 = '" & chkString(request.form("pollAns2"), "SQLString") & "'"
		strSql = strSql & ", ANSWER3 = '" & chkString(request.form("pollAns3"), "SQLString") & "'"
		strSql = strSql & ", ANSWER4 = '" & chkString(request.form("pollAns4"), "SQLString") & "'"
		strSql = strSql & ", ANSWER5 = '" & chkString(request.form("pollAns5"), "SQLString") & "'"
		strSql = strSql & ", ANSWER6 = '" & chkString(request.form("pollAns6"), "SQLString") & "'"
		strSql = strSql & ", ANSWER7 = '" & chkString(request.form("pollAns7"), "SQLString") & "'"
		strSql = strSql & ", ANSWER8 = '" & chkString(request.form("pollAns8"), "SQLString") & "'"
		strSql = strSql & ", ANSWER9 = '" & chkString(request.form("pollAns9"), "SQLString") & "'"
		strSql = strSql & ", ANSWER10 = '" & chkString(request.form("pollAns10"), "SQLString") & "'"
		strSql = strSql & ", ANSWER11 = '" & chkString(request.form("pollAns11"), "SQLString") & "'"
		strSql = strSql & ", ANSWER12 = '" & chkString(request.form("pollAns12"), "SQLString") & "'"

		strSql = strSql & ", RESULT1 = '" & chkString(request.form("pollRes1"), "SQLString") & "'"
		strSql = strSql & ", RESULT2 = '" & chkString(request.form("pollRes2"), "SQLString") & "'"
		strSql = strSql & ", RESULT3 = '" & chkString(request.form("pollRes3"), "SQLString") & "'"
		strSql = strSql & ", RESULT4 = '" & chkString(request.form("pollRes4"), "SQLString") & "'"
		strSql = strSql & ", RESULT5 = '" & chkString(request.form("pollRes5"), "SQLString") & "'"
		strSql = strSql & ", RESULT6 = '" & chkString(request.form("pollRes6"), "SQLString") & "'"
		strSql = strSql & ", RESULT7 = '" & chkString(request.form("pollRes7"), "SQLString") & "'"
		strSql = strSql & ", RESULT8 = '" & chkString(request.form("pollRes8"), "SQLString") & "'"
		strSql = strSql & ", RESULT9 = '" & chkString(request.form("pollRes9"), "SQLString") & "'"
		strSql = strSql & ", RESULT10 = '" & chkString(request.form("pollRes10"), "SQLString") & "'"
		strSql = strSql & ", RESULT11 = '" & chkString(request.form("pollRes11"), "SQLString") & "'"
		strSql = strSql & ", RESULT12 = '" & chkString(request.form("pollRes12"), "SQLString") & "'"
		
		strSql = strSql & ", POST_DATE = '" & strCurDateString & "'"
		strSql = strSql & ", END_DATE = '" & datetostr(DateAdd("d","+" & chkString(request.form("pollExpire"),"sqlstring"),strCurDateAdjust)) & "'"
       		strSql = strSql & " WHERE POLL_ID = " & chkstring(request.form("POLL_ID"), "sqlstring")
       
       'response.Write(strSql & "<br />")		
		my_Conn.Execute (strSql)
		
		if hasAccess(1) then
		  if request.form("featuredPoll") = "1" then
            strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_FEATUREDPOLL = " & cint(request.form("POLL_ID")) & ""
			strSql = strSql & " WHERE CONFIG_ID = 1"
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		  else
		    if cint(strFeaturedPoll) = cint(request.form("POLL_ID")) then
                strSql = "UPDATE " & strTablePrefix & "CONFIG "
			    strSql = strSql & " SET C_FEATUREDPOLL = 0"
			    strSql = strSql & " WHERE CONFIG_ID = 1"
			    executeThis(strSql)
			    Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
			end if
		  end if
		end if
		Go_Result err_Msg, 1
	end if

	else
		Go_Result  "You are not authorized to create or edit polls", 0
	end if
end if

'******************************************************************************************
'***		EDIT reply
'******************************************************************************************
if MethodType = "Edit" then
	err.clear
	'error.clear
	member = cint(ChkUser(strDBNTUserName, chkString(Request.Form("Password"),"sqlstring")))
	Select Case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only Admins, Moderators and the Author can change this post", 0
		case 3 '## Moderator so OK - check the Moderator of this forum
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only Admins, Moderators and the Author can change this post", 0
			end if
		case 4 '## Admin so OK
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	'txMessage = chkString(chkHtmlCode(Request.Form("Message")),"message")
	Err_Msg = ""
	if txMessage = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for your Reply</li>"
	end if
	if Err_Msg = "" then
		if strEditedByDate = "1" and not hasAccess(1) then
			txMessage = txMessage & "<br /><br />Edited by - "
			txMessage = txMessage & ChkString(STRdbntUserName, "display") & " on " & ChkDate(strCurDateString) & " " & ChkTime(strCurDateString)
		end if

		' - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "REPLY "
		strSql = strSql & " SET R_MESSAGE = '" & txMessage & "'"
		if lcase(strEmail) = "1" then '**
			if Request.Form("rmail") <> "1" then
				TF = "0"
			else 
				TF = "1"
			end if
			strSql = strSql & ", R_MAIL = " & TF
		end if
		strSql = strSql & " WHERE REPLY_ID=" & Reply_ID

		executeThis(strSql)
	  
		'Check if they want subscribed to topic
	  if Request.Form("rmail") = 1 then
	    if intSubscriptions = 1 and strEmail = 1 then
		  sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
		  set rsAP = my_Conn.execute(sSql)
		  if not rsAP.eof then
	  	    intAppID = rsAP("APP_ID")
		  end if
		  set rsAP = nothing
		
		    sSql = "SELECT TOPIC_ID FROM "& strTablePrefix & "REPLY WHERE REPLY_ID=" & Reply_ID
		    set rsA = my_Conn.execute(sSql)
		    if not rsA.eof then 
		      topic_id = rsA("TOPIC_ID")
		      itmTitle = chkString(Request.Form("Topic_Title"),"display")
		    end if
			set rsA = nothing
	        sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=" & topic_id
	        set rsX = my_Conn.execute(sSql)
	        If rsX.BOF or rsX.EOF Then
	          ' subscription doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & strUserMemberID & ", " & intAppID & ", 0, 0, " & topic_id & ", '" & itmTitle & "')"
	          executeThis(insSql)
			  strSubTxt = "You are now subscribed to this topic<br /><br />"
			  strSubTxt = strSubTxt & "You will recieve an email whenever<br />"
			  strSubTxt = strSubTxt & "anyone replies to this topic<br />"
		    else
			  strSubTxt = "You are already subscribed to this topic<br /><br />"
	        End If
	        set rsX = nothing
		end if
	  end if

		if not hasAccess(1) then
			' - Update Last Post
			strSql = " UPDATE " & strTablePrefix & "FORUM"
			strSql = strSql & " SET F_LAST_POST = '" & strCurDateString & "'"
			strSql = strSql & ",    F_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

			executeThis(strSql)
			
			' - Update Last Post

			strSql = " UPDATE " & strTablePrefix & "TOPICS"
			strSql = strSql & " SET T_LAST_POST = '" & strCurDateString & "'"
			strSql = strSql & ",    T_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			executeThis(strSql)
		
		end if

		err_Msg = ""

		'
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_LAST_POST = '" & strCurDateString & "'"
		strSql = strSql & ",    T_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

		executeThis(strSql)

		err_Msg = ""
		'if Err.description <> "" then 
		'	Go_Result "There was an error = " & Err.description, 0
		'else
			Go_Result  "Updated OK", 1
		'end if
	else ' there is an error message %>
	<p align=center><span class="fTitle">There Was A Problem</span></p>
	<table align=center border="0">
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>
	<p align=center><a href="JavaScript:history.go(-1)">Go Back</a></p>
<%
	end if
end if

'******************************************************************************************
'***		END EDIT REPLY
'******************************************************************************************

'******************************************************************************************
'***		EDIT TOPIC
'******************************************************************************************
if MethodType = "EditTopic" then
	member = cint(ChkUser(STRdbntUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member 
		case 0 '## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only Admins, Moderators and the Author can change this post", 0
		case 3 '## Moderator so 
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only Admins, Moderators and the Author can change this post", 0
			end if
		case 4 '## Admin so OK
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	'txMessage = chkString(chkHtmlCode(Request.Form("Message")),"message")
	txSubject = left(replace(ChkString(Request.Form("Subject"),"sqlstring"),"'","''"),50)
	Err_Msg = ""

	if len(trim(txSubject) & "x") = 1 then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Topic</li>"
	end if
	if len(trim(txMessage) & "x") = 1 then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the Topic</li>"
	end if
	if Err_Msg = "" then
		if strEditedByDate = "1" and not hasAccess(1) then
			txMessage = txMessage & "<br /><br />Edited by - "
			txMessage = txMessage & ChkString(STRdbntUserName, "display") & " on " & ChkDate(strCurDateString) & " " & ChkTime(strCurDateString)
		end if

		'## Set array to pull out CAT_ID and FORUM_ID from dropdown values in post.asp
		aryForum = split(Request.Form("Forum"), "|")
		aryForum(0) = chkString(aryForum(0),"numeric")
		aryForum(1) = chkString(aryForum(1),"numeric")

		'
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_MESSAGE = '" & txMessage & "'"
		strSql = strSql & ",     T_SUBJECT = '" & txSubject & "'"
		if Forum_ID <> "" and Forum_ID <> aryForum(1) then
			strSql = strSql & ", CAT_ID = " & aryForum(0)
			strSql = strSql & ", FORUM_ID = " & aryForum(1)
		end if
		if Request.Form("news") = 1 and mlev >=3 then
			strsql = strsql & ", T_NEWS = 1 "		
		else
			strsql = strsql & ", T_NEWS = 0 "
		end if
		if Request.Form("sig") = "yes" then
			strsql = strsql & ", T_SIG = 1 "
		else
			strsql = strsql & ", T_SIG = 0 "
		end if
		if lcase(strEmail) = "1" then
			if Request.Form("rmail") <> "1" then
				TF = "0"
			else 
				TF = "1"
			end if
			strSql = strSql & ", T_MAIL = " & TF
		end if
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID
			'response.Write(strSql & "<br />")

		executeThis(strSql)
	  
		'Check if they want subscribed to topic
	  if Request.Form("rmail") = 1 then
	    if intSubscriptions = 1 and strEmail = 1 then
		  sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
		  set rsAP = my_Conn.execute(sSql)
		  if not rsAP.eof then
	  	    intAppID = rsAP("APP_ID")
		  end if
		  set rsAP = nothing
		    
			itmTitle = txSubject
	        sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=" & topic_id
	        set rsX = my_Conn.execute(sSql)
	        If rsX.BOF or rsX.EOF Then
	          ' subscription doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & strUserMemberID & ", " & intAppID & ", 0, 0, " & topic_id & ", '" & itmTitle & "')"
	          executeThis(insSql)
			  strSubTxt = "You are now subscribed to this topic<br /><br />"
			  strSubTxt = strSubTxt & "You will recieve an email whenever<br />"
			  strSubTxt = strSubTxt & "anyone replies to this topic<br />"
		    else
			  strSubTxt = "You are already subscribed to this topic<br /><br />"
	        End If
	        set rsX = nothing
		end if
	  end if

		if Forum_ID <> aryForum(1) then 'they have moved the topic
			'
			strSql = "UPDATE " & strTablePrefix & "REPLY "
			strSql = strSql & " SET CAT_ID = " & aryForum(0)
			strSql = strSql & ", FORUM_ID = " & aryForum(1)
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

			executeThis(strSql)
			
			'set rs = Server.CreateObject("ADODB.Recordset")
			
			' - count total number of replies in Topics table
			strSql = "SELECT T_REPLIES, T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID
			set rs = my_Conn.Execute (strSql)
			
			intResetCount = rs("T_REPLIES") + 1
			strT_Last_Post = rs("T_LAST_POST")
			strT_Last_Post_Author = rs("T_LAST_POST_AUTHOR")
			'rs.close
			set rs = nothing

			' - Get last_post and last_post_author for MoveFrom-Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC;"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			set rs = nothing

			' - Update count of replies to a topic in Forum table
			strSql = "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_COUNT = F_COUNT - " & intResetCount
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
			executeThis(strSql)

			'
			strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_TOPICS = F_TOPICS - 1 "
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
			executeThis(strSql)

			' - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1) & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC;"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			set rs = nothing

			' - Update count of replies to a topic in Forum table
			strSql = "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_COUNT = (F_COUNT + " & intResetCount & ")"
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1)
			executeThis(strSql)

			'
			strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_TOPICS = F_TOPICS + 1 "
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1)
			executeThis(strSql)

		end if ' end move topic
		err_Msg = ""
		aryForum = ""
		'if Err.description <> "" then 
		'	Go_Result "There was an error = " & Err.description, 0
		'else
		  if request.form("poll") = "1" then
		  	pollTopic = chkString(request.form("pollTopic_ID"),"numeric")
			%><!--#INCLUDE FILE="inc_footer.asp" --><%
			response.redirect "forum_post_info.asp?Method_Type=CreatePoll&TOPIC_IDP=" & pollTopic
			'response.end
		  end if
		  Go_Result  "Updated OK", 1
		'end if
	else ' there was an error %>
		<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

		<table align=center border=0>
	  	  <tr>
	    	<td>
			<ul><% =Err_Msg %></ul>
			</td>
	  	  </tr>
		</table>
		<p align=center>
		<a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p><%
	end if
end if
'******************************************************************************************
'***		END EDIT TOPIC
'******************************************************************************************

'******************************************************************************************
'***		NEW TOPIC
'******************************************************************************************
if MethodType = "Topic" then
	'
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_SIG, M_LASTPOSTDATE, "&Strdbntsqlname
	if strAuthType = "db" then
		strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	if strAuthType = "db" then
		strSql = strSql & " AND M_PASSWORD = '" & strTmpPassword &"'"
		QuoteOk = (ChkQuoteOk(STRdbntUserName) and ChkQuoteOk(strTmpPassword))
	else
		QuoteOk = ChkQuoteOk(Session(strCookieURL & "userid"))
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) then '##  Invalid Password
		Go_Result "Invalid UserName or Password", 0
	else

		if not(chkForumAccess(strUserMemberID,Forum_ID)) then
			Go_Result "You are not allowed to post in this forum !", 0
		end if
		if strFloodCheck = 1 then
			if rs("M_LASTPOSTDATE") > datetostr(DateAdd("s",strFloodCheckTime,strCurDateAdjust)) and not hasAccess(1) then
				strTimeLimit = replace(strFloodCheckTime, "-", "")
				Go_Result "Sorry! Flood Control activated.<br />You cannot post within " & strTimeLimit & " seconds of your last post.<br />Please try again after this period of time elapses.", 0
			end if
		end if
		'txMessage = ChkString(Request.Form("Message"),"message")
		txMessage = chkString(chkHtmlCode(Request.Form("Message")),"message")
		txSubject = left(replace(ChkString(Request.Form("Subject"),"sqlstring"),"'","''"),50)
		if len(trim(txMessage) & "x") = 1 then
			Go_Result "You must post a message!", 0
		end if
			
		if len(trim(txSubject) & "x") = 1 then
			Go_Result "You must post a subject!", 0
		end if
		sig = 0        
		if ((hasAccess(1)) or (chkForumModerator(FORUM_ID, strDBNTUserName) = "1")) then
		lockPermission = 1
		end if
		
		if not Request.Form("news") = 1 then
		if Request.Form("sig") = "yes" and rs("M_SIG") <> "" then
		     sig = 1
		end if
		end if
'
		if Request.Form("rmail") <> "1" then
			TF = "0"
		else 
			TF = "1"
		end if

		' - Add new post to Topics Table
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS (FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_DATE"
		strSql = strSql & ", T_STATUS"
		strSql = strSql & ", T_NEWS"		
		if strIPLogging <> "0" then
			strSql = strSql & ", T_IP"
		end if
		strSql = strSql & ", T_MSGICON"
		strSql = strSql & ", T_MAIL"
		strSql = strSql & ", T_SIG"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Forum_ID
		strSql = strSql & ", " & Cat_ID
		strSql = strSql & ", '" & txSubject & "'"
		strSql = strSql & ", '" & txMessage & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & strCurDateString & "'"
		if Request.Form("lock") = 1 and lockPermission = 1 then
			strSql = strSql & ", 0 "
		else
		 	strSql = strSql & ", 1 "
		end if
		if Request.Form("news") = 1 and mlev >= 3 then
		 	strSql = strSql & ", 1 "
		else
		 	strSql = strSql & ", 0 "	
		end if
				
		if strIPLogging <> "0" then
			strSql = strSql & ", '" & request.ServerVariables("REMOTE_ADDR") & "'"
		end if
		strSql = strSql & ", " & chkstring(request.form("strMessageIcon"), "sqlstring")
		strSql = strSql & ", " & TF & "," & sig & ")"

		executeThis(strSql)

		if Err.description <> "" then 
			err_Msg = "There was an error = " & Err.description
		else
			err_Msg = "Updated OK"
		end if

		' - Increase count of topics and replies in Forum table by 1
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_LAST_POST = '" & strCurDateString & "'"
		strSql = strSql & ",    F_TOPICS = F_TOPICS + 1"
		strSql = strSql & ",    F_COUNT = F_COUNT + 1"
		strSql = strSql & ",    F_LAST_POST_AUTHOR = " & rs("MEMBER_ID") & ""
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		executeThis(strSql)
	  
	  if intSubscriptions = 1 and strEmail = 1 then
	    sSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS WHERE FORUM_ID = " & Forum_ID & " AND CAT_ID = " & Cat_ID & " AND T_SUBJECT = '" & txSubject & "'"
		set rsT = my_Conn.execute(sSql)
		    topic_id = rsT(0)
		set rsT = nothing
		
		sSql = "SELECT APP_ID FROM " & strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
		set rsAP = my_Conn.execute(sSql)
		if not rsAP.eof then
	  	  intAppID = rsAP("APP_ID")
	      'send subscriptions emails
	      eSubject = strSiteTitle & " - New Forum Topic"
		  eMsg = "A new Forum Topic has been submitted at " & strSiteTitle & vbCrLf
		  eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		  eMsg = eMsg & "You can view the topic by visiting " & strHomeUrl & "link.asp?topic_id=" & topic_ID & vbCrLf
	      sendSubscriptionEmails intAppID,"0",Forum_ID,"0",eSubject,eMsg
		  'response.Write("<br />Email sent<br />" )
		end if
		set rsAP = nothing
		
		'Check if they want subscribed to topic
		if Request.Form("rmail") = 1 then
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=" & topic_id
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
		      itmTitle = txSubject
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & strUserMemberID & ", " & intAppID & ", 0, 0, " & topic_id & ", '" & itmTitle & "')"
	          executeThis(insSql)
			  strSubTxt = "You are now subscribed to this topic<br /><br />"
			  strSubTxt = strSubTxt & "You will recieve an email whenever<br />"
			  strSubTxt = strSubTxt & "anyone replies to this topic<br />"
		  else
			  strSubTxt = "You are already subscribed to this topic<br /><br />"
	      End If
	      set rs = nothing
		end if
	  end if

		if request.form("poll") = "1" then
			strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS WHERE FORUM_ID = " & Forum_ID
			strSql = strSql & "AND  T_AUTHOR = " & rs("MEMBER_ID") & " ORDER BY T_DATE DESC"
			set rsPoll = my_Conn.Execute (strSql)
			if not rsPoll.eof or not rsPoll.bof then
				pollTopic = rsPoll("TOPIC_ID")
			end if
			set rsPoll = nothing
			%><!--INCLUDE FILE="inc_footer.asp" --><%
			response.redirect "forum_post_info.asp?Method_Type=CreatePoll&TOPIC_IDP=" & pollTopic
			'response.end
		end if

		Go_Result err_Msg, 1
	end if	
end if
'******************************************************************************************
'***		END NEW TOPIC
'******************************************************************************************

'*************************************************************************************
'***		REPLY										**********************
'*************************************************************************************
if MethodType = "Reply" or MethodType = "ReplyQuote" or MethodType = "TopicQuote" then
	'
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_LASTPOSTDATE, M_SIG, "&Strdbntsqlname
	if strAuthType = "db" then
	strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	if strAuthType = "db" then
		strSql = strSql & " AND M_PASSWORD = '" & strTmpPassword &"'"
		QuoteOk = (ChkQuoteOk(STRdbntUserName) and ChkQuoteOk(strTmpPassword))
	else
		QuoteOk = ChkQuoteOk(STRdbntUserName)
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) then '##  Invalid Password
		err_Msg = "Invalid Password or User Name"
		Go_Result err_Msg, 0
	else
		if not(chkForumAccess(strUserMemberID,Forum_ID)) then
			Go_Result "You are not allowed to post in this forum !", 0
		end if
		if strFloodCheck = 1 then
			if rs("M_LASTPOSTDATE") > datetostr(DateAdd("s",strFloodCheckTime,strCurDateAdjust)) and not hasAccess(1) then
				strTimeLimit = replace(strFloodCheckTime, "-", "")
				Go_Result "Sorry! Flood Control activated.<br />You cannot post within " & strTimeLimit & " seconds of your last post.<br />Please try again after this period of time elapses.", 0
			end if
		end if

		txMessage = ChkString(chkHtmlCode(Request.Form("Message")),"message")

		if trim(txMessage) = "" then
			Go_Result "You must post a message!", 0
		end if
		if ((hasAccess(1)) or (chkForumModerator(FORUM_ID, strDBNTUserName) = "1")) then
		lockPermission = 1
		end if
		sig = 0
		if not Request.Form("news") = 1 then
		if Request.Form("sig") = "yes" and rs("M_SIG") <> "" then
		     sig = 1
		end if
		end if

		if strEmail = 1 then
		'DoReplyEmail Topic_ID, rs("MEMBER_ID"), chkString(Request.Form("UserName"),"sqlstring")
		end if

		if Request.Form("rmail") <> "1" then
			RF  = "0"
		else
			RF = "1"
		end if

		'
		strSql = "INSERT INTO " & strTablePrefix & "REPLY "
		strSql = strSql & "(TOPIC_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", R_AUTHOR"
		strSql = strSql & ", R_DATE "
		if strIPLogging <> "0" then
			strSql = strSql & ", R_IP"
		end if
		strSql = strSql & ", R_MAIL"
		strSql = strSql & ", R_MESSAGE"
		strSql = strSql & ", R_SIG "
		strSql = strSql & ") VALUES ("
		strSql = strSql & Topic_ID
		strSql = strSql & ", " & Forum_ID
		strSql = strSql & ", " & Cat_ID
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		if strIPLogging <> "0" then
			strSql = strSql & ", " & "'" & strOnlineUserIP & "'"
		end if
		strSql = strSql & ", " & RF
		strSql = strSql & ", " & "'" & txMessage & "'," & sig & ")"

		executeThis(strSql)

		' - Update Last Post and count
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_LAST_POST = '" & strCurDateString & "'"
		strSql = strSql & ",    T_REPLIES = T_REPLIES + 1 "
		strSql = strSql & ",    T_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
		if Request.Form("lock") = 1 and lockPermission = 1 then
			strSql = strSql & ",	T_STATUS = 0 "
		end if
		strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

		executeThis(strSql)

		'
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_LAST_POST = '" & strCurDateString & "'"
		strSql = strSql & ",	F_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
		strSql = strSql & ",    F_COUNT = F_COUNT + 1 "
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		executeThis(strSql)
	  
	  if intSubscriptions = 1 and strEmail = 1 then
		sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
		set rsAP = my_Conn.execute(sSql)
		if not rsAP.eof then
	  	  intAppID = rsAP("APP_ID")
		  strSub = chkString(Request.Form("Topic_Title"),"display")
	      'send subscriptions emails
	      eSubject = strSiteTitle & " - New Reply to: " & strSub
		  eMsg = strDBNTUserName & " has replied to a topic on " & strSiteTitle & " that you requested notification to. "
		  eMsg = eMsg & "Regarding the subject - " & strSub & "." & vbCrLf & vbCrLf
		  eMsg = eMsg & "You can view the reply at " & strHomeURL & "link.asp?TOPIC_ID=" & Topic_ID & "&view=lasttopic" & vbCrLf
		  'response.Write("<br />Email here: " & Topic_ID & "<br />" )
	      sendSubscriptionEmails intAppID,"0","0",Topic_ID,eSubject,eMsg
		end if
		set rsAP = nothing

		if Err.description <> "" then 
			Go_Result  "There was an error 1 = " & Err.description, 0
		end if
		
		'Check if they want subscribed to topic
		if Request.Form("rmail") = "1" then
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=" & topic_id
	      set rsX = my_Conn.execute(sSql)
	      If rsX.BOF or rsX.EOF Then
		    sSql = "SELECT T_SUBJECT FROM "& strTablePrefix & "TOPICS WHERE TOPIC_ID=" & topic_id
		    set rsA = my_Conn.execute(sSql)
		    if not rsA.eof then 'topic does exist
		      itmTitle = rsA("T_SUBJECT")
		      'itmTitle = txSubject
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & strUserMemberID & ", " & intAppID & ", 0, 0, " & topic_id & ", '" & itmTitle & "')"
	          executeThis(insSql)
			  strSubTxt = "You are now subscribed to this topic<br /><br />"
			  strSubTxt = strSubTxt & "You will recieve an email whenever<br />"
			  strSubTxt = strSubTxt & "anyone replies to this topic<br />"
			end if
			set rsA = nothing
		  else
			strSubTxt = "You are already subscribed to this topic<br /><br />"
	      End If
	      set rsX = nothing
		end if
	  end if

		if Err.description <> "" then 
			Go_Result  "There was an error 2 = " & Err.description, 0
		else
			if Request.Form("M") = "1" then 
				'
				strSql  = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
				strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " & Topic_ID

				'set rs2 = my_Conn.Execute (strSql)
				
				'DoEmail  rs2("M_EMAIL"), rs2("M_NAME")
				'set rs2 = nothing
			end if
			Go_Result  "Updated OK", 1
	     end if
	end if
end if
'******************************************************************************************
'***		END REPLY																								**********************
'******************************************************************************************

'******************************************************************************************
'***		NEW FORUM																							**********************
'******************************************************************************************
if MethodType = "Forum" then
	member = cint(ChkUser(strDBNTUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member
		case 0 
			'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only the Moderator can create a Forum", 0
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only the Moderator can create a Forum", 0
			end if

		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	txSubject = left(ChkString(Request.Form("Subject"),"message"),50)
	Err_Msg = ""

	if txSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New Forum</li>"
	end if
	if txMessage = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the New Forum</li>"
	end if
	if Err_Msg = "" then
		' - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
		end if
		if cint(Request.Form("AuthType")) = 2 or cint(Request.Form("AuthType")) = 3 or cint(Request.Form("AuthType")) = 7 then
			strSql = strSql & ", F_PASSWORD_NEW"
		end if
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & Cat_ID
		if strPrivateForums = "1" then
			strSql = strSql & ", " & cint(request.form("AuthType")) & ""
		end if
		if cint(Request.Form("AuthType")) = 2 or cint(Request.Form("AuthType")) = 3 or cint(Request.Form("AuthType")) = 7  or cint(Request.Form("AuthType")) = 13 or cint(Request.Form("AuthType")) = 14 then
			strSql = strSql & ", '" & ChkString(Request.Form("AuthPassword"),"sqlstring") & "'"
		end if
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & txSubject & "'"
		strSql = strSql & ", '" & txMessage & "'"
		strSql = strSql & ", " & chkstring(request.form("Type"), "sqlstring")
		strSql = strSql & ")"

		my_Conn.Execute (strSql)

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		Else
'######## Update allowed user list##################################		
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")		
			set rsCount = nothing
'##################################################################
			Go_Result  "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if

if MethodType = "URL" then
	member = cint(ChkUser(strDBNTUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member
		case 0'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only the Moderator can create a web link", 0
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only the Moderator can create a web link", 0
			end if
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	txAddress = ChkString(Request.Form("Address"),"sqlstring")
	txSubject = left(ChkString(Request.Form("Subject"),"message"),50)
	Err_Msg = ""

	if txSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New URL</li>"
	end if
	if txAddress = " " or lcase(txAddress) = "http://" or lcase(txAddress) = "https://" or lcase(txAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>You Must Enter an Address for the New URL</li>"
	end if
	if (left(lcase(txAddress), 7) <> "http://" and left(lcase(txAddress), 8) <> "https://") and txAddress <> "" then
		Err_Msg = Err_Msg & "<li>You Must prefix the Address with <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if txMessage = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the New URL</li>"
	end if
	if Err_Msg = "" then
		' - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
		end if
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_LAST_POST_AUTHOR"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_URL"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE"
		strSql = strSql & ")  VALUES ("
		strSql = strSql & Cat_ID
		if strPrivateForums = "1" then
			strSql = strSql & ", " & chkstring(request.form("AuthType"), "sqlstring") & ""
		end if
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & getMemberID(chkString(Request.Form("UserName"),"sqlstring")) & " "
		strSql = strSql & ", " & "'" & txSubject & "'"
		strSql = strSql & ", " & "'" & txAddress & "'"
		strSql = strSql & ", " & "'" & txMessage & "'"
		strSql = strSql & ", " & chkstring(request.form("Type"), "sqlstring")
		strSql = strSql & ") "

		my_Conn.Execute (strSql)

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		else
			'########### Update allowed user list ##############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")                   
			'##################################################################
			Go_Result  "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if

if MethodType = "EditForum" then
	member = cint(ChkUser(STRdbntUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only Admins and Moderators can change this Forum", 0
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only Admins and Moderators change this Forum", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	txSubject = left(ChkString(Request.Form("Subject"),"sqlstring"),50)
	Err_Msg = ""

	if txSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Forum</li>"
	end if
	if txMessage = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the Forum</li>"
	end if
	if Err_Msg = "" then
		' - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & chkstring(request.form("category"), "sqlstring")
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS = " & chkstring(request.form("AuthType"), "sqlstring") & ""
		end if
		if ChkString(Request.Form("AuthType"),"sqlstring") = 2 or ChkString(Request.Form("AuthType"),"sqlstring") = 3 or ChkString(Request.Form("AuthType"),"sqlstring") = 7 or ChkString(Request.Form("AuthType"),"sqlstring") = 13 or ChkString(Request.Form("AuthType"),"sqlstring") = 14 then
			strSql = strSql & ", F_PASSWORD_NEW = '" & trim(ChkString(Request.Form("AuthPassword"),"sqlstring")) & "'"
		end if
		strSql = strSql & ",    F_SUBJECT = '" & txSubject & "'"
		strSql = strSql & ",    F_DESCRIPTION = '" & txMessage & "'"
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		else
'########## Update Allowed user List ###############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			updateForumMembers Forum_ID
'###################################################################
			Go_Result  "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if

if MethodType = "EditURL" then
	member = cint(ChkUser(strDBNTUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only Admins and Moderators can change this Forum", 0
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only Admins and Moderators change this Forum", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select

	txMessage = ChkString(Request.Form("Message"),"message")
	txAddress = ChkString(Request.Form("Address"),"sqlstring")
	txSubject = left(ChkString(Request.Form("Subject"),"message"),50)
	Err_Msg = ""

	if txSubject = " " then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New URL</li>"
	end if
	if txAddress = " " or lcase(txAddress) = "http://" or lcase(txAddress) = "https://" or lcase(txAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>You Must Enter an Address for the New URL</li>"
	end if
	if (left(lcase(txAddress), 7) <> "http://" and left(lcase(txAddress), 8) <> "https://" and left(lcase(txAddress), 8) <> "file:///") and (txAddress <> "") then
		Err_Msg = Err_Msg & "<li>You Must prefix the Address with <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if txMessage = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Message for the New URL</li>"
	end if
	if Err_Msg = "" then

		' - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & chkstring(request.form("Category"), "sqlstring")
		if strPrivateForums = "1" then
			strSql = strSql & ",    F_PRIVATEFORUMS = " & chkstring(request.form("AuthType"), "sqlstring") & ""
		end if
		strSql = strSql & ",    F_SUBJECT = '" & txSubject & "'"
		strSql = strSql & ",    F_URL = '" & txAddress & "'"
		strSql = strSql & ",    F_DESCRIPTION = '" & txMessage & "'"
		strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		else
'########## Update Allowed user List ###############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			updateForumMembers Forum_ID 
'###################################################################
			Go_Result  "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if

if MethodType = "Category" then
	member = cint(ChkUser(STRdbntUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an administrator can create a category", 0
		case 3 '## Moderator
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only an administrator can create a category", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select
	Err_Msg = ""
	if trim(Request.Form("Subject")) = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the New Category</li>"
	end if
	if Err_Msg = "" then

		' - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "CATEGORY (CAT_NAME) "
		strSql = strSql & " VALUES ('" & left(ChkString(Request.Form("Subject"),"message"),50) & "')"

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		else
			Go_Result  "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if

if MethodType = "EditCategory" then
	member = cint(ChkUser(STRdbntUserName, chkString(Request.Form("Password"),"sqlstring")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "Invalid Password or UserName", 0
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "Only an administrator can change a category", 0
		case 3 '## Moderator
			'## Do Nothing
			if chkForumModerator(Forum_ID, STRdbntUserName) = "0" then
				Go_Result "Only an administrator can change a category", 0
			end if
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
	end select
	Err_Msg = ""
	if Request.Form("Subject") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Subject for the Category</li>"
	end if
	if Err_Msg = "" then
		' - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "CATEGORY "
		strSql = strSql & " SET CAT_NAME = '" & left(ChkString(Request.Form("Subject"),"message"),50) & "'"
		strSql = strSql & " WHERE CAT_ID = " & Cat_ID

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "There was an error = " & Err.description, 0
		else
			Go_Result "Updated OK", 1
		end if
	else 
%>
	<p align=center><span class="fTitle">There Was A Problem With Your Details</span></p>

	<table align=center border=0>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p align=center><a href="JavaScript:history.go(-1)">Go Back To Enter Data</a></p>
<%
	end if
end if
%>
<% set rs = nothing %>
<!--INCLUDE FILE="inc_footer.asp" -->
<%
sub DoEmail(email, user_name)
		strSub = chkString(Request.Form("Topic_Title"),"display")
	'## Emails Topic Author if Requested.  
	if lcase(strEmail) = "1" then
		strRecipientsName = user_name
		strRecipients = email
		strSubject = strSiteTitle & " - Reply to: " & strSub
		strMessage = "Hello " & user_name & vbCrLf & vbCrLf
		strMessage = strMessage & "You have received a reply to your posting on " & strSiteTitle & "." & vbCrLf
		strMessage = strMessage & "Regarding the subject - " & strSub & "." & vbCrLf & vbCrLf
		strMessage = strMessage & "You can view the reply at " & chkString(Request.Form("Refer"),"refer") & vbCrLf
		
		sendOutEmail strRecipients,strSubject,strMessage,2,0
	end if
end sub

sub DoReplyEmail(TopicNum, PostedBy, PostedByName)
	'## Emails all users who wish to receive a mail if topic
	'## has a reply but only send one per member.
	'
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & TopicNum 
	strSql = strSql & " AND   R_MAIL = 1 "
	strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"

	set rsReply = my_Conn.Execute (strSql)

	'
	strSql = " SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_SUBJECT "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS,  "
	strSql = strSql & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND " & strTablePrefix & "TOPICS.TOPIC_ID = " & TopicNum

	set rsTopicAuthor = my_Conn.Execute (strSql)
	
	strSub = chkString(Request.Form("Topic_Title"),"display")
	
	MailSendToAuthor = false

	if (rsTopicAuthor("T_MAIL") = 1) and (PostedBy <> rsTopicAuthor("MEMBER_ID")) then
		strRecipientsName = rsTopicAuthor("M_NAME")
		strRecipients = rsTopicAuthor("M_EMAIL")
		strSubject = strSiteTitle & " - Reply to: " & strSub
		strMessage = "Hello " & rsTopicAuthor("M_NAME") & vbCrLf & vbCrLf
		strMessage = strMessage & PostedByName & " has replied to a topic on " & strSiteTitle & " that you requested notification to. "
		strMessage = strMessage & "Regarding the subject - " & strSub & "." & vbCrLf & vbCrLf
		strMessage = strMessage & "You can view the reply at " & strHomeURL & "link.asp?TOPIC_ID=" & TopicNum & "&view=lasttopic" & vbCrLf
		
		sendOutEmail strRecipients,strSubject,strMessage,2,0

		MailSendToAuthor = true
	end if
	
	prevMember = ""
	
	do while (not rsReply.EOF) and (not rsReply.BOF)
		if (prevMember <> rsReply("MEMBER_ID")) and (PostedBy <> rsReply("MEMBER_ID")) then
			if (rsTopicAuthor("MEMBER_ID") = rsReply("MEMBER_ID")) and (MailSendToAuthor) then
				'## Do Nothing
				'## The reply was done by the author, and he/she allready has got a mail
			else
				if (rsTopicAuthor("MEMBER_ID") = rsReply("MEMBER_ID")) then
					MailSendToAuthor = true
				end if
				strRecipientsName = rsReply("M_Name")
				strRecipients = rsReply("M_EMAIL")
				strSubject = strSiteTitle & " - Reply to: " & strSub
				strMessage = "Hello " & rsReply("M_NAME") & vbCrLf & vbCrLf
				strMessage = strMessage & PostedByName & " has replied to a topic on " & strSiteTitle & " that you requested notification to. "
				strMessage = strMessage & "Regarding the subject - " & chkString(Request.Form("Topic_Title"),"display") & "." & vbCrLf & vbCrLf
				strMessage = strMessage & "You can view the reply at " & strHomeURL & "link.asp?TOPIC_ID=" & TopicNum & "&view=lasttopic" & vbCrLf
				sendOutEmail strRecipients,strSubject,strMessage,2,0
			end if
		end if
		prevMember = rsReply("MEMBER_ID")
		rsReply.MoveNext
	loop

	set rsReply = nothing

	set rsTopicAuthor = nothing
end sub

sub Go_Result(str_err_Msg, boolOk)
%>
<table border="0" width="100%">
  <tr>
	<td width="33%" align="left">
	<img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
<% 
	if MethodType = "Topic" or _
		MethodType = "Reply" or _
		MethodType = "EditTopic" then 
%>
	<img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open.gif" border="0">&nbsp;<a href="FORUM.asp?FORUM_ID=<% = Forum_ID %>&CAT_ID=<% = Cat_ID %>&Forum_Title=<% = ChkString(Request.Form("FORUM_Title"),"urlpath")%>"><% = chkString(Request.Form("FORUM_Title"),"sqlstring") %></a><br />
<% 
	end if 
	if MethodType = "Replyxx" or _
		MethodType = "EditTopicxx" then 
%>
	<img src="images/icons/icon_blank.gif" border="0"><img src="images/icons/icon_bar.gif" border="0"><img src="images/icons/icon_folder_open_topic.gif" border="0">&nbsp;<a href="<% =chkString(Request.Form("refer"),"refer") %>"><%= chkstring(replace(replace(Request.Form("Topic_Title"),"&amp;","&"),"&59;",";"),"sqlstring") %></a>
<% 
	end if 
%>
    </td>
  </tr>
</table>
<%  if boolOk = 1 then %>
		<p align="center"><span class="fTitle">
<%		select case MethodType
			case "Edit"
				Response.Write("Your Reply Was Changed Successfully!")
			case "EditCategory"
				Response.Write("Category Name Changed Successfully!")
			case "EditForum"
				Response.Write("FORUM Information Updated Successfully!")
			case "EditTopic"
				Response.Write("Topic Changed Successfully!")
			case "EditURL"
				Response.Write("URL Information Updated Successfully!")
			case "Reply"
				Iuser = chkString(Request.Form("UserName"),"sqlstring")
				Response.Write("New Reply Posted!")
				DoRepAdd Iuser
				DoPCount
				DoUCount Iuser
				DoULastPost Iuser
			case "ReplyQuote"
				Iuser = chkString(Request.Form("UserName"),"sqlstring")
				Response.Write("New Reply Posted!")
				DoRepAdd Iuser
				DoPCount
				DoUCount Iuser
				DoULastPost Iuser
			case "TopicQuote"
				Iuser = chkString(Request.Form("UserName"),"sqlstring")
				Response.Write("New Reply Posted!")
				DoRepAdd Iuser
				DoPCount
				DoUCount Iuser
				DoULastPost Iuser
			case "Topic"
				Iuser = chkString(Request.Form("UserName"),"sqlstring")
				DoRepAdd Iuser
				DoTCount
				DoPCount
				DoUCount Iuser
				DoULastPost Iuser
				'chkSubscriptions(intAppID,"","","", emailTopic)
				Response.Write("New Topic Posted!") 
			case "Forum"
				Response.Write("New Forum Created!")
			case "URL"
				Response.Write("New URL Created!")
			case "Category"
				Response.Write("New Category Created!")
			case "makePoll"
			 	Response.Write("New poll has been created!")
			case "updatePoll"
			 	Response.Write("Poll has been updated!")
			case else
				Iuser = chkString(Request.Form("UserName"),"sqlstring")
				Response.Write("Complete!")
				DoRepAdd Iuser
				DoPCount
				DoUCount Iuser
				DoULastPost Iuser
		end select
%>
		</span></p>
		<meta http-equiv="Refresh" content="3; URL=<%= chkString(Request.Form("Refer"),"refer") %>">
		<p align="center"><span class="fTitle">
	<%	select case MethodType
			case "Category"
				Response.Write("Remember to create at least one new forum in this category.")
			case "EditCategory"
				Response.Write("Thank you for your contribution!")
			case "Forum"
				Response.Write("The new forum is ready for users to begin posting!")
			case "EditForum"
				Response.Write("Thank you for your contribution!")
			case "URL"
				Response.Write("The new URL is in place!")
			case "EditURL"
				Response.Write("Cheers! Have a nice day!")
			case "Topic"
				Response.Write("Thank you for your contribution!")
			case "TopicQuote"
				Response.Write("Thank you for your contribution!")
			case "EditTopic"
				Response.Write("Thank you for your contribution!")
			case "Reply"
				Response.Write("Thank you for your contribution!")
			case "ReplyQuote"
				Response.Write("Thank you for your contribution!")
			case "Edit"
				Response.Write("Thank you for your contribution!")
			case "makePoll"
				Response.Write("Thank you for your contribution!")
			case "updatePoll"
				Response.Write("Thank you for your contribution!")
			case else
				Response.Write("Have a nice day!")
		end select
				Response.Write("<br /><br />" & strSubTxt & "<br /><br />") %>
		</span></p>
	  <p align="center">
	  <a href="<%= chkString(Request.Form("refer"),"refer")%>">Back To Forum</a></p>
	  <p>&nbsp;</p>
<%	else %>
	  <p align="center">
	  <span class="fTitle">There has been a problem!</span></p>
	  <p align="center"><span class="fSubTitle"><% =str_err_Msg %></span></p>
	  <p align="center">
	  <a href="JavaScript:history.go(-1)">Go back to correct the problem.</a></p>
<%  end if %>
	<!--#INCLUDE FILE="inc_footer.asp" --><%
	Response.End()
end sub

sub newForumMembers(fForumID)
		on error resume next
		if Request.Form("AuthUsers") = "" then
			exit Sub
		end if
	Users = split(chkString(Request.Form("AuthUsers"),"sqlstring"),",")
	for count = Lbound(Users) to Ubound(Users)
		curCnt = chkString(Users(count),"numeric")
		strSql = "INSERT INTO " & strMemberTablePrefix & "ALLOWED_MEMBERS ("
		strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& curCnt & ", " & fForumID & ")"

		my_conn.execute (strSql)
		if err.number <> 0 then
			Go_REsult err.description, 0
		end if
	next

end sub

sub updateForumMembers(fForumID)
		my_Conn.execute ("DELETE FROM " & strMemberTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID = " & fForumId)
		newForumMembers(fForumID)
end sub

function emailTopic()
CAT_ID = chkString(request.form("CAT_ID"),"sqlstring")
FORUM_ID = chkString(request.form("FORUM_ID"),"sqlstring")
FORUM_TITLE = chkString(request.form("Forum_Title"),"sqlstring")
PostedByName = chkString(request.form("UserName"),"sqlstring")
subject = chkString(Request.Form("Subject"),"message")
Set rs = my_Conn.Execute("SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS WHERE T_SUBJECT = '" & subject & "' AND FORUM_ID="& FORUM_ID & " AND CAT_ID=" & CAT_ID & " ORDER BY T_DATE DESC")
TopicNum = rs(0)
rs.close
set rs = Nothing
if lcase(strEmail) = "1" then
	strSubject = FORUM_TITLE & " - New Topic"
	strMessage = "Hello! " & vbCrLf & vbCrLf
	strMessage = strMessage & PostedByName & " posted a new topic on the forum " & FORUM_TITLE & " (that you subscribed to). " & vbCrLf & vbCrLf
	strMessage = strMessage & "Regarding the subject - " & chkString(Request.Form("Subject"),"sqlstring") & "." & vbCrLf & vbCrLf
	strMessage = strMessage & "You can view the topic at " & strHomeURL & "link.asp?TOPIC_ID=" & TopicNum & vbCrLf
	'get subscribed members
	strRecipients = ""
	sql = "select M_EMAIL from forums_subscriptions, " & strMemberTablePrefix & "members where f_id="& forum_id & " and c_id=" & cat_id &" and m_id=member_id and m_id<>" & getMemberId(strDBNTUserName)
	set rs = my_Conn.execute(sql)
	do while not rs.eof
	     'concatenate all the emails with a ";" probably works only with CDONTS
	     if emailfield(rs(0))<>0 then
	     	strRecipients = strRecipients & rs(0) & "; "
	     end if	
	     rs.movenext
	loop
	if strRecipients<>"" then
		SubscriptionsForumsMod = "1"
		'sendOutEmail strRecipients,strSubject,strMessage,2,0

		SubscriptionsForumsMod = "0"
	end if	
end if
  emailTopic = strMessage
end function
%>