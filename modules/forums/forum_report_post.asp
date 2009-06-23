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

dim intTid, intRid, strReporterIP, strReporterName, strCkPassWord, Msg, Dtbl, CurPageType
CurPageType="forums"
%>
<!--#include file="config.asp" --> 
<!-- #include file="lang/en/forum_core.asp" -->
<!--#include file="inc_functions.asp" -->
<%
'#################################################################################
'## Initialise variables 
'#################################################################################
intTid = ""
intRid = ""
Msg = ""
strSql = ""
intPage = ""
rptItem = ""

if Request.QueryString("tid") <> "" then
	if IsNumeric(Request.QueryString("tid")) = True then
		intTid = cLng(Request.QueryString("tid"))
		rptItem = "topic"
	end if
end if
if Request.QueryString("rid") <> "" then
	if IsNumeric(Request.QueryString("rid")) = True then
		intRid = cLng(Request.QueryString("rid"))
		rptItem = "reply"
	end if
end if
if Request.QueryString("page") <> "" or  Request.QueryString("page") <> " " then
	if IsNumeric(Request.QueryString("page")) = True then
		intPage = cLng(Request.QueryString("page"))
	end if
end if
'#################################################################################
'## Page-code start
'#################################################################################
%>
<!--#include file="inc_top_short.asp" -->
<% 
if strDBNTUsername <> "" then

  if Request.form("method") = "report" then
	intTid = chkString(Request.form("intTid"),"sqlstring")
	intRid = split(Request.form("intRid"),":")(0)
	intRid = chkString(intRid,"sqlstring")
	intPage = split(Request.form("intRid"),":")(1)
	intPage = chkString(intPage,"sqlstring")
	strReporterIP = chkString(Request.form("strReporterIP"),"sqlstring")
	strReporterID = chkString(Request.form("strReporterID"),"sqlstring")
	strReason = trim(chkString(Request.form("strReason"),"sqlstring"))
	strAuthor = chkString(Request.form("strAuthor"),"sqlstring")
	strRptItem = chkString(Request.form("strRptItem"),"sqlstring")
	if strReason <> "" and len(strReason) > 20 then
		if strRptItem = "topic" then
			strSql = "SELECT T_MESSAGE, FORUM_ID, CAT_ID FROM " & strTablePrefix & "TOPICS  WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & intTid
			set rs = my_Conn.Execute (strSql)
			intTid = intTid & ":" & rs(1) & ":" & rs(2) & ":" & intPage
			intRid = "0"
		else
			strSql = "SELECT R_MESSAGE, TOPIC_ID, FORUM_ID, CAT_ID FROM " & strTablePrefix & "REPLY  WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & intRid
			set rs = my_Conn.Execute (strSql)
			intRid = intRid & ":" & rs(1) & ":" & rs(2) & ":" & rs(3) & ":" & intPage
			intTid = "0"
		end if
	
		strReportedTxt = cleancode(rs(0))
	
			strSql = "INSERT INTO " & strTablePrefix & "REPORTED_POST ("
			strSql = strSql & "R_REPORTER_ID"
			strSql = strSql & ", R_REPORTER_IP"
			strSql = strSql & ", R_TOPIC_ID"
			strSql = strSql & ", R_REPLY_ID"
			strSql = strSql & ", R_REASON"
			strSql = strSql & ", R_REPORTED_DATE"
			strSql = strSql & ", R_POST"
   			strSql = strSql & ") VALUES ("
			strSql = strSql & " '" & strReporterID & ":" & strAuthor & "'"
			strSql = strSql & ", '" & strReporterIP & "'"
			strSql = strSql & ", '" & intTid & "'"
			strSql = strSql & ", '" & intRid & "'"
			strSql = strSql & ", '" & chkString(strReason,"message") & "'"
			strSql = strSql & ", " & "'" & strCurDateString & "'"
			strSql = strSql & ", '" & chkString(strReportedTxt,"message") & "')"
			
'			Response.Write(strSql)
			my_Conn.Execute (strSql)
			
			Response.Write("<br /><br />")
			spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"" width=""100%"">")
			Response.Write("<tr><td width=""100%"" align=""center""><br />")
			Response.Write("<b>")
			Response.Write("The post has been reported!</b><br />")
			Response.Write("<br />The next administrator or moderator<br />")
			Response.Write("that visits the site will be notified<br /><br /><br /><br /></td></tr>")
			Response.Write("</table>")
spThemeBlock1_close(intSkin)
			
	else
			Response.Write("<br /><b>You didn't enter in a reason<br />")
			Response.Write("for reporting this post.<br />Or the reason wasn't long enough.<br />")
			Response.Write("<br />Please try again</b><br />&nbsp;")
	end if
  else


	if intTid <> "" then
		strSql = "SELECT T_MESSAGE, T_AUTHOR FROM " & strTablePrefix & "TOPICS  WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & intTid
	else
 		if intRid <> "" then
			strSql = "SELECT R_MESSAGE, R_AUTHOR, TOPIC_ID FROM " & strTablePrefix & "REPLY  WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & intRid
		else
			Msg = "Invalid Topic or Reply ID"
		end if
	end if
	
  	if Msg = "" then
	set rs1 = my_Conn.Execute (strSql)

	strReportedTxt = cleancode(rs1(0))
	strReportedTxt = replace(strReportedTxt,"tiny_mce/",strHomeURL & "tiny_mce/")
	strAuthor = rs1(1)
	strReporterIP = Request.ServerVariables("REMOTE_HOST")
	strReporterID = getmemberid(strDBNTUsername)
	if intRid <> "" then
		intTid = rs1(2)
	end if
%>
<center>
<script type="text/javascript">
function checkfrm(){
 if (document.forms.report_post.strReason.value == "") {
 alert("You must enter a reason for reporting this post!");
 return;
 }
 else{
 document.forms.report_post.submit();
 }
 }
</script>
<form action="forum_report_post.asp" method="post" name="report_post" id="report_post">
<input name="method" type="hidden" value="report" ID="method">
<input name="intTid" type="hidden" value="<% =intTid %>" ID="intTid">
<input name="intRid" type="hidden" value="<% =intRid %>:<% =intPage %>" ID="intRid">
<input name="strReporterIP" type="hidden" value="<% =strReporterIP %>" ID="strReporterIP">
<input name="strReporterID" type="hidden" value="<% =strReporterID %>" ID="strReporterID">
<input name="strAuthor" type="hidden" Value="<% =strAuthor %>">
<input name="strRptItem" type="hidden" Value="<% =rptItem %>">
<%
spThemeTitle = "Report post to Moderator"
spThemeTableCustomCode = "align=""center"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">") %>
                <tr align="center"> 
                  <td height="25" colspan="2"><br /><b>To 
                    prevent abuse, your information below<br />is being sent along with the report<br /><br />20 characters minimum please.<br /></b></font><br /></td>
                </tr>
                <tr> 
                  <td width="50%" align="right">Your 
                    user name:&nbsp; </font></td>
                  <td width="50%"><%= strDBNTUserName %></font></td>
                </tr>
                <tr> 
                  <td align="right">Your 
                    IP:&nbsp; </font></td>
                  <td><%= strReporterIP %></font></td>
                </tr>
                <tr> 
                  <td align="right">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr align="center"> 
                  <td height="15" colspan="2">Your 
                    reason for reporting this post:&nbsp; </font> </td>
                </tr>
                <tr> 
                  <td colspan="2" align="center" valign="top"> 
                    <textarea class="textbox" name="strReason" cols="50" rows="5" wrap="VIRTUAL"></textarea>
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" align="center"> 
                    <input class="button" type="button" name="Submit" value="Submit" onClick="checkfrm()">
                  </td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" align="center">
				  <table width="75%" border="0" cellspacing="0" cellpadding="0"><tr><td>
				  <% spThemeTableCustomCode = " border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" align=""center"" width=""93%"""
				  spThemeSmallBlock_open()%>
				  	<tr><td <%= spThemeBlock_subTitleCell %> colspan="2" align="center"><b>Reported Post</font></b></td></tr>
					<tr><td valign="top"><%= chkString(strReportedTxt,"") %></font></td></tr>
					<% spThemeSmallBlock_close() %><BR><BR></td></tr></table>
                  </td>
                </tr>
              <% Response.Write("</table>")
spThemeBlock1_close(intSkin) %>
</form>
	</center><br />
<% 
	  else %>
<p align="center"><% =Msg %></p>
<%	  
	  end if
	end if
else %>

<p align="center"><span class="fTitle">Please login or register.</span></p>

<%
end if
%><!--#include file="inc_footer_short.asp" -->