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
intPage = ""
intStatus = 0
intTid = "0"
intRid = "0"
'breadcrumb here
  arg1 = "Reported Post|forum_report_post_moderate.asp"
  arg2 = "Open Reports|forum_report_post_moderate.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
if request.Form("status") <> "" then
  strStatusView = chkString(request.Form("status"),"sqlstring")
	select case strStatusView
		case "all"
			arg2 = "View All Reports|forum_report_post_moderate.asp"
			intStatus = 2
		case "open"
			arg2 = "Open Reports|forum_report_post_moderate.asp"
			intStatus = 0
		case "closed"
			arg2 = "Closed Reports|forum_report_post_moderate.asp"
			intStatus = 1
		case "archive"
			arg2 = "Archived Reports|forum_report_post_moderate.asp"
			intStatus = 3
	end select
end if
if Session.Contents("rptMSG") = "<br /><b>Report Deleted!</b><br /><br />" then

end if
if Session.Contents("rptMSG") = "<br /><b>Report Archived!</b><br /><br />" then
	intStatus = 1
end if
%>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<% 
If mlev >= 3 Then 
'	response.Flush()
  If mlev >= 3 and request.Form("method") = "archive" Then  
	strReportID = chkString(request.Form("intID"),"sqlstring")
	strSQL = "update " & strTablePrefix & "REPORTED_POST "
	strSQL = strSQL & "set " & strTablePrefix & "REPORTED_POST.R_STATUS = 3 "
	strSQL = strSQL & "where " & strTablePrefix & "REPORTED_POST.ID = " & strReportID
	my_Conn.Execute (strSQL)
	my_conn.close
	set my_Conn = nothing
	Session.Contents("rptMSG") = "<br /><b>Report Archived!</b><br /><br />"
	response.Redirect("forum_report_post_moderate.asp")
  else
	if request.Form("method") = "archive" Then  
	my_conn.close
	set my_Conn = nothing
	Session.Contents("rptMSG") = "<br /><b>No permission to Archive reports</b><br /><br />"
	response.Redirect("forum_report_post_moderate.asp")
	end if
  End if
  If mlev >= 3 and request.Form("method") = "delete" Then  
	strReportID = chkString(request.Form("intID"),"sqlstring")
	strSQL = "delete from " & strTablePrefix & "REPORTED_POST where " & strTablePrefix & "REPORTED_POST.ID =" & strReportID
	my_Conn.Execute (strSQL)
	my_conn.close
	set my_Conn = nothing
	Session.Contents("rptMSG") = "<br /><b>Report Deleted!</b><br /><br />"
	response.Redirect("forum_report_post_moderate.asp")
  else
	if request.Form("method") = "delete" Then  
	my_conn.close
	set my_Conn = nothing
	Session.Contents("rptMSG") = "<br /><b>No permission to Delete reports</b><br /><br />"
	response.Redirect("forum_report_post_moderate.asp")
	end if
  End if

  If mlev >= 3 and request.Form("method") = "takeaction" Then  
	strAction = chkString(request.Form("strActionTaken"),"message")
	strReportID = chkString(request.Form("reportID"),"sqlstring")
	strStatus = 1
	
	if trim(strAction) <> "" then
	
	strSQL = "update " & strTablePrefix & "REPORTED_POST set"
	strSQL = strSQL & " R_ACTION_BY = " & getmemberid(strdbntusername)
	strSQL = strSQL & ", R_ACTION_DATE = " & strCurDateString
	strSQL = strSQL & ", R_ACTION_TAKEN = '" & strAction & "'"
	strSQL = strSQL & ", R_STATUS = " & strStatus
	strSQL = strSQL & " where ID = " & strReportID
	
'	Response.Write(strSql)
	my_Conn.Execute (strSQL)
	
	Session.Contents("rptMSG") = "<br /><b>Action taken!</b><br /><br />"
	else
	Session.Contents("rptMSG") = "<br /><b>ERROR! Report not updated!</b><br />Please enter the action you took.<br /><br />"
	end if
	my_conn.close
	set my_Conn = nothing
	response.Redirect("forum_report_post_moderate.asp")
  else
	if request.Form("method") = "takeaction" Then  
	my_conn.close
	set my_Conn = nothing
	Session.Contents("rptMSG") = "<br /><b>No permission to Take Action</b><br /><br />"
	response.Redirect("forum_report_post_moderate.asp")
	end if
  end if
%>
	<script language="JavaScript" type="text/JavaScript">
	function mode(cmd, id) {
		var strid; strid = id
		var comd; comd = cmd
		switch (cmd) {
			case "delete":
				if (!confirm('Are you sure to want delete this report?\nThis cannot be undone')) return;
					document.forms.buffer.method.value = comd;
					document.forms.buffer.intID.value = strid;
					document.forms.buffer.submit();
	//			break;
			case "archive":
					document.forms.buffer.method.value = comd;
					document.forms.buffer.intID.value = strid;
					document.forms.buffer.submit();				
	//			break;
			default:
				document.forms.buffer.method.value = comd;
				document.forms.buffer.intID.value = strid;
		}
	}
	</script>

	<FORM name="buffer" action="forum_report_post_moderate.asp" method="post">
	<input name="method" type="hidden" Value="">
	<input name="intID" type="hidden" Value="">
	</FORM>

	<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td valign="top" class="leftPgCol" rowspan="2" nowrap>
	<% 
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %>
	</td>
	<td valign="top" class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<form action="forum_report_post_moderate.asp" method="post" name="formEle" id="formEle">
	<table border="0" width="100%" align="center">
	<tr><td width="200" align="center" valign="baseline">
	<select name="status" ID="status" onChange="submit()">
	<option value="all"<% If intStatus = 2 Then response.Write(" selected") else 'nothing %>>View All Reports</option>
	<option value="open"<% If intStatus = 0 Then response.Write(" selected") else 'nothing %>>Open Reports</option>
	<option value="closed"<% If intStatus = 1 Then response.Write(" selected") else 'nothing %>>Closed Reports</option>
	<option value="archive"<% If intStatus = 3 Then response.Write(" selected") else 'nothing %>>Archived Reports</option>
	</select>
	<input name="method" type="hidden" Value="view">
	</td><td>
	<% If trim(session.Contents("rptMSG")) <> "" Then 
			response.Write("<p>" & session.Contents("rptMSG") & "</p>")
			session.Contents("rptMSG") = ""
		  end if %>
	</td></tr></table>
	</form>
      <% 
	spThemeTitle="View Report"
	spThemeBlock1_open(intSkin) %>
	<table cellpadding="0" cellspacing="0" width="100%"><tr><td valign="top">
	<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
<% 
	if intStatus <> 2 then
	strSQL = "SELECT * from PORTAL_REPORTED_POST WHERE R_STATUS = " & intStatus & " order by ID desc"
	else
	strSQL = "SELECT * from PORTAL_REPORTED_POST order by ID desc"
	end if
	set rs = my_Conn.Execute (strSQL)
	if rs.eof then
	  if intStatus = 3 then
		response.Write("<tr><td width=""100%""><center><br /><br /><b>No Archived Reports</b><br /><br /><br /><br /><br /></center></td></tr></table>")
	  else 
	  	if intStatus = 1 then
			response.Write("<tr><td width=""100%""><center><br /><br /><b>No Closed Reports</b><br /><br /><br /><br /><br /></center></td></tr></table>")
	  	else
			response.Write("<tr><td width=""100%""><center><br /><br /><b>No Reported Posts</b><br /><br /><br /><br /><br /></center></td></tr></table>")
		end if
	  end if
	else
	  do until rs.eof
		rptID = rs("ID")
		if rs("R_TOPIC_ID") <> "0" then
			intTid = split(rs("R_TOPIC_ID"),":")(0)
			intFid = split(rs("R_TOPIC_ID"),":")(1)
			intCid = split(rs("R_TOPIC_ID"),":")(2)
			intPage = split(rs("R_TOPIC_ID"),":")(3)
		else
		end if
		if rs("R_REPLY_ID") <> "0" then
		intRid = split(rs("R_REPLY_ID"),":")(0)
		intTid = split(rs("R_REPLY_ID"),":")(1)
		intFid = split(rs("R_REPLY_ID"),":")(2)
		intCid = split(rs("R_REPLY_ID"),":")(3)
		intPage = split(rs("R_REPLY_ID"),":")(4)
		else
		end if
		strStatus = rs("R_STATUS")
		strReporter = getmembername(split(rs("R_REPORTER_ID"),":")(0))
		strAuthor = getmembername(split(rs("R_REPORTER_ID"),":")(1))
		strReporterIP = rs("R_REPORTER_IP")
		strReportDate = strtodate(rs("R_REPORTED_DATE"))
		strReason = rs("R_REASON")
		strPost = rs("R_POST")
		if rs("R_COMMENTS") <> "0" then 
			strComments = rs("R_COMMENTS")
		else
			strComments = ""
		end if
		if rs("R_ACTION_TAKEN") <> "0" then 
			strActionTaken = rs("R_ACTION_TAKEN")
		else
			strActionTaken = ""
		end if
		if rs("R_ACTION_BY") <> "0" then 
			strActionBy = rs("R_ACTION_BY")
			strActionByName = getmembername(rs("R_ACTION_BY"))
		else
			strActionBy = ""
		end if
		if rs("R_ACTION_DATE") <> "0" then
			strActionDate = strtodate(rs("R_ACTION_DATE"))
		else
			strActionDate = ""
		end if
		%>
		<tr><td align="center">
	<form action="forum_report_post_moderate.asp" method="post" name="report_post<%= rptID %>" id="report_post<%= rptID %>">
      <table width="100%" border="1" cellspacing="5" cellpadding="2" style="border-collapse: collapse" class="tCellAlt2">
        <tr> 
          <td align="right" width="25%">
		  <% If hasAccess(2) Then %>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		    <tr><td align="left" width="25%">
		  	<% If strStatus = 1 or strStatus = 3 Then %>
				<a href="javascript:mode('delete', &#34;<%=rptID%>&#34;);"><img src="images/icons/icon_delete_reply.gif" border="0" width="15" height="15" alt="Delete report - admin only"></a>
				<% If strStatus = 1 Then %>
				<a href="javascript:mode('archive', &#34;<%=rptID%>&#34;);"><img src="images/icons/icon_go_right.gif" onClick="mode('archive','<%= rptID %>')" border="0" width="15" height="15" alt="Archive report - admin only"></a>
				<% End If %>
			<% Else %>
				<a href="javascript:mode('delete', &#34;<%=rptID%>&#34;);"><img src="images/icons/icon_delete_reply.gif" onClick="mode('delete','<%= rptID %>')" border="0" width="15" height="15" alt="Delete report - admin only"></a>
			<% End If %>
			</td><td align="right" width="75%">
			Reported 
            by:&nbsp;&nbsp; 
			</td></tr></table>
		  <% Else %>
		  Reported 
            by:&nbsp;&nbsp; 
		  <% End If %>
		  </td>
          <td width="25%">&nbsp;<b><%= strReporter %></b></td>
          <td align="right" colspan="2">Date 
            reported:&nbsp;&nbsp;&nbsp;
			<span class="fAltSubTitle"><b><%= strReportDate %></b></span></td>
        </tr>
        <tr> 
          <td align="right" valign="top">Reason 
            submitted:&nbsp; &nbsp; </td>
          <td colspan="3" valign="middle">&nbsp;<%= strReason %></td>
        </tr>
        <tr> 
          <td colspan="4" align="center"><br />
             	<table width="90%" border="0" cellspacing="0" cellpadding="0"><tr><td>
				  <% spThemeTableCustomCode = " border=""1"" cellpadding=""3"" cellspacing=""0"" style=""border-collapse: collapse"" align=""center"" width=""93%"""
				  spThemeSmallBlock_open()%>
				  	<tr><td <%= spThemeBlock_subTitleCell %>><b>Reported Post posted by: <span class="fAltSubTitle"><%= strAuthor %></span></b></td><td <%= spThemeBlock_subTitleCell %> align="center">
			<% If intRid <> "0" Then %>
				<% If intPage > "1" Then %>
						<a href="forum_topic.asp?cat_ID=<%= intCid %>&forum_ID=<%= intFid %>&topic_ID=<%= intTid %>&whichpage=<%= intPage %>&tmp=1#pid<%= intRid %>" target="_blank">view post</a>
				<% Else %>
						<a href="forum_topic.asp?cat_ID=<%= intCid %>&forum_ID=<%= intFid %>&topic_ID=<%= intTid %>&tmp=1#pid<%= intRid %>" target="_blank">view post</a>
				<% End If %>
			<% Else %>
			<a href="forum_topic.asp?cat_ID=<%= intCid %>&forum_ID=<%= intFid %>&topic_ID=<%= intTid %>" target="_blank">view post</a>
			<% End If %>
					&nbsp;&nbsp;</td></tr>
					<tr><td valign="top" colspan="2"><span class="fSmall"><%= chkString(strPost,"") %></span></td></tr>
					<% spThemeSmallBlock_close() %><BR></td></tr></table>
          </td>
        </tr>
        <tr> 
		<% If strActionBy <> "" Then %>
          <td align="right" valign="top">Action By:&nbsp; </td>
          <td align="left" valign="top">&nbsp;<%= strActionByName %></td>
		<% Else %>
          <td align="right" valign="top">Action By:&nbsp; </td>
          <td align="left" valign="top"><span class="fAltSubTitle">&nbsp;<b>NEEDS ATTENTION!</b></span></td>
		<% End If %>
		<% If strActionDate <> "" Then %>
          <td align="right" valign="top" colspan="2">Action date:&nbsp;&nbsp;&nbsp;<span class="fAltSubTitle"><b>&nbsp;<%= strActionDate %></b></span></td>
		<% Else %>
          <td align="right" valign="top" colspan="2">&nbsp;</td>
		<% End If %>
		</tr>
		<% If strActionTaken <> "" Then %>
        <tr> 
          <td align="right" valign="top">Action Taken:&nbsp; </td>
          <td align="left" valign="top" colspan="3">&nbsp;<%= strActionTaken %></td>
		</tr>
		<% Else %>
        <tr> 
          <td align="right" valign="top">Action Taken:&nbsp; </td>
          <td align="left" valign="top" colspan="3"><textarea class="textbox" name="strActionTaken" cols="50" rows="5" wrap="VIRTUAL"></textarea></td>
		</tr>
        <tr> 
          <td align="center" colspan="4" height="30" valign="middle"><input class="button" type="submit" name="Submit" value="Submit"><input name="method" type="hidden" Value="takeaction"><input name="reportID" type="hidden" Value="<%= rptID %>"></td>
		</tr>
		<% End If %>
      </table></form>
		</td></tr>
<%  rs.movenext 
			response.Write("<tr><td>&nbsp;</td></tr>")
	  loop %>
	  </table>
<% 
	  end if
	  response.write("</td></tr>")
		Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
	
	</td></tr>
	</table>
<% 
Else
	response.Redirect("default.asp")
End If %>
<!--#INCLUDE FILE="inc_footer.asp" -->
