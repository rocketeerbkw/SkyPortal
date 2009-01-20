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
pgType = "manager"
  modPgType = "addForm"
  uploadPg = false
  hasEditor = true
  strEditorElements = "Message"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<% 
iPgType = 0
sMode = 0
a_id = 0
if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		sMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("w_id") <> "" or  Request("w_id") <> " " then
	if IsNumeric(Request("w_id")) = True then
		w_id = cLng(Request("w_id"))
	else
		closeAndGo("default.asp")
	end if
end if


select case sMode
  case 1 'delete from db
   if w_id = 1 or w_id = 2 then
	strMsg = "<li><b>" & txtWel1 & "</b></li>"
   else
    sSql = "DELETE FROM " & strTablePrefix & "WELCOME WHERE W_ID=" & w_id
	executeThis(sSql)
	strMsg = "<li><b>" & txtWel2 & "</b></li>"
   end if
  case 2 'edit db
    a_id = cLng(request.Form("W_ID"))
    a_title = chkString(request.Form("W_TITLE"),"message")
    a_subject = chkString(request.Form("W_SUBJECT"),"message")
    a_active = chkString(request.Form("W_ACTIVE"),"message")
    a_message = chkString(request.Form("Message"),"message")
    'a_message = request.Form("W_MESSAGE")
	a_message = replace(a_message,"</p><p>","<br /><br />")
	a_message = replace(a_message,"<p>","")
	a_message = replace(a_message,"</p>","")
	'response.Write("a_message: " & a_message & "<br />")
	sSql = "UPDATE " & strTablePrefix & "WELCOME SET "
	sSql = sSql & "W_SUBJECT='" & a_subject & "'"
	sSql = sSql & ",W_MESSAGE='" & a_message & "'"
	sSql = sSql & ",W_TITLE='" & a_title & "'"
	sSql = sSql & ",W_ACTIVE=" & a_active & ""
	sSql = sSql & " WHERE W_ID=" & a_id
	'response.Write(sSql & "<br />")
	executeThis(sSql)
	strMsg = "<li><b>X " & txtWel3 & "</b></li>"
  case 3 'add to db
    a_title = chkString(request.Form("W_TITLE"),"message")
    a_subject = chkString(request.Form("W_SUBJECT"),"message")
    a_active = chkString(request.Form("W_ACTIVE"),"message")
    a_message = chkString(request.Form("W_MESSAGE"),"message")
	a_message = replace(a_message,"</p><p>","<br /><br />")
	a_message = replace(a_message,"<p>","")
	a_message = replace(a_message,"</p>","")
	sSql = "INSERT INTO " & strTablePrefix & "WELCOME ("
	sSql = sSql & "W_SUBJECT,W_MESSAGE,W_TITLE,W_ACTIVE"
	sSql = sSql & ") VALUES ("
	sSql = sSql & "'" & a_subject & "','" & a_message & "','" & a_title & "'," & a_active & ")"
	executeThis(sSql)
	strMsg = "<li><b>" & txtWel4 & "</b></li>"
  case else
    'do nothing
end select
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<% 
	  intSkin = getSkin(intSubSkin,2)
	  'breadcrumb here
  	  arg1 = txtAdminHome & "|admin_home.asp"
  	  arg2 = txtWel5 & "|admin_welcome.asp"
  	  arg3 = ""
  	  arg4 = ""
  	  arg5 = ""
  	  arg6 = ""
  
  	  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
spThemeTitle = ""
spThemeBlock1_open(intSkin)
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr><td width="100%">
<table class="grid" cellpadding="0" cellspacing="0" border="0">
<tr><td>
<span class="tCellAlt1" style="padding:6px;margin:1px;" onmouseout="this.className='tCellAlt1';" onmouseover="this.className='tCellHover';">
<a href="admin_welcome.asp"><b><%= txtViewAll %></b></a></span>
<!--span class="tCellAlt1" style="padding:6px;margin:1px;" onmouseout="this.className='tCellAlt1';" onmouseover="this.className='tCellHover';"><a href="admin_welcome.asp?cmd=2"><b><%= txtWel9 %></b></a></span-->
</td></tr></table>
</td></tr>
<tr><td width="100%">
<% 
select case iPgType
  case 1 'edit
	sSQL = "SELECT * FROM " & strTablePrefix & "WELCOME WHERE W_ID = " & w_id
	set rsAnn = my_Conn.execute(sSQL)
	if rsAnn.eof then
  	  response.Write("<center><b>" & txtWel6 & "</b></center>")
	else%>
	  <form name="PostTopic" id="PostTopic" method="post" action="admin_welcome.asp?mode=2">
	  <table width="100%" cellpadding="5" cellspacing="0" border="0">
	  <tr><td width="25%" class="tSubTitle" align="right"></td><td width="75%" class="tSubTitle">Edit Welcome Message</td></tr>
	  <tr><td align="right"><b><%= txtActive %></b></td>
	  <td><select name="W_ACTIVE" id="W_ACTIVE">
	  <option value="1"<%= chkSelect(rsAnn("W_ACTIVE"),1) %>><%= txtYes %></option>
	  <option value="0"<%= chkSelect(rsAnn("W_ACTIVE"),0) %>><%= txtNo %></option>
	  </select></td></tr>
	  <tr><td align="right"><b><%= txtTitle %></b></td>
	  <td><input type="text" size="50" class="textbox" maxlength="200" name="W_TITLE" id="W_TITLE" value="<%= rsAnn("W_TITLE") %>" /></td></tr>
	  <tr><td align="right"><b><%= txtSubject %></b></td>
	  <td><input type="text" class="textbox" size="50" maxlength="200" name="W_SUBJECT" id="W_SUBJECT" value="<%= rsAnn("W_SUBJECT") %>" /></td></tr>
  <% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b>" & txtMsg & "</b>", "" & rsAnn("W_MESSAGE") & ""
  else
  	displayPLAINeditor 1,Trim(CleanCode(rsAnn("W_MESSAGE")))
  end if
  %>  
	  <tr><td width="25%" align="right"></td>
	  <td><input type="submit" name="submit" id="submit" value=" <%= txtSubmit %> " class="button" />
	  <input type="hidden" name="W_ID" id="W_ID" value="<%= rsAnn("W_ID") %>" />
	  </td></tr>
	  </table>
	  </form>
  	<%
	end if
	set rsAnn = nothing
  case 2 'Add %>
	  <form name="addAnn" id="addAnn" method="post" action="admin_welcome.asp?mode=3">
	  <table width="100%" cellpadding="5" cellspacing="0" border="0">
	  <tr><td width="25%" class="tSubTitle" align="right"></td><td width="75%" class="tSubTitle"><%= txtWel7 %></td></tr>
	  <tr><td align="right"><b><%= txtTitle %></b></td>
	  <td><input type="text" size="50" maxlength="200" class="textbox" name="W_TITLE" id="W_TITLE" value="" /></td></tr>
	  <tr><td align="right"><b><%= txtActive %></b></td>
	  <td><select name="W_ACTIVE" id="W_ACTIVE">
	  <option value="1" selected="selected"><%= txtYes %></option>
	  <option value="0"><%= txtNo %></option>
	  </select></td></tr>
	  <tr><td align="right"><b><%= txtSubject %></b></td>
	  <td><input type="text" class="textbox" maxlength="200" size="50" name="W_SUBJECT" id="W_SUBJECT" value="" /></td></tr>
  <% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "W_MESSAGE", "<b>" & txtMsg & "</b>", ""
  else
  	displayPLAINeditor 1,""
  end if
  %>  
	  <tr><td width="25%" align="right"></td>
	  <td><input type="submit" name="submit" id="submit" value=" <%= txtSubmit %> " class="button" />
	  <input type="hidden" name="W_ID" id="W_ID" value="<%= rsAnn("W_ID") %>" />
	  </td></tr>
	  </table>
	  </form>
  <%
  case 3
  case else
sSQL = "SELECT * FROM " & strTablePrefix & "WELCOME ORDER BY W_ID DESC"
set rsAnn = my_Conn.execute(sSQL)
if rsAnn.eof then
  response.Write("<center><b>" & txtWel8 & "</b></center>")
else
  response.Write("<table width=""100%"" cellpadding=""5"" cellspacing=""0"" border=""0"">")
	  if strMsg <> "" then
	    response.Write("<tr><td width=""100%""><ul>" & strMsg & "</ul></td></tr>")
		strMsg = ""
	  end if
  response.Write("<tr><td width=""100%""><hr align=""center"" width=""100%""></td></tr>")
  do until rsAnn.eof %>
    <tr>
	<td align="left">
    <a href="admin_welcome.asp?cmd=1&amp;w_id=<%= rsAnn("W_ID") %>">
	<%= icon(icnEdit,txtEdit,"","","") %></a>
	<% If rsAnn("W_DELETE") = 1 Then %>
    <a href="javascript:delAnn('<%= rsAnn("W_ID") %>');">
	<%= icon(icnDelete,txtDel,"","","") %></a>
	<% End If %>&nbsp;
	<b><%= txtTitle %>:</b> <%= rsAnn("W_TITLE") %></td></tr>
	<tr><td><b><%= txtActive %>:</b> <% if rsAnn("W_ACTIVE") = 1 then response.Write(txtYes) else response.Write(txtNo) end if %></td></tr>
	<tr><td><b><%= txtSubject %>:</b> <%= rsAnn("W_SUBJECT") %></td></tr>
	<tr><td><b><%= txtMsg %>:</b><br /><%= rsAnn("W_MESSAGE") %></td></tr>
    <%
	response.Write("<tr><td><hr align=""center"" width=""100%""></td></tr>")
    rsAnn.movenext
  loop
  response.Write("</table>")
end if
set rsAnn = nothing
end select
%>
</td></tr>
</table>
<%
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp?target=admin_welcome.asp" %><% end if %>