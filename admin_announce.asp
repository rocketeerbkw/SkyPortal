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
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<% 
iPgType = 0
sMode = 0
a_id = 0
strMsg = ""

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
if Request("a_id") <> "" or  Request("a_id") <> " " then
	if IsNumeric(Request("a_id")) = True then
		a_id = cLng(Request("a_id"))
	else
		closeAndGo("default.asp")
	end if
end if

select case sMode
  case 1 'delete from db
    sSql = "DELETE FROM " & strTablePrefix & "ANNOUNCEMENTS WHERE A_ID=" & a_id
	executeThis(sSql)
	strMsg = "<li><b>" & txtAnn1 & "</b></li>"
  case 2 'edit db
    a_id = request.Form("A_ID")
    a_start = DateToStr(request.Form("START_DATE"))
    a_end = DateToStr(request.Form("END_DATE"))
    a_subject = chkString(request.Form("A_SUBJECT"),"message")
    a_message = chkString(request.Form("Message"),"message")
	a_message = replace(a_message,"</p><p>","<br /><br />")
	a_message = replace(a_message,"<p>","")
	a_message = replace(a_message,"</p>","")
	sSql = "UPDATE " & strTablePrefix & "ANNOUNCEMENTS SET "
	sSql = sSql & "A_SUBJECT='" & a_subject & "'"
	sSql = sSql & ",A_MESSAGE='" & a_message & "'"
	sSql = sSql & ",A_START_DATE='" & a_start & "'"
	sSql = sSql & ",A_END_DATE='" & a_end & "'"
	sSql = sSql & " WHERE A_ID=" & a_id
	executeThis(sSql)
	strMsg = "<li><b>" & txtAnn2 & "</b></li>"
  case 3 'add db
    a_start = DateToStr(request.Form("START_DATE"))
    a_end = DateToStr(request.Form("END_DATE"))
    a_subject = chkString(request.Form("A_SUBJECT"),"message")
    a_message = chkString(request.Form("MESSAGE"),"message")
	a_message = replace(a_message,"</p><p>","<br /><br />")
	a_message = replace(a_message,"<p>","")
	a_message = replace(a_message,"</p>","")
	sSql = "INSERT INTO " & strTablePrefix & "ANNOUNCEMENTS ("
	sSql = sSql & "A_SUBJECT,A_MESSAGE,A_START_DATE,A_END_DATE"
	sSql = sSql & ") VALUES ("
	sSql = sSql & "'" & a_subject & "','" & a_message & "','" & a_start & "','" & a_end & "')"
	executeThis(sSql)
	strMsg = "<li><b>" & txtAnn3 & "</b></li>"
  case else
    'do nothing
end select
%>
<script type="text/javascript">
function delAnn(rid){
var stM
stM = "<%= txtAnn4 %>\n";
stM += "\n<%= txtAnn5 %>\n";
var del=confirm(stM);
if (del==true){
window.location="<%= strHomeURL %>admin_announce.asp?mode=1&a_id="+rid;
}
}
</script>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td width="190" class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
<script type="text/javascript">
// define calendars used on page
 addCalendar("Calendar1", "Select Date", "START_DATE", "PostTopic");
 addCalendar("Calendar2", "Select Date", "END_DATE", "PostTopic");
</script>
</td>
<td class="mainPgCol">
<% 
	intSkin = getSkin(intSubSkin,2)
	  'breadcrumb here
  	  arg1 = txtAdminHome & "|admin_home.asp"
  	  arg2 = txtAnnouncements & "|admin_announce.asp"
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
<span class="tCellAlt1" style="padding:6px;margin:0px;" onmouseout="this.className='tCellAlt1';" onmouseover="this.className='tCellHover';">
<a href="admin_announce.asp"><b><%= txtViewAll %></b></a></span>
</td><td>
<span class="tCellAlt1" style="padding:6px;margin:0px;" onmouseout="this.className='tCellAlt1';" onmouseover="this.className='tCellHover';"><a href="admin_announce.asp?cmd=2"><b><%= txtAnn6 %></b></a></span>
</td></tr></table>
</td></tr>
<tr><td width="100%">
<% 
select case iPgType
  case 1 'edit
	sSQL = "SELECT * FROM " & strTablePrefix & "ANNOUNCEMENTS WHERE A_ID = " & a_id
	set rsAnn = my_Conn.execute(sSQL)
	if rsAnn.eof then
  	  response.Write("<center><b>" & txtAnn7 & "</b></center>")
	else%>
	  <form name="PostTopic" id="PostTopic" method="post" action="admin_announce.asp?mode=2">
	  <table width="100%" cellpadding="5" cellspacing="0" border="0">
	  <tr><td width="25%" class="tSubTitle" align="right"></td><td width="75%" class="tSubTitle"><%= txtAnn8 %></td></tr>
	  <tr><td align="right"><b><%= txtSubject %></b></td>
	  <td><input type="text" class="textbox" name="A_SUBJECT" size="50" maxlength="200" id="A_SUBJECT" value="<%= rsAnn("A_SUBJECT") %>" /></td></tr>
	  <tr><td align="right"><b><%= txtStDt %></b></td>
	  <% 
	  s_date = ChkDate2(rsAnn("A_START_DATE")) 
	  's_date = doublenum(month(s_date)) & "/" & doublenum(day(s_date)) & "/" & year(s_date) '@@KG
	  e_date = ChkDate2(rsAnn("A_END_DATE"))
	  'e_date = doublenum(month(e_date)) & "/" & doublenum(day(e_date)) & "/" & year(e_date)'@@KG
	  %>
	  <td valign="top"><input type="text" class="textbox" name="START_DATE" id="START_DATE" value="<%= s_date %>" readonly />&nbsp;<a href="javascript:showCal('Calendar1')"><img border="0" src="images/icons/SmallCalendar.gif" width="16" height="16"></a>
	  </td></tr>
	  <tr><td align="right"><b><%= txtEndDt %></b></td>
	  <td valign="top"><input type="text" class="textbox" name="END_DATE" id="END_DATE" value="<%= e_date %>" readonly />&nbsp;<a href="javascript:showCal('Calendar2')"><img border="0" src="images/icons/SmallCalendar.gif" width="16" height="16"></a>
	  </td></tr>
  <% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b>" & txtMsg & "</b>", "" & rsAnn("A_MESSAGE") & ""
  else
  	displayPLAINeditor 1,Trim(CleanCode(rsAnn("A_MESSAGE")))
  end if
  %>  
	  <tr><td align="right"></td>
	  <td><input type="submit" name="submit" id="submit" value=" <%= txtSubmit %> " class="button" />
	  <input type="hidden" name="a_id" id="a_id" value="<%= rsAnn("A_ID") %>" />
	  </td></tr>
	  </table>
	  </form>
  	<%
	end if
	set rsAnn = nothing
  case 2 'Add %>
	  <form name="PostTopic" id="PostTopic" method="post" action="admin_announce.asp?mode=3">
	  <table width="100%" cellpadding="5" cellspacing="0" border="0">
	  <tr><td width="25%" class="tSubTitle" align="right"></td><td width="75%" class="tSubTitle"><%= txtAnn9 %></td></tr>
	  <tr><td align="right"><b><%= txtSubject %></b></td>
	  <td><input type="text" class="textbox" maxlength="200" name="A_SUBJECT" size="50" id="A_SUBJECT" value="" /></td></tr>
	  <tr><td align="right"><b><%= txtStDt %></b></td>
	  <td valign="top"><input type="text" class="textbox" name="START_DATE" id="START_DATE" value="<%= date() %>" readonly />&nbsp;<a href="javascript:showCal('Calendar1')"><img border="0" src="images/icons/SmallCalendar.gif" width="16" height="16"></a>
	  </td></tr>
	  <tr><td align="right"><b><%= txtEndDt %></b></td>
	  <td valign="top"><input type="text" class="textbox" name="END_DATE" id="END_DATE" value="<%= date() %>" readonly />&nbsp;<a href="javascript:showCal('Calendar2')"><img border="0" src="images/icons/SmallCalendar.gif" width="16" height="16"></a>
	  </td></tr>
  <% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b>" & txtMsg & "</b>", ""
  else
  	displayPLAINeditor 1,""
  end if
  %>  
	  <tr><td align="right"></td>
	  <td><input type="submit" name="submit" id="submit" value=" <%= txtSubmit %> " class="button" />
	  </td></tr>
	  </table>
	  </form>
  <%
  case 3
  case else
	sSQL = "SELECT * FROM " & strTablePrefix & "ANNOUNCEMENTS ORDER BY A_ID DESC"
	set rsAnn = my_Conn.execute(sSQL)
	if rsAnn.eof then
  	  response.Write("<center><b>" & txtAnn10 & "</b></center>")
	else
  	  response.Write("<table width=""100%"" cellpadding=""0"" cellspacing=""4"" border=""0"">")
	  if strMsg <> "" then
	    response.Write("<tr><td width=""100%""><ul>" & strMsg & "</ul></td></tr>")
		strMsg = ""
	  end if
	  response.Write("<tr><td width=""100%""><hr align=""center"" width=""100%""></td></tr>")
  	  do until rsAnn.eof
	    isExpired = ""
	    if rsAnn("A_END_DATE") < strCurDateString then
		  isExpired = "<span class=""fAlert""><b>" & txtExpired & "</b></span>"
		end if %>
    	<tr>
		<td align="left">
    	<a href="admin_announce.asp?cmd=1&amp;a_id=<%= rsAnn("A_ID") %>">
		<%= icon(icnEdit,txtEdit,"","","") %></a>
    	<a href="javascript:delAnn('<%= rsAnn("A_ID") %>');">
		<%= icon(icnDelete,txtDel,"","","") %></a>
		&nbsp;<%= isExpired %><br />
		<b><%= txtSubject %>:</b> <%= rsAnn("A_SUBJECT") %></td></tr>
		<tr><td><b><%= txtMsg %>:</b><br /><%= rsAnn("A_MESSAGE") %></td></tr>
		<tr><td><b><%= txtStDt %>:</b> <%= ChkDate2(rsAnn("A_START_DATE")) %>
		<br /><b><%= txtEndDt %>:</b> <%= ChkDate2(rsAnn("A_END_DATE")) %>
		</td></tr>
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
<% 
else
  Response.Redirect "admin_login.asp?target=admin_announce.asp"
end if %>