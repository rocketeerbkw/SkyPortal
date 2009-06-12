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

pgType = "SiteConfig"

My_ID = cLng(request.querystring("id"))
My_Mode = request.querystring("mode")

if My_Mode = "edit" or My_Mode = "compose" then
  hasEditor = true
  strEditorElements = "Message, Message2"
end if
%>
<% server.scripttimeout = 6000 %>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If hasAccess(1) Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	menu_admin()
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdmin & "|admin_home.asp"
  arg2 = txtemUserEmailList & "|admin_emaillist.asp"
  arg3 = txtemEmailManager & "|admin_emailmanager.asp"
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<% spThemeBlock1_open(intSkin) %>
<% 
if My_Mode = "" then
%>
<% 
strSql = "SELECT * FROM " & strTablePrefix & "SPAM ORDER BY ARCHIVE ASC"
set rs = Server.CreateObject("ADODB.Recordset")
rs.open  strSql, My_Conn, 3
%>
<TABLE BORDER=1 class="grid" CELLSPACING=0 align="center" width="100%">
<TR ALIGN="CENTER">
<TD class="tTitle"><B><%=txtStatus%></B></TD>
<TD class="tTitle"><B><%=txtemMsgTitle%></B></TD>
<TD class="tTitle"><B><%=txtemComposed%></B></TD>
<TD class="tTitle"><a href="admin_emailmanager.asp?mode=compose"><img src="images/icons/icon_folder_new_topic.gif" alt="<%=txtemAddNewMsg%>" title="<%=txtemAddNewMsg%>" border="0" hspace="0">&nbsp;<%=txtemAddNewMsg%></a></TD>
</TR>
<%
if RS.eof or RS.bof then
  response.write "<b>No messages found!</b>"
else
  RS.MoveFirst
  do while Not RS.eof                       
  ARCHIVED = rs("ARCHIVE")
  if ARCHIVED = "1" then
  ARCHIVED = "ARCHIVED"
  else
  ARCHIVED = "LIVE"
  end if
  if rs("F_SENT") <> "" then
  F_SENT = ChkDate(rs("F_SENT"))
  else
  F_SENT = "-" 
  end if
 %>
<TR VALIGN=TOP>
<td class="tCellAlt1"><%= ARCHIVED %>&nbsp;</TD>
<td class="tCellAlt1"><input type="hidden" name="ID" value="<%=RS("ID")%>"><a href="admin_emailmanager.asp?mode=edit&id=<%=RS("id")%>"><%=RS("SUBJECT")%></a>&nbsp;</TD>
<td ALIGN="CENTER" class="tCellAlt1"><% =F_SENT %>&nbsp;</TD>
<td class="tCellAlt1" align="right"> <a href="admin_emailmanager.asp?mode=edit&ID=<% =rs("ID") %>"><%= icon(icnEdit,txtemEditMsg,"","","") %></a>
  <a href="admin_emailmanager.asp?mode=update&ID=<% =rs("ID") %>&ARCHIVE=2"><%= icon(icnDelete,txtemDelMsg,"","","") %></a></td>
</TR>
<%
RS.MoveNext
loop%>
</table>
<%
end if
set rs = nothing



elseif My_Mode = "update" then%>
<%
if request.querystring("ARCHIVE")= "2" then
 
		set conn = server.createobject("adodb.connection")
	      	conn.Open My_Conn
		For each record in request("ID")
    		sqlstmt = "DELETE * from " & strTablePrefix & "SPAM WHERE ID=" & My_ID
			Set kRS = conn.execute(sqlstmt)
		Next
 
	set kRS = nothing
%>

<%=replace(replace(txtemMsgDeleted,"[%marker_end%]","</a>"),"[%marker_href%]","<a href=""admin_emailmanager.asp"">")%>

	<%
else

  My_ID = request("ID") 
	strSQL3="select * from " & strTablePrefix & "SPAM where id=" & My_ID
	set kRS=Server.CreateObject("ADODB.Recordset")
	kRS.Open strSQL3, My_Conn, 1, 3
  kRS("SUBJECT") = request("SUBJECT")
  kRS("MESSAGE") = chkString(request("MESSAGE"),"message")	
  kRS("ARCHIVE") = request("ARCHIVE")

	kRS.Update
  %>

  <%=replace(replace(txtemMsgUpdated,"[%marker_end%]","</a>"),"[%marker_href%]","<a href=""admin_emailmanager.asp"">")%>
  <%
  set kRS = nothing

end if

elseif My_Mode = "compose" then 
%>
<h3>Compose New Message</h3>
<form action="admin_emailmanager.asp">
<input type="hidden" name="mode" value="save">
<table class="grid" border="0" cellspacing="0" cellpadding="5">
<tr>
<td><%=txtSubject%>:</td><td><input type="text" name="SUBJECT" size="50"></td>
</tr>
<tr>
<td colspan="2"><%= txtMsg%>:</td>
</tr>
<% 

  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message","", ""
  else
  	displayPLAINeditor 1, ""
  end if
  %>
</table>

Save this message in: &nbsp; 
 
 <select name="ARCHIVE" size="1">
  <option value="0" selected="selected">&nbsp;<%=txtemLiveList%></option>
  <option value="1">&nbsp;<%=txtemArchive%></option>
</select>
 

 &nbsp;<input type="Submit" value="<%=txtSave%>" class="button">&nbsp;<input type="reset" class="button">
 
<%
elseif My_Mode = "save" then

strSubject = chkString(request("SUBJECT"),"message")
strMessage = chkString(request("MESSAGE"),"message")
strArchive = cLng(request("ARCHIVE"))

	executeThis("insert into " & strTablePrefix & "SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) values (" _
		& "'" & strSubject & "', " _
		& "'" & strMessage & "', " _ 
		& "'" & strCurDateString & "', " _		
		& "'" & strArchive & "')")
%>

<%=replace(replace(txtemMsgSaved,"[%marker_end%]","</a>"),"[%marker_href%]","<a href=""admin_emailmanager.asp"">")%>

<%
elseif My_Mode = "edit" then

strSql2 = "SELECT * FROM " & strTablePrefix & "SPAM WHERE ID =" & My_ID
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql2, My_Conn, 3
mySUBJECT = Server.HTMLEncode(rsSP("SUBJECT"))
myMESSAGE = rsSP("MESSAGE")
%>
<h2><%=txtemModifyMsg%></h2>
<form action="admin_emailmanager.asp"><input type="hidden" name="mode" value="update"><input type="hidden" name="ID" value="<%= rsSP("ID") %>">
<table class="grid" border="0" cellspacing="0" cellpadding="5">
<tr>
<td><%=txtSubject%>:</td><td><input type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
</tr>
<tr>
<td colspan="2"><%=txtMsg%>:</td>
</tr>
<% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message","", myMESSAGE
  else
  	displayPLAINeditor 1, CleanCode(myMESSAGE)
  end if %>
</table>
Message Status:&nbsp; 
 <select name="ARCHIVE" size="1">
  <option value="0" selected="selected">&nbsp;<%=txtemLiveList%></option>
  <option value="1">&nbsp;<%=txtemArchive%></option>
  <option value="2">&nbsp;<%=txtDel%></option>
</select>
 &nbsp;<input type="Submit" value="<%=txtemModify%>" class="button">&nbsp;<input type="reset" class="button">
<%
set rsSP = nothing
%>
<br /><br />
</form>
<% end if
spThemeBlock1_close(intSkin) %>
</td></tr></table>
<!--#include file="inc_footer.asp" -->

<% else %>
<%Response.Redirect "admin_login.asp" %>
<% end iF %>