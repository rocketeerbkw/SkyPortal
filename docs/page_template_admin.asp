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

%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
<tr><td class="leftPgCol">
<% intSkin = getSkin(intSubSkin,1) %>
<% 
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<% intSkin = getSkin(intSubSkin,2) %>
<%
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<% 
spThemeTitle = "Admin Template"
spThemeBlock1_open(intSkin)
%>
<table width="100%">
<tr><td width="100%">
Rename this page before modifying it
</td></tr></table>
<%
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp" %><% end if %>