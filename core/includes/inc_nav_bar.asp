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
strWebsiteDVersion = "20080519"
strWebsiteBVersion = " 1.0"
strWebsiteVersion = strWebSiteMVersion & "-" & strWebsiteDVersion & "-" & strWebsiteBVersion %>
<script language="JavaScript" type="text/JavaScript">
	var index= 0;
	var contentWin = null;
	function openLoginDialog() {
	    Dialog.alert($('dogin').innerHTML, {windowParameters: {width:250, height:350}, 
        okLabel: "cancel"});
	}
	function openLoginDialog345() {
		var win = new Window('modal_window', {className: "dialog", title: "Login",top:100, left:100,  width:300, height:200, zIndex:150, opacity:1, resizable: true})
		win.getContent().innerHTML = "Hi"
		win.setDestroyOnClose();
		win.show();	
	}
		
	function openLoginDialog123() {
		if (contentWin != null) {
			Dialog.alert("Close the window 'Login' before opening it again!", {windowParameters:{ width:200, height:130}}); 
		}
		else {
			contentWin = new Window('content_win', {className: "dialog", title: "Login", width:300, height:200, zIndex:150, opacity:1, resizable: true, maximizable: false})
			contentWin.setContent('login_form', true, false);
			contentWin.toFront();
			contentWin.setDestroyOnClose();
			contentWin.showCenter();	
		}		
	}
</script>
<% if not hasAccess(2) then 'guest
spThemeNavBar_open()%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr><td align="left" valign="top" width="75%" height="23">
<!--#include file="menu_com.asp" -->
  </td>
<td valign="top" align="right">
<% If strLoginType = 1 and strLockDown = 0 and strAuthType = "db" Then %>
<div id="dogin" class="spThemeNavLog" style="display:none;">
<% login_box() %>
</div>
	<% If not isMAC Then %>
		<!-- <input type="button" value="<%= txtLogin %>" id="submit1" name="submit1" class="btnLogin" onclick="javascript:openLoginDialog();<% If varBrowser = "ie" Then response.Write(" hidFm('formEle');") %>" /> -->
		<input type="button" value="<%= txtLogin %>" id="submit1" name="submit1" class="btnLogin" onclick="javascript:openJsLayer('dogin','250','350');" />
	<% Else %>
		<input type="button" value="<%= txtLogin %>" id="submit1" name="submit1" class="btnLogin" onclick="javascript:mwpHSs('dogin','1');<% If varBrowser = "ie" Then response.Write(" hidFm('formEle');") %>" />
	<% End If %>
<% Else %>
<div id="dogin" style="width:1px; height:1px; display:none; position:absolute; right:5px;"></div>
<% End If %>
</td></tr>
</table>
<%else 'they ARE a member AND logged in 
rptdPosts = getReported
spThemeNavBar_open()%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<form action="<% =Request.ServerVariables("URL") %>" method="post" id="logme" name="logme">
<tr><td align="left" width="1"><%= rptdPosts %></td>
<td align="left">
<!--#include file="menu_com.asp" -->
  </td>
<td align="right"><%= pmimage %>
<% 
If strLoginType <> 2 Then %>
<input type="hidden" name="Method_Type" value="logout" />&nbsp;&nbsp;<span class="fNavMember"><%if strAuthType<>"db" then %> <% =Session(strUniqueID & "username")%> (<% =Session(strUniqueID & "userID") %>)</span>&nbsp;</td>
<%else 
	if strAuthType = "db" then %> <b><% = ChkString(strDBNTUserName, "display") %></b>
</span>&nbsp;</td><td width="1">
<input class="btnLogin" type="submit" value="<%= txtLogout %>" id="logout" name="logout" /><%
	end if 
  end if 
end if %><div id="dogin" style="width:1px; height:1px; display:none; position:absolute; right:1px;"></div></td></tr></form></table>
<%
end if
spThemeNavBar_close()
%>