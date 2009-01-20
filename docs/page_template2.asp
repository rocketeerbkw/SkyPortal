<!--#include file="config.asp" --><%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
CurPageType = "core"
PageName = ""
%>
<!--#include file="inc_functions.asp" -->
<%
CurPageInfoChk = "1"
function CurPageInfo () 
	PageName = "New Page" 
	PageAction = "Browsing<br />" 
	PageLocation = "Default.asp" 
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function 
%>
<!--#include file="inc_top.asp" -->
<% 
':: Insert your functions and processing in this area or lower on the page
':: Placing them here will allow you access to all the available variables
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
<td valign="top" class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1) %>
<!-- insert first column content here -->
<% spThemeTitle = "Title here"
spThemeBlock1_open(intSkin) %>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td align="center" valign="middle">
		Place your left column content here
        </td>
      </tr>
    </table>
<% spThemeBlock1_close(intSkin) %>
		<% 
		':: The function below will put the default main menu in this column
		':: You can find a whole list of available functions that you can use 
		':: in the Admin Area >> Managers >> Layout Manager >> Active Blocks
		menu_fp()
		%>
</td>

<td valign="top" class="mainPgCol">
<% intSkin = getSkin(intSubSkin,2)
  arg1 = "SkyPortal Breadcrumb|default.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6 %>
  
<!-- insert main column content here -->
<% spThemeTitle = "Title here"
spThemeBlock1_open(intSkin) %>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td align="center" valign="middle">
		Any text or anything here will appear inside a themebox
        </td>
      </tr>
    </table>
<% spThemeBlock1_close(intSkin) %>
</td>

<!-- start of the 3rd column. Keep or delete as you need -->
<td valign="top" class="rightPgCol">
<% intSkin = getSkin(intSubSkin,3) %>
<!-- insert third column content here -->
</td>
<!-- end of 3rd column. Keep or delete as you need -->

</tr>
</table>
<!--#include file="inc_footer.asp" -->
