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
%>
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<%
if IsNumeric(Request("mode")) = True then
	iMode = cLng(Request("mode"))
end if

spThemeBlock1_open(intSkin)
select case iMode
  case 1
%>
 <table class=""tPlain"">
  <tr>
    <td class="tSubTitle"><a name="mode"></a><b><%= txtWhtIsModeFunct %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt1">
    <li><b><%= txtBasic %>:</b><br />
    <%= txtBasicTxt %></li>
    <li><b><%= txtPrompt %>:</b><br />
    <%= txtPromptTxt %></li>
    <li><b><%= txtHelp %>:</b><br />
    <%= txtHelpTxt %></li>
    <a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a>
    </td>
  </tr>
 </table>
<%
  case 2 %>
    <!--#include file="includes/faq/pop_config_help.asp" -->
  <%
  case 3 %>
    <!--#include file="includes/faq/pop_ipgate_help.asp" -->
  <%
  case 4 %>
  <%
  case 5 %>
  <%
  case 6 %>
  <%
  case 7 %>
  <%
  case 8 %>
  <%
  case 9 %>
  <%
  case else %>
  <%
end select
spThemeBlock1_close(intSkin)%>

<!--#include file="inc_footer_short.asp" -->
