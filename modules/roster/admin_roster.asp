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

'/**
' * SkyPortal Roster Module
' *
' * This file is a container for all the admin functions
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */

pgType = "manager"
  modPgType = "addForm"
  uploadPg = false
  hasEditor = true
  strEditorElements = "Message,descrip"
%>

<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<% 
strView = ""
iCmd = 0
iID = 0
intPage = 0
strMsg = ""

if Request("v") <> "" or  Request("v") <> " " then
	if len(Request("v")) > 2 then
		closeAndGo("default.asp")
	end if
	strView = Request("v")
end if
if Request("c") <> "" or  Request("c") <> " " then
	if IsNumeric(Request("c")) = True then
		iCmd = cLng(Request("c"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("i") <> "" or  Request("i") <> " " then
	if IsNumeric(Request("i")) = True then
		iID = cLng(Request("i"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("page") <> "" or  Request("page") <> " " then
	if IsNumeric(Request("page")) = True then
		intPage = cLng(Request("page"))
	else
		closeAndGo("stop")
	end if
end if



%>
<script type="text/javascript">
function askDelete(go){
	var del = confirm('Are you sure you want to delete that?');
	if (del==true) {
		window.location = go;
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
</td>
<td class="mainPgCol">
<% 
	intSkin = getSkin(intSubSkin,2)
	rosterBreadcrumbs()
  	  
spThemeTitle = ""
spThemeBlock1_open(intSkin)

select case strView
	case "d"
		rosterDivisions()
	case "l"
		rosterLeagues()
	case "pr"
		rosterPrograms()
	case "pp"
		rosterPlayerPositions()
	case "pl"
		rosterPlayers()
    case "v"
        rosterVolunteers()
	case "s"
		rosterSponsors()
	case "t"
        rosterChkDependencies("t")
		rosterTeams()
	case "tp"
        rosterChkDependencies("tp")
		rosterTeamPhotos()
	case "r"
        rosterChkDependencies("r")
		rosterRoster()
	case "y"
		rosterYears()
	case else
        rosterChkDependencies("all")
		rosterListMenu()
end select
%>

<%
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% 
else
  Response.Redirect "admin_login.asp?target=admin_roster.asp"
end if %>