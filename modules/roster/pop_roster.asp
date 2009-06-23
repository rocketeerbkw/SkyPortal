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
' * This file is a container for viewing and editing team, player,
' *   roster, etc data
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */

CurPageType = "core"
hasEditor = true
strEditorElements = "descrip"
%>
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="modules/roster/roster_functions.asp" -->
<%
sMsg = ""
hasPerm = false

'mode
'command
'section
'cid
'sid

'roster/contact/team/player
'view/add/edit/delete
strMode = ""
strCmd = ""
strSec = ""
c_id = 0
s_id = 0
x_id = 0
intPage = 0

strMode = iif(isBarren(Request("mode")),"",Request("mode"))
strCmd  = iif(isBarren(Request("cmd")) ,"",Request("cmd"))
strSec  = iif(isBarren(Request("sec")) ,"",Request("sec"))

if Request("cid") <> "" or  Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		c_id = cLng(Request("cid"))
	else
		closeAndGo("stop")
	end if
end if
if Request("sid") <> "" or  Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		s_id = cLng(Request("sid"))
	else
		closeAndGo("stop")
	end if
end if
if Request("xid") <> "" or  Request("xid") <> " " then
	if IsNumeric(Request("xid")) = True then
		x_id = cLng(Request("xid"))
	else
		closeAndGo("stop")
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
<!--#include file="inc_top_short.asp" -->
<%
setAppPerms "roster","iName"

spThemeBlock1_Open(intSkin)
Select case strMode
    case "roster"
        Select case strCmd
            case "view"
                if bAppRead then
                    pop_roster("view")
                else
                    showMsg "warn","You don't have permission to do that."
                end if
                
            case "cview"
                pop_roster("cview")
            
            case "add","cadd"
                if bAppFull then
                    pop_roster(strCmd)
                else
                    showMsg "warn","You don't have permission to do that."
                end if
            
            case "edit","cedit"
                if bAppWrite then
                    pop_roster(strCmd)
                else
                    showMsg "warn","You don't have permission to do that."
                end if
            
            case "delete"
                if bAppFull then
                    pop_roster("delete")
                else
                    showMsg "warn","You don't have permission to do that."
                end if
            
            case else
                showMsg "warn","You can't do that! >.<"
            
        End Select
        
    case "team"
        select case strCmd
            case "view"
                pop_team_xtras("")
                
            case "edit"
                if bAppWrite then
                    pop_team()
                else
                    showMsg "warn","You don't have permission to do that."
                end if
            
            case else
                pop_team_xtras(strCmd)
        
        end select
        
    case "player"
        select case strCmd
            case "list"
                listPlayers()
                
            case else
                showMsg "info","Nothing here."
                
        end select
    
    case else
        showMsg "warn","You can't do that! >.<"

End Select
spThemeBlock1_Close(intSkin)

%>
<!--#include file="inc_footer_short.asp" -->