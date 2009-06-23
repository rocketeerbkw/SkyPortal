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

'/**
' * SkyPortal Roster Module
' *
' * This file contains constants used throughout the module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */

Dim PlayerT1, PlayerT2, PlayerT3, PlayerT4, PlayerT5, PlayerT6, PlayerT7, PlayerT8, PlayerT10
Dim VolunteerT1, VolunteerT2, VolunteerT3, VolunteerT4, VolunteerT5, VolunteerT6, VolunteerT7, VolunteerT8, VolunteerT9, VolunteerT10
Dim SponsorT1, SponsorT2, SponsorT3, SponsorT4, SponsorT5, SponsorT6, SponsorT7, SponsorT8, SponsorT9, SponsorT10
Dim rosterEmptyFieldCharacter, rosterMasterPagesize, rstrPlayerPageSize, rstrVolunteerPageSize, rosterIDCurrentYear

PlayerT1 = "Shoots"
PlayerT2 = ""
PlayerT3 = ""
PlayerT4 = ""
PlayerT5 = ""
PlayerT6 = ""
PlayerT7 = ""
PlayerT8 = ""
PlayerT9 = ""
PlayerT10 = ""

dim rstrArrRosterExt(0)
rstrArrRosterExt(0) = "T1"

VolunteerT1 = ""
VolunteerT2 = ""
VolunteerT3 = ""
VolunteerT4 = ""
VolunteerT5 = ""
VolunteerT6 = ""
VolunteerT7 = ""
VolunteerT8 = ""
VolunteerT9 = ""
VolunteerT10 = ""


SponsorT1 = ""
SponsorT2 = ""
SponsorT3 = ""
SponsorT4 = ""
SponsorT5 = ""
SponsorT6 = ""
SponsorT7 = ""
SponsorT8 = ""
SponsorT9 = ""
SponsorT10 = ""

rosterEmptyFieldCharacter = "&nbsp;"
rosterMasterPagesize = 25
rstrPlayerPageSize = rosterMasterPagesize
rstrVolunteerPageSize = rosterMasterPagesize

strSql = "SELECT [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'yearCurrent'"
set rs_Current_Year = my_conn.execute(strSql)

if rs_Current_Year.EOF or rs_Current_Year.BOF then
	rosterIDCurrentYear = 0
else
	rosterIDCurrentYear = rs_Current_Year.Fields("M_VALUE")
end if

set rs_Current_Year = nothing

%>