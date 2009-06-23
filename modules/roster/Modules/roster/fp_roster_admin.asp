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
' * This file contains all subs used for administrative purposes
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */


sub rosterBreadcrumbs()
	arg1 = txtAdminHome & "|admin_home.asp"
	arg2 = "Roster|admin_roster.asp"
	select case strView
		case "d"
			arg3 = "Divisions|?v=d"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "l"
			arg3 = "Leagues|?v=l"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "pr"
			arg3 = "Programs|?v=pr"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "pp"
			arg3 = "Positions|?v=pp"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "pl"
			arg3 = "Players|?v=pl"
			arg4 = ""
			arg5 = ""
			arg6 = ""
        case "v"
            arg3 = "Volunteers|?v=v"
            arg4 = ""
            arg5 = ""
            arg6 = ""
		case "s"
			arg3 = "Sponsors|?v=s"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "t"
			arg3 = "Teams|?v=t"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "tp"
			arg3 = "Team Photos|?v=tp"
			arg4 = ""
			arg5 = ""
			arg6 = ""
		case "r"
			arg3 = "Team Roster|?v=r"
			arg4 = ""
			arg5 = ""
			arg6 = ""
        case "y"
			arg3 = "Years|?v=y"
			arg4 = ""
			arg5 = ""
			arg6 = ""
	end select
	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
end sub

sub rosterListMenu()
%>
<a href="?v=l">Leagues</a><br />
<a href="?v=pr">Programs</a><br />
<a href="?v=d">Divisions</a><br />
<a href="?v=s">Sponsors</a><br />
<a href="?v=pp">Positions</a><br />
<a href="?v=pl">Players</a><br />
<!--<a href="?v=r">Roster</a><br />-->
<a href="?v=t">Teams</a><br />
<!--<a href="?v=tp">Team Photos</a><br />-->
<a href="?v=v">Volunteers</a><br />
<a href="?v=y">Years</a><br />
<%
end sub

sub rosterChkDependencies(strSection)
    strsql = ""
    set rs = nothing
    errDep = ""
    blNoDep = true
    'League | None
    'Program | None
    'Division | None
    'Sponsor | None
    'Position | None
    'Player | None
    'Roster | team,player,position,year
    'Team | program,division
    'Team Photo | year,team
    'Year | default year
    
    if strSection = "t" or strSection = "all" then
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "PROGRAM"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=pr"">create a program</a></li>"
        end if
        
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "DIVISION"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=d"">create a division</a></li>"
        end if
    end if
        
    if strSection = "r" or strSection = "v" or strSection = "all" then
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=pp"">create a position</a></li>"
        end if
        
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "PLAYER"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=pl"">create a player</a></li>"
        end if
        
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "VOLUNTEER"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=v"">create a volunteer</a></li>"
        end if
    end if
        
    if strSection = "tp" or strSection = "r" or strSection = "all" then
        strsql = "SELECT [ID] FROM " & STRTABLEPREFIX & "TEAM"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=t"">create a team</a></li>"
        end if
        
        
        strSql = "SELECT [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'year'"
        set rs = my_conn.execute(strsql)
        
        if rs.EOF or rs.BOF then
            errDep = errDep & "<li>You must <a href=""?v=y"">create a year</a></li>"
        elseif rosterIDCurrentYear = 0 then
            errDep = errDep & "<li>You must <a href=""?v=y"">set one of the years as default</a></li>"
        end if
    end if

    if len(errDep) > 0 then
        showMsg "warn","<ul>" & errDep & "</ul>"
        blNoDep = false
    end if
    
    set rs = nothing
    strsql = ""
    errDep = ""

end sub

sub rosterDivisions()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "DIVISION WHERE [ID] = " & iID
				set rs_edit_division = my_conn.execute(strSql)

				rstrDivision = rs_edit_division.fields("DIVISION")
				rstrDescrip = rs_edit_division.fields("DESCRIP")
				rstrStartage = rs_edit_division.fields("STARTAGE")
				rstrEndage = rs_edit_division.fields("ENDAGE")

				set rs_edit_division = nothing
				%>
				<form action="?v=d&c=2&i=<% =iID %>" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Division</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=d&c=1" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Division</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterDivisionForm") = "true" then
				rstrDivision = chkString(request.form("division"), "sqlstring")
				rstrDescrip = chkString(request.form("descrip"), "sqlstring")
				rstrStartage = chkString(request.form("startage"), "sqlstring")
				rstrEndage = chkString(request.form("endage"), "sqlstring")

				if not len(rstrDivision) > 0 then
					errmsg = "<li>Name must not be empty</li>"
				end if
				if not len(rstrStartage) > 0 then
					errmsg = errmsg & "<li>Start Age must not be empty</li>"
				elseif not isNumeric(rstrStartage) then
					errmsg = errmsg & "<li>Start Age must be a number</li>"
				end if
				if not len(rstrEndage) > 0 then
					errmsg = errmsg & "<li>End Age must not be empty</li>"
				elseif not isNumeric(rstrEndage) then
					errmsg = errmsg & "<li>End Age must be a number</li>"
				end if
				if len(errmsg) = 0 then
					if cLng(rstrStartage) > cLng(rstrEndage) then
						errmsg = errmsg & "<li>Start Age cannot be greater than End Age</li>"
					end if
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "DIVISION ([DIVISION],[DESCRIP],[STARTAGE],[ENDAGE],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrDivision & "','" & chkString(rstrDescrip,"message") & "'," & rstrStartage & "," & rstrEndage & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrDivision = ""
						rstrDescrip = ""
						rstrStartage = ""
						rstrEndage = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "DIVISION SET [DIVISION] = '" & rstrDivision & "', [DESCRIP] = '" & chkString(rstrDescrip,"message") & "', [STARTAGE] = " & rstrStartage & ", [ENDAGE] = " & rstrEndage & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="division" id="division" value="<% =rstrDivision %>" /></td>
					</tr>
					<tr>
						<td align="right">Description:</td>
						<td>&nbsp;</td>
						<td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rstrDescrip %></textarea></td>
					</tr>
					<tr>
						<td align="right">Start Age:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="startage" id="startage" size="3" value="<% =rstrStartage %>" /></td>
					</tr>
					<tr>
						<td align="right">End Age:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="endage" id="endage" size="3" value="<% =rstrEndage %>" /></td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterDivisionForm" id="rosterDivisionForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
                strSql = "SELECT [ID] FROM " & STRTABLEPREFIX & "TEAM WHERE [DIVISION_ID] = " & iID
                set rs = my_conn.execute(strSql)
                
                if rs.BOF and rs.EOF then
    				strSql = "DELETE FROM " & STRTABLEPREFIX & "DIVISION WHERE [ID] = " & iID
    				executeThis(strSql)
    				showMsg "success","Deleted"
                else
                    showMsg "err","That division can't be deleted because it has teams associated with it"
                end if
                
                set rs = nothing
			end if

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "DIVISION ORDER BY [STARTAGE]"
			set rs_divisions = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="divisions">
				<tr>
					<td colspan="5" align="center"><h2>Divisions</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=d&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Name</td>
					<td>Start Age</td>
					<td>End Age</td>
				</tr>
				<%
				if rs_divisions.eof or rs_divisions.bof then
					response.write "<tr><td colspan=""5"" align=""center"">No divisions</td></tr>"
				else
					while not rs_divisions.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=d&c=3&i=" & rs_divisions.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=d&c=2&i=" & rs_divisions.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_divisions.fields("DIVISION") & "</td>"
							response.write "<td>" & rs_divisions.fields("STARTAGE") & "</td>"
							response.write "<td>" & rs_divisions.fields("ENDAGE") & "</td>"
						response.write "</tr>"
						rs_divisions.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_divisions = nothing
	end select
end sub

sub rosterLeagues()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "LEAGUE WHERE [ID] = " & iID
				set rs_edit_league = my_conn.execute(strSql)

				rstrLeague = rs_edit_league.fields("LEAGUE")
				rstrDescrip = rs_edit_league.fields("DESCRIP")
				rstrWebsite = rs_edit_league.fields("WEBSITE")

				set rs_edit_league = nothing
				%>
				<form action="?v=l&c=2&i=<% =iID %>" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit League</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=l&c=1" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add League</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterLeagueForm") = "true" then
				rstrLeague = chkString(request.form("league"), "sqlstring")
				rstrDescrip = chkString(request.form("descrip"), "sqlstring")
				rstrWebsite = chkString(request.form("website"), "sqlstring")

				if not len(rstrLeague) > 0 then
					errmsg = "<li>Name must not be empty</li>"
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "LEAGUE ([LEAGUE],[DESCRIP],[WEBSITE],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrLeague & "','" & chkString(rstrDescrip,"message") & "','" & rstrWebsite & "'," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrLeague = ""
						rstrDescrip = ""
						rstrWebsite = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "LEAGUE SET [LEAGUE] = '" & rstrLeague & "', [DESCRIP] = '" & chkString(rstrDescrip,"message") & "', [WEBSITE] = '" & rstrWebsite & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="league" id="league" value="<% =rstrLeague %>" /></td>
					</tr>
					<tr>
						<td align="right">Description:</td>
						<td>&nbsp;</td>
						<td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rstrDescrip %></textarea></td>
					</tr>
					<tr>
						<td align="right">Website:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="website" id="website" value="<% =rstrWebsite %>" /></td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterLeagueForm" id="rosterLeagueForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
				strSql = "DELETE FROM " & STRTABLEPREFIX & "LEAGUE WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "LEAGUE"
			set rs_leagues = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="leagues">
				<tr>
					<td colspan="4" align="center"><h2>Leagues</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=l&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Name</td>
					<td>Website</td>
				</tr>
				<%
				if rs_leagues.eof or rs_leagues.bof then
					response.write "<tr><td colspan=""4"" align=""center"">No leagues</td></tr>"
				else
					while not rs_leagues.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=l&c=3&i=" & rs_leagues.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=l&c=2&i=" & rs_leagues.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_leagues.fields("LEAGUE") & "</td>"
							response.write "<td>" & rs_leagues.fields("WEBSITE") & "</td>"
						response.write "</tr>"
						rs_leagues.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_leagues = nothing
	end select
end sub

sub rosterPrograms()
	if request.form("rosterProgramForm") = "true" then
		rstrProgram = chkString(request.form("program"), "sqlstring")
		rstrDescrip = chkString(request.form("descrip"), "sqlstring")

		if not len(rstrProgram) > 0 then
			errmsg = "<li>Name must not be empty</li>"
		end if

		if len(errmsg) = 0 then
			if iCMD = 1 then
				strSql = "INSERT INTO " & STRTABLEPREFIX & "PROGRAM ([PROGRAM],[DESCRIP],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrProgram & "','" & chkString(rstrDescrip,"message") & "'," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
				executeThis(strSql)
				showMsg "success","Added"
			elseif iCMD = 2 then
				strSql = "UPDATE " & STRTABLEPREFIX & "PROGRAM SET [PROGRAM] = '" & rstrProgram & "', [DESCRIP] = '" & chkString(rstrDescrip,"message") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Edited"
			end if
		else
			showMsg "validation","<ul>" & errmsg & "</ul>"
		end if
	end if
	if iCMD = 3 then
        strSql = "SELECT [ID] FROM " & STRTABLEPREFIX & "TEAM WHERE [PROGRAM_ID] = " & iID
        set rs = my_conn.execute(strSql)
        
        if rs.BOF and rs.EOF then
    		strSql = "DELETE FROM " & STRTABLEPREFIX & "PROGRAM WHERE [ID] = " & iID
    		executeThis(strSql)
    		showMsg "success","Deleted"
        else
            showMsg "err","That program can't be deleted because it has teams associate with it"
        end if
        
        set rs = nothing
	end if

	strSql = "SELECT * FROM " & STRTABLEPREFIX & "PROGRAM"
	set rs_programs = my_conn.execute(strSql)

	%>
	<script type="text/javascript">
		function toggleEdit(rID,bAdd) {
			var strEmpty = '<span id="check" style="display:none;"></span>';
			if (bAdd) {
				$('program').action = '?v=pr&c=1';
				var cntnt = '<input type="text" name="program" id="program" value=""<br /><textarea name="descrip" id="descrip" rows="15" cols="70"></textarea><br /><input type="submit" class="button" value="Submit" /><input type="hidden" name="rosterProgramForm" value="true" /><span id="check" style="display:none;">add</span>';

				if ($('check').innerHTML !== '') {
					if ($('check').innerHTML == 'add') {
						toggleEditor('descrip');
						$('edit').update(strEmpty);
					}
					else {
						toggleEditor('descrip');
						$('edit').update(cntnt);
						toggleEditor('descrip');
					}
				}
				else {
					$('edit').update(cntnt);
					toggleEditor('descrip');
				}

			}
			else {
				$('program').action = '?v=pr&c=2&i=' + rID;
				var cntnt = '<input type="text" name="program" id="program" value="'+$('p1'+rID).innerHTML+'"<br /><textarea name="descrip" id="descrip" rows="15" cols="70">'+$('p2'+rID).innerHTML+'</textarea><br /><input type="submit" class="button" value="Submit" /><input type="hidden" name="rosterProgramForm" value="true" /><span id="check" style="display:none;">'+rID+'</span>';

				if ($('check').innerHTML == rID) {
					toggleEditor('descrip');
					$('edit').update(strEmpty);
				}
				else {
					if ($('check').innerHTML !== '')
						toggleEditor('descrip');

					$('edit').update(cntnt);
					toggleEditor('descrip');
				}
			}
		}
	</script>
	<form action="?v=pr&c=1" method="post" id="program">
		<table border="0" cellpadding="0" cellspacing="3">
			<tr>
				<td>
					<table border="1" cellpadding="2" cellspacing="0" id="programs">
						<tr>
							<td colspan="3" align="center"><h2>Programs</h2></td>
						</tr>
						<tr>
							<td width="50"><a href="#" onClick="toggleEdit(0,true)" ><% =icon(icnPlus,txtAdd,"","","") %></a></td>
							<td>Name</td>
							<td>Description</td>
						</tr>
						<%
						if rs_programs.eof or rs_programs.bof then
							response.write "<tr><td colspan=""3"" align=""center"">No programs</td></tr>"
						else
							while not rs_programs.eof
								response.write "<tr>"
									response.write "<td>"
										response.write "<a href=""javascript:askDelete('?v=pr&c=3&i=" & rs_programs.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
										response.write "<a href=""#"" onClick=""toggleEdit(" & rs_programs.fields("ID") & ",false)"" >" & icon(icnEdit,txtEdit,"","","") & "</a>"
									response.write "</td>"
									response.write "<td><span id=""p1" & rs_programs.fields("ID") & """ >" & rs_programs.fields("PROGRAM") & "</span></td>"
									response.write "<td><span id=""p2" & rs_programs.fields("ID") & """ >" & rs_programs.fields("DESCRIP") & "</span></td>"
								response.write "</tr>"
								rs_programs.movenext
							wend
						end if
						%>
					</table>
				</td>
				<td>
					<div id="edit"><span id="check" style="display:none;"></span></div>
				</td>
			</tr>
		</table>
	</form>
	<%
	set rs_programs = nothing
end sub

sub rosterPlayerPositions()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [ID] = " & iID
				set rs_edit_player_position = my_conn.execute(strSql)

				rstrPlayerPosition = rs_edit_player_position.fields("POSITION")
				rstrDescrip = rs_edit_player_position.fields("DESCRIP")
				rstrSort = rs_edit_player_position.fields("SORT")
                rstrType = rs_edit_player_position.fields("TYPE")

				set rs_edit_player_position = nothing
				%>
				<form action="?v=pp&c=2&i=<% =iID %>" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Position</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=pp&c=1" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Position</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterPlayerPositionForm") = "true" then
				rstrPlayerPosition = chkString(request.form("position"), "sqlstring")
				rstrDescrip = chkString(request.form("descrip"), "sqlstring")
				rstrSort = chkString(request.form("sort"), "sqlstring")
                rstrType = chkString(request.form("type"), "sqlstring")

				if not len(rstrPlayerPosition) > 0 then
					errmsg = "<li>Name must not be empty</li>"
				end if
				if not len(rstrSort) > 0 then
					errmsg = errmsg & "<li>Sort must not be empty</li>"
				elseif not IsNumeric(rstrSort) then
					errmsg = errmsg & "<li>Sort must be a number</li>"
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "PLAYER_POSITION ([POSITION],[DESCRIP],[SORT],[TYPE],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrPlayerPosition & "','" & chkString(rstrDescrip,"message") & "'," & rstrSort & ",'" & rstrType & "'," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrPlayerPosition = ""
						rstrDescrip = ""
						rstrSort = ""
                        rstrType = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "PLAYER_POSITION SET [POSITION] = '" & rstrPlayerPosition & "', [DESCRIP] = '" & chkString(rstrDescrip,"message") & "', [SORT] = " & rstrSort & ", [TYPE] = '" & rstrType & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID]= " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="position" id="position" value="<% =rstrPlayerPosition %>" /></td>
					</tr>
					<tr>
						<td align="right">Description:</td>
						<td>&nbsp;</td>
						<td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rstrDescrip %></textarea></td>
					</tr>
                    <tr>
                        <td align="right">Type:</td>
                        <td>&nbsp;</td>
                        <td align="left">
                            <select name="type" id="type">
                                <option value="player" <%=chkSelect("player",rstrType)%>>Player</option>
                                <option value="vol" <%=chkSelect("vol",rstrType)%>>Volunteer</option>
                            </select>
                        </td>
                    </tr>
					<tr>
						<td align="right">Sort Order:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="sort" id="sort" value="<% =rstrSort %>" size="3" /></td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterPlayerPositionForm" id="rosterPlayerPositionForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
				strSql = "DELETE FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "PLAYER_POSITION ORDER BY [SORT]"
			set rs_player_positions = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="player_positions">
				<tr>
					<td colspan="4" align="center"><h2>Positions</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=pp&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Name</td>
					<td>Description</td>
					<td>Sort Order</td>
				</tr>
				<%
				if rs_player_positions.eof or rs_player_positions.bof then
					response.write "<tr><td colspan=""4"" align=""center"">No positions</td></tr>"
				else
					while not rs_player_positions.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=pp&c=3&i=" & rs_player_positions.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=pp&c=2&i=" & rs_player_positions.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_player_positions.fields("POSITION") & "</td>"
							response.write "<td>" & rs_player_positions.fields("DESCRIP") & "</td>"
							response.write "<td>" & rs_player_positions.fields("SORT") & "</td>"
						response.write "</tr>"
						rs_player_positions.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_player_positions = nothing
	end select
end sub

sub rosterPlayers()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "PLAYER WHERE [ID] = " & iID
				set rs_edit_player = my_conn.execute(strSql)

				rstrFirstName = rs_edit_player.fields("FIRSTNAME")
				rstrLastName = rs_edit_player.fields("LASTNAME")
				rstrSex = rs_edit_player.fields("SEX")
				rstrBirthdate = split(rs_edit_player.fields("BIRTHDATE"), "/")
				rstrBd1 = rstrBirthdate(0)
				rstrBd2 = rstrBirthdate(1)
				rstrBd3 = rstrBirthdate(2)
                rstrPhone = rs_edit_player.fields("PHONE")
                rstrCell = rs_edit_player.fields("CELL")
                rstrEmail = rs_edit_player.fields("EMAIL")
                rstrPic = rs_edit_player.fields("PIC")

				for i=1 to 10
					if len(eval("PlayerT" & i)) > 0 then
						execute("rstrT" & i & " = rs_edit_player.fields(""T" & i & """)")
					end if
				next

				set rs_edit_player = nothing
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Player</h2></td>
					</tr>
                    <form id="playerForm" action="?v=pl&c=2&i=<% =iID %>" method="post">
				<%
			else
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Player</h2></td>
					</tr>
                    <form id="playerForm" action="?v=pl&c=1" method="post">
				<%
			end if
			if request.form("rosterPlayerForm") = "true" then
				rstrFirstName = chkString(request.form("firstname"), "sqlstring")
				rstrLastName = chkString(request.form("lastname"), "sqlstring")
				rstrSex = chkString(request.form("sex"), "sqlstring")
				rstrBd1 = chkString(request.form("bd1"), "sqlstring")
				rstrBd2 = chkString(request.form("bd2"), "sqlstring")
				rstrBd3 = chkString(request.form("bd3"), "sqlstring")
                rstrPhone = chkString(request.form("phone"), "sqlstring")
                rstrCell = chkString(request.form("cell"), "sqlstring")
                rstrEmail = chkString(request.form("email"), "sqlstring")
                rstrPic = chkString(request.form("pic"), "sqlstring")
                
                response.write request.form("firstname")
                response.write rstrFirstName

				for i=1 to 10
					if len(eval("PlayerT" & i)) > 0 then
						execute("rstrT" & i & " = chkString(request.form(""t" & i & """), ""sqlstring"")")
					end if
				next

				if not len(rstrFirstName) > 0 then
					errmsg = "<li>First Name must not be empty</li>"
				end if
				if not len(rstrLastName) > 0 then
					errmsg = errmsg & "<li>Last Name must not be empty</li>"
				end if
                if len(rstrPhone) > 0 and not isNumeric(rstrPhone) then
                    errmsg = errmsg & "<li>Only use numbers for the phone number</li>"
                end if
                if len(rstrCell) > 0 and not isNumeric(rstrCell) then
                    errmsg = errmsg & "<li>Only use numbers for the cell number</li>"
                end if
                if len(rstrEmail) > 0 and not IsValidEmail(rstrEmail) then
                    errmsg = errmsg & "<li>Email must be valid</li>"
                end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "PLAYER ([FIRSTNAME],[LASTNAME],[SEX],[BIRTHDATE],[PHONE],[CELL],[EMAIL],[PIC],"
						for i=1 to 10
							if len(eval("PlayerT" & i)) > 0 then
								strSql = strSql & "[T" & i & "],"
							end if
						next
						strSql = strSql & "[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrFirstName & "','" & rstrLastName & "','" & rstrSex & "','" & rstrBd1 & "/" & rstrBd2 & "/" & rstrBd3 & "','" & rstrPhone & "','" & rstrCell & "','" & rstrEmail & "','" & rstrPic & "',"
						for i=1 to 10
							if len(eval("PlayerT" & i)) > 0 then
								strSql = strSql & "'" & eval("rstrT" & i) & "',"
							end if
						next

						strSql = strSql & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrFirstName = ""
						rstrLastName = ""
                        rstrSex = ""
                        rstrPhone = ""
                        rstrCell = ""
                        rstrEmail = ""
                        rstrPic = ""

						for i=1 to 10
							if len(eval("PlayerT" & i)) > 0 then
								execute("rstrT" & i & " = """"")
							end if
						next
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "PLAYER SET [FIRSTNAME] = '" & rstrFirstName & "', [LASTNAME] = '" & rstrLastName & "', [SEX] = '" & rstrSex & "', [BIRTHDATE] = '" & rstrBd1 & "/" & rstrBd2 & "/" & rstrBd3 & "', [PHONE] = '" & rstrPhone & "', [CELL] = '" & rstrCell & "', [EMAIL] = '" & rstrEmail & "', [PIC] = '" & rstrPic & "', "
						for i=1 to 10
							if len(eval("PlayerT" & i)) > 0 then
								strSql = strSql & "[T" & i & "] = '" & eval("rstrT" & i) & "', "
							end if
						next


						strSql = strSql & "[EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			elseif request.querystring("upPhoto") = "upped" then
                returnErr = chkString(request.querystring("err"), "display")
                if returnErr = "true" then
                    showMsg "validation","<b>There was a problem uploading your picture</b><br /><ul>" & chkString(Session.Contents("rosterErr"), "message") & "</ul>"
                    Session.Contents("rosterErr") = ""
                else
                    strSql = "UPDATE " & STRTABLEPREFIX & "PLAYER SET [PIC] = '" & ChkString(request.querystring("photourl"),"sqlstring") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
                    executeThis(strSql)
                    rstrPic = ChkString(request.querystring("photourl"),"sqlstring")
                    showMsg "success","Picture Uploaded"
                end if
            end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">First Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="firstname" id="firstname" value="<% =rstrFirstName %>" /></td>
					</tr>
					<tr>
						<td align="right">Last Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="lastname" id="lastname" value="<% =rstrLastName %>" /></td>
					</tr>
					<tr>
						<td align="right"><%= txtSex %>:</td>
						<td>&nbsp;</td>
						<td align="left">
							<select name="sex">
								<option value="M" <% =chkSelect(rstrSex,"M") %>>Male</option>
								<option value="F" <% =chkSelect(rstrSex,"F") %>>Female</option>
							</select>
						</td>
					</tr>
					<tr>
						<td align="right">Birth Date:</td>
						<td>&nbsp;</td>
						<td align="left">
							<select name="bd1">
								<%
								for i=1 to 12
									response.write "<option value=""" & doublenum(i) & """"
									if clng(rstrBd1)=clng(i) then
										response.write " selected=""selected"""
									end if
									response.write ">" & doublenum(i) & "</option>" & vbcrlf
								next
								%>
							</select>
							<select name="bd2">
								<%
								for i=1 to 31
									response.write "<option value=""" & doublenum(i) & """"
									if clng(rstrBd2)=clng(i) then
										response.write " selected=""selected"""
									end if
									response.write ">" & doublenum(i) & "</option>" & vbcrlf
								next
								%>
							</select>
							<select name="bd3">
								<%
								for i=1908 to 2008
									response.write "<option value=""" & doublenum(i) & """"
									if clng(rstrBd3)=clng(i) then
										response.write " selected=""selected"""
									end if
									response.write ">" & doublenum(i) & "</option>" & vbcrlf
								next
								%>
							</select>
						</td>
					</tr>
                    <tr>
                        <td align="right">Phone:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="phone" id="phone" value="<% =rstrPhone %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Cell:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="cell" id="cell" value="<% =rstrCell %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Email:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="email" id="email" value="<% =rstrEmail %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">URL of Picture:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="pic" id="pic" value="<% =rstrPic %>" /></td>
                    </tr>
					<%
					for i=1 to 10
						if len(eval("PlayerT" & i)) > 0 then
							response.write "<tr>"
							response.write "<td align=""right"">" & eval("PlayerT" & i) & ":</td>"
							response.write "<td>&nbsp;</td>"
							response.write "<td align=""left""><input type=""text"" name=""t" & i & """ id=""t" & i & """ value=""" & eval("rstrT" & i) & """ /></td>"
							response.write "</tr>"
						end if
					next
					%>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterPlayerForm" id="rosterPlayerForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    <tr>
                        <td colspan="3"><hr /></td>
                    </tr>
                    </form>
                    <% if iCMD = 1 then %>
                    <tr>
                        <td colspan="3">You may upload pictures in edit mode only</td>
                    </tr>
                    <% elseif iCMD = 2 then %>
                    <form name="formUpload" action="rosterUpload.asp" method="post" enctype="multipart/form-data">
                    <%
        			strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'roster'"
        			set rsUload = my_Conn.execute(strSQL)
        			uActive = rsUload("UP_ACTIVE")
        			uUpGrps = rsUload("UP_ALLOWEDGROUPS")
        			uSize = rsUload("UP_SIZELIMIT")
        			uExt = rsUload("UP_ALLOWEDEXT")
        			uID = rsUload("ID")
        			set rsUload = nothing
        			session.Contents("uploadType") = uID
        			session.Contents("loggedUser") = strdbntusername
        			%>
                    <tr>
                        <td align="right">Upload Picture:</td>
                        <td>&nbsp;</td>
						<td align="left">
                            <input type="file" name="file1" id="file1" />
    						<input type="hidden" name="folder" value="player" />
    						<input type="hidden" name="cid" value="0" />
    						<input type="hidden" name="sid" value="<%=iID%>" />
                        </td>
                    </tr>
                    <tr>
						<td colspan="3" align="left">
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    </form>
                    <% end if %>
				</table>
			<%
		case else
			if iCMD = 3 then
                'Delete player from all rosters
                strSql = "DELETE FROM " & STRTABLEPREFIX & "ROSTER WHERE [PLAYER_ID] = " & iID
                executeThis(strSql)
				strSql = "DELETE FROM " & STRTABLEPREFIX & "PLAYER WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if
            
            'Let's paginate this stuffs...
            set rs_players = Server.CreateObject("ADODB.Recordset")

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "PLAYER "
            
            if not isBarren(request.querystring("search")) then
                strSql = strSql & " WHERE ([FIRSTNAME] + ' ' + [LASTNAME] LIKE '%" & ChkString(request.querystring("search"),"sqlstring") & "%' OR [LASTNAME] + ' ' + [FIRSTNAME] LIKE '%" & ChkString(request.querystring("search"),"sqlstring") & "%') "
            end if
            
            strSql = strSql & " ORDER BY [LASTNAME],[FIRSTNAME]"
                        
			rs_players.PageSize = rstrPlayerPageSize
            rs_players.CacheSize = rstrPlayerPageSize * 3
            
            rs_players.open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
            
            if not (rs_players.eof or rs_players.bof) then
                if intPage = 0 then
                    rs_players.AbsolutePage = 1
                else
                    rs_players.AbsolutePage = intPage
                end if
                
                rstrPageSize = rs_players.PageSize
                rstrAbsPage = rs_players.AbsolutePage
                rstrPageCount = rs_players.PageCount
            end if

			%>
            <form action="" method="get">
                <input type="hidden" name="v" value="pl" />
			<table border="1" cellpadding="2" cellspacing="0" id="players">
				<tr>
					<td colspan="14" align="center"><h2>Players</h2></td>
				</tr>
                <tr>
                        <td colspan="14" align="center" valign="middle">Search For:&nbsp;<input type="text" name="search" />&nbsp;<input type="submit" value="Go" class="button" /></td>
                    </tr>
				<tr>
					<td width="50"><a href="?v=pl&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>First Name</td>
					<td>Last Name</td>
					<td><%= txtSex %></td>
					<td>Birth Date</td>
				</tr>
				<%
				if rs_players.eof or rs_players.bof then
					response.write "<tr><td colspan=""14"" align=""center"">No players</td></tr>"
				else
					for curRecCnt = 1 to rstrPageSize
                        if not rs_players.EOF then
    						response.write "<tr>"
    							response.write "<td>"
    								response.write "<a href=""javascript:askDelete('?v=pl&c=3&i=" & rs_players.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
    								response.write "<a href=""?v=pl&c=2&i=" & rs_players.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
    							response.write "</td>"
    							response.write "<td>" & rs_players.fields("FIRSTNAME") & "</td>"
    							response.write "<td>" & rs_players.fields("LASTNAME") & "</td>"
    							response.write "<td>" & rs_players.fields("SEX") & "</td>"
    							response.write "<td>" & rs_players.fields("BIRTHDATE") & "</td>"
    						response.write "</tr>"
    						rs_players.movenext
                        end if
					next
				end if
                
                response.write "<tr><td colspan=""14"" align=""center"">"
                for i=rstrAbsPage-3 to rstrAbsPage+3
                    if i <= 0 then
                        'next
                    else
                        if (i=rstrAbsPage-3) and (rstrAbsPage-3 > 1) then
                            response.write "<a href=""?v=" & strView & """>[First]</a> ... "
                        end if
                        if i = rstrAbsPage then
                            response.write "[" & i & "] "
                        elseif i > rstrPageCount then
                            'next
                        else
                            response.write "<a href=""?v=" & strView & "&page=" & i & """>" & i & "</a> "
                            if (i = rstrAbsPage+3) and (rstrAbsPage+3 < rstrPageCount) then
                                response.write "... <a href=""?v=" & strView & "&page=" & rstrPageCount & """>[Last]</a>"
                            end if
                        end if
                    end if
                next
                response.write "</td></tr>"
				%>
			</table>
            </form>
			<%
		set rs_players = nothing
	end select
end sub

sub rosterVolunteers()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "VOLUNTEER WHERE [ID] = " & iID
				set rs_edit_vol = my_conn.execute(strSql)

				rstrFirstName = rs_edit_vol.fields("FIRSTNAME")
				rstrLastName = rs_edit_vol.fields("LASTNAME")
                rstrPhone = rs_edit_vol.fields("PHONE")
                rstrCell = rs_edit_vol.fields("CELL")
                rstrEmail = rs_edit_vol.fields("EMAIL")
                rstrPic = rs_edit_vol.fields("PIC")

				for i=1 to 10
					if len(eval("VolunteerT" & i)) > 0 then
						execute("rstrT" & i & " = rs_edit_vol.fields(""T" & i & """)")
					end if
				next

				set rs_edit_vol = nothing
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Volunteer</h2></td>
					</tr>
                    <form action="?v=v&c=2&i=<% =iID %>" method="post">
				<%
			else
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Volunteer</h2></td>
					</tr>
                    <form action="?v=v&c=1" method="post">
				<%
			end if
			if request.form("rosterVolForm") = "true" then
				rstrFirstName = chkString(request.form("firstname"), "sqlstring")
				rstrLastName = chkString(request.form("lastname"), "sqlstring")
                rstrPhone = chkString(request.form("phone"), "sqlstring")
                rstrCell = chkString(request.form("cell"), "sqlstring")
                rstrEmail = chkString(request.form("email"), "sqlstring")
                rstrPic = chkString(request.form("pic"), "sqlstring")

				for i=1 to 10
					if len(eval("VolunteerT" & i)) > 0 then
						execute("rstrT" & i & " = chkString(request.form(""t" & i & """), ""sqlstring"")")
					end if
				next

				if not len(rstrFirstName) > 0 then
					errmsg = "<li>First Name must not be empty</li>"
				end if
				if not len(rstrLastName) > 0 then
					errmsg = errmsg & "<li>Last Name must not be empty</li>"
				end if
                if len(rstrPhone) > 0 and not isNumeric(rstrPhone) then
                    errmsg = errmsg & "<li>Only use numbers for the phone number</li>"
                end if
                if len(rstrCell) > 0 and not isNumeric(rstrCell) then
                    errmsg = errmsg & "<li>Only use numbers for the cell number</li>"
                end if
                if len(rstrEmail) > 0 and not IsValidEmail(rstrEmail) then
                    errmsg = errmsg & "<li>Email must be valid</li>"
                end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "VOLUNTEER ([FIRSTNAME],[LASTNAME],[PHONE],[CELL],[EMAIL],[PIC],"
						for i=1 to 10
							if len(eval("VolunteerT" & i)) > 0 then
								strSql = strSql & "[T" & i & "],"
							end if
						next
						strSql = strSql & "[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrFirstName & "','" & rstrLastName & "','" & rstrPhone & "','" & rstrCell & "','" & rstrEmail & "','" & rstrPic & "',"
						for i=1 to 10
							if len(eval("VolunteerT" & i)) > 0 then
								strSql = strSql & "'" & eval("rstrT" & i) & "',"
							end if
						next

						strSql = strSql & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrFirstName = ""
						rstrLastName = ""
						rstrPhone = ""
                        rstrCell = ""
                        rstrEmail = ""
                        rstrPic = ""

						for i=1 to 10
							if len(eval("VolunteerT" & i)) > 0 then
								execute("rstrT" & i & " = """"")
							end if
						next
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "VOLUNTEER SET [FIRSTNAME] = '" & rstrFirstName & "', [LASTNAME] = '" & rstrLastName & "', [PHONE] = '" & rstrPhone & "', [CELL] = '" & rstrCell & "', [EMAIL] = '" & rstrEmail & "', [PIC] = '" & rstrPic & "', "
						for i=1 to 10
							if len(eval("VolunteerT" & i)) > 0 then
								strSql = strSql & "[T" & i & "] = '" & eval("rstrT" & i) & "', "
							end if
						next


						strSql = strSql & "[EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			elseif request.querystring("upPhoto") = "upped" then
                returnErr = chkString(request.querystring("err"), "display")
                if returnErr = "true" then
                    showMsg "validation","<b>There was a problem uploading your picture</b><br /><ul>" & chkString(Session.Contents("rosterErr"), "message") & "</ul>"
                    Session.Contents("rosterErr") = ""
                else
                    strSql = "UPDATE " & STRTABLEPREFIX & "VOLUNTEER SET [PIC] = '" & ChkString(request.querystring("photourl"),"sqlstring") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
                    executeThis(strSql)
                    rstrPic = ChkString(request.querystring("photourl"),"sqlstring")
                    showMsg "success","Picture Uploaded"
                end if
            end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">First Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="firstname" id="firstname" value="<% =rstrFirstName %>" /></td>
					</tr>
					<tr>
						<td align="right">Last Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="lastname" id="lastname" value="<% =rstrLastName %>" /></td>
					</tr>
                    <tr>
                        <td align="right">Phone:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="phone" id="phone" value="<% =rstrPhone %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Cell:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="cell" id="cell" value="<% =rstrCell %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Email:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="email" id="email" value="<% =rstrEmail %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">URL of Picture:</td>
                        <td>&nbsp;</td>
						<td align="left"><input type="text" name="pic" id="pic" value="<% =rstrPic %>" /></td>
                    </tr>
					<%
					for i=1 to 10
						if len(eval("VolunteerT" & i)) > 0 then
							response.write "<tr>"
							response.write "<td align=""right"">" & eval("VolunteerT" & i) & ":</td>"
							response.write "<td>&nbsp;</td>"
							response.write "<td align=""left""><input type=""text"" name=""t" & i & """ id=""t" & i & """ value=""" & eval("rstrT" & i) & """ /></td>"
							response.write "</tr>"
						end if
					next
					%>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterVolForm" id="rosterVolForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    </form>
                    <tr>
                        <td colspan="3"><hr /></td>
                    </tr>
                    <% if iCMD = 1 then %>
                    <tr>
                        <td colspan="3">You may upload pictures in edit mode only</td>
                    </tr>
                    <% elseif iCMD = 2 then %>
                    <form name="formUpload" action="rosterUpload.asp" method="post" enctype="multipart/form-data">
                    <%
        			strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'roster'"
        			set rsUload = my_Conn.execute(strSQL)
        			uActive = rsUload("UP_ACTIVE")
        			uUpGrps = rsUload("UP_ALLOWEDGROUPS")
        			uSize = rsUload("UP_SIZELIMIT")
        			uExt = rsUload("UP_ALLOWEDEXT")
        			uID = rsUload("ID")
        			set rsUload = nothing
        			session.Contents("uploadType") = uID
        			session.Contents("loggedUser") = strdbntusername
        			%>
                    <tr>
                        <td align="right">Upload Picture:</td>
                        <td>&nbsp;</td>
						<td align="left">
                            <input type="file" name="file1" id="file1" />
    						<input type="hidden" name="folder" value="volunteer" />
    						<input type="hidden" name="cid" value="0" />
    						<input type="hidden" name="sid" value="<%=iID%>" />
                        </td>
                    </tr>
                    <tr>
						<td colspan="3" align="left">
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    </form>
                    <% end if %>
				</table>
			<%
		case else
			if iCMD = 3 then
                'Delete player from all rosters
                strSql = "DELETE FROM " & STRTABLEPREFIX & "ROSTER WHERE [PLAYER_ID] = " & iID
                executeThis(strSql)
				strSql = "DELETE FROM " & STRTABLEPREFIX & "VOLUNTEER WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if
            
            'Let's paginate this stuffs...
            set rs_vols = Server.CreateObject("ADODB.Recordset")

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "VOLUNTEER "
            
            if not isBarren(request.querystring("search")) then
                strSql = strSql & " WHERE ([FIRSTNAME] + ' ' + [LASTNAME] LIKE '%" & ChkString(request.querystring("search"),"sqlstring") & "%' OR [LASTNAME] + ' ' + [FIRSTNAME] LIKE '%" & ChkString(request.querystring("search"),"sqlstring") & "%') "
            end if
            
            strSql = strSql & " ORDER BY [LASTNAME],[FIRSTNAME]"
                        
			rs_vols.PageSize = rstrVolunteerPageSize
            rs_vols.CacheSize = rstrVolunteerPageSize * 3
            
            rs_vols.open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
            
            if not (rs_vols.eof or rs_vols.bof) then
                if intPage = 0 then
                    rs_vols.AbsolutePage = 1
                else
                    rs_vols.AbsolutePage = intPage
                end if
                
                rstrPageSize = rs_vols.PageSize
                rstrAbsPage = rs_vols.AbsolutePage
                rstrPageCount = rs_vols.PageCount
            end if

			%>
            <form action="" method="get">
                <input type="hidden" name="v" value="v" />
			<table border="1" cellpadding="2" cellspacing="0" id="players">
				<tr>
					<td colspan="14" align="center"><h2>Volunteers</h2></td>
				</tr>
                <tr>
                        <td colspan="14" align="center" valign="middle">Search For:&nbsp;<input type="text" name="search" />&nbsp;<input type="submit" value="Go" class="button" /></td>
                    </tr>
				<tr>
					<td width="50"><a href="?v=v&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>First Name</td>
					<td>Last Name</td>
				</tr>
				<%
				if rs_vols.eof or rs_vols.bof then
					response.write "<tr><td colspan=""14"" align=""center"">No volunteers</td></tr>"
				else
					for curRecCnt = 1 to rstrPageSize
                        if not rs_vols.EOF then
    						response.write "<tr>"
    							response.write "<td>"
    								response.write "<a href=""javascript:askDelete('?v=v&c=3&i=" & rs_vols.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
    								response.write "<a href=""?v=v&c=2&i=" & rs_vols.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
    							response.write "</td>"
    							response.write "<td>" & rs_vols.fields("FIRSTNAME") & "</td>"
    							response.write "<td>" & rs_vols.fields("LASTNAME") & "</td>"
    						response.write "</tr>"
    						rs_vols.movenext
                        end if
					next
				end if
                
                response.write "<tr><td colspan=""14"" align=""center"">"
                for i=rstrAbsPage-3 to rstrAbsPage+3
                    if i <= 0 or rstrPageCount = 1 then
                        'next
                    else
                        if (i=rstrAbsPage-3) and (rstrAbsPage-3 > 1) then
                            response.write "<a href=""?v=" & strView & """>[First]</a> ... "
                        end if
                        if i = rstrAbsPage then
                            response.write "[" & i & "] "
                        elseif i > rstrPageCount then
                            'next
                        else
                            response.write "<a href=""?v=" & strView & "&page=" & i & """>" & i & "</a> "
                            if (i = rstrAbsPage+3) and (rstrAbsPage+3 < rstrPageCount) then
                                response.write "... <a href=""?v=" & strView & "&page=" & rstrPageCount & """>[Last]</a>"
                            end if
                        end if
                    end if
                next
                response.write "</td></tr>"

				%>
			</table>
            </form>
			<%
		set rs_vols = nothing
	end select
end sub

sub rosterSponsors()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "SPONSOR WHERE [ID] = " & iID
				set rs_edit_sponsor = my_conn.execute(strSql)

				rstrSponsor = rs_edit_sponsor.fields("SPONSOR")
                rstrEmail = rs_edit_sponsor.fields("EMAIL")
                rstrURL = rs_edit_sponsor.fields("URL")
                rstrAddress = rs_edit_sponsor.fields("ADDRESS")
                rstrPhone = rs_edit_sponsor.fields("PHONE")
                rstrCell = rs_edit_sponsor.fields("CELL")
                rstrFax = rs_edit_sponsor.fields("FAX")
                rstrDescrip = rs_edit_sponsor.fields("DESCRIP")
                rstrPic = rs_edit_sponsor.fields("PIC")

				for i=1 to 10
					if len(eval("SponsorT" & i)) > 0 then
						execute("rstrT" & i & " = rs_edit_sponsor.fields(""T" & i & """)")
					end if
				next

				set rs_edit_sponsor = nothing
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Sponsor</h2></td>
					</tr>
                    <form action="?v=s&c=2&i=<% =iID %>" method="post">
				<%
			else
				%>
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Sponsor</h2></td>
					</tr>
                    <form action="?v=s&c=1" method="post">
				<%
			end if
			if request.form("rosterSponsorForm") = "true" then
				rstrSponsor = chkString(request.form("sponsor"), "sqlstring")
                rstrEmail = chkString(request.form("email"), "sqlstring")
                rstrURL = chkString(request.form("url"), "sqlstring")
                rstrAddress = chkString(request.form("address"), "sqlstring")
                rstrPhone = chkString(request.form("phone"), "sqlstring")
                rstrCell = chkString(request.form("cell"), "sqlstring")
                rstrFax = chkString(request.form("fax"), "sqlstring")
                rstrDescrip = chkString(request.form("descrip"), "sqlstring")
                rstrPic = chkString(request.form("pic"), "sqlstring")

				for i=1 to 10
					if len(eval("SponsorT" & i)) > 0 then
						execute("rstrT" & i & " = chkString(request.form(""t" & i & """), ""sqlstring"")")
					end if
				next

				if not len(rstrSponsor) > 0 then
					errmsg = "<li>Name must not be empty</li>"
				end if
                if len(rstrEmail) > 0 and not IsValidEmail(rstrEmail) then
                    errmsg = errmsg & "<li>Email must be valid</li>"
                end if
                if len(rstrPhone) > 0 and not isNumeric(rstrPhone) then
                    errmsg = errmsg & "<li>Only use numbers for the phone number</li>"
                end if
                if len(rstrCell) > 0 and not isNumeric(rstrCell) then
                    errmsg = errmsg & "<li>Only use numbers for the cell number</li>"
                end if
                if len(rstrFax) > 0 and not isNumeric(rstrFax) then
                    errmsg = errmsg & "<li>Only use numbers for the fax number</li>"
                end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "SPONSOR ([SPONSOR],[EMAIL],[URL],[ADDRESS],[PHONE],[CELL],[FAX],[DESCRIP],[PIC],"
						for i=1 to 10
							if len(eval("SponsorT" & i)) > 0 then
								strSql = strSql & "[T" & i & "],"
							end if
						next
						strSql = strSql & "[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrSponsor & "', '" & rstrEmail & "', '" & rstrURL & "', '" & chkString(rstrAddress,"message") & "', '" & rstrPhone & "', '" & rstrCell & "', '" & rstrFax & "', '" & ChkString(rstrDescrip,"message") & "', '" & rstrPic & "',"
						for i=1 to 10
							if len(eval("SponsorT" & i)) > 0 then
								strSql = strSql & "'" & eval("rstrT" & i) & "',"
							end if
						next

						strSql = strSql & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrSponsor = ""

						for i=1 to 10
							if len(eval("SponsorT" & i)) > 0 then
								execute("rstrT" & i & " = """"")
							end if
						next
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "SPONSOR SET [SPONSOR] = '" & rstrSponsor & "', [EMAIL] = '" & rstrEmail & "', [URL] = '" & rstrURL & "', [ADDRESS] = '" & chkString(rstrAddress,"message") & "', [PHONE] = '" & rstrPhone & "', [CELL] = '" & rstrCell & "', [FAX] = '" & rstrFax & "', [DESCRIP] = '" & ChkString(rstrDescrip,"message") & "', [PIC] = '" & rstrPic & "', "
						for i=1 to 10
							if len(eval("SponsorT" & i)) > 0 then
								strSql = strSql & "[T" & i & "] = '" & eval("rstrT" & i) & "', "
							end if
						next


						strSql = strSql & "[EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			elseif request.querystring("upPhoto") = "upped" then
                returnErr = chkString(request.querystring("err"), "display")
                if returnErr = "true" then
                    showMsg "validation","<b>There was a problem uploading your logo</b><br /><ul>" & chkString(Session.Contents("rosterErr"), "message") & "</ul>"
                    Session.Contents("rosterErr") = ""
                else
                    strSql = "UPDATE " & STRTABLEPREFIX & "SPONSOR SET [PIC] = '" & ChkString(request.querystring("photourl"),"sqlstring") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
                    executeThis(strSql)
                    rstrPic = ChkString(request.querystring("photourl"),"sqlstring")
                    showMsg "success","Logo Uploaded"
                end if
            end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="sponsor" id="sponsor" value="<% =rstrSponsor %>" /></td>
					</tr>
                    <tr>
                        <td align="right">Description:</td>
						<td>&nbsp;</td>
						<td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rstrDescrip %></textarea></td>
					</tr>
                    <tr>
                        <td align="right">Website:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="url" id="url" value="<% =rstrURL %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Email:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="email" id="email" value="<% =rstrEmail %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Address:</td>
                        <td>&nbsp;</td>
                        <td align="left"><textarea name="address" id="address" rows="4" ><% =rstrAddress %></textarea></td>
                    </tr>
                    <tr>
                        <td align="right">Phone:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="phone" id="phone" value="<% =rstrPhone %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Cell:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="cell" id="cell" value="<% =rstrCell %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Fax:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="fax" id="fax" value="<% =rstrFax %>" /></td>
                    </tr>
					<%
					for i=1 to 10
						if len(eval("SponsorT" & i)) > 0 then
							response.write "<tr>"
							response.write "<td align=""right"">" & eval("SponsorT" & i) & ":</td>"
							response.write "<td>&nbsp;</td>"
							response.write "<td align=""left""><input type=""text"" name=""t" & i & """ id=""t" & i & """ value=""" & eval("rstrT" & i) & """ /></td>"
							response.write "</tr>"
						end if
					next
					%>
                    <tr>
                        <td align="right">URL of Logo:</td>
                        <td>&nbsp;</td>
                        <td align="left"><input type="text" name="pic" id="pic" value="<% =rstrPic %>" /></td>
                    </tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterSponsorForm" id="rosterSponsorForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    </form>
                    <tr>
                        <td colspan="3"><hr /></td>
                    </tr>
                    <% if iCMD = 1 then %>
                    <tr>
                        <td colspan="3">You may upload logos in edit mode only</td>
                    </tr>
                    <% elseif iCMD = 2 then %>
                    <form name="formUpload" action="rosterUpload.asp" method="post" enctype="multipart/form-data">
                    <%
        			strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'roster'"
        			set rsUload = my_Conn.execute(strSQL)
        			uActive = rsUload("UP_ACTIVE")
        			uUpGrps = rsUload("UP_ALLOWEDGROUPS")
        			uSize = rsUload("UP_SIZELIMIT")
        			uExt = rsUload("UP_ALLOWEDEXT")
        			uID = rsUload("ID")
        			set rsUload = nothing
        			session.Contents("uploadType") = uID
        			session.Contents("loggedUser") = strdbntusername
        			%>
                    <tr>
                        <td align="right">Upload Logo:</td>
                        <td>&nbsp;</td>
						<td align="left">
                            <input type="file" name="file1" id="file1" />
    						<input type="hidden" name="folder" value="sponsor" />
    						<input type="hidden" name="cid" value="0" />
    						<input type="hidden" name="sid" value="<%=iID%>" />
                        </td>
                    </tr>
                    <tr>
						<td colspan="3" align="left">
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
                    </form>
                    <% end if %>
				</table>
			<%
		case else
			if iCMD = 3 then
				strSql = "DELETE FROM " & STRTABLEPREFIX & "SPONSOR WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

			strSql = "SELECT * FROM " & STRTABLEPREFIX & "SPONSOR ORDER BY [SPONSOR]"
			set rs_sponsors = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="players">
				<tr>
					<td colspan="12" align="center"><h2>Sponsors</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=s&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Name</td>
					<td>Website</td>
                    <td>Email</td>
				</tr>
				<%
				if rs_sponsors.eof or rs_sponsors.bof then
					response.write "<tr><td colspan=""12"" align=""center"">No sponsors</td></tr>"
				else
					while not rs_sponsors.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=s&c=3&i=" & rs_sponsors.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=s&c=2&i=" & rs_sponsors.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_sponsors.fields("SPONSOR") & "</td>"
                            response.write "<td>" & rs_sponsors.fields("URL") & "</td>"
                            response.write "<td>" & rs_sponsors.fields("EMAIL") & "</td>"
						response.write "</tr>"
						rs_sponsors.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_sponsors = nothing
	end select
end sub

sub rosterTeams()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "TEAM WHERE [ID] = " & iID
				set rs_edit_team = my_conn.execute(strSql)

				rstrTeam = rs_edit_team.fields("TEAM")
				rstrDescrip = rs_edit_team.fields("DESCRIP")
				rstrLeague = rs_edit_team.fields("LEAGUE_ID")
				rstrProgram = rs_edit_team.fields("PROGRAM_ID")
				rstrDivision = rs_edit_team.fields("DIVISION_ID")
				rstrSponsor = rs_edit_team.fields("SPONSOR_ID")
				rstrColorsHome = rs_edit_team.fields("COLORS_HOME")
				rstrColorsAway = rs_edit_team.fields("COLORS_AWAY")
				rstrActive = rs_edit_team.fields("ACTIVE")

				set rs_edit_team = nothing
				%>
				<form action="?v=t&c=2&i=<% =iID %>" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Team</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=t&c=1" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Team</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterTeamForm") = "true" then
				rstrTeam = chkString(request.form("team"), "sqlstring")
				rstrDescrip = chkString(request.form("descrip"), "sqlstring")
				rstrLeague = chkString(request.form("league"), "sqlstring")
				rstrProgram = chkString(request.form("program"), "sqlstring")
				rstrDivision = chkString(request.form("division"), "sqlstring")
				rstrSponsor = chkString(request.form("sponsor"), "sqlstring")
				rstrColorsHome = chkString(request.form("colorshome"), "sqlstring")
				rstrColorsAway = chkString(request.form("colorsaway"), "sqlstring")
				rstrActive = chkString(request.form("active"), "sqlstring")

				if not len(rstrTeam) > 0 then
					errmsg = "<li>Name must not be empty</li>"
				end if
				if rstrProgram = 0 then
					errmsg = errmsg & "<li>You must choose a program</li>"
				end if
				if rstrDivision = 0 then
					errmsg = errmsg & "<li>You must choose a division</li>"
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM ([TEAM],[DESCRIP],[LEAGUE_ID],[PROGRAM_ID],[DIVISION_ID],[SPONSOR_ID],[COLORS_HOME],[COLORS_AWAY],[ACTIVE],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('" & rstrTeam & "','" & chkString(rstrDescrip,"message") & "'," & rstrLeague & "," & rstrProgram & "," & rstrDivision & "," & rstrSponsor & ",'" & rstrColorsHome & "','" & rstrColorsAway & "'," & rstrActive & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrTeam = ""
						rstrDescrip = ""
						rstrLeague = ""
						rstrProgram = ""
						rstrDivision = ""
						rstrSponsor = ""
						rstrColorsHome = ""
						rstrColorsAway = ""
						rstrActive = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "TEAM SET [TEAM] = '" & rstrTeam & "', [DESCRIP] = '" & chkString(rstrDescrip,"message") & "', [LEAGUE_ID] = " & rstrLeague & ", [PROGRAM_ID] = " & rstrProgram & ", [DIVISION_ID] = " & rstrDivision & ", [SPONSOR_ID] = " & rstrSponsor & ", [COLORS_HOME] = '" & rstrColorsHome & "', [COLORS_AWAY] = '" & rstrColorsAway & "', [ACTIVE] = " & rstrActive & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Name:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="team" id="team" value="<% =rstrTeam %>" /></td>
					</tr>
					<tr>
						<td align="right">Description:</td>
						<td>&nbsp;</td>
						<td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rstrDescrip %></textarea></td>
					</tr>
					<tr>
						<td align="right">League:</td>
						<td>&nbsp;</td>
						<td align="left"><% = DoRosterDropDownSm(STRTABLEPREFIX & "LEAGUE","LEAGUE","ID",rstrLeague,"league","","None","","ID") %></td>
					</tr>
					<tr>
						<td align="right">Program:</td>
						<td>&nbsp;</td>
						<td align="left"><% = DoRosterDropDownSm(STRTABLEPREFIX & "PROGRAM","PROGRAM","ID",rstrProgram,"program","","None","","ID") %></td>
					</tr>
					<tr>
						<td align="right">Division:</td>
						<td>&nbsp;</td>
						<td align="left"><% = DoRosterDropDownSm(STRTABLEPREFIX & "DIVISION","DIVISION","ID",rstrDivision,"division","","None","","ID") %></td>
					</tr>
					<tr>
						<td align="right">Sponsor:</td>
						<td>&nbsp;</td>
						<td align="left"><% = DoRosterDropDownSm(STRTABLEPREFIX & "SPONSOR","SPONSOR","ID",rstrSponsor,"sponsor","","None","","ID") %></td>
					</tr>
					<tr>
						<td align="right">Colors Home:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="colorshome" id="colorshome" value="<% =rstrColorsHome %>" /></td>
					</tr>
					<tr>
						<td align="right">Colors Away:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="colorsaway" id="colorsaway" value="<% =rstrColorsAway %>" /></td>
					</tr>
					<tr>
						<td align="right">Active:</td>
						<td>&nbsp;</td>
						<td align="left">
							<select name="active" id="active">
								<option value="1" <%= chkSelect(rstrActive,1) %>>Yes</option>
								<option value="0" <%= chkSelect(rstrActive,0) %>>No</option>
							</select>
						</td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterTeamForm" id="rosterTeamForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
                'Delete from roster
                strSql = "DELETE FROM " & STRTABLEPREFIX & "ROSTER WHERE [TEAM_ID] = " & iID
                executeThis(strSql)
                'Delte team photos
                strSql = "DELETE FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [TEAM_ID] = " & iID
                executeThis(strSql)
                'Delete the team
				strSql = "DELETE FROM " & STRTABLEPREFIX & "TEAM WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

strSql = "	SELECT T.[ID], T.[TEAM], T.[DESCRIP], L.[LEAGUE], P.[PROGRAM], D.[DIVISION], S.[SPONSOR], T.[COLORS_HOME], T.[COLORS_AWAY], T.[ACTIVE]		" &_
"			  FROM (((" & STRTABLEPREFIX & "TEAM AS T																									" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "LEAGUE AS L																							" &_
"			    ON T.[LEAGUE_ID] = L.[ID])																												" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PROGRAM AS P																						" &_
"			    ON T.[PROGRAM_ID] = P.[ID])																												" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION AS D																						" &_
"			    ON T.[DIVISION_ID] = D.[ID])																											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "SPONSOR AS S																						" &_
"				ON T.[SPONSOR_ID] = S.[ID]																												" &_
"			ORDER BY P.[PROGRAM], D.[STARTAGE], T.[TEAM]"

			set rs_teams = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="teams">
				<tr>
					<td colspan="10" align="center"><h2>Teams</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=t&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>


					<td>Program</td>
					<td>Division</td>
					<td>Team</td>
					<td>League</td>
					<td>Active</td>
				</tr>
				<%
				if rs_teams.eof or rs_teams.bof then
					response.write "<tr><td colspan=""10"" align=""center"">No teams</td></tr>"
				else
					while not rs_teams.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=t&c=3&i=" & rs_teams.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=t&c=2&i=" & rs_teams.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"

							response.write "<td>" & rs_teams.fields("PROGRAM") & "</td>"
							response.write "<td>" & rs_teams.fields("DIVISION") & "</td>"
							response.write "<td>" & rs_teams.fields("TEAM") & "</td>"
							response.write "<td>" & rs_teams.fields("LEAGUE") & "</td>"
							response.write "<td>"
							IF rs_teams.fields("ACTIVE") = 1 THEN
							response.write "Yes"
							else
							response.write "No"
							end if
							response.write "</td>"
						response.write "</tr>"
						rs_teams.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_teams = nothing
	end select
end sub

sub rosterTeamPhotos()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [ID] = " & iID
				set rs_edit_team_photo = my_conn.execute(strSql)

				rstrPhoto = rs_edit_team_photo.fields("VALUE")
				rstrYear = rs_edit_team_photo.fields("YEAR")
				rstrTeam = rs_edit_team_photo.fields("TEAM_ID")

				set rs_edit_team_photo = nothing
				%>
				<form action="?v=tp&c=2&i=<% =iID %>" method="post" enctype="multipart/form-data">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Team Photo</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=tp&c=1" method="post" enctype="multipart/form-data">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Team Photo</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterTeamPhotoForm") = "true" then
				rstrPhoto = chkString(request.form("photo_url"), "sqlstring")
                rstrPhotoUp = chkString(request.form("photo_upload"), "sqlstring")
				rstrYear = chkString(request.form("year"), "sqlstring")
				rstrTeam = chkString(request.form("team"), "sqlstring")
                                
                if not len(rstrPhoto) > 0 and not len(rstrPhotoUp) > 0 then
                    errmsg = "<li>You must enter a photo URL, or choose a photo to upload</li>"
                end if
                if len(rstrPhoto) > 0 and len(rstrPhotoUp) > 0 then
                    errmsg = errmsg & "<li>You may enter a URL, or upload a photo, not both</li>"
                end if
				if rstrYear = 0 then
					errmsg = errmsg & "<li>You must choose a year</li>"
				end if
				if rstrTeam = 0 then
					errmsg = errmsg & "<li>You must choose a team</li>"
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','" & rstrPhoto & "'," & rstrTeam & ", " & rstrYear & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrPhoto = ""
						rstrYear = ""
						rstrTeam = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "TEAM_YEARLIES SET [VALUE] = '" & rstrPhoto & "', [TEAM_ID] = " & rstrTeam & ", [YEAR] = " & rstrYear & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Photo URL:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="photo_url" id="photo_url" value="<% =rstrPhoto %>" /></td>
					</tr>
                    <tr>
                        <td colspan="3" align="center">Or</td>
                    </tr>
                    <tr>
                    	<td align="right">Photo Upload:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="file" name="photo_upload" id="photo_upload" /></td>
					</tr>
					<tr>
						<td align="right">Year:</td>
						<td>&nbsp;</td>
						<td align="left">
						<%
						strSql = "SELECT [ID], [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'year'"
						response.write DoRosterDropDown(strSql,"M_VALUE","ID",rstrYear,"year","","None")
						%>
						</td>
					</tr>
					<tr>
						<td align="right">Team:</td>
						<td>&nbsp;</td>
						<td align="left"><%= DoRosterDropDownSm(STRTABLEPREFIX & "TEAM","TEAM","ID",rstrTeam,"team","","None","","ID") %></td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterTeamPhotoForm" id="rosterTeamPhotoForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
				strSql = "DELETE FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

			strSql = "SELECT TY.[ID], T.[TEAM], M.[M_VALUE] AS [YEAR], TY.[VALUE] FROM (" & STRTABLEPREFIX & "TEAM_YEARLIES TY LEFT OUTER JOIN " & STRTABLEPREFIX & "TEAM T ON TY.[TEAM_ID] = T.[ID]) LEFT OUTER JOIN " & STRTABLEPREFIX & "MODS M ON TY.[YEAR] = M.[ID] WHERE TY.[NAME] = 'photo'"
			set rs_team_photos = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="programs">
				<tr>
					<td colspan="4" align="center"><h2>Team Photos</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=tp&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Team</td>
					<td>Year</td>
					<td>Photo</td>
				</tr>
				<%
				if rs_team_photos.eof or rs_team_photos.bof then
					response.write "<tr><td colspan=""4"" align=""center"">No photos</td></tr>"
				else
					while not rs_team_photos.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=tp&c=3&i=" & rs_team_photos.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=tp&c=2&i=" & rs_team_photos.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_team_photos.fields("TEAM") & "</td>"
							response.write "<td>" & rs_team_photos.fields("YEAR") & "</td>"
							response.write "<td>" & rs_team_photos.fields("VALUE") & "</td>"
						response.write "</tr>"
						rs_team_photos.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_team_photos = nothing
	end select
end sub

sub rosterRoster()
	select case iCMD
		case 1,2 'add/edit
			if iCMD = 2 then
				strSql = "SELECT * FROM " & STRTABLEPREFIX & "ROSTER WHERE [ID] = " & iID
				set rs_edit_roster = my_conn.execute(strSql)

				rstrTeam = rs_edit_roster.fields("TEAM_ID")
				rstrPlayer = rs_edit_roster.fields("PLAYER_ID")
				rstrPosition = rs_edit_roster.fields("POSITION_ID")
				rstrRank = rs_edit_roster.fields("RANK")
				rstrYear = rs_edit_roster.fields("YEAR")

				set rs_edit_roster = nothing
				%>
				<form action="?v=r&c=2&i=<% =iID %>" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Edit Roster</h2></td>
					</tr>
				<%
			else
				%>
				<form action="?v=r&c=1" method="post">
				<table border="0" cellpadding="2" cellspacing="0">
					<tr>
						<td colspan="3" align="center"><h2>Add Roster</h2></td>
					</tr>
				<%
			end if
			if request.form("rosterRosterForm") = "true" then
				rstrTeam = chkString(request.form("team"), "sqlstring")
				rstrPlayer = chkString(request.form("player"), "sqlstring")
				rstrPosition = chkString(request.form("position"), "sqlstring")
				rstrRank = chkString(request.form("rank"), "sqlstring")
				rstrYear = chkString(request.form("year"), "sqlstring")

				if rstrTeam = 0 then
					errmsg = "<li>You must choose a team</li>"
				end if
				if rstrPlayer = 0 then
					errmsg = errmsg & "<li>You must choose a player</li>"
				end if
				if rstrPosition = 0 then
					errmsg = errmsg & "<li>You must choose a position</li>"
				end if
				if len(rstrRank) > 0 and not isNumeric(rstrRank) then
					errmsg = errmsg & "<li>Jersey # must be a number</li>"
				end if
				if rstrYear = 0 then
					errmsg = errmsg & "<li>You must choose a year</li>"
				end if

				if len(errmsg) = 0 then
					if iCMD = 1 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "ROSTER ([TEAM_ID],[PLAYER_ID],[POSITION_ID],"
						if len(rstrRank) > 0 then
						strSql = strSql & "[RANK],"
						end if
						strSql = strSql & "[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES (" & rstrTeam & "," & rstrPlayer & "," & rstrPosition & ","
						if len(rstrRank) > 0 then
							strSql = strSql & rstrRank & ","
						end if
						strSql = strSql & rstrYear & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
						executeThis(strSql)
						showMsg "success","Added"

						rstrTeam = ""
						rstrPlayer = ""
						rstrPosition = ""
						rstrRank = ""
						rstrYear = ""
					elseif iCMD = 2 then
						strSql = "UPDATE " & STRTABLEPREFIX & "ROSTER SET [TEAM_ID] = " & rstrTeam & ", [PLAYER_ID] = " & rstrPlayer & ", [POSITION_ID] = " & rstrPosition & ", "
						if len(rstrRank) > 0 then
							strSql = strSql & "[RANK] = " & rstrRank & ", "
						else
							strSql = strSql & "[RANK] = NULL, "
						end if
						strSql = strSql & "[YEAR] = " & rstrYear & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & iID
						executeThis(strSql)
						showMsg "success","Edited"
					end if
				end if
			end if

			if len(errmsg) > 0 then
				showMsg "validation","<ul>" & errmsg & "</ul>"
			end if
			%>
					<tr>
						<td align="right">Team:</td>
						<td>&nbsp;</td>
						<td align="left"><% = DoRosterDropDownSm(STRTABLEPREFIX & "TEAM","TEAM","ID",rstrTeam,"team","","None","","ID") %></td>
					</tr>
					<tr>
						<td align="right">Player:</td>
						<td>&nbsp;</td>
						<td align="left"><%
						strSql = "SELECT [LASTNAME] + ', ' + [FIRSTNAME] AS PLAYER, [ID] FROM " & STRTABLEPREFIX & "PLAYER ORDER BY [LASTNAME],[FIRSTNAME],[ID]"
						response.write DoRosterDropDown(strSql,"PLAYER","ID",rstrPlayer,"player","","None") %></td>
					</tr>
					<tr>
						<td align="right">Position:</td>
						<td>&nbsp;</td>
						<td align="left"><%
						strSql = "SELECT [POSITION], [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION ORDER BY [SORT]"
						response.write DoRosterDropDown(strSql,"POSITION","ID",rstrPosition,"position","","None") %></td>
					</tr>
					<tr>
						<td align="right">Jersey #:</td>
						<td>&nbsp;</td>
						<td align="left"><input type="text" name="rank" id="rank" value="<% =rstrRank %>" /></td>
					</tr>
					<tr>
						<td align="right">Year:</td>
						<td>&nbsp;</td>
						<td align="left">
						<%
						strSql = "SELECT [ID], [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'year'"
						response.write DoRosterDropDown(strSql,"M_VALUE","ID",rstrYear,"year","","None")
						%>
						</td>
					</tr>
					<tr>
						<td colspan="3" align="left">
							<input type="hidden" name="rosterRosterForm" id="rosterRosterForm" value="true" />
							<input type="submit" class="button" name="submit" value="Submit" />
						</td>
					</tr>
				</table>
			</form>
			<%
		case else
			if iCMD = 3 then
				strSql = "DELETE FROM " & STRTABLEPREFIX & "ROSTER WHERE [ID] = " & iID
				executeThis(strSql)
				showMsg "success","Deleted"
			end if

strSql = "	SELECT R.[ID], P.[FIRSTNAME], P.[LASTNAME], T.[TEAM], PP.[POSITION], R.[RANK], M.[M_VALUE] AS [YEAR]		" &_
"			  FROM (((" & STRTABLEPREFIX & "ROSTER AS R													" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER AS P											" &_
"				ON R.[PLAYER_ID] = P.[ID])																" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "TEAM AS T											" &_
"			    ON R.[TEAM_ID] = T.[ID])																" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER_POSITION AS PP								" &_
"			    ON R.[POSITION_ID] = PP.[ID])															" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "MODS AS M											" &_
"				ON R.[YEAR] = M.[ID]																	" &_
"			  ORDER BY T.[TEAM],P.[LASTNAME],P.[FIRSTNAME]"

			set rs_rosters = my_conn.execute(strSql)

			%>
			<table border="1" cellpadding="2" cellspacing="0" id="rosters">
				<tr>
					<td colspan="10" align="center"><h2>Rosters</h2></td>
				</tr>
				<tr>
					<td width="50"><a href="?v=r&c=1"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
					<td>Last Name</td>
					<td>First Name</td>
					<td>Team</td>
					<td>Position</td>
					<td>Jersey #</td>
					<td>Year</td>
				</tr>
				<%
				if rs_rosters.eof or rs_rosters.bof then
					response.write "<tr><td colspan=""10"" align=""center"">No rosters</td></tr>"
				else
					while not rs_rosters.eof
						response.write "<tr>"
							response.write "<td>"
								response.write "<a href=""javascript:askDelete('?v=r&c=3&i=" & rs_rosters.fields("ID") & "');"">" & icon(icnDelete,txtDel,"","","") & "</a>"
								response.write "<a href=""?v=r&c=2&i=" & rs_rosters.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>"
							response.write "</td>"
							response.write "<td>" & rs_rosters.fields("LASTNAME") & "</td>"
							response.write "<td>" & rs_rosters.fields("FIRSTNAME") & "</td>"
							response.write "<td>" & rs_rosters.fields("TEAM") & "</td>"
							response.write "<td>" & rs_rosters.fields("POSITION") & "</td>"
							response.write "<td>" & rs_rosters.fields("RANK") & "</td>"
							response.write "<td>" & rs_rosters.fields("YEAR") & "</td>"
						response.write "</tr>"
						rs_rosters.movenext
					wend
				end if
				%>
			</table>
			<%
		set rs_rosters = nothing
	end select
end sub

sub rosterYears()
	if request.form("rosterYearForm") = "true" then
		rstrYear = chkString(request.form("year"), "sqlstring")

		if not len(rstrYear) > 0 then
			errmsg = "<li>Year must not be empty</li>"
		end if

		if len(errmsg) = 0 then
			if iCMD = 1 then
				strSql = "INSERT INTO " & STRTABLEPREFIX & "MODS ([M_CODE],[M_NAME],[M_VALUE]) VALUES ('year','roster','" & rstrYear & "')"
				executeThis(strSql)
				showMsg "success","Added"
			elseif iCMD = 2 then
				strSql = "UPDATE " & STRTABLEPREFIX & "MODS SET [M_VALUE] = '" & rstrYear & "' WHERE [ID] = " & iID
				executeThis(strSql)

				if request.form("currentYear") = "1" then
					if rosterIDCurrentYear = 0 then
						strSql = "INSERT INTO " & STRTABLEPREFIX & "MODS ([M_CODE],[M_NAME],[M_VALUE]) VALUES ('yearCurrent','roster','" & iID & "')"
					else
						strSql = "UPDATE " & STRTABLEPREFIX & "MODS SET [M_VALUE] = '" & iID & "' WHERE [M_NAME] = 'roster' AND [M_CODE] = 'yearCurrent'"
					end if
					executeThis(strSql)
					rosterIDCurrentYear = iID
				end if
				showMsg "success","Edited"
			end if
		else
			showMsg "validation","<ul>" & errmsg & "</ul>"
		end if
	end if
	if iCMD = 3 then
		strSql = "DELETE FROM " & STRTABLEPREFIX & "MODS WHERE [ID] = " & iID
		executeThis(strSql)
		showMsg "success","Deleted"
	end if

	strSql = "SELECT Y.[ID], Y.[M_VALUE] AS [YEAR] FROM " & STRTABLEPREFIX & "MODS Y WHERE Y.[M_NAME] = 'roster' AND Y.[M_CODE] = 'year'"
	set rs_years = my_conn.execute(strSql)

	%>
	<script type="text/javascript">
		function rosterToggleEdit(rID,bAdd) {
			var form = $('year');
			var inputs = form.getInputs('text');

			if (bAdd) {
				if ($('y'+rID).innerHTML !== "") {
					$('y'+rID).update();
				}
				else {
					var cntnt = '<input type="text" name="year" id="year" value="" />&nbsp;<input type="submit" value="Submit" class="button" />';
					for(i=0;i<=inputs.length-1;i++) {
						//console.log(inputs[i]);
						$(inputs[i]).up().update($(inputs[i]).up().attributes.getNamedItem('orig').value);
						//alert($(inputs[i]).up());
					}
					$('year').action = '?v=y&c=1';
					$('y'+rID).innerHTML = cntnt;
				}
			}
			else {
				if ($('y'+rID).innerHTML !== $('y'+rID).attributes.getNamedItem('orig').value) {
					$('y'+rID).update($('y'+rID).attributes.getNamedItem('orig').value);
				}
				else {
					for(i=0;i<=inputs.length-1;i++) {
						//console.log(inputs[i]);
						$(inputs[i]).up().update($(inputs[i]).up().attributes.getNamedItem('orig').value);
						//alert($(inputs[i]).up());
					}
					$('year').action = '?v=y&c=2&i=' + rID;
					var cntnt = '<input type="text" name="year" id="year" value="'+$('y'+rID).innerHTML+'" />&nbsp;<input type="checkbox" name="currentYear" value="1" />&nbsp;Make this the current year?&nbsp;<input type="submit" value="Submit" class="button" />';
					$('y'+rID).innerHTML = cntnt;
				}
			}
		}
	</script>
	<form action="?v=y&c=1" method="post" id="year">
		<input type="hidden" name="rosterYearForm" id="rosterYearForm" value="true" />
		<table border="1" cellpadding="2" cellspacing="0" id="years">
			<tr>
				<td colspan="2" align="center"><h2>Years</h2></td>
			</tr>
			<tr>
				<td width="50"><a href="#" onClick="rosterToggleEdit(0,true);"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
				<td>Year</td>
			</tr>
			<tr>
				<td colspan="2"><span id="y0" orig=""></span></td>
			</tr>
			<% if rs_years.eof or rs_yearsbof then %>
			<tr>
				<td colspan="2" align="center">No years</td>
			</tr>
			<% else
				while not rs_years.eof
					rstrID = rs_years.fields("ID")
					rstrYear = rs_years.fields("YEAR")
					if cLng(rstrID) = cLng(rosterIDCurrentYear) then
						rstrCurrentYear = "* "
					else
						rstrCurrentYear = ""
					end if
					%>
					<tr>
						<td>
							<a href="javascript:askDelete('?v=tp&c=3&i=<% =rstrID %>');"><% =icon(icnDelete,txtDel,"","","") %></a>
							<a href="#" onClick="rosterToggleEdit(<% =rstrID %>,false)"><% =icon(icnEdit,txtEdit,"","","") %></a>
						</td>
						<td><% =rstrCurrentYear %><span id="y<% =rstrID %>" orig="<% =rstrYear %>"><% =rstrYear %></span></td>
					</tr>
			<%		rs_years.movenext
				wend
			end if
			%>
		</table>* Indicates the current year
	</form>
	<%
	set rs_years = nothing
end sub
%>