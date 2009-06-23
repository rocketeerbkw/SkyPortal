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
' * This file contains all subs used displaying and editing data
' *   from the user side
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */
%>
<!--#include file="roster_config.asp"-->
<!--#include file="roster_functions.asp"-->
<!--#include file="fp_roster_admin.asp"-->
<%

sub searchTeam()
    strSearchSql = ""
    blAdvanced = false
    blSearchPlayer = false
    strSearchSql = "SELECT T.[ID], T.[TEAM], T.[PROGRAM_ID], P.[PROGRAM], T.[DIVISION_ID], D.[DIVISION], D.[STARTAGE], T.[ACTIVE], 0 AS [YEAR] " &_
                    " FROM (" & STRTABLEPREFIX & "TEAM AS T	" &_
                    " LEFT OUTER JOIN " & STRTABLEPREFIX & "PROGRAM AS P ON T.[PROGRAM_ID] = P.[ID]) " &_
                    " LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION AS D ON T.[DIVISION_ID] = D.[ID] " &_
                    "WHERE T.[ACTIVE] = 1 " &_
                    "ORDER BY P.[PROGRAM], D.[STARTAGE], D.[DIVISION], T.[TEAM] "
    
    if request.form("submit") = "Search" then
        if request.form("sFor") = "team" then
            strSearchSql = "SELECT T.[ID], T.[TEAM], T.[PROGRAM_ID], P.[PROGRAM], T.[DIVISION_ID], D.[DIVISION], D.[STARTAGE], T.[ACTIVE], 0 AS [YEAR] " &_
                            " FROM (" & STRTABLEPREFIX & "TEAM AS T " &_
                            " LEFT OUTER JOIN " & STRTABLEPREFIX & "PROGRAM AS P ON T.[PROGRAM_ID] = P.[ID]) " &_
                            " LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION AS D ON T.[DIVISION_ID] = D.[ID] " &_
                            "WHERE "
            
            rstrSearchInactive = iif(request.form("sInactive") = "1", true, false)
            if rstrSearchInactive then
                strSearchSql = strSearchSql & " (T.[ACTIVE] = 0 OR T.[ACTIVE] = 1)"
                blAdvanced = true
            else
                strSearchSql = strSearchSql & " T.[ACTIVE] = 1"
            end if
            strSearchSql = strSearchSql & " AND T.[TEAM] LIKE '%" & chkString(request.form("sPhrase"),"sqlstring") & "%'"
    		if request.form("sProgram") > 0 then
    			strSearchSql = strSearchSql & " AND T.[PROGRAM_ID] IN (" & chkString(request.form("sProgram"),"sqlstring")  & ")"
                blAdvanced = true
    		end if
    		if request.form("sDivision") > 0 then
    			strSearchSql = strSearchSql & " AND T.[DIVISION_ID] IN (" & chkString(request.form("sDivision"),"sqlstring") & ")"
                blAdvanced = true
    		end if
            
            strSearchSql = strSearchSql & " ORDER BY P.[PROGRAM], D.[STARTAGE], D.[DIVISION], T.[TEAM] "
        
        elseif request.form("sFor") = "player" then
            blSearchPlayer = true
            if isBarren(request.form("sPhrase")) then
                errmsg = errmsg & "<li>Please enter the name of a player to search for</li>"
                blSearchPlayer = false
            end if
            
            if len(errmsg) = 0 then
                strSearchSql = "SELECT T.[ID], T.[TEAM], T.[ACTIVE], R.[YEAR], M.[M_VALUE], P.[POSITION], PL.[FIRSTNAME], PL.[LASTNAME] " &_
                                " FROM (((" & STRTABLEPREFIX & "ROSTER R " &_
                                " LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER PL ON R.[PLAYER_ID] = PL.[ID]) " &_
                                " LEFT OUTER JOIN " & STRTABLEPREFIX & "TEAM T ON R.[TEAM_ID] = T.[ID]) " &_
                                " LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER_POSITION P ON R.[POSITION_ID] = P.[ID]) " &_
                                " LEFT OUTER JOIN " & STRTABLEPREFIX & "MODS M ON R.[YEAR] = M.[ID] " &_
                                "WHERE "
         
                rstrSearchInactive = iif(request.form("sInactive") = "1", true, false)
                if rstrSearchInactive then
                    strSearchSql = strSearchSql & " (T.[ACTIVE] = 0 OR T.[ACTIVE] = 1) "
                    blAdvanced = true
                else
                    strSearchSql = strSearchSql & " T.[ACTIVE] = 1 "
                end if
         
                strSearchSql = strSearchSql & " AND (PL.[FIRSTNAME] + ' ' + PL.[LASTNAME] LIKE '%" & ChkString(request.form("sPhrase"),"sqlstring") & "%' OR PL.[LASTNAME] + ' ' + PL.[FIRSTNAME] LIKE '%" & ChkString(request.form("sPhrase"),"sqlstring") & "%') "
            else
                showMsg "validation","<ul>" & errmsg & "</ul>"
            end if
        
        end if
    end if


	spThemeTitle = "Search For Team"
    spThemeBlock1_open(intSkin)
	%>
    <script type="text/javascript">
        function showAdvSearch() {
            var sFor = document.forms['blah'].sFor;
            for (var i=0; i < sFor.length; i++) {
                if (sFor[i].checked) {
                    var sForValue = sFor[i].value;
                }
            }
            
            switch(sForValue) {
                case 'team':
                    $('sAdvanced').update($('advTeam').innerHTML);
                    break;
                    
                case 'player':
                    $('sAdvanced').update($('advPlayer').innerHTML);
                    break;
                    
                default:
                    break;
            }
            
            if (!$('2check').visible())
                $('sAdvanced').show();
        }
        
        function teamHS(div) {
            if($(div).style.display != "none"){
        		swap(div,0);
        	}else{
        		swap(div,1);
        	}
            $(div).toggle();
        }
    </script>
	<form method="post" action="" name="blah">
	<table border="0" cellpadding="0" cellspacing="3" align="center">
        <tr>
            <td><label for="sPhrase">Search For:</label></td><td><input type="text" name="sPhrase" id="sPhrase" value="<% = request.form("sPhrase") %>" /></td>
        </tr>
        <tr>
            <td>&nbsp;</td><td><label for="sForTeam">Team:</label> <input type="radio" name="sFor" id="sForTeam" value="team" <%= iif(isBarren(request.form("sFor")),"checked=""checked""",chkRadio(request.form("sFor"),"team")) %> onChange="showAdvSearch()" /> <label for="sForPlayer">Player:</label> <input type="radio" name="sFor" id="sForPlayer" value="player" <%= chkRadio(request.form("sFor"),"player") %>  onChange="showAdvSearch()" /></td>
        </tr>
        <tr>
            <td colspan="2">
                <a href="#" id="2check" onClick="$(this).hide();showAdvSearch();" style="<% =iif(blAdvanced,"display:none;","") %>">View Advanced Search Options</a>
            </td>
		</tr>
        <tbody id="sAdvanced" style="display:none;">
        </tbody>
		<tr>	
            <td><input type="submit" name="submit" value="Search" class="button" /></td>
		</tr>
	</table>
	</form>
    
    <table style="display:none;">
        <tbody id="advTeam">
            <tr>
                <td><label for="sInactive">Show Inactive<br />teams?</label></td>
                <td><input type="checkbox" name="sInactive" id="sInactive" value="1" <%=chkCheckbox(request.form("sInactive"),1,true) %> /></td>
            </tr>
            <tr>
                <td><label for="sProgram">Program:</label></td><td><% = DoRosterDropDownSm(STRTABLEPREFIX & "PROGRAM","PROGRAM","ID",request.form("sProgram"),"sProgram","id=""sProgram""","All","","ID") %></td>
            </tr>
            <tr>
                <td><label for="sDivision">Division:</label></td><td><% = DoRosterDropDownSm(STRTABLEPREFIX & "DIVISION","DIVISION","ID",request.form("sDivision"),"sDivision","id=""sDivision""","All","","ID") %></td>
            </tr>
        </tbody>
    </table>
    <table style="display:none;">
        <tbody id="advPlayer">
            <tr>
                <td><label for="sInactive">Show Inactive<br />teams?</label></td>
                <td><input type="checkbox" name="sInactive" id="sInactive" value="1" <%=chkCheckbox(request.form("sInactive"),1,true) %> /></td>
            </tr>
        </tbody>
    </table>
    <script type="text/javascript">
    showAdvSearch();
    </script>
	<%
	spThemeBlock1_Close(intSkin)
        	
	'response.write strSearchSql
	'response.end
	set rsSearchTeams = my_conn.execute(strSearchSql)
	
	spThemeTitle = ""
	spThemeBlock1_open(intSkin)
	
	rstrRowClass = "tCellAlt0"
    
    if blSearchPlayer then
	%>
    <table border="0" cellpadding="3" cellspacing="1" align="center">
        <tr>
            <td class="tSubTitle">Name</td>
            <td class="tSubTitle">Team</td>
            <td class="tSubTitle">Position</td>
            <td class="tSubTitle">Year</td>
        </tr>
        <%
        if rsSearchTeams.eof or rsSearchTeams.bof then
            response.write "<tr class=""" & rstrRowClass & """><td colspan=""9"" align=""center"" class=""fNorm"">No players found</td></tr>"
        else
            while not rsSearchTeams.eof
                response.write "<tr class=""" & rstrRowClass & """>"
                response.write "<td class=""fNorm"">" & rsSearchTeams.fields("FIRSTNAME") & " " & rsSearchTeams.fields("LASTNAME") & "</td>"
                response.write "<td class=""fNorm""><a href=""?view=team&team=" & rsSearchTeams.Fields("ID") & iif(blSearchPlayer, "&year=" & rsSearchTeams.Fields("YEAR"), "") & """>" & rsSearchTeams.fields("TEAM") & "</a></td>"
                response.write "<td class=""fNorm"">" & rsSearchTeams.fields("POSITION") & "</td>"
                response.write "<td class=""fNorm"">" & rsSearchTeams.fields("M_VALUE") & "</td>"
                response.write "<tr>"
                if rstrRowClass = "tCellAlt0" then
                    rstrRowClass = "tCellAlt1"
                else
                    rstrRowClass = "tCellAlt0"
                end if
                rsSearchTeams.movenext
            wend
        end if
        %>
    </table>
    <% else %>
    <table border="0" cellpadding="3" cellspacing="0" align="center" width="200">
    <%
        strCurProg = ""
        strCurDiv = ""
        cntDiv = 0
        rstrOpenMinMax = false
        rstrFirstTeam = true
        if rsSearchTeams.eof or rsSearchTeams.bof then
            response.write vbTab & "<tr class=""" & rstrRowClass & """><td colspan=""9"" align=""center"" class=""fNorm"">No teams found</td></tr>"
        else
            while not rsSearchTeams.EOF
                if strCurProg <> rsSearchTeams.fields("PROGRAM_ID") then
                    if rstrOpenMinMax then
                        response.write vbTab & "</tbody>" & vbCrlf
                        rstrOpenMinMax = false
                    end if
                    response.write vbTab & vbTab & "<tr>" & vbCrlf
                    response.write vbTab & vbTab & vbTab & "<td colspan=""3"" class=""tSubTitle"">" & rsSearchTeams.fields("PROGRAM") & "</td>" & vbCrlf
                    response.write vbTab & vbTab & "</tr>" & vbCrlf
                    strCurProg = rsSearchTeams.fields("PROGRAM_ID")
                    cntDiv = 0
                end if
                if strCurDiv <> rsSearchTeams.fields("DIVISION_ID") then
                    if rstrOpenMinMax then
                        response.write vbTab & "</tbody>" & vbCrlf
                        rstrOpenMinMax = false
                    end if
                    response.write vbTab & vbTab & "<tr>" & vbCrlf
                    response.write vbTab & vbTab & vbTab & "<td></td>" & vbCrlf
                    response.write vbTab & vbTab & vbTab & "<td colspan=""2"" class=""tAltSubTitle"">"
                    if rstrFirstTeam then
                        %>
                        <img name="divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>Img" id="divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>Img" src="Themes/<%= strTheme %>/icon_min.gif" onclick="javascript:teamHS('divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>');" style="cursor:pointer;" alt="<%= txtCollapse %>" title="<%= txtCollapse %>" />
                        <%
                    else
                        %>
                        <img name="divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>Img" id="divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>Img" src="Themes/<%= strTheme %>/icon_max.gif" onclick="javascript:teamHS('divminmax<%= rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv %>');" style="cursor:pointer;" alt="<%= txtExpand %>" title="<%= txtExpand %>" />
                        <%
                    end if
                    response.write rsSearchTeams.fields("DIVISION") & "</td>" & vbCrlf
                    response.write vbTab & vbTab & "</tr>" & vbCrlf
                    'This is the opening container for teams under a division
                    response.write vbTab & "<tbody id=""divminmax" & rsSearchTeams.fields("PROGRAM_ID") & rsSearchTeams.fields("DIVISION_ID") & cntDiv & """"
                    if not rstrFirstTeam then
                        response.write " style=""display:none;"" "
                        rstrFirstTeam = false
                    else
                        rstrFirstTeam = false
                    end if
                    response.write ">" & vbCrlf
                    rstrOpenMinMax = true
                    strCurDiv = rsSearchTeams.fields("DIVISION_ID")
                    cntDiv = cntDiv + 1
                end if
                response.write vbTab & vbTab & "<tr>" & vbCrlf
                response.write vbTab & vbTab & vbTab & "<td width=""20""></td>" & vbCrlf
                response.write vbTab & vbTab & vbTab & "<td width=""20""></td>" & vbCrlf
                response.write vbTab & vbTab & vbTab & "<td><a href=""?view=team&team=" & rsSearchTeams.Fields("ID") & iif(blSearchPlayer, "&year=" & rsSearchTeams.Fields("YEAR"), "") & """>" & rsSearchTeams.fields("TEAM") & "</a></td>" & vbCrlf
                response.write vbTab & vbTab & "</tr>" & vbCrlf
                rsSearchTeams.MoveNext
            wend
        end if
    end if
    if rstrOpenMinMax then
        response.write vbTab & "</tbody>" & vbCrlf
        rstrOpenMinMax = false
    end if
    %>
    </table>
    <%
    
	spThemeBlock1_Close(intSkin)
end sub

function viewTeam(rTeamId)
	searchYear = rosterIDCurrentYear
	counterid=0
	if not isBarren(Request("year")) then
		if IsNumeric(Request("year")) = True then
			searchYear = cLng(Request("year"))
		else
			closeAndGo("stop")
		end if
	end if
	
strSql = "	SELECT T.[ID], T.[TEAM], T.[DESCRIP], T.[LEAGUE_ID], L.[LEAGUE], T.[PROGRAM_ID], P.[PROGRAM], T.[DIVISION_ID], D.[DIVISION], T.[SPONSOR_ID], S.[SPONSOR], T.[COLORS_HOME], T.[COLORS_AWAY], T.[ACTIVE]		" &_
"			  FROM (((" & STRTABLEPREFIX & "TEAM AS T																									" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "LEAGUE AS L																							" &_
"			    ON T.[LEAGUE_ID] = L.[ID])																												" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PROGRAM AS P																						" &_
"			    ON T.[PROGRAM_ID] = P.[ID])																												" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION AS D																						" &_
"			    ON T.[DIVISION_ID] = D.[ID])																											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "SPONSOR AS S																						" & _
"				ON T.[SPONSOR_ID] = S.[ID]" &_
"			 WHERE T.[ID] = " & rTeamId

	set rsTeam = my_conn.execute(strSql)
	
strSql = "	SELECT P.*, PP.[POSITION], R.[RANK], R.[ID] as [ROSTER_ID]		" &_
"			  FROM ((" & STRTABLEPREFIX & "ROSTER AS R								" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER AS P						" &_
"				ON R.[PLAYER_ID] = P.[ID])											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "TEAM AS T						" &_
"			    ON R.[TEAM_ID] = T.[ID])											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER_POSITION AS PP			" &_
"			    ON R.[POSITION_ID] = PP.[ID]										" &_
"			 WHERE PP.[TYPE] = 'player' AND R.[TEAM_ID] = " & rTeamId & " AND R.[YEAR] = " & chkString(searchYear,"sqlstring")		  &_
"			  ORDER BY PP.[SORT],P.[LASTNAME],P.[FIRSTNAME]"

	set rsRoster = my_conn.execute(strSql)
    
strSql = "	SELECT V.[ID], V.[FIRSTNAME], V.[LASTNAME], PP.[POSITION], R.[RANK], R.[ID] as [ROSTER_ID], R.[PERMS]		" &_
"			  FROM ((" & STRTABLEPREFIX & "ROSTER AS R								" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "VOLUNTEER AS V						" &_
"				ON R.[PLAYER_ID] = V.[ID])											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "TEAM AS T						" &_
"			    ON R.[TEAM_ID] = T.[ID])											" &_
"			  LEFT OUTER JOIN " & STRTABLEPREFIX & "PLAYER_POSITION AS PP			" &_
"			    ON R.[POSITION_ID] = PP.[ID]										" &_
"			 WHERE PP.[TYPE] = 'vol' AND R.[TEAM_ID] = " & rTeamId & " AND R.[YEAR] = " & chkString(searchYear,"sqlstring")		  &_
"			  ORDER BY PP.[SORT],V.[LASTNAME],V.[FIRSTNAME]"

	set rsVolunteer = my_conn.execute(strSql)

	

	spThemeTitle = ""
	spThemeBlock1_Open(intSkin)
    
    if rsTeam.Fields("ACTIVE") = 0 and not bAppFull then
        rw("<p>This team is not active</p>")
    else
        if rsTeam.Fields("ACTIVE") = 0 then
            showMsg "note","This team is currently innactive. Regular users cannot see it."
        end if
    	%>
        <table border="0" class="tBorder">
            <tr>
                <td rowspan="2" width=300 align="center"><%
                    strSql = "SELECT [VALUE] FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [NAME] = 'photo' AND [TEAM_ID] = " & rTeamId & " AND [YEAR] = " & searchYear
                    set rsPhoto = my_conn.execute(strSql)
                    
                    if rsPhoto.EOF or rsPhoto.BOF then
                        response.write "<img src=""images/no_photo.gif"" alt=""Team picture for this year not available"" title=""Team picture for this year not available"" />"
                    else
                        response.write "<img src=""" & rsPhoto.fields("VALUE") & """ alt="""" title="""" />"
                    end if
                    
                    set rsPhoto = nothing
                    %></td>
                <td valign=top>
                    <span class="fTitle">
                        <%
                        if bAppWrite then
                            response.write "<a href=""JavaScript:openWindow6('pop_roster.asp?mode=team&cmd=edit&cid=" & rsTeam.Fields("ID") & "&sid=" & searchYear & "')"">" & icon(icnEdit,txtEdit,"","","") & "</a>"
                        end if
                        response.write rsTeam.Fields("TEAM") %>
                    </span><br /><br />
					
					   <table border="0" cellpadding="2" cellspacing="0" class="tBorder">
            <tr class="tTitle">
			<td>Program</td>
			<td>Division</td>
			<td>League</td>
			</tr>
			<tr>
			<td>
                    <a href="JavaScript:openWindow('pop_roster.asp?mode=team&cmd=view&cid=<% =rsTeam.Fields("PROGRAM_ID") %>&sid=2')"><%= rsTeam.Fields("PROGRAM") %></a><br />
			</td>
			<td>
                    <a href="JavaScript:openWindow('pop_roster.asp?mode=team&cmd=view&cid=<% =rsTeam.Fields("DIVISION_ID") %>&sid=3')"><%= rsTeam.Fields("DIVISION") %></a><br />
			</td>
			<td>
                    <% if not isNull(rsTeam.Fields("LEAGUE")) then %><a href="JavaScript:openWindow('pop_roster.asp?mode=team&cmd=view&cid=<% =rsTeam.Fields("LEAGUE_ID") %>&sid=1')"><%= rsTeam.Fields("LEAGUE") %></a><br /><% end if %>
			</td>
			</tr>
			</table>
                    Home Jersey:&nbsp;<%= rsTeam.Fields("COLORS_HOME") %><br />
                    Away Jersey:&nbsp;<%= rsTeam.Fields("COLORS_AWAY") %><br />
                    <% if not isNull(rsTeam.Fields("SPONSOR")) then %><span class="fBold"><br>Sponsored By:</span> <a href="JavaScript:openWindowCT('pop_roster.asp?mode=team&cmd=sponsor&cid=<% =rsTeam.Fields("SPONSOR_ID") %>&sid=4')"><%= rsTeam.Fields("SPONSOR") %></a><% end if %>
                
				<br>
				<br>
				<table class="tBorder">
                        <tr class="tTitle">
                            <td><% if bAppFull then %><a href="JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=cadd&cid=<% =rsTeam.Fields("ID") %>&sid=<% =searchYear %>')"><% =icon(icnPlus,txtAdd,"","","") %></a>&nbsp;<% end if %>Team Contacts:</td>
                        </tr>
                        <tr>
                            <td>
                        <%
                        if rsVolunteer.EOF or rsVolunteer.BOF then
                        
                        else
                            while not rsVolunteer.EOF
                                if bAppFull or bAppWrite then
                                    if bAppFull then
                                        response.write "<a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=delete&cid=" & rsVolunteer.Fields("ROSTER_ID") & "')"">" & icon(icnDelete,txtDel,"","","") & "</a>"
                                    end if
                                    if bAppWrite then
                                        response.write "<a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=cedit&cid=" & rsVolunteer.Fields("ROSTER_ID") & "')"">" & icon(icnEdit,txtEdit,"","","") & "</a>"
                                    end if
                                    response.write "&nbsp;"
                                end if
                                if hasAccess(1) or rsVolunteer.Fields("PERMS") = 1 then
                                    response.write rsVolunteer.Fields("POSITION") & " - <a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=cview&cid=" & rsVolunteer.Fields("ID") & "&sid=" & searchYear & "')"">" & rsVolunteer.Fields("FIRSTNAME") & " " & rsVolunteer.Fields("LASTNAME") & "</a><br />"
                                else
                                    response.write rsVolunteer.Fields("POSITION") & " - " & rsVolunteer.Fields("FIRSTNAME") & " " & rsVolunteer.Fields("LASTNAME") & "<br />"
                                end if
                                rsVolunteer.MoveNext
                            wend
                        end if
                        %>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2"><% = rsTeam.Fields("DESCRIP") %></td>
            </tr>
            <tr>
                <td colspan="3" align="center">
                    <% strSql = "SELECT [ID], [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'year'"
                    response.write DoRosterDropDown(strSql,"M_VALUE","ID",searchYear,"position","onChange=""window.location='?view=team&team=" & rTeamId & "&year='+this.options[this.options.selectedIndex].value;""","") %>
                </td>
            </tr>
        </table>
        <br /><br />
        <table border="0" cellpadding="2" cellspacing="0" class="tBorder tPlain grid">
            <tr class="tTitle">
                <% if bAppFull then %>
                <td class="fBold" width="50"><a href="JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=add&cid=<% =rsTeam.Fields("ID") %>&sid=<% =searchYear %>')"><% =icon(icnPlus,txtAdd,"","","") %></a></td>
                <% end if %>
                <td class="fbold" width=2>&nbsp; </td>
				<td class="fBold">Name</td>
                <td class="fBold">Position</td>
                <td class="fBold" align="right">Jersey #</td>
           
			   <%
                for i=0 to ubound(rstrArrRosterExt)
                    response.write "<td class='fBold'>"
                    response.write eval("Player" & rstrArrRosterExt(i))
                    response.write "</td>"
                next
                %>
            </tr>
            <tbody>
            <%
            if rsRoster.EOF or rsRoster.BOF then
                response.write "<tr><td colspan=""15"">There aren't any members on this team yet</td></tr>"
            else
                while not rsRoster.EOF
				    response.write "<tr>"
                    if bAppFull or bAppWrite then
                        response.write "<td>"
                        if bAppFull then
                            response.write "<a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=delete&cid=" & rsRoster.Fields("ROSTER_ID") & "')"">" & icon(icnDelete,txtDel,"","","") & "</a>"
                        end if
                        if bAppWrite then
                            response.write "<a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=edit&cid=" & rsRoster.Fields("ROSTER_ID") & "')"">" & icon(icnEdit,txtEdit,"","","") & "</a>"
                        end if
                        response.write "</td>"
                    end if
					response.write "<td align='right'>"
					counterid = counterid + 1
					response.write counterid 
					response.write "</td>"
                    if bAppRead then
                        response.write "<td><a href=""JavaScript:openWindow('pop_roster.asp?mode=roster&cmd=view&cid=" & rsRoster.Fields("ID") & "')"">" & rsRoster.Fields("FIRSTNAME") & " " & rsRoster.Fields("LASTNAME") & "</a></td>"
                    else
                        response.write "<td>" & rsRoster.Fields("FIRSTNAME") & " " & rsRoster.Fields("LASTNAME") & "</td>"
                    end if
                    response.write "<td>" & iif(isBarren(rsRoster.Fields("POSITION")),rosterEmptyFieldCharacter,rsRoster.Fields("POSITION")) & "</td>"
                    response.write "<td align='right'>" & iif(isBarren(rsRoster.Fields("RANK")),rosterEmptyFieldCharacter,rsRoster.Fields("RANK")) & "</td>"
                    for i=0 to ubound(rstrArrRosterExt)
                        response.write "<td>" & iif(isBarren(rsRoster.Fields(rstrArrRosterExt(i))),rosterEmptyFieldCharacter,rsRoster.Fields(rstrArrRosterExt(i))) & "</td>"
                    next

                    response.write "</tr>"
                    rsRoster.MoveNext
                wend
            end if
            %>
            </tbody>
        </table>
    	<%
    end if
	spThemeBlock1_Close(intSkin)

end function

sub pop_roster(strCmd)
    Select case strCmd
        case "view"
            strSql = "SELECT [FIRSTNAME] + ' ' + [LASTNAME] AS [XNAME], [BIRTHDATE], [PHONE], [CELL], [EMAIL], [PIC] "
            for i=1 to 10
                if len(eval("PlayerT" & i)) > 0 then
                    strSql = strSql & ", [T" & i & "]"
                end if
            next
            strSql = strSql & " FROM " & STRTABLEPREFIX & "PLAYER WHERE ID = " & c_id

            set rs = my_conn.execute(strSql)

            if rs.EOF or rs.BOF then
            	showMsg "note","Player not found"
            else
            	%>
            	<table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr class="tSubTitle">
                        <td colspan="2" align="center"><% = rs.Fields("XNAME") %></td>
                    </tr>
                    <tr>
                        <td>
                            <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                                <tr>
                                    <td align="right">Birth Date:</td>
                                    <td align="left"><% = rs.Fields("BIRTHDATE") %></td>
                                </tr>
                                <tr>
                                    <td align="right">Phone:</td>
                                    <td align="left"><% if isBarren(rs.Fields("PHONE")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("PHONE")) end if %></td>
                                </tr>
                                <tr>
                                    <td align="right">Cell:</td>
                                    <td align="left"><% if isBarren(rs.Fields("CELL")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("CELL")) end if %></td>
                                </tr>
                                <tr>
                                    <td align="right">Email:</td>
                                    <td align="left"><% if isBarren(rs.Fields("EMAIL")) then response.write rosterEmptyFieldCharacter else response.write plain2HTMLtxt(rs.Fields("EMAIL")) end if %></td>
                                </tr>
                        			<%
                        			for i=1 to 10
                                        if len(eval("PlayerT" & i)) > 0 then
                                            response.write "<tr>"
                                            response.write "<td align=""right"">" & eval("PlayerT" & i) & ":</td>"
                                            response.write "<td align=""left"">"
                                            if isBarren(rs.Fields("T"&i)) then
                                                response.write rosterEmptyFieldCharacter
                                            else
                                                response.write rs.Fields("T"&i)
                                            end if
                                            response.write "</td>"
                                            response.write "</tr>"
                                        end if
                                    next
                        			%>
                            </table>
                        </td>
						<tr>
						<td align="center">
						<img src="<% = iif(isBarren(rs.Fields("PIC")), "images/no_photo.gif", rs.Fields("PIC")) %>" /></td>
                    </tr>
            	</table>
            	<%
            end if
            
            set rs = nothing
            
        case "cview"
            set rs = nothing
            strSql = "SELECT [FIRSTNAME] + ' ' + [LASTNAME] AS [NAME], [PHONE], [CELL], [EMAIL], [PIC] "
            for i=1 to 10
                if len(eval("VolunteerT" & i)) > 0 then
                    strSql = strSql & ", [T" & i & "]"
                end if
            next
            strSql = strSql & " FROM " & STRTABLEPREFIX & "VOLUNTEER WHERE ID = " & c_id

            set rs = my_conn.execute(strSql)

            if rs.EOF or rs.BOF then
                showMsg "note","Volunteer not found"
            else
                
                %>
                <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr class="tSubTitle">
                        <td colspan="2" align="center"><% = rs.Fields("NAME") %></td>
                    </tr>
                    <tr>
                        <td>
                            <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                                <tr>
                                    <td align="right">Phone:</td>
                                    <td align="left"><% if isBarren(rs.Fields("PHONE")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("PHONE")) end if %></td>
                                </tr>
                                <tr>
                                    <td align="right">Cell:</td>
                                    <td align="left"><% if isBarren(rs.Fields("CELL")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("CELL")) end if %></td>
                                </tr>
                                <tr>
                                    <td align="right">Email:</td>
                                    <td align="left"><% if isBarren(rs.Fields("EMAIL")) then response.write rosterEmptyFieldCharacter else response.write plain2HTMLtxt(rs.Fields("EMAIL")) end if %></td>
                                </tr>
                                <%
                                for i=1 to 10
                                    if len(eval("VolunteerT" & i)) > 0 then
                                        response.write "<tr>"
                                        response.write "<td align=""right"">" & eval("VolunteerT" & i) & ":</td>"
                                        response.write "<td align=""left"">" & rs.Fields("T"&i) & "</td>"
                                        response.write "</tr>"
                                    end if
                                next
                                %>
                            </table>
                        </td>
						</tr>
						<tr>
						<td align="center"><img src="<% = iif(isBarren(rs.Fields("PIC")), "images/no_photo.gif", rs.Fields("PIC")) %>" /></td>
                    </tr>
                </table>
                <%
            end if
            
            set rs = nothing
            
            
        case "add"
            errmsg = ""
            if request.form("confEdit") = "confirmed" then
                if request.form("player") = 0 then
                    errmsg = errmsg & "<li>You must choose a player</li>"
                end if
                if request.form("position") = 0 then
                    errmsg = errmsg & "<li>You must choose a position</li>"
                end if
                if len(request.form("rank")) > 0 and not isNumeric(request.form("rank")) then
					errmsg = errmsg & "<li>Jersey # must be a number</li>"
				end if
                
                if len(errmsg) = 0 then
                    strSql = "INSERT INTO " & STRTABLEPREFIX & "ROSTER ([TEAM_ID],[PLAYER_ID],[POSITION_ID],[RANK],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES (" & c_id & "," & chkString(request.form("player"),"sqlstring") & "," & chkString(request.form("position"),"sqlstring") & "," & iif(isBarren(request.form("rank")), "NULL", ChkString(request.form("rank"),"sqlstring"))  & "," & s_id & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
                    executeThis(strSql)
                    showMsg "success",txtAdded
                    response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
                else
                    showMsg "validation", "<ul>" & errmsg & "</ul>"
                end if
            end if
            
            if not request.querystring("sAge") = "0" then
                strSql = "SELECT D.[STARTAGE], D.[ENDAGE] FROM " & STRTABLEPREFIX & "TEAM T LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION D ON T.[DIVISION_ID] = D.[ID] WHERE T.[ID] = " & c_id
                set xRs = my_conn.execute(strSql)
                            
                STARTAGE = Date()
                ENDAGE = cDate("01/01/" & datepart("yyyy",Date()))
                STARTAGE = dateadd("yyyy",-xRs.Fields("STARTAGE"),STARTAGE)
                ENDAGE = dateadd("yyyy",-XRS.FIELDS("ENDAGE"),ENDAGE)
                'rw("BETWEEN #" & ENDAGE & "# AND #" & STARTAGE & "#")
                
                showMsg "note","Only players between the ages of " & xRs.Fields("STARTAGE") & " and " & xRs.Fields("ENDAGE") & " are listed.<br /><a href=""javascript:showMorePlayers();"">Click here</a> to pick from a list of all players."
            end if
            %>
                <script type="text/javascript">
                function openWindowR(url) {
                    LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
                    TopPosition = (screen.height) ? (screen.height-450)/2 : 0;
                    popupWin = window.open(url,'rosterPage','width=400,height=450,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes');
                }
                function showMorePlayers() {
                    var messWith = document.getElementById('players');
                    messWith.innerHTML = '';
                    messWith.innerHTML = '<input type="text" name="playerShow" id="playerShow" readonly="readonly" /><input type="hidden" name="PLAYER" id="PLAYER" />';
                    LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
                    TopPosition = (screen.height) ? (screen.height-450)/2 : 0;
                    popupWin = window.open('pop_roster.asp?mode=player&cmd=list','rosterPage','width=400,height=450,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes');
                }
                </script>
            <form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
                <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr>
                        <td align="right">Player:</td>
                        <td id="players"><%
						strSql = "SELECT [LASTNAME] + ', ' + [FIRSTNAME] AS PLAYER, [ID] FROM " & STRTABLEPREFIX & "PLAYER "
                        if not request.querystring("sAge") = "0" then
                            strSql = strSql & " WHERE [BIRTHDATE] BETWEEN #" & ENDAGE & "# AND #" & STARTAGE & "# "
                        end if
                        strSql = strSql & " ORDER BY [LASTNAME],[FIRSTNAME],[ID]"
						response.write DoRosterDropDown(strSql,"PLAYER","ID",request.form("player"),"player","","None") %>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">Position:</td>
                        <td align="left"><%
                        strSql = "SELECT [POSITION], [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [TYPE] = 'player' ORDER BY [SORT]"
                        response.write DoRosterDropDown(strSql,"POSITION","ID",request.form("position"),"position","","None") %>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">Jersey #:</td>
                        <td align="left"><input type="text" name="rank" id="rank" value="<% =request.form("rank") %>" /></td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center">
                            <input type="hidden" name="confEdit" value="confirmed" />
                            <input type="submit" value="Submit" class="button" />&nbsp;<input type="button" value="Cancel" class="button" onClick="window.close()" />
                        </td>
                    </tr>
                </table>
            </form>
            <%
            
        case "cadd"
            errmsg = ""
            if request.form("confEdit") = "confirmed" then
                if isBarren(request.form("perms")) then
                    rstrPerms = 0
                elseif request.form("perms") = "1" then
                    rstrPerms = 1
                else
                    rstrPerms = 0
                end if
                
                if request.form("player") = 0 then
                    errmsg = errmsg & "<li>You must choose a volunteer</li>"
                end if
                if request.form("position") = 0 then
                    errmsg = errmsg & "<li>You must choose a position</li>"
                end if
                
                if len(errmsg) = 0 then
                    strSql = "INSERT INTO " & STRTABLEPREFIX & "ROSTER ([TEAM_ID],[PLAYER_ID],[POSITION_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE],[PERMS]) VALUES (" & c_id & "," & chkString(request.form("player"),"sqlstring") & "," & chkString(request.form("position"),"sqlstring") & "," & s_id & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "','" & rstrPerms & "')"
                    executeThis(strSql)
                    showMsg "success",txtAdded
                    response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
                else
                    showMsg "validation", "<ul>" & errmsg & "</ul>"
                end if
            end if
            
            %>
            <form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
                <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr>
                        <td align="right">Volunteer:</td>
                        <td><%
						strSql = "SELECT [LASTNAME] + ', ' + [FIRSTNAME] AS VOLUNTEER, [ID] FROM " & STRTABLEPREFIX & "VOLUNTEER "
                        strSql = strSql & " ORDER BY [LASTNAME],[FIRSTNAME],[ID]"
						response.write DoRosterDropDown(strSql,"VOLUNTEER","ID",request.form("player"),"player","","None") %>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">Position:</td>
                        <td align="left"><%
                        strSql = "SELECT [POSITION], [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [TYPE] = 'vol' ORDER BY [SORT]"
                        response.write DoRosterDropDown(strSql,"POSITION","ID",request.form("position"),"position","","None") %>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">Make info public?</td>
                        <td align="left"><input type="checkbox" name="perms" id="perms" value="1" /></td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center">
                            <input type="hidden" name="confEdit" value="confirmed" />
                            <input type="submit" value="Submit" class="button" />&nbsp;<input type="button" value="Cancel" class="button" onClick="window.close()" />
                        </td>
                    </tr>
                </table>
            </form>
            <%
        
        case "edit"
            errmsg = ""
            if request.form("confEdit") = "confirmed" then
                if request.form("position") = 0 then
                    errmsg = errmsg & "<li>You must choose a position</li>"
                end if
                if len(request.form("rank")) > 0 and not isNumeric(request.form("rank")) then
					errmsg = errmsg & "<li>Jersey # must be a number</li>"
				end if
                
                if len(errmsg) = 0 then
                    strSql = "UPDATE " & STRTABLEPREFIX & "ROSTER SET [POSITION_ID] = " & ChkString(request.form("position"),"sqlstring") & ", [RANK] = " & iif(isBarren(request.form("rank")), "NULL", ChkString(request.form("rank"),"sqlstring"))  & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & c_id
                    executeThis(strSql)
                    showMsg "success","Edited"
                    response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
                else
                    showMsg "validation","<ul>" & errmsg & "</ul>"
                end if
            end if
            
            strSql = "SELECT * FROM " & STRTABLEPREFIX & "ROSTER WHERE [ID] = " & c_id
            set rs = my_conn.execute(strSql)
            
            if rs.EOF or rs.BOF then
                showMsg "note","Record not found"
            else
                %>
                <form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
                    <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                        <tr>
                            <td align="right">Position:</td>
                            <td align="left"><%
						strSql = "SELECT [POSITION], [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [TYPE] = 'player' ORDER BY [SORT]"
						response.write DoRosterDropDown(strSql,"POSITION","ID",rs.Fields("POSITION_ID"),"position","","") %></td>
                        </tr>
                        <tr>
                            <td align="right">Jersey #:</td>
                            <td align="left"><input type="text" name="rank" id="rank" value="<% =rs.Fields("RANK") %>" /></td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <input type="hidden" name="confEdit" value="confirmed" />
                                <input type="submit" value="Submit" class="button" />&nbsp;<input type="button" value="Cancel" class="button" onClick="window.close()" />
                            </td>
                        </tr>
                    </table>
                </form>
                <%
            end if
            
        case "cedit"
            errmsg = ""
            if request.form("confEdit") = "confirmed" then
                if isBarren(request.form("perms")) then
                    rstrPerms = 0
                elseif request.form("perms") = "1" then
                    rstrPerms = 1
                else
                    rstrPerms = 0
                end if
                
                if request.form("position") = 0 then
                    errmsg = errmsg & "<li>You must choose a position</li>"
                end if
                
                if len(errmsg) = 0 then
                    strSql = "UPDATE " & STRTABLEPREFIX & "ROSTER SET [POSITION_ID] = " & ChkString(request.form("position"),"sqlstring") & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "', [PERMS] = '" & rstrPerms & "' WHERE [ID] = " & c_id
                    executeThis(strSql)
                    showMsg "success","Edited"
                    response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
                else
                    showMsg "validation","<ul>" & errmsg & "</ul>"
                end if
            end if
            
            strSql = "SELECT * FROM " & STRTABLEPREFIX & "ROSTER WHERE [ID] = " & c_id
            set rs = my_conn.execute(strSql)
            
            if rs.EOF or rs.BOF then
                showMsg "note","Record not found"
            else
                %>
                <form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
                    <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                        <tr>
                            <td align="right">Position:</td>
                            <td align="left"><%
						strSql = "SELECT [POSITION], [ID] FROM " & STRTABLEPREFIX & "PLAYER_POSITION WHERE [TYPE]='vol' ORDER BY [SORT]"
						response.write DoRosterDropDown(strSql,"POSITION","ID",rs.Fields("POSITION_ID"),"position","","") %></td>
                        </tr>
                        <tr>
                            <td align="right">Make info public?</td>
                            <td align="left"><input type="checkbox" name="perms" id="perms" value="1" <%=chkCheckbox(rs.Fields("PERMS"),1,true) %> /></td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <input type="hidden" name="confEdit" value="confirmed" />
                                <input type="submit" value="Submit" class="button" />&nbsp;<input type="button" value="Cancel" class="button" onClick="window.close()" />
                            </td>
                        </tr>
                    </table>
                </form>
                <%
            end if
        
        case "delete"
            if request.form("confDelete") = "confirmed" then
                strSql = "DELETE FROM " & STRTABLEPREFIX & "ROSTER WHERE [ID] = " & c_id
                executeThis(strSql)
                showMsg "success","Deleted"
                response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
            else
                %>
                <form name="delete" method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
                    <p>Are you sure you want to delete that?</p>
                    <p><input type="submit" value="<%=txtYes%>" class="button" />&nbsp;<input type="button" value="<%=txtNo%>" class="button" onClick="window.close()" /></p>
                    <input type="hidden" name="confDelete" value="confirmed" />
                </form>
                <%
            end if
            
        case else
            showMsg "warn","You can't do that! >.<"
        
    End Select

end sub

sub pop_team()
    if request.form("confEdit") = "confirmed" then
        errmsg = ""
        if not len(request.form("team")) > 0 then
            errmsg = errmsg & "<li>Name must not be empty</li>"
        end if
        if request.form("program") = 0 then
            errmsg = errmsg & "<li>You must choose a program</li>"
        end if
        if request.form("division") = 0 then
            errmsg = errmsg & "<li>You must choose a division</li>"
        end if
        
        if not isNumeric(request.form("league")) or not isNumeric(request.form("program")) or not isNumeric(request.form("division")) or not isNumeric(request.form("sponsor")) or not isNumeric(request.form("active")) then
            showMsg "warn","You can't do that! >.<"
            closeAndGo("stop")
        end if

        if len(errmsg) = 0 then
            strSql = "UPDATE " & STRTABLEPREFIX & "TEAM SET [TEAM] = '" & ChkString(request.form("team"),"sqlstring") & "', [DESCRIP] = '" & chkString(request.form("descrip"),"message") & "', [LEAGUE_ID] = " & chkString(request.form("league"), "sqlstring") & ", [PROGRAM_ID] = " & chkString(request.form("program"), "sqlstring") & ", [DIVISION_ID] = " & chkString(request.form("division"), "sqlstring") & ", [SPONSOR_ID] = " & chkString(request.form("sponsor"), "sqlstring") & ", [COLORS_HOME] = '" & ChkString(request.form("colorshome"),"sqlstring") & "', [COLORS_AWAY] = '" & ChkString(request.form("colorsaway"),"sqlstring") & "', [ACTIVE] = " & chkString(request.form("active"), "sqlstring") & ", [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [ID] = " & c_id
            executeThis(strSql)
            showMsg "success","Team Edited"
            response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
        else
            showMsg "validation","<ul>" & errmsg & "</ul>"
        end if
    
    end if
    
    if request.form("confEditPhoto") = "confirmed" then
        errmsg = ""
        pDelete = iif(isBarren(request.form("pDelete")), false, true)
        if not len(request.form("nPhoto")) > 0 and pDelete = false then
            errmsg = errmsg & "<li>Photo location must not be empty</li>"
        end if
        
        if len(errmsg) = 0 then
            if pDelete then
                strSql = "DELETE FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [TEAM_ID] = " & c_id & " AND [YEAR] = " & s_id
                executeThis(strSql)
                showMsg "success","Image Deleted"
            elseif request.form("pMode") = "new" then
                strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','" & ChkString(request.form("nPhoto"),"sqlstring") & "'," & c_id & ", " & s_id & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
                executeThis(strSql)
                showMsg "success","Image Added"
            elseif request.form("pMode") = "edit" then
                strSql = "UPDATE " & STRTABLEPREFIX & "TEAM_YEARLIES SET [VALUE] = '" & ChkString(request.form("nPhoto"),"sqlstring") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [TEAM_ID] = " & c_id & " AND [YEAR] = " & s_id
                executeThis(strSql)
                showMsg "success","Image Edited"
            end if
            response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
        else
            showMsg "validation","<ul>" & errmsg & "</ul>"
        end if
    end if
    
    if request.querystring("upPhoto") = "upped" then
        returnErr = chkString(request.querystring("err"), "display")
        if returnErr = "true" then
            showMsg "validation","<b>There was a problem uploading your image</b><br /><ul>" & chkString(Session.Contents("rosterErr"), "message") & "</ul>"
            Session.Contents("rosterErr") = ""
        else
            strSql = "SELECT [VALUE] FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [NAME] = 'photo' AND [TEAM_ID] = " & c_id & " AND [YEAR] = " & s_id
            set rsCheck = my_conn.Execute(strSql)
            
            if rsCheck.EOF or rsCheck.BOF then
                strSql = "INSERT INTO " & STRTABLEPREFIX & "TEAM_YEARLIES ([NAME],[VALUE],[TEAM_ID],[YEAR],[AUSER],[ADATE],[EUSER],[EDATE]) VALUES ('photo','" & ChkString(request.querystring("photourl"),"sqlstring") & "'," & c_id & ", " & s_id & "," & strUserMemberID & ",'" & now() & "'," & strUserMemberID & ",'" & now() & "')"
                executeThis(strSql)
            else
                strSql = "UPDATE " & STRTABLEPREFIX & "TEAM_YEARLIES SET [VALUE] = '" & ChkString(request.querystring("photourl"),"sqlstring") & "', [EUSER] = " & strUserMemberID & ", [EDATE] = '" & now() & "' WHERE [TEAM_ID] = " & c_id & " AND [YEAR] = " & s_id
                executeThis(strSql)
            end if
            set rsCheck = nothing
            showMsg "success","Image Uploaded"
            response.write "<script type=""text/javascript"">opener.document.location.reload();</script>"
        end if
    end if
    
    strSql = "SELECT * FROM " & STRTABLEPREFIX & "TEAM WHERE [ID] = " & c_id

    set rsTeam = my_conn.execute(strSql)
    
    if rsTeam.EOF or rsTeam.BOF then
        showMsg "note","Team not found"
    else
        %>
        <form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>">
            <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                <tr>
                    <td align="right">Name:</td>
                    <td align="left"><input type="text" name="team" id="team" value="<% =rsTeam.Fields("TEAM") %>" /></td>
                </tr>
                <tr>
                    <td align="right">Description:</td>
                    <td align="left"><textarea name="descrip" id="descrip" cols="70" rows="15"><% =rsTeam.Fields("DESCRIP") %></textarea></td>
                </tr>
                <tr>
                    <td align="right">League:</td>
                    <td align="left"><% = DoDropDown(STRTABLEPREFIX & "LEAGUE","LEAGUE","ID",rsTeam.Fields("LEAGUE_ID"),"league","None","","ID") %></td>
                </tr>
                <tr>
                    <td align="right">Program:</td>
                    <td align="left"><% = DoDropDown(STRTABLEPREFIX & "PROGRAM","PROGRAM","ID",rsTeam.Fields("PROGRAM_ID"),"program","None","","ID") %></td>
                </tr>
                <tr>
                    <td align="right">Division:</td>
                    <td align="left"><% = DoDropDown(STRTABLEPREFIX & "DIVISION","DIVISION","ID",rsTeam.Fields("DIVISION_ID"),"division","None","","ID") %></td>
                </tr>
                <tr>
                    <td align="right">Sponsor:</td>
                    <td align="left"><% = DoDropDown(STRTABLEPREFIX & "SPONSOR","SPONSOR","ID",rsTeam.Fields("SPONSOR_ID"),"sponsor","None","","ID") %></td>
                </tr>
                <tr>
                    <td align="right">Colors Home:</td>
                    <td align="left"><input type="text" name="colorshome" id="colorshome" value="<% =rsTeam.Fields("COLORS_HOME") %>" /></td>
                </tr>
                <tr>
                    <td align="right">Colors Away:</td>
                    <td align="left"><input type="text" name="colorsaway" id="colorsaway" value="<% =rsTeam.Fields("COLORS_AWAY") %>" /></td>
                </tr>
                <tr>
                    <td align="right">Active:</td>
                    <td align="left">
                        <select name="active" id="active">
                            <option value="1" <%= chkSelect(rsTeam.Fields("ACTIVE"),1) %>>Yes</option>
                            <option value="0" <%= chkSelect(rsTeam.Fields("ACTIVE"),0) %>>No</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" align="center">
                        <input type="hidden" name="confEdit" value="confirmed" />
                        <input type="submit" value="Submit" class="button" />&nbsp;<input type="button" value="Cancel" class="button" onClick="window.close()" />
                    </td>
                </tr>
            </table>
        </form>
        <hr />
        <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
            <tr>
                <td colspan="2" align="center" class="fSubTitle">Current team image:</td>
            </tr>
            <tr>
                <td colspan="2" align="center"><%
                rstrPmode = ""
                
                strSql = "SELECT [VALUE] FROM " & STRTABLEPREFIX & "TEAM_YEARLIES WHERE [NAME] = 'photo' AND [TEAM_ID] = " & c_id & " AND [YEAR] = " & s_id
                set rsPhoto = my_conn.execute(strSql)
                
                if rsPhoto.EOF or rsPhoto.BOF then
                    response.write "<img src=""images/no_photo.gif"" alt=""Team picture for this year not available"" title=""Team picture for this year not available"" />"
                    rstrPmode = "<input type=""hidden"" name=""pMode"" id=""pMode"" value=""new"" />"
                else
                    curImage = rsPhoto.fields("VALUE")
                    response.write "<img src=""" & curImage & """ alt="""" title="""" /><br />"
                    response.write "<form name=""delete"" id=""delete"" method=""post"" action=""pop_roster.asp?mode=" & strMode & "&cmd=" & strCmd & "&cid=" & c_id & "&sid=" & s_id & """><input type=""hidden"" name=""pDelete"" value=""true"" /><a href=""#"" onClick=""$('delete').submit()""><img src=""images/icons/icon_trashcan.gif"" alt=""Delete Image"" title=""Delete Image"" style=""border: 0px;"" /> Show Default</a><input type=""hidden"" name=""confEditPhoto"" value=""confirmed"" /></form>"
                    rstrPmode = "<input type=""hidden"" name=""pMode"" id=""pMode"" value=""edit"" />"
                end if
                
                set rsPhoto = nothing
                %></td>
            </tr>
            <tr>
                <td colspan="2">&nbsp;</td>
            </tr>
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
			<form name="formUpload" action="rosterUpload.asp" method="post" enctype="multipart/form-data">
				<tr>
					<td align="right">Upload Image: </td>
					<td align="left">
						<input type="file" name="file1" id="file1" />
						<input type="hidden" name="folder" value="team" />
						<input type="hidden" name="cid" value="<%=c_id%>" />
						<input type="hidden" name="sid" value="<%=s_id%>" />
					</td>
				</tr>
			</form>
			<form name="formURL" action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring %>" method="post">
				<tr>
					<td align="right">URL of Image: </td>
					<td align="left">
						<input type="text" name="nPhoto" id="nPhoto" value="<%=curImage%>" />
						<input type="hidden" name="confEditPhoto" value="confirmed" />
						<%= rstrPmode %>
					</td>
				</tr>
			</form>
            <tr>
                <td colspan="2">&nbsp;</td>
            </tr>
            <tr>
                <td colspan="2" align="center" id="submitButtons"><input type="button" value="Submit" class="button" onClick="submitForm()" /> <input type="button" value="Cancel" class="button" onClick="window.close()" /><img src="images/icons/icon_ajax_loading.gif" alt="Loading..." title="Loading..." style="display:none;"/></td>
            </tr>
        </table>
		<script type="text/javascript">
		function submitForm() {
			var urlBox = document.getElementById('nPhoto');
            var uplBox = document.getElementById('file1');
            
            if (urlBox.value.length == 0 && uplBox.value.length == 0) {
                alert('You must upload an image or provide a URL to an image.');
                return false;
            }
			
			if (uplBox.value.length > 0) {
                //submit formDownloads
                document.getElementById('submitButtons').innerHTML = 'Uploading&nbsp;<img src="images/icons/icon_ajax_loading.gif" alt="Loading..." title="Loading..." />';
				document.formUpload.submit();
			}
			else {
				//submit formURL
				document.formURL.submit();
			}
		}
		</script>
        <%
    end if

end sub

sub pop_team_xtras(strCmd)
    select case strCmd
        case "sponsor"
            strSql = "SELECT * FROM " & STRTABLEPREFIX & "SPONSOR WHERE [ID] = " & c_id
            
            set rs = my_conn.execute(strSql)
            
            if rs.EOF or rs.BOF then
                
            else
                %>
                <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr class="tSubTitle">
                        <td colspan="2" align="center"><% = rs.Fields("SPONSOR") %></td>
                    </tr>
                    <tr>
                        <td colspan="2" align="center"><img src="<% = iif(isBarren(rs.Fields("PIC")), "images/no_photo.gif", rs.Fields("PIC")) %>" /></td>
                    </tr>
                    <tr>
                        <td align="right">Website:</td>
                        <td align="left"><% if isBarren(rs.Fields("URL")) then response.write rosterEmptyFieldCharacter else response.write plain2HTMLtxt(rs.Fields("URL")) end if %></td>
                    </tr>
                    <tr>
                        <td align="right">Email:</td>
                        <td align="left"><% if isBarren(rs.Fields("EMAIL")) then response.write rosterEmptyFieldCharacter else response.write plain2HTMLtxt(rs.Fields("EMAIL")) end if %></td>
                    </tr>
                    <tr>
                        <td valign="top" align="right">Address:</td>
                        <td valign="top" align="left"><pre><% if isBarren(rs.Fields("ADDRESS")) then response.write rosterEmptyFieldCharacter else response.write rs.Fields("ADDRESS") end if %></pre></td>
                    </tr>
                    <tr>
                        <td align="right">Phone:</td>
                        <td align="left"><% if isBarren(rs.Fields("PHONE")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("PHONE")) end if %></td>
                    </tr>
                    <tr>
                        <td align="right">Cell:</td>
                        <td align="left"><% if isBarren(rs.Fields("CELL")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("CELL")) end if %></td>
                    </tr>
                    <tr>
                        <td align="right">Fax:</td>
                        <td align="left"><% if isBarren(rs.Fields("FAX")) then response.write rosterEmptyFieldCharacter else response.write FormatPhoneNumber(rs.Fields("FAX")) end if %></td>
                    </tr>
                    <%
                    for i=1 to 10
                        if len(eval("SponsorT" & i)) > 0 then
                            response.write "<tr>"
                            response.write "<td align=""right"">" & eval("SponsorT" & i) & "</td>"
                            response.write "<td align=""left"">" & eval("rs.Fields(""T" & i & """)") & "</td>"
                            response.write "</tr>"
                        end if
                    next
                    %>
                    <tr>
                        <td colspan="2"><% = rs.Fields("DESCRIP") %></td>
                    </tr>
                <table>
                <%
            end if
        
        case else
            select case s_id
                case 1
                    reDim arrFields(1,3)
                    arrFields(0,0) = "League"
                    arrFields(1,0) = "LEAGUE"
                    arrFields(0,1) = "Name: "
                    arrFields(1,1) = "LEAGUE"
                    arrFields(0,2) = "Website: "
                    arrFields(1,2) = "WEBSITE"
                    arrFields(0,3) = "Description: "
                    arrFields(1,3) = "DESCRIP"
                case 2
                    reDim arrFields(1,2)
                    arrFields(0,0) = "Program"
                    arrFields(1,0) = "PROGRAM"
                    arrFields(0,1) = "Name: "
                    arrFields(1,1) = "PROGRAM"
                    arrFields(0,2) = "Description: "
                    arrFields(1,2) = "DESCRIP"
                case 3
                    reDim arrFields(1,4)
                    arrFields(0,0) = "Division"
                    arrFields(1,0) = "DIVISION"
                    arrFields(0,1) = "Name: "
                    arrFields(1,1) = "DIVISION"
                    arrFields(0,2) = "Description: "
                    arrFields(1,2) = "DESCRIP"
                    arrFields(0,3) = "Start Age: "
                    arrFields(1,3) = "STARTAGE"
                    arrFields(0,4) = "End Age: "
                    arrFields(1,4) = "ENDAGE"
                case 4
                    reDim arrFields(1,1)
                    arrFields(0,0) = "Sponsor"
                    arrFields(1,0) = "SPONSOR"
                    arrFields(0,1) = "Name: "
                    arrFields(1,1) = "SPONSOR"
                    for i=1 to 10
                        x=i+1
                        if len(eval("SponsorT" & i)) > 0 then
                            reDim PRESERVE arrFields(1,x)
                            execute("arrFields(0," & x & ") = SponsorT"&i&"&"": """)
                            execute("arrFields(1," & x & ") = ""T"&i&"""")
                        end if
                    next
                case else
                    showMsg "warn","You can't do that! >.<"
                    closeAndGo("stop")
            end select

            strSql = "SELECT "
            for i=1 to Ubound(arrFields,2)
                            strSql = strSql & "[" & arrFields(1,i) & "]"
                            if i<>uBound(arrFields,2) then strSql = strSql & ", "
                        next
            strSql = strSql & " FROM " & STRTABLEPREFIX & arrFields(1,0) & " WHERE [ID] = " & c_id

            set rs = my_conn.execute(strSql)

            if rs.EOF or rs.BOF then
                spThemeTitle = ""
                spThemeBlock1_Open(intSkin)
                response.write "<p align=""center"">" & arrFields(0,0) & " Not Found</p>"
                spThemeBlock1_Close(intSkin)
            else
                %>
                <table border="0" cellpadding="3" cellspacing="0" width="100%" align="center">
                    <tr class="tSubTitle">
                        <td colspan="2" align="center">
                            <% if inStr(arrFields(1,1), "]") then
                                response.write rs.Fields(left(arrFields(1,1),inStr(arrFields(1,1), "]")-1))
                            else
                                response.write rs.Fields(arrFields(1,1))
                            end if %>
                        </td>
                    </tr>
                        <%
                        for i=2 to Ubound(arrFields,2)
                            response.write "<tr>"
                            response.write "<td align=""right"">" & arrFields(0,i) & "</td>"
                            response.write "<td align=""left"">"
                            if inStr(arrFields(1,i), "]") then
                                if isBarren(rs.Fields(left(arrFields(1,i),inStr(arrFields(1,i), "]")-1))) then
                                    response.write rosterEmptyFieldCharacter
                                else
                                    response.write plain2HTMLtxt(rs.Fields(left(arrFields(1,i),inStr(arrFields(1,i), "]")-1))) & "&nbsp;"
                                end if
                            else
                                if isBarren(rs.Fields(arrFields(1,i))) then
                                    response.write rosterEmptyFieldCharacter
                                else
                                    response.write plain2HTMLtxt(rs.Fields(arrFields(1,i))) & "&nbsp;"
                                end if
                            end if
                            response.write "</td>"
                            response.write "</tr>"
                            
                        next
                        %>
                </table>
                <%
            end if
    
    end select
end sub

sub listPlayers()
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
    <script type="text/javascript">
    function insertPlayer(xName,xID) {
        opener.document.getElementById('playerShow').value = xName;
        opener.document.getElementById('PLAYER').value = xID;
        window.close()
    }
    </script>
    <form action="" method="get">
        <input type="hidden" name="mode" value="player" />
        <input type="hidden" name="cmd" value="list" />
    <table border="1" cellpadding="2" cellspacing="0" id="players">
        <tr>
            <td colspan="14" align="center"><h2>Players</h2></td>
        </tr>
        <tr>
                <td colspan="14" align="center" valign="middle">Search For:&nbsp;<input type="text" name="search" />&nbsp;<input type="submit" value="Go" class="button" /></td>
            </tr>
        <tr>
            <td>Name</td>
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
                        response.write "<a href=""javascript:insertPlayer('" & rs_players.fields("LASTNAME") & ", " & rs_players.fields("FIRSTNAME") & "'," & rs_players.fields("ID") & ");"">"
                        response.write rs_players.fields("LASTNAME") & ", " & rs_players.fields("FIRSTNAME")
                        response.write "</a></td>"
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
                    response.write "<a href=""?mode=" & strMode & "&cmd=" & strCmd & """>[First]</a> ... "
                end if
                if i = rstrAbsPage then
                    response.write "[" & i & "] "
                elseif i > rstrPageCount then
                    'next
                else
                    response.write "<a href=""?mode=" & strMode & "&cmd=" & strCmd & "&page=" & i & """>" & i & "</a> "
                    if (i = rstrAbsPage+3) and (rstrAbsPage+3 < rstrPageCount) then
                        response.write "... <a href=""?mode=" & strMode & "&cmd=" & strCmd & "&page=" & rstrPageCount & """>[Last]</a>"
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
end sub

sub config_roster()
end sub

incRosterFp = true
%>