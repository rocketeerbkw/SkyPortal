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

%>
<script type="text/javascript">
<!-- Begin script
function CheckNav(Netscape, Explorer) {
  if ((navigator.appVersion.substring(0,3) >= Netscape && navigator.appName == 'Netscape') ||
      (navigator.appVersion.substring(0,3) >= Explorer && navigator.appName.substring(0,9) == 'Microsoft'))
    return true;
  else return false;
}
//  End script -->
</SCRIPT>
<tr>
  <td align=center class="tSubTitle" colspan="2">
  <b><%= txtAvatar %>&nbsp;</b></td>
</tr>
<tr>
  <td class="fNorm" align=right nowrap valign=top><b><%= txtAvatar %>:&nbsp;</b></td>
  <td>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" height="64">
      <tr>
	<td width="35%" valign="top" align="left">
	<select name="url2" size="4" onchange ="if (CheckNav(3.0,4.0)) URL.src=form.url2.options[form.url2.options.selectedIndex].value;">
        	<OPTION <% if IsNull(rs("M_AVATAR_URL")) or rs("M_AVATAR_URL") = "" or rs("M_AVATAR_URL") = " " or rs("M_AVATAR_URL") = "files/avatars/noavatar.gif" or Request.QueryString("mode") = "Register" then %>selected value="<%= strSiteUrl %>files/avatars/noavatar.gif"> <%= txtNone %></OPTION><%else%>selected value="<% =rs("M_AVATAR_URL")%>"> <%= txtCurrent %></OPTION><option value="<%= strSiteUrl %>files/avatars/noavatar.gif"> <%= txtNone %></OPTION><% end if %>

<%		' - Get Avatar Settings from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR2.A_HSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_WSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"

		set rsavx = my_Conn.Execute (strSql)

		' - Get Avatars from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR.A_ID" 
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_URL"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_NAME"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_MEMBER_ID"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_MEMBER_ID = 0"
		if Request.Querystring("mode") <> "Register" then
			strSql = strSql & " OR " & strTablePrefix & "AVATAR.A_MEMBER_ID = " & rs("MEMBER_ID")
		end if
		strSql = strSql & " ORDER BY " & strTablePrefix & "AVATAR.A_ID ASC;"

		set rsav = Server.CreateObject("ADODB.Recordset")
		rsav.cachesize = 20
		rsav.open  strSql, my_Conn, 3

		if not(rs.EOF or rs.BOF) then  '## Avatars found in DB
			rsav.movefirst
			rsav.pagesize = strPageSize
			maxpages = cint(rsav.pagecount)
			howmanyrecs = 0
			rec = 1

			do until rsav.EOF '**
				if Request.Querystring("mode") <> "Register" then
%>
	               			<OPTION <% if rsav("A_URL") = rs("M_AVATAR_URL") then response.write("selected") %> VALUE="<% =rsav("A_URL") %>">&nbsp;<% =rsav("A_NAME") %></OPTION>
<%				else %>
	               			<OPTION VALUE="<% =rsav("A_URL") %>">&nbsp;<% =rsav("A_NAME") %></OPTION>
<%       			end if

			        rsav.MoveNext
	 			rec = rec + 1
			loop
		end if
		rsav.close
		set rsav = nothing
%>
		</select><br /><b><a href="pop_portal.asp?cmd=8" target="_blank">
    <font size="1"><%= txtAvLgnd %></font></a>
    </b>    </td>
	<td width="65%" valign="top" align="left"><img name="URL" src="<% if IsNull(rs("M_AVATAR_URL")) or rs("M_AVATAR_URL") = "" or rs("M_AVATAR_URL") = " " or Request.QueryString("mode") = "Register" then %><%= strHomeUrl %>files/avatars/noavatar.gif<% else %><% =rs("M_AVATAR_URL")%><% end if %>" border=<% =rsavx("A_BORDER") %> width=<% =rsavx("A_WSIZE") %> height=<% =rsavx("A_HSIZE") %>></td>
      </tr>
    </table>
    
    <% set rsavx = nothing %></td></tr>
