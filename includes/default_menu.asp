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

img_new = "<img src=""themes/" & strTheme & "/new.gif"" border=""0"" alt="""" title=""New Items"" />"

function menu_fp()

' Count the number of MEMBERS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Members] "
else
	strSqL = "SELECT count(UserID) Members  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) <> '" & txtGuest & "' "

Set rsMembers = my_Conn.Execute(strSql)
if not rsMembers.eof then
iolMembers = rsMembers("Members")
strOnlineMembersCount = rsMembers("Members")
else
iolMembers = 0
strOnlineMembersCount = 0
end if
Set rsMembers = nothing

' Count the number of GUESTS online
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) = '" & txtGuest & "' "

Set rsGuests = my_Conn.Execute(strSql)
if not rsGuests.eof then
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")
else
Guests = 0
strOnlineGuestsCount = 0
end if
Set rsGuests = nothing

'::::::::::::::::::::::: Start the menu HTML ::::::::::::::::::::::::::::::
spThemeTitle= txtMenu
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)

defaultMenu()
%>

<table>
<tr><td width="100%"><hr /></td></tr>
<% if hasAccess(2) then
strSql = "SELECT " & strTablePrefix & "TOTALS.U_COUNT "
strSql = strSql & " FROM " & strTablePrefix & "TOTALS"
set rs1 = my_Conn.Execute(strSql)
Users = rs1("U_COUNT")
rs1.Close
set rs1 = nothing
%>
<tr><td width="100%"><span class="fSmall"><a href="members.asp"><%= txtMembers %>: <% =Users%></a></span></td></tr>
<% End If %>
<tr><td width="100%"><a href="active_users.asp"><span class="fSmall"><%= txtActvUsrs %>: <br /><%=strOnlineMembersCount & " " & txtMembers & " " & txtAnd & " " & strOnlineGuestsCount & " " & txtGuests %></span></a></td></tr></table>
<% 
spThemeBlock1_close(intSkin)
end function
%>