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

if strAuthType = "db" then
	if (Request.Cookies(strUniqueID & "User")("Name") <> "" and Request.Cookies(strUniqueID & "User")("PWord") <> "") then

		strSql = "SELECT MEMBER_ID, M_NAME, M_PASSWORD "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(Request.Cookies(strUniqueID & "User")("Name"), "SQLString") & "' "
		strSql = strSql & " AND M_PASSWORD = '" & ChkString(Request.Cookies(strUniqueID & "User")("Pword"), "SQLString") &"'"
		Set rsCheck = my_Conn.Execute(strSql)
		if rsCheck.BOF or rsCheck.EOF then
			Call ClearCookies()
			strAdmin1UserName = ""
		else
			strAdmin1UserName = rsCheck("M_NAME")
		end if
		rsCheck.close
		set rsCheck = nothing
	else
		strAdmin1UserName = ""
	end if
end if
strAdmin1FUserName = chkString(Request.Form("Name"),"sqlstring")
if strAuthType <> "db" then
	strAdmin1UserName = Session(strUniqueID & "userID")
	strAdmin1FUserName = Session(strUniqueID & "userID")
end if
if strAdmin1UserName <> "" then strUserMemberID = getMemberID(strAdmin1UserName) else strUserMemberID = 0

%>