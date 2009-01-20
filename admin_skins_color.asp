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
<!-- #include file="lang/en/core_admin.asp" -->
<!-- #include file="inc_functions.asp" -->
<!-- #include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" Then %>
<%
Session("strMsg") = ""

if request("tName") <> "" and request("tFolder") <> "" then

newThm = replace(replace(request("tName"),"<",""),">","")
thmFolder = replace(replace(request("tFolder"),"<",""),">","")

thmAuthor = Session.Contents("thmAuthor")
thmDescription = Session.Contents("thmDescription")
thmLogoImage = Session.Contents("thmLogoImage")
thmSubSkin = Session("thmSubSkin")

Session.Contents("thmFolder") = ""
Session.Contents("thmAuthor") = ""
Session.Contents("thmDescription") = ""
Session.Contents("thmLogoImage") = ""
Session("thmSubSkin") = ""



'set my_Conn= Server.CreateObject("ADODB.Connection")
'my_Conn.Open strConnString
		
	set rs1 = my_conn.execute ( "Select * from " & strTablePrefix & "COLORS where C_TEMPLATE = '" & newThm & "'")
	if not rs1.eof then
		Session("strMsg") = txtSknAlrSkinNam & ": <b>" & newThm & "</b><br />"
	end if
	set rs2 = my_conn.execute ( "Select * from " & strTablePrefix & "COLORS where C_STRFOLDER = '" & thmFolder & "'")
	if not rs2.eof then
		if Session("strMsg") <> "" then
			Session("strMsg") = Session("strMsg") & "<br />" & txtSknAlrFoldNam & ": <b>" & thmFolder & "</b>"
		else
			Session("strMsg") = txtSknAlrFoldNam & ": <b>" & thmFolder & "</b>"
		end if
	end if
	set rs2 = nothing
	' Skin Folder Bug Fix
	tmpFolder = chkString(thmFolder,"display")
	if trim(lcase(tmpFolder)) <> trim(lcase(thmFolder)) then
		Session("strMsg") = "There was a problem with the Folder Name: " & thmFolder & ".  Please rename the folder and try to re-add it."
	end if
	
	if Session("strMsg") = "" then
		executeThis( "INSERT INTO " & strTablePrefix & "COLORS (C_TEMPLATE) VALUES ('" & newThm & "')")
		
		strSql = "UPDATE " & strTablePrefix & "COLORS "
		strSql = strSql & " SET C_STRFOLDER = '" & tmpFolder & "', "
		strSql = strSql & " C_STRDESCRIPTION = '" & chkString(thmDescription,"display") & "', "
		strSql = strSql & " C_STRAUTHOR = '" & thmAuthor & "',"
		strSql = strSql & " C_STRTITLEIMAGE = '" & chkString(thmLogoImage,"display") & "', "
		strSql = strSql & " C_INTSUBSKIN = " & cLng(thmSubSkin) & ", "
		strSQL = strSQL & " C_SKINLEVEL = '1'"
		strSql = strSql & " WHERE C_TEMPLATE = '" & newThm & "'"
'response.Write(strSQL)
		executeThis(strSql)
		
		Session("strMsg") = replace(txtSknNewSknAdded,"[%skin%]",newThm)
		
	end if
		
my_Conn.close
set my_Conn = nothing
else
Session("strMsg") = txtSknNoSelect
end if
where = "admin_skins_config.asp"
Response.Redirect(where)

Else
	Response.Redirect("admin_login.asp")
End IF		
%>