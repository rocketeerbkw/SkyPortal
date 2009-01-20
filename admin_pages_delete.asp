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

<% If Session(strCookieURL & "Approval") = "256697926329" Then %>

	<!--#include file="inc_functions.asp" -->
	<!--#include file="inc_top.asp" -->
	<%
p_id=request.querystring("p_id")
grp_id=request.querystring("grp_id")

	if p_id <> "" then
			dSQL = "select * from PORTAL_PAGES where p_id =" & p_id
			set rsD = my_Conn.execute(dSQL)
				if rsD("P_CAN_DELETE") = "1" then
						strSql2 = "DELETE from PORTAL_PAGES where P_ID=" & p_id
						Set rs = my_Conn.Execute (strSql2)
					else
						response.write "error - you can't delete this file, it is locked.  unlock file then delete"
					end if
			set rs= nothing
			set rsD= nothing
			response.redirect("admin_config_cp.asp")
	else
		response.redirect "admin_home.asp"
	end if

else
	response.redirect "admin_login.asp"
end if
 %>