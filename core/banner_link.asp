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
<!--#include file="inc_functions.asp" -->
<% 
intBanner = ""
whereto = ""
' Check for valid querystring
if Request.QueryString("id") <> "" or  Request.QueryString("id") <> " " then
	if IsNumeric(Request.QueryString("id")) = True then
		intBanner = cLng(Request.QueryString("id"))
	end if
end if

if intBanner <> "" then
  if intBanner <> 0 then
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	' get the Banner ID
	sSQL = "SELECT B_LINKTO FROM " & strTablePrefix & "BANNERS  WHERE " & strTablePrefix & "BANNERS.ID = " & intBanner & " AND " & strTablePrefix & "BANNERS.B_ACTIVE = 1"
	set rs = my_Conn.Execute (sSQL)
	
	if not rs.eof then
	' Put the link into a variable
	whereto = chkString(rs("B_LINKTO"),"urlstring")
	whereto = replace(rs("B_LINKTO"),"&amp;","&")
	' Update the hit count
	sSQL = "UPDATE " & strTablePrefix & "BANNERS SET " & strTablePrefix & "BANNERS.B_HITS = " & strTablePrefix & "BANNERS.B_HITS + 1  WHERE " & strTablePrefix & "BANNERS.ID = " & intBanner
	my_Conn.Execute (sSQL)
	end if
	
	set rs = nothing
	my_Conn.close
	set my_Conn = nothing
	
	if whereto <> "" then
		response.Redirect(whereto)
	end if
  else
	response.Redirect("http://www.SkyPortal.net")
  end if
end if
%>