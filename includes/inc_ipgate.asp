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
'#################################################################################
'## NET IPGATE v2.3.0 Orig Idea by alex042@aol.com(c)Aug 2002, 
'## rewritten by www.gpctexas.net Patrick@gpctexas2.net Aug, 2004
'##
'## rewritten by Hawk92 Nov 2004 for compatability with SkyPortal ver 1.50
'#################################################################################
'################################################################################# %>
<%'user changeable variable**********************************************************************
cookiename="MWPipgate" 'must change for each forum you host on the same domain!
psw="6f6^&f5924@f3e6969"  ' change this to an encryption key for your site

headcss="class=""fSubTitle"""
headnocss=""
headfontnocss=""

catcss="class=""fSubTitle"""
catnocss=""
catfontnocss=""
tableshadow=""

forumcss=""
forumnocss=""
fontnocss=""

' no edit below this point!********************************************************************
strIPGateCss = 1
pagereq=Request.ServerVariables("Path_Info")
PathLen = InStrRev(pagereq,"/",-1)
pagreq = lcase(Right(pagereq,(len(pagereq)-PathLen)))
	
userhost=request.servervariables("REMOTE_HOST")
userdate=strCurDateString
pageqryd=chkString(Request.ServerVariables("QUERY_STRING"),"sqlstring")
'pagreq=request.servervariables("SCRIPT_NAME")
pagereqtemp=pagreq & "?" & pageqryd 
getthecookie=EnDeCrypt(Request.Cookies(cookiename), psw)
notindb=1
FoundIP=0
FoundName=0
	
pagename=pagreq
if pageqryd <> "" and Len(pagereqtemp) <= 245 then 'database field length check for access db users with a little leway
	pagreq=pagereqtemp
end if

If (Request.ServerVariables("HTTP_CLIENT_IP") <> "") and (lcase(Request.ServerVariables("HTTP_CLIENT_IP")) <> "unknown") then
   userip=Request.ServerVariables("HTTP_CLIENT_IP")
else
   if (Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "") and (lcase(Request.ServerVariables("HTTP_X_FORWARDED_FOR")) <> "unknown") then
      userip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
   else
      if (Request.ServerVariables("REMOTE_ADDR") <> "") and (lcase(Request.ServerVariables("REMOTE_ADDR")) <> "unknown")  then
         userip=Request.ServerVariables("REMOTE_ADDR")
      else
         userip="999.999.999.999"
      end if
   end if
end if

if strIPGateBan = 1 and strIPGateLck = 0 then
	'get db records
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	StrSql = "SELECT * FROM " & strTablePrefix & "IPLIST"
	count = 0
	rs1.Open StrSql, strConnString
	
	do until ((rs1.eof) or (Foundip=1) or (FoundName=1))
		dbrecord = rs1("IPLIST_STARTIP") & "."
		dbrecordarr = split(dbrecord,".")
		useriparr = split(userip,".")

		
		'start matching record to current username or ip
		if lcase(trim(rs1("IPLIST_MEMBERID"))) = lcase(trim(strDBNTUsername)) then FoundName = 1
		if rs1("IPLIST_STARTIP") = userip then FoundIP = 1
		if FoundIP <> 1 and (dbrecordarr(2) = "") and ((dbrecordarr(0) =  useriparr(0)) and (dbrecordarr(1) =  useriparr(1))) then FoundIP = 1
		if FoundIP <> 1 and (dbrecordarr(3) = "") and ((dbrecordarr(0) =  useriparr(0)) and (dbrecordarr(1) =  useriparr(1)) and (dbrecordarr(2) =  useriparr(2))) then FoundIP = 1

				
		if (FoundName = 1 or FoundIP = 1) then
			select case rs1("IPLIST_STATUS")
				case "0"
					if ((userdate >= rs1("IPLIST_STARTDATE")) and (userdate <= rs1("IPLIST_ENDDATE"))) then
						if strIPGateCok = 1 then call writethecookie ("Banned", ChkDate(rs1("IPLIST_ENDDATE")))
						if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
						if not hasAccess(1) then call processstatus ("banned")
					end if
				case "1"
					if ((userdate >= rs1("IPLIST_STARTDATE")) and (userdate <= rs1("IPLIST_ENDDATE"))) then
						if strIPGateCok = 1 then call writethecookie ("Watched", ChkDate(rs1("IPLIST_ENDDATE")))
						if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
						if not hasAccess(1) then call processstatus ("watched")
					end if
				case "2"
					if ((userdate >= rs1("IPLIST_STARTDATE")) and (userdate <= rs1("IPLIST_ENDDATE"))) then
						pagefound = 0
						Set rs3 = Server.CreateObject("ADODB.Recordset")
						PageSql = "SELECT * FROM " & strTablePrefix & "PageKeys"
		
						rs3.Open PageSql, strConnString
						do until ((rs3.eof) or (pagefound = 1))
						
							if lcase(rs3("PAGEKEYS_PAGEKEY")) = lcase(pagename) then
							   if strIPGateCok = 1 then call writethecookie ("Blocked", ChkDate(rs1("IPLIST_ENDDATE")))
							   pagefound = 1
							   if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
							   if not hasAccess(1) then call processstatus ("noaccess")
							end if
					
						rs3.MoveNext
    					Loop
    					if rs3.state=1 then rs3.close
    					set rs3 = nothing
					end if
				case "3"
					call writethecookie ("Expired", DateAdd("d",-14,Date)) 'sets cookie date today-14 days thus expiring it
			end select
		end if
		
	rs1.MoveNext
	loop
	
	if rs1.state=1 then rs1.close
    set rs1 = nothing
	
	'not in db at this point lets search for cookies
	if strIPGatecok = 1 and not hasAccess(1) and FoundIP = 0 and FoundName = 0 then
		select case getthecookie
			case "Banned"
				if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
				call autoinsert("Banned", "0")
				call processstatus ("banned")
			case "Watched"
				if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
				call autoinsert("watched", "1")
				call processstatus ("watched")
			case "Blocked"
				Set rs3 = Server.CreateObject("ADODB.Recordset")
				PageSql = "SELECT * FROM " & strTablePrefix & "PageKeys"
		  		
		  		rs3.Open PageSql, strConnString
 				do until ((rs3.eof) or (pagefound=1))
 					if lcase(rs3("PAGEKEYS_PAGEKEY")) = lcase(pagename) then 
						pagefound = 1
						call autoinsert("Blocked", "2")						
						if strIPGatetype = 0 and strIPGatelog = 1 then call loguser()
						call processstatus ("noaccess")
					end if
				rs3.MoveNext
    			Loop
    			if rs3.state=1 then rs3.close
    			set rs3 = nothing
    	end select
	end if
end if
	

if StrIPGateTyp = 1 and strIPGateBan = 1 then
 call loguser()
end if

if strIPGateLck = 1 and not hasAccess(1) and strIPGateBan = 1 then
  call processstatus ("lockdown")
end if

' functions ******************************************************

sub loguser()
	
	Set rs2 = Server.CreateObject("ADODB.Recordset")
	StrSql = "INSERT into " & strTablePrefix & "IPLOG (IPLOG_MEMBERID, IPLOG_IP, IPLOG_PATHINFO, IPLOG_DATE) "
	StrSql = StrSql & "values ('" & strDBNTUserName & "','" & userip & "','" & pagreq & "','" & userdate & "')"
	rs2.Open StrSql, strConnString 
	if rs2.State = 1 then rs2.Close
	set rs2 = nothing	
	
end sub

sub autoinsert(message, code)
	if strDBNTUsername = "" then 
	   tempname = "0"
	else
	   	tempname = strDBNTUsername
	end if
	Set rs = Server.CreateObject("ADODB.Recordset")
	strSql = "INSERT into " & strTablePrefix & "IPLIST (IPLIST_MEMBERID, IPLIST_STARTIP, IPLIST_STARTDATE, IPLIST_ENDDATE, IPLIST_COMMENT, IPLIST_STATUS)"
	strSql = strSql & "values ('" & tempname& "','" & userip & "','" & strCurDateString & "','" & DateToStr(dateadd("m", 1, date())) & "','" & "Auto inserted: " & message & " cookie detected." & "','" & code & "')"
	rs.Open StrSql, strConnString 
	if rs.State = 1 then rs.Close
end sub

sub writethecookie (cookietype, cookiedate)
	
	if strIPGateCok = 1  Then
	    cookietype=EnDeCrypt(cookietype, psw)
	    Response.Cookies(trim(cookiename))=cookietype
	    Response.Cookies(trim(cookiename)).Expires = cookiedate
	end if

end sub

sub processstatus (statustype)
	Select Case StrIPGateMet
		case "0"
			select case statustype
				case "banned"
					%>
					<div align="center" <% if strIPGateCss = 1 then response.write(forumcss) else response.write(forumnocss) end if %>>
					<table border="0" width="100%" id="table1" style="border-collapse: collapse">
						<tr>
							<td align="center" <% if strIPGateCss = 1 then response.write(forumcss) else response.write(forumnocss) end if %>>
								<% if strIPGateCss = 0 then response.write(fontnocss)%>
								<b>
									<br />
								<br />
								<br />
								<br />
									<%= txtMsg %>: <%=StrIPGateMsg%>
								<br />
								<br />
								<br />
								<br />
								</b>
								<% if strIPGatecCss = 0 then response.write ("") %>
							</td>
						</tr>
					</table>
					</div>
					<%
					closeandgo("stop")
				case "watched"
					%>
					<table border="0" width="100%" id="table1" style="border-collapse: collapse">
						<tr>
							<td align="center" <% if strIPGateCss = 1 then response.write(forumcss) else response.write(forumnocss) end if %>>
								<% if strIPGateCss = 0 then response.write(fontnocss)%>
								<b>
									<br />
								<br />
								<br />
								<br />
									<%= txtMsg %>: <%=StrIPGateWarnMsg%>
								<br />
								<br />
								<br />
								<br />
								</b>
								<% if strIPGatecCss = 0 then response.write ("") %>
							</td>
						</tr>
					</table>
					<%
				case "noaccess"
					%>
					<table border="0" width="100%" id="table1" style="border-collapse: collapse">
						<tr>
							<td align="center" <% if strIPGateCss = 1 then response.write(forumcss) else response.write(forumnocss) end if %>>
								<% if strIPGateCss = 0 then response.write(fontnocss)%>
								<b>
									<br />
								<br />
								<br />
								<br />
									<%= txtMsg %>: <%=strIPGateNoAcMsg%>
								<br />
								<br />
								<br />
								<br />
								</b>
								<% if strIPGatecCss = 0 then response.write ("") %>
							</td>
						</tr>
					</table>
					<%
					closeandgo("stop")
				case "lockdown"
					%>
					<table border="0" width="100%" id="table1" style="border-collapse: collapse">
						<tr>
							<td align="center" <% if strIPGateCss = 1 then response.write(forumcss) else response.write(forumnocss) end if %>>
								<% if strIPGateCss = 0 then response.write(fontnocss)%>
								<b>
									<br />
								<br />
								<br />
								<br />
									<%= txtMsg %>: <%=StrIPGateLkMsg%>
								<br />
								<br />
								<br />
								<br />
								</b>
								<% if strIPGatecCss = 0 then response.write ("") %>
							</td>
						</tr>
					</table>
					<%
					closeandgo("stop")
			end select
		case "1"
			select case statustype
				case "banned"
					ipgate_banned()
				case "noaccess"
					ipgate_noaccess()
				case "lockdown"
					ipgate_lockdown()			
			end select
			closeandgo("stop")
	end select
end sub
%>
