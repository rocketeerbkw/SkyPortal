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
CurPageType="PM"
%>
<!--#include virtual="config.asp" -->
<!--#include virtual="inc_functions.asp" -->
<script type="text/javascript">
<!--
function autoReload()
{
	document.ReloadFrm.submit()
}
//-->
</script>

<script type="text/javascript"> 
<!-- 
<% If strAllowHtml <> 1 Then %>
function openWindowPM(url) {
  popupWin = window.open(url,'pm_pop_send','resizable,width=500,height=340,top=75,left=220,scrollbars=yes')
}
<% Else %>
function openWindowPM(url) {
  popupWin = window.open(url,'pm_pop_send','resizable,width=590,height=510,top=30,left=120,scrollbars=yes')
}
<% End If %>
//--> 
</SCRIPT>
<!--#include virtual="inc_top_short.asp" -->
<%
'## Do Cookie stuffs with reload
nRefreshTime = trim(chkString(Request.Cookies(strCookieURL & "Reload"),"sqlstring"))

if Request.form("cookie") = "1" then	
    Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	Response.Cookies(strCookieURL & "Reload") = chkString(Request.Form("RefreshTime"),"sqlstring")
	Response.Cookies(strCookieURL & "Reload").expires = dateAdd("d", 360, now())
	nRefreshTime = chkString(Request.Form("RefreshTime"),"sqlstring")
end if

if nRefreshTime = "" then
	nRefreshTime = 6
end if
%>

<%
' Get Guest count for display on Default.asp
set rsGuests = Server.CreateObject("ADODB.Recordset")

if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) = 'Guest' "

Set rsGuests = my_Conn.Execute(strSql)
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")

spThemeTitle= "Pager"

spThemeTableCustomCode = "width=""95%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
	<tr>
		<td valign="middle" align="left" nowrap>
<%
        mypage = trim(chkString(request("whichpage"),"sqlstring"))

	If  mypage = "" then
	   mypage = 1
	end if
	mypagesize = ""
	If  mypagesize = "" then
	   mypagesize = 15
	end if
	set rs = Server.CreateObject("ADODB.Recordset")
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated "
	strSql = strSql & " FROM " & strMemberTablePrefix & "ONLINE "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated DESC"
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3
	i = 0 
	If rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write ""
	Else
		rs.movefirst
		num = 0
		rs.pagesize = mypagesize
		maxpages = cint(rs.pagecount)
		maxrecs = cint(rs.pagesize)
		rs.absolutepage = mypage
		howmanyrecs = 0
		rec = 1
		do until rs.EOF or (rec = mypagesize+1)
			if Right(rs("UserID"), 5) <> "Guest" then 
				strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "ONLINE.UserID "
				strSql = strSql & " FROM " & strTablePrefix & "MEMBERS, " & strTablePrefix & "ONLINE "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & rs("UserID") & "' "
				set rsMember =  my_Conn.Execute (strSql)
			end if
			if Right(rs("UserID"), 5) <> "Guest" then
				Response.Write("&nbsp;<a href=""Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=" & rsMember("MEMBER_ID") & "');"">")
				Response.Write(rs("UserID") & "</a> ")
				Response.Write("<a href=""Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=" & rsMember("MEMBER_ID") & "');"">")
				Response.Write("<img src=" & strHomeUrl & "images/icons/pm.gif border=0 width=11 height=17 align=absmiddle hspace=6>" & "</a><br />")			          
			end if
			rs.MoveNext
			rec = rec + 1
		loop
 %>
		&nbsp;<% =Guests %><% if Guests=1 then %>&nbsp;guest<% else %>&nbsp;guests<br /><% end if %>
<%	end if %>
		</td>
	</tr>
	<tr>
		<td valign="middle" align="left" nowrap>
<%      ' Get Private Message count
	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSqL = "SELECT count(M_TO) as pmcount " 
    end if
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
		strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 

	Set rsPM = my_Conn.Execute(strSql)
		pmcount = rsPM("pmcount")
%>
<%	if strDBNTUserName = "" Then %>
		<IMG SRC="<%= strHomeUrl %>images/icons/icon_pmdead.gif" align=absmiddle border=0 hspace=6> Please Login
<%	else
		if pmcount = 0 then %>
		<A HREF="pm.asp" target="_new"><IMG SRC="<%= strHomeUrl %>images/icons/icon_pm.gif" align=absmiddle border=0 hspace=6></a>(<% =pmcount %>) new <A HREF="pm.asp" target="_new">messages</a>.
<%		end if
        if pmcount >= 1 then %>
        <EMBED SRC="<%= strHomeUrl %>images/newmsg.wav" WIDTH=1 HEIGHT=1 HIDDEN="true" AUTOSTART="true" LOOP="false" volume="100"></EMBED>
<%
      	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
      	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
      	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
      	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
      	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 
      	Set rsMessage = my_Conn.Execute(strSql)
      	i = 0
      		do Until rsMessage.EOF or (i = 3)
	  	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID  "
	  	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	  	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsMessage("M_FROM") & ""
          	Set rsFrom = my_Conn.Execute(strSql)	
          	if rsMessage("M_READ") = "0" then %>
			<script language=javascript>window.open('<%= strHomeUrl %>pm_pop.asp?mode=1&amp;cid=<% =rsMessage("M_ID") %>','_blank','width=490,height=340,top=75,left=220,scrollbars=yes')</script>
<%        		i = i + 1
			end if
	    	rsMessage.MoveNext	
       		Loop
%>
        <A HREF="pm.asp" target="_new"><IMG SRC="<%= strHomeUrl %>images/icons/icon_pm_new.gif" align=absmiddle border=0 hspace=6></a>(<b><% =pmcount %></b>) new <A HREF="pm.asp" target="_new"><% if pmcount = 1 then %>message<% else %>messages<% end if %></a>.
<%		end if 
	end if %>	
		</td>
	</tr>
	<tr>
	<form name="ReloadFrm" action="pm_pop_pager.asp" method="post"> 	
		<td height="10" valign="middle" align="right" nowrap>
    			<select name="RefreshTime" size="1" onchange="autoReload();" style="font-size:10px;">
        			<option value="0"  <% if nRefreshTime = "0" then Response.Write(" SELECTED")%>>Don't
        			auto refresh</option>
        			<option value="3"  <% if nRefreshTime = "3" then Response.Write(" SELECTED")%>>30 second 
        			refresh</option>
        			<option value="4.5"  <% if nRefreshTime = "4.5" then Response.Write(" SELECTED")%>>45 second
        			refresh</option>
        			<option value="6" <% if nRefreshTime = "6" then Response.Write(" SELECTED")%>>1 minute
        			refresh</option>
        			<option value="12" <% if nRefreshTime = "12" then Response.Write(" SELECTED")%>>2 minute
        			refresh</option>
        			<option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>>5 minute
    	    			refresh</option>
    			</select>
    		</td>	
	<input type="hidden" name="Cookie" value="1">
	</form>
    	</tr>
<% if maxpages > 1 then %>    	
    	<tr>
		<td valign="middle" align="left"><b><% Call Paging() %></b></td>
  	</tr>
<% end if
Response.Write("</table>")
spThemeBlock1_close(intSkin)%>

<%
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = chkString(Request.QueryString("whichpage"),"sqlstring")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 width=95% cellspacing=0 cellpadding=1 align=top><tr><td align=left>Pages: </td>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then
				ref = "<td align=right>" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname
				ref = ref & "?whichpage=" & counter
				ref = ref & "&pagesize=" & mypagesize
				if top = "1" then
					ref = ref & "'><b>" & counter & "</b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=right>&nbsp;" & widenum(counter) & " <b>" & counter & "</b></td>")
			end if
			if counter mod 15 = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub 
%>

<SCRIPT>
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 5000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 10000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</SCRIPT>
<div align=center><center>
<!--#include virtual="inc_footer_short.asp" -->

