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
<!--#include file="inc_top_short.asp" -->

<%
if Request("cmd") <> "" then
	if IsNumeric(Request("cmd")) = True then cmd = cLng(Request("cmd")) else cmd = 0
end if
if Request("id") <> "" then
	if IsNumeric(Request("id")) = True then id = cLng(Request("id")) else id = 0
end if

if (id + cmd) < 1 then
	Response.Write	"      <p align=""center""><span class=""fTitle""><b>The URL has been modified!</b></span></p>" & vbNewLine & _
			"      <p align=""center""><b>" & txtPosHack & "</b></p>" & vbNewLine
  closeAndGo("stop")
end if
 
if strAuthType = "db" then
	strDBNTUserName = chkString(Request.Form("User"),"sqlstring")
end if
Select case cmd
	case 1
		strType = "add"
		Call showLogin()
	case 2
		strType = "del"
		Call showLogin()
	case 3
		if request.Form("strCmd") = "del" then
			Call delGlow()
		end if
		if request.Form("strCmd") = "add" then
			strType = "add"
			Call getColor()
		end if
	case 4
		if request.Form("strCmd") = "add" then
			Call addGlow()
		end if	
	case else
	
end select %>
<!--#include file="inc_footer_short.asp" -->

<%
Sub getColor()	' Set glow for member
	spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">") %>
<form action="pop_glow.asp?cmd=4" method="post" id="Form1" name="Form1">
<input type="hidden" name="id" value="<%= id %>">
<input type="hidden" name="strCmd" value="<%= strType %>">
<%
	Session.Contents("usr") = strDBNTUserName
	Session.Contents("pss") = chkString(Request.Form("pass"),"sqlstring")
	if hasAccess(2) then  ' is Member
	  if hasAccess(1) or strUserMemberID = cint(request.form("id")) then ' is Admin
		
			strSql = "SELECT " & strTablePrefix & "MEMBERS.M_GLOW, " & strTablePrefix & "MEMBERS.M_NAME FROM " & strTablePrefix & "MEMBERS WHERE " & strTablePrefix & "MEMBERS.MEMBER_ID = " & cint(request.form("id"))
'			response.Write(strSql & "<br />")
			set rs = my_Conn.Execute(strSql)
			
			if not rs.eof then
			  if rs("M_GLOW") <> "" then
				GlowColor = split(rs("M_GLOW"),":")(0)
				TxColor = split(rs("M_GLOW"),":")(1)
			  else
				GlowColor = split(def_glow,":")(0)
				TxColor = split(def_glow,":")(1)
			  end if
			end if
			
			if len(GlowColor) < 6 then
			  GlowColor = GlowColor
			else
			  if left(GlowColor,1) <> "#" then
			    GlowColor = "#" & GlowColor
			  end if
			end if
			if len(TxColor) < 6 then
			  TxColor = TxColor
			else
			  if left(TxColor,1) <> "#" then
			    TxColor = "#" & TxColor
			  end if
			end if
%>
      <tr>
        <td height="25" align="center" valign="middle" colspan="2"><b><%= txtCGloClr %>!</b>
		<%'"<br />" & rs("M_GLOW") %></td>
      </tr>
      <tr>
        <td height="30" align="right" valign="middle" nowrap><b><%= txtMember %>:&nbsp; </b></td>
		<% If GlowColor <> "" Then %>
        <td>&nbsp;&nbsp;<b><font id="glowname" style="filter:glow(color:<%= GlowColor %>,strength:4); width:100%" color="<% =TxColor %>"><%= getmembername(id) %></font></b></td>
		<% Else %>
        <td>&nbsp;&nbsp;<b><font id="glowname"><%= getmembername(id) %></font></b></td>
		<% End If %>
      </tr>
      <tr>
        <td align=right nowrap><b><%= txtGloClr %>:&nbsp; </b></td>
        <td><input type="text" name="strGlowColor" onBlur="shoGlow();" onFocus="shoGlow();" size="10" value="<% if GlowColor <> "" then Response.Write(GlowColor) else '## do nothing%>">
              <a href="JavaScript:;"><img src="<%= strHomeURL %>images/icons/icon_color.gif" border="0" onclick="openWindow3('includes/pop_colorwheel.asp?box=strGlowColor&form=Form1')"></a></td>
      </tr>
      <tr>
        <td align=right nowrap><b><%= txtTxtClr %>:&nbsp; </b></td>
        <td><input type="text" name="strTxColor" onBlur="shoText();" onFocus="shoText();" size="10" value="<% if TxColor <> "" then Response.Write(TxColor) else '## do nothing%>">
              <a href="JavaScript:;"><img src="<%= strHomeURL %>images/icons/icon_color.gif" border="0" onclick="openWindow3('includes/pop_colorwheel.asp?box=strTxColor&form=Form1')"></a></td>
      </tr>
	  <tr>
        <td height="25" colspan="2" align="center" valign="middle"><Input class="button" type="Submit" value=" <%= txtSubmit %> " id="Submit1" name="Submit1"></td>
      </tr>
<%		Else %>
<P><span class="fSubTitle"><B><%= txtNoPermAdGlo %></B></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<%		end if %>
<%	Else %>
<P><span class="fSubTitle"><B><%= txtNoPermAdGlo %></B></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<%	end if
	Response.Write("</table>")
spThemeBlock1_close(intSkin) %>
	</form>
<%
end sub

Sub delGlow()	' Delete glow from member
	spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
	if hasAccess(2) then  ' is Member
	  if hasAccess(1) then ' is Admin
		
			strSql = "UPDATE " & strTablePrefix & "MEMBERS "
			strSql = strSql & " SET " & strTablePrefix & "MEMBERS.M_GLOW = ''"
			strSql = strSql & " WHERE " & strTablePrefix & "MEMBERS.MEMBER_ID = " & chkstring(request.form("id"), "sqlstring")
'			response.Write(strSql & "<br />")
			my_Conn.Execute strSql
%>
<P><span class="fSubTitle"><B><%= txtGloRem %></B></span></p>
<script type="text/javascript"> opener.document.location.reload();</script>
<%		Else %>
<P><span class="fSubTitle"><B><%= txtNoPermRemGlo %></B></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<%		end if %>
<%	Else %>
<P><span class="fSubTitle"><B><%= txtNoPermRemGlo %></B></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<%	end if
	Response.Write("</table>")
spThemeBlock1_close(intSkin)
end sub

Sub addGlow()
	spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
	
	strDBNTUserName = Session.Contents("usr")
	strPass = Session.Contents("pss") 
	Session.Contents.Remove("usr")
	Session.Contents.Remove("pss")
	
	if hasAccess(2) then  ' is Member
	  if hasAccess(1) or strUserMemberID = cint(request.form("id")) then ' is Admin

			strGlow = chkstring(request.form("strGlowColor"), "sqlstring")
			strColor = chkstring(request.form("strTxColor"), "sqlstring")
			glowColor = strGlow & ":" & strColor
			
			' Add glow colors to member
			strSql = "UPDATE " & strTablePrefix & "MEMBERS "
			strSql = strSql & " SET " & strTablePrefix & "MEMBERS.M_GLOW = '" & glowcolor & "'"
			strSql = strSql & " WHERE " & strTablePrefix & "MEMBERS.MEMBER_ID = " & chkstring(request.form("id"), "sqlstring")

'			response.Write(strSql & "<br />")
			my_Conn.Execute strSql
%>
<P><span class="fSubTitle"><%= txtGloAdd %></span></p>
<script type="text/javascript"> opener.document.location.reload();</script>
<%	Else %>
<P><span class="fSubTitle"><%= txtNoPermAdGlo %></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<%	end if %>
<% Else %>
<P><span class="fSubTitle"><%= txtNoPermAdGlo %> : <%= strDBNTUserName %></span></p>
<p><a href="JavaScript: onclick= history.go(-1)"><%= txtGoAuth %></a></p>
<% end if
	Response.Write("</table>")
spThemeBlock1_close(intSkin)
end sub

' :::::::::::::: LOGIN FORM ::::::::::::::::::::::::
Sub showLogin() %>
<P> <span class="fTitle">
<%				if Request.QueryString("cmd") = "1" then %>
						<%= txtAddEdtGlo %>
<%				else %>
<%					if Request.QueryString("cmd") = "2" then %>
							<%= txtRemGlo %>
<%					end if %>
<%				end if %>
</span></p>

<p><span class="fSubTitle"><b><%= txtAdminAdGlo %>.</b></span>
</p>

<form action="pop_glow.asp?cmd=3" method="post" id="Form1" name="Form1">
<input type=hidden name="id" value="<%= id %>">

<%
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
	if strAuthType="db" then %>
      <tr>
        <td align=right nowrap><b><%= txtUsrNam %>:&nbsp; </b></td>
        <td><p><input type=text name="User" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size=20></p></td>
      </tr>
      <tr>
        <td align=right nowrap><b><%= txtPass %>:&nbsp; </b></td>
        <td><input type=Password name="Pass" size=20></td>
      </tr>
	  <% Else %>
<%					if strAuthType <> "db" then %>
      <tr>
        <td align=right nowrap><b><%= txtNTacct %>:</b></td>
        <td><%=Session(strUniqueID & "userID")%></td>
      </tr>
<%					end if %>
<% end if %>
	  <% If strType = "del" Then %>
      <tr>
        <td align=right nowrap><b><%= txtRemGlow %>:&nbsp; </b></td>
        <td><input type=radio name="strCmd" value="del"></td>
      </tr>
      <tr>
        <td align=right nowrap><b><%= txtEditGlo %>:&nbsp; </b></td>
        <td><input type=radio name="strCmd" value="add" checked></td>
      </tr>
	  <% Else %>
	  <input type=hidden name="strCmd" value="<%= strType %>">
	  <% End If %>
      <tr>
        <td colspan=2 align=center><Input class="button" type=Submit value=" Submit " id=Submit1 name=Submit1></td>
      </tr></table>
<%
spThemeBlock1_close(intSkin)%>
</form>
<% end sub %>