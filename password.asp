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
curPageType = "register"
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%
dim memKey, actKey, strPass, strPass2, showForm, showForm2
memKey = ""
actKey = ""
strPass = ""
strPass2 = ""
memId = ""
showForm = true
showForm2 = false

if strLoginStatus = 1 then
  'closeAndGo("default.asp")
end if	
%>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td class="leftPgCol" align="center" valign="top">
	<% 
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %>
	</td>
	<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtForgotPass & "?|password.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

if lcase(strEmail) = "1" then
  if request.querystring("mode") = "reset" then
	if session.Contents("memKey") <> "" and (session.Contents("memKey") = request.form("memKey")) then
	  session.Contents("memKey") = ""
	  session.Contents.Remove("memKey")
	  'validate form variable
	  if request.form("memId") <> "" then
  	    if IsNumeric(request.form("memId")) = True then
		  memId = cLng(request.form("memId"))
	  	  if memId <> "" and memId < 1 then
	  	  	raiseHackAttempt("")
	  		showForm = false
	  		'closeAndGo("default.asp")
	      end if
  	  	else
	      raiseHackAttempt("")
	      showForm = false
	      'closeAndGo("default.asp")
  		end if
	  end if
	  if request.form("memKey") <> "" then
  	    if len(request.form("memKey")) = 10 then
		  memKey = replace(replace(chkString(trim(Request.form("memKey")),"sqlstring")," ",""),"=","")
  	    else
	      raiseHackAttempt("")
	      showForm = false
		  'closeAndGo("default.asp")
  	    end if
	  end if
	 if showForm = true then
	  strPass = trim(chkString(Request.Form("pass"), "sqlstring"))
	  strPass2 = trim(chkString(Request.Form("pass2"), "sqlstring"))
	  if strPass = "" or strPass2 = "" then
		Err_Msg = Err_Msg & "<b>" & txtPassNoMatch & "</b><br />"
		showForm = false
		showForm2 = true
		actKey = memKey
	  elseif strPass <> strPass2 then
		Err_Msg = Err_Msg & "<b>" & txtPassNoMatch & "</b><br />"
		showForm = false
		showForm2 = true
		actKey = memKey
	  elseif len(memKey) <> 10 then
		Err_Msg = Err_Msg & "<b>" & txtValNoMatch & "</b><br />"
		raiseHackAttempt("")
	    showForm = false
		'closeAndGo("default.asp")
	  else
		memPass = pEncrypt(pEnPrefix & strPass)
		Err_Msg = txtPassChgSuccess & ".<br /><br />" & txtLoginNewPass & "."
		'verKey = GetKey("passemail")
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_KEY = '', M_PASSWORD = '" & memPass & "' WHERE MEMBER_ID = " & memId & " AND M_KEY = '" & memKey & "'"
		executeThis(strSql)
        response.Write("<table width=""350"" align=""center""><tr>")
        response.Write("<td align=""center"">")
	    spThemeBlock1_open(intSkin)
		response.Write("<br /><center><b>" & Err_Msg & "</b><br /><br />")
	    spThemeBlock1_close(intSkin)
		response.Write("</td></tr></table>")
		showForm = false
	  end if
     end if
	else
	  'raiseHackAttempt("[THIS IS ONLY A TEST - DO NOT REPLY TO THIS EMAIL]")
	  showForm = false
	  closeAndGo("default.asp")
	end if

  elseif request.querystring("mode") = "validateEmail" then
	actKey = ""
	if request.QueryString("actKey") <> "" then
  	  if len(request.QueryString("actKey")) = 10 then
		actKey = replace(replace(chkString(trim(Request.QueryString("actKey")),"sqlstring")," ",""),"=","")
  	  else
	    raiseHackAttempt("")
	    showForm = false
  	  end if
	end if
	if showForm = true then
	  strSql = "SELECT MEMBER_ID, M_NAME, M_KEY"
	  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_KEY = '" & actKey & "'"
	  set rsKey = my_Conn.Execute (strSql)
	  if rsKey.EOF or rsKey.BOF then
		Err_Msg = txtThereIsProb & "<br />"
		Err_Msg = Err_Msg & txtValNoMatch & ".<br />"
	   showForm = true	
	  else 
	    session.Contents("memKey") = rsKey("M_KEY")
	    showForm = false %>
        <table width="350" align="center"><tr>
        <td align="center">	
	    <form action="password.asp?mode=reset" method="post">
	    <input type="hidden" name="memId" value="<%= rsKey("MEMBER_ID") %>">
	    <input type="hidden" name="memKey" value="<%= rsKey("M_KEY") %>">
	    <%
	    spThemeTitle= txtChNewPass
	    spThemeBlock1_open(intSkin)
	    %>
       <table><tr>
       <td class="tCellAlt1"><b><br /><%= txtUsrName %>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =rsKey("M_NAME") %><br /><br /><%= txtNewPass %>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <input class="textbox" type="password" name="pass" size="25"><br /><br />
	   <%= txtRepPass %>: <input class="textbox" type="password" name="pass2" size="25">
	   <br /><br /><input type="submit" value="<%= txtSubmit %>" class="button"></b>
	   </td></tr></table>
	   <%spThemeBlock1_close(intSkin)%>
	   </form>
	   </td></tr></table>	
<%	  end if
	  rsKey.close
	  set rsKey = nothing
	end if


  elseif request.querystring("mode") = "sendEmail" then
	emailAddress = chkString(Trim(Request.Form("Email")), "sqlstring")
	memberName = chkString(Trim(Request.Form("mName")), "sqlstring")
	browserIP = request.ServerVariables("REMOTE_ADDR")
	  strSql = "SELECT M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
	  strSql = strSql & " WHERE M_EMAIL = '" & emailAddress &"' or M_NAME='" & memberName & "'"
	  set rs = my_Conn.Execute(strSql)
	  if rs.BOF and rs.EOF then 
		Err_Msg =  txtEmlNoExist & ".<br />"
	  else 
		memName = rs("M_NAME")
		memEmail = rs("M_EMAIL")
		Err_Msg = txtEmlPassSent & ".<br /><br />"
	       	verKey = GetKey("passemail")
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_KEY = '" & verKey & "' WHERE M_EMAIL = '" & memEmail & "'"
		executeThis(strSql)
	  rs.close
	  set rs = nothing
	end if
  end if
end if
if showForm then
  'raiseHackAttempt("[THIS IS ONLY A TEST - DO NOT REPLY TO THIS EMAIL]")
  if strEmail = 1  then %>
    <table width="350" align="center" cellpadding="0" cellspacing="0"><tr>
    <td align="center">
	<form action="password.asp?mode=sendEmail" method="post">
	<%
	spThemeTitle= txtForgotPass
	spThemeBlock1_open(intSkin)
	%><table border="0" cellpadding="6" cellspacing="0">
	<% if Err_Msg <> "" then %>
	<tr><td class="tCellAlt1" align="center">
	<b><% =Err_Msg %></b><hr />
	</td></tr>
	<% end if %>
	<tr>
    <td class="tCellAlt1" align="center"><%= txtPassText %><br /><br />
	<%= txtUsrName %>: <input class="textbox" type="text" name="mName" size="25"><br />
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <%= txtEmail %>: <input class="textbox" type="text" name="Email" size="25"><br />
	<input type="submit" value="<%= txtSubmit %>" class="button"><br /><br /></td>
    </tr></table>
	<% spThemeBlock1_close(intSkin) %>
	</form>
	</td></tr></table>
	<%
  else 'email is OFF, PM the admin

	spThemeTitle= txtForgotPass
	spThemeTableCustomCode = "align=""center"" width=""95%"""
	spThemeBlock1_open(intSkin)
	%><table cellpadding="0" cellspacing="0">
	<% if Err_Msg <> "" then %>
	<tr><td class="tCellAlt1" align="center">
	<% =Err_Msg %>
	</td></tr>
	<% end if %>
    <tr>
    <td class="tCellAlt1" align="center"><br /><%= txtPassText2 %><br /><br />
	<input type="button" value="<%= txtSubmit %>" class="button"></td>
  </tr></table>
<%spThemeBlock1_close(intSkin)%>
<%
  end if
end if

if showForm2 then
	  strSql = "SELECT MEMBER_ID, M_NAME, M_KEY"
	  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_KEY = '" & actKey & "'"
	  set rsKey = my_Conn.Execute (strSql)
	  if rsKey.EOF or rsKey.BOF then
		Err_Msg = txtThereIsProb & "<br />"
		Err_Msg = Err_Msg & txtValNoMatch & ".<br />"
	   showForm = true	
	  else 
	    session.Contents("memKey") = rsKey("M_KEY")
	    showForm = false %>
        <table width="350" align="center"><tr>
        <td align="center">	
	    <form action="password.asp?mode=reset" method="post">
	    <input type="hidden" name="memId" value="<%= rsKey("MEMBER_ID") %>">
	    <input type="hidden" name="memKey" value="<%= rsKey("M_KEY") %>">
	    <%
	    spThemeTitle= txtChNewPass
	    spThemeBlock1_open(intSkin)
	    %>
       <table cellpadding="0" cellspacing="0" border="0" width="100%">
	<% if Err_Msg <> "" then %>
	<tr><td class="tCellAlt1" align="center">
	<b><% =Err_Msg %></b><hr />
	</td></tr>
	<% end if %><tr>
       <td class="tCellAlt1"><b><br /><%= txtUsrName %>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<% =rsKey("M_NAME") %><br /><br /><%= txtNewPass %>:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	   <input class="textbox" type="password" name="pass" size="25"><br /><br />
	   <%= txtRepPass %>: <input class="textbox" type="password" name="pass2" size="25">
	   <br /><br /><input type="submit" value="<%= txtSubmit %>" class="button"></b>
	   </td></tr></table>
	   <%spThemeBlock1_close(intSkin)%>
	   </form>
	   </td></tr></table>	
<%	  end if
	  rsKey.close
	  set rsKey = nothing
end if %>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->