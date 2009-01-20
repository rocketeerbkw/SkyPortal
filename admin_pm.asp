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
pgType = "manager"

  modPgType = "addForm"
  uploadPg = false
  hasEditor = true
  strEditorElements = "automessage"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then
Err_Msg = "" %>
<!--#include file="includes/inc_admin_functions.asp" -->
<%
if Request.Form("Method_Type") = "newUserPM" then 
		Err_Msg = ""
		if (Request.Form("autopmonoff") = "1" and autopmonoff = "1") or (Request.Form("autopmonoff") = "1" and autopmonoff = "0") then
			if Request.Form("autosubjectline") = "" then 
				Err_Msg = Err_Msg & "<li>" & txtNoSubLine & "</li>"
			end if
		end if

	if Err_Msg = "" then
		
			'
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET AUTOPM_ON = " & Request.Form("autopmonoff") & ""
			if (Request.Form("autopmonoff") = "1") then
				strSql = strSql & ",    AUTOPM_SUBJECTLINE = '" & ChkString(Request.Form("autosubjectline"),"SQLString") & "'"
				strSql = strSql & ",    AUTOPM_MESSAGE = '" & ChkString(Request.Form("automessage"),"message") & "'"
			end if
			strSql = strSql & " WHERE CONFIG_ID = 1"
			
			executeThis(strSql)
			
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		
		Err_Msg = "<li><span class=""fTitle"">"
		Err_Msg = Err_Msg & txtMemPMMsgUpd & "</span></li>"
%>
<%	else
		Err_Msg1 = "<li><span class=""fTitle"">"
		Err_Msg1 = Err_Msg1 & txtThereIsProb & "</span></li>"
		Err_Msg = Err_Msg1 & Err_Msg
	 %>

<%	end if
end if

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Err_Msg = "" then

			'## Delete private messages
			strMessageDays = DateToStr(dateAdd("d",-Request.Form("strMessageDays"),now()))
			strSql = "DELETE FROM " & strTablePrefix & "PM "
			strSql = strSql & "WHERE M_SENT < '" & strMessageDays & "'"
			if Request.Form("UnRead") = "yes" then
				strSql = strSql & " AND M_READ = 0"
			else
				strSql = strSql & " AND M_READ = 1"
			end if

			executeThis(strSql)
			Err_Msg = Err_Msg & "<li><span class=""fTitle"">" & txtMsgsDeleted & "</span></li>"
	else
			Err_Msg1 = "<li><span class=""fTitle"">" & txtThereIsProb & "</span></li>"
			Err_Msg = Err_Msg1 & Err_Msg
	end if	
end if

	strMessage15Days = DateToStr(dateAdd("d",-split(arrPMadmin1(1),",")(2),now()))
	strMessage30Days = DateToStr(dateAdd("d",-split(arrPMadmin1(2),",")(2),now()))
	strMessage45Days = DateToStr(dateAdd("d",-split(arrPMadmin1(3),",")(2),now()))
	strMessage60Days = DateToStr(dateAdd("d",-split(arrPMadmin1(4),",")(2),now()))

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_READ = 1"

	Set rsPM = my_Conn.Execute(strSql)
	pmcountr = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_READ = 0"

	Set rsPM = my_Conn.Execute(strSql)
	pmcountu = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage15Days & "' AND M_READ = 1"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount15r = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage15Days & "' AND M_READ = 0"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount15u = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage30Days & "' AND M_READ = 1"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount30r = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage30Days & "' AND M_READ = 0"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount30u = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage45Days & "' AND M_READ = 1"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount45r = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage45Days & "' AND M_READ = 0"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount45u = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage60Days & "' AND M_READ = 1"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount60r = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing

	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(M_TO) as pmcount " 
    	end if
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE M_SENT < '" & strMessage60Days & "' AND M_READ = 0"

	Set rsPM = my_Conn.Execute(strSql)
	pmcount60u = rsPM("pmcount")

	rsPM.close
	set rsPM = nothing
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol" valign="top">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	pmConfigMenu("1") %>
  <hr />
	<%
	menu_admin()
	%>
	<% spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtPvtMsgMngr & "|admin_pm.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<% 
	if Err_Msg <> "" then
	spThemeBlock1_open(intSkin) %>
	<table align="center" border="0">
	  <tr>
	    <td>
		<ul>
		<%= Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>
	<% spThemeBlock1_close(intSkin)
	end if %>
	
<table border="0" cellspacing="0" cellpadding="0" align=center width="100%">
  <tr>
    <td>
	<DIV id="paa" style="display:block;">
	<% spThemeBlock1_open(intSkin) %>
      <table border="0" class="grid" cellspacing="1" cellpadding="4" width="100%" align="center">
	<tr valign="top">
	  <td class="tTitle" align="center" colspan="3"><span class="fTitle"><b><%= txtPvtMsgCnts %></b></span></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><b><%= TxtTimePrd %></b></td>
	  <td class="tCellAlt1" align="center"><b><%= txtRead %></b></td>
	  <td class="tCellAlt1" align="center"><b><%= txtUnread %></b></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><%= split(arrPMadmin1(0),",")(0) %></td>
	  <td class="tCellAlt1" align="center"><% =pmcountr %></td>
	  <td class="tCellAlt1" align="center"><% =pmcountu %></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><%= split(arrPMadmin1(1),",")(0) %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount15r %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount15u %></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><%= split(arrPMadmin1(2),",")(0) %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount30r %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount30u %></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><%= split(arrPMadmin1(3),",")(0) %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount45r %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount45u %></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt1" align="center"><%= split(arrPMadmin1(4),",")(0) %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount60r %></td>
	  <td class="tCellAlt1" align="center"><% =pmcount60u %></td>
	</tr>
    </table>
	<% spThemeBlock1_close(intSkin) %>
	</div>
	<DIV id="pbb" style="display:block;">
	<% spThemeBlock1_open(intSkin) %>
<form action="admin_pm.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td align="center">
      <table border="0" class="grid" cellspacing="1" cellpadding="4">
	<tr valign="top">
	  <td class="tTitle" align="center" colspan="2"><span class="fTitle"><b>Private Messages Maintenance</b></span></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt0" align="right"><b>Delete Private Messages Older than:</b>&nbsp;</td>
	  <td class="tCellAlt0">
	  <select name="strMessageDays" size=1>
	  	<option value="<%= split(arrPMadmin1(0),",")(2) %>"><%= split(arrPMadmin1(0),",")(1) %></option>
	  	<option value="<%= split(arrPMadmin1(1),",")(2) %>"><%= split(arrPMadmin1(1),",")(1) %></option>
	  	<option value="<%= split(arrPMadmin1(2),",")(2) %>"><%= split(arrPMadmin1(2),",")(1) %></option>
	  	<option value="<%= split(arrPMadmin1(3),",")(2) %>" selected="selected"><%= split(arrPMadmin1(3),",")(1) %></option>
	  	<option value="<%= split(arrPMadmin1(4),",")(2) %>"><%= split(arrPMadmin1(4),",")(1) %></option>
	  </select></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt0" colspan="2">
          <input name="UnRead" type="checkbox" value="yes"><%= txtChkDwlUnred %><br /></td>
	</tr>
	<tr valign="top">
	  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" class="button"></td>
	</tr>
      </table>
    </td>
  </tr>
</table>
</form>
	<% spThemeBlock1_close(intSkin) %>
	</div>
	<DIV id="pcc" style="display:none;">
	<%

	strSql = "SELECT AUTOPM_ON, AUTOPM_SUBJECTLINE, AUTOPM_MESSAGE FROM "
	strSql = strSql & strTablePrefix & "CONFIG"
	strSql = strSql & " WHERE CONFIG_ID = 1"
	
	Set rs = my_Conn.Execute (strSql)

autopmonoff = rs("AUTOPM_ON")
autosubjectline = rs("AUTOPM_SUBJECTLINE")
automessage = rs("AUTOPM_MESSAGE")
%>
	<% spThemeBlock1_open(intSkin) %>
<form action="admin_pm.asp" method="post" id="FormA" name="FormA">
<input type="hidden" name="Method_Type" value="newUserPM">
        <table border="0" cellspacing="0" cellpadding="3" align="center" class="tBorder">
	<tr valign="top">
	  <td class="tTitle" align="center" colspan="3"><span class="fTitle"><b><%= txtPMNewMemConf %></b></span></td>
	</tr>
          <tr valign="top"> 
            <td align="right" width="200"><b><%= txtPmAutoNewMem %>:</b>&nbsp;</td>
            <td align="left"> 
              <input type="radio" class="radio" name="autopmonoff" value="1" <% if autopmonoff <> 0 then Response.Write("checked")%>><%= txtYES %>&nbsp;(<%= txtOn %>)&nbsp;&nbsp;&nbsp;<input type="radio" class="radio" name="autopmonoff" value="0" <% if autopmonoff = 0 then Response.Write("checked")%>><%= txtNO %>&nbsp;(<%= txtOff %>)</td>
            <td align="left"> 
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right"><b><%= txtSubject %>:</b>&nbsp; 
              </td>
            <td colspan="2"> 
              <input class="textbox" type="text" name="autosubjectline" size="20" value="<%= autosubjectline %>">
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right"><b><%= txtMsg %>:&nbsp;</b> 
              </td>
            <td colspan="2"> 
              <textarea name="automessage" cols="30" rows="8" style="width:100%;"><%=automessage%></textarea>
            </td>
          </tr>
          <tr valign="top"> 
            <td colspan="3" height="6" align="right"></td>
          </tr>
          <tr valign="top"> 
            <td colspan="3" align="center"> 
              <input class="button" type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1">
              &nbsp;&nbsp; 
              <input class="button" type="reset" value="<%= txtReset %>" id="reset1" name="reset1">
            </td>
          </tr>
        </table>
		</form>
	<% spThemeBlock1_close(intSkin) %>
	</div>
    </td>
  </tr>
</table>
</td></tr></table>
<!--#include file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp?target=admin_pm.asp" %>
<% End IF %>