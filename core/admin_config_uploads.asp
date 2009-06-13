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
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<%
  strUCResult = ""
  Err_Msg = ""
  intMode = 0
' Check for valid querystring
  if Request.QueryString("mode") <> "" or Request.QueryString("mode") <> " " then
	if IsNumeric(Request.QueryString("mode")) = True then
		intMode = cInt(Request.QueryString("mode"))
	end if
  end if 
  
	if Request.Form("Method_Type") = "upload_comp" then
		sComp = request.Form("upComp")
		strSQL = "UPDATE " & strTablePrefix & "CONFIG SET C_COMP_UPLOAD = '" & sComp & "'"
		executeThis(strSQL)
		Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		strUploadComp = sComp
		strUCResult = "<span class=""fTitle"">" & txtCU03 & "</span>"	
	end if
  
	if Request.Form("Method_Type") = "upload_allow" then
		intAllow = cLng(request.Form("allow"))
		strSQL = "UPDATE " & strTablePrefix & "CONFIG SET C_ALLOWUPLOADS = " & intAllow
		executeThis(strSQL)
		Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		strAllowUploads = intAllow
		strUCResult = "<span class=""fTitle"">" & txtCU03 & "</span>"	
	end if
	
	if Request.Form("Method_Type") = "upload_config" then 
		if Request.Form("upLocation") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtCU01 & "</li>"
		end if
		if Request.Form("upExt") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtCU02 & "</li>"
		end if

		if Err_Msg = "" then

			strSql = "UPDATE " & strTablePrefix & "UPLOAD_CONFIG"
			strSql = strSql & " SET UP_SIZELIMIT = " & cLng(Request.Form("upSize"))
			strSql = strSql & ", UP_ALLOWEDEXT = '" & ChkString(Request.Form("upExt"),"sqlstring") & "'"
			strSql = strSql & ", UP_LOGFILE = '" & ChkString(Request.Form("upFile"),"") & "'"
			strSql = strSql & ", UP_ACTIVE = " & cLng(Request.Form("upAllow"))
			strSql = strSql & ", UP_ALLOWEDGROUPS = '" & Request.Form("upGrps") & "'"
			strSql = strSql & ", UP_LOGUSERS = " & cLng(Request.Form("upLog"))
			
			strSql = strSql & ", UP_RESIZE = " & cLng(Request.Form("intResize"))
			strSql = strSql & ", UP_THUMB_MAX_W = " & cLng(Request.Form("strMaxTW"))
			strSql = strSql & ", UP_THUMB_MAX_H = " & cLng(Request.Form("strMaxTH"))
			strSql = strSql & ", UP_NORM_MAX_W = " & cLng(Request.Form("strMaxW"))
			strSql = strSql & ", UP_NORM_MAX_H = " & cLng(Request.Form("strMaxH"))
			strSql = strSql & ", UP_CREATE_THUMB = " & cLng(Request.Form("strDoThumb"))
			strSql = strSql & ", UP_FOLDER = '" & Request.Form("strFolder") & "'"
			
			strSql = strSql & " WHERE ID = " & cLng(Request.Form("strID"))
'			response.Write(strsql)
'			response.End()
			executeThis(strSql) 
			
			strUCResult = "<span class=""fTitle"">" & txtCU03 & "</span>"
			

		else
			strUCResult = "<span class=""fTitle"">" & txtThereIsProb & "</span>"
			strUCResult = strUCResult & "<ul>" & Err_Msg & "</ul>"
			strUCResult = strUCResult & "<a href=""JavaScript:history.go(-1)"">" & txtGoBack & "</a>"
 %>
<%		end if %>
<%	end if %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol" width="190">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtCU04 & "|admin_config_uploads.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if strUCResult <> "" then
    call showMsgBlock(1,strUCResult)
  end if

  spThemeBlock1_open(intSkin)
		select case intMode
			case 1 ' edit
				call displayForm("edit")
			case 2 ' add
				call displayForm("edit")
			case else ' list all
				call listAll()
		end select
spThemeBlock1_close(intSkin) %>
</td></tr></table>
<!--#include file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp?target=admin_config_uploads.asp" %>
<% End If

Sub listAll() %>
  <table style="width:570px" border="0" cellspacing="0" cellpadding="0" align=center>
    <tr> 
      <td class="tCellAlt2"> 
        <table border="0" cellspacing="1" cellpadding="3" width="100%">
          <tr> 
            <td class="tTitle" colspan="6"><span class="fTitle"><b><%= txtCU05 %></b></span></td>
          </tr>
          <tr valign="top"> 
            <td class="tCellAlt0" colspan="6">
  <table border="0" cellspacing="0" cellpadding="0" align=center width="100%">
          <tr> 
            <td align="right"><b><%= txtCU06 %>:</b>&nbsp;</td>
            <td>
			<form name="form2" id="form2" method="post" action="admin_config_uploads.asp">
  			<select name="allow" onchange="submit()"<% if intUploads = 0 then %> disabled="disabled"<% end if %>>
    			<option value="1"<% if strAllowUploads = 1 then Response.Write(" selected") %>><%= txtYes %></option>
    			<option value="0"<% if strAllowUploads = 0 then Response.Write(" selected") %>><%= txtNo %></option>
  			</select>
			<input type="hidden" name="Method_Type" value="upload_allow">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#AllowUploads')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
			  </form>
              </td>
          </tr>
          <tr> 
            <td align="right"><b><%= txtUpComp %>:</b>&nbsp;</td>
            <td>
			<form name="form3" id="form3" method="post" action="admin_config_uploads.asp">
  			<select name="upComp" onchange="submit()"<% if intUploads = 0 then %> disabled="disabled"<% end if %>>
    			<option value="none"<% if strUploadComp = "NONE" then Response.Write(" selected") %>><%= txtNONE %></option>
			<% If bFso Then %>
    			<option value="aspnet"<% if strUploadComp = "aspnet" then Response.Write(" selected") %>>ASP</option>
			<% End If %>
    			<!-- <option value="dundas"<% if strUploadComp = "dundas" then Response.Write(" selected") %>>Dundas</option> -->
  			</select>
			<input type="hidden" name="Method_Type" value="upload_comp">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#UploadComp')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
			  </form>
              </td>
          </tr>
  </table>
			</td>
          </tr>
    <tr align="center" valign="middle" class="tSubTitle"> 
      <td>edit</td>
      <td width="14%" height="25"><%= txtLocation %></td>
      <td width="12%" height="25"><%= txtActive %></td>
      <!-- <td width="15%" height="25"><%= txtUsrLvl %></td> -->
      <td width="24%" height="25"><%= txtCU07 %></td>
      <td width="16%" height="25"><%= txtCU08 %></td>
      <td width="19%" height="25"><%= txtCU09 %></td>
    </tr>
 <% set rsUp = my_Conn.execute("select * from " & strTablePrefix & "UPLOAD_CONFIG")
	if not rsUp.eof then
		do until rsUp.eof %>
    <tr align="center" class="tCellAlt0">
      <td><a href="admin_config_uploads.asp?mode=1&loc=<%= cLng(rsUp("ID")) %>"><%= icon(icnEdit,txtEdit,"","","") %></a></td>
      <td><%= chkString(rsUp("UP_LOCATION"),"display") %></td>
      <td><% if chkString(rsUp("UP_ACTIVE"),"display") <> 0 then Response.Write(txtYes) else Response.Write(txtNo) %></td>
      <td><%= chkString(rsUp("UP_ALLOWEDEXT"),"display") %></td>
      <td><%= chkString(rsUp("UP_SIZELIMIT"),"display") %>&nbsp;<%= txtKB %></td>
      <td><% if chkString(rsUp("UP_LOGUSERS"),"display") <> 0 then Response.Write(txtYes) else Response.Write(txtNo) %></td>
    </tr>
	<% 		rsUp.movenext
		loop 
		set rsUp = nothing %>
 <% Else %>
    <tr align="center" class="tCellAlt0"> 
      <td colspan="6"><b><%= txtCU10 %></b></td>
    </tr>
 <% End If %>
  </table></td></tr></table>
<%
End Sub 

sub displayForm(typ)
	if typ = "edit" then
	    if request.QueryString("loc") > 2 then
		  sSql = "SELECT PORTAL_UPLOAD_CONFIG.*, PORTAL_APPS.APP_GROUPS_USERS "
		  sSql = sSql & "FROM PORTAL_UPLOAD_CONFIG INNER JOIN PORTAL_APPS ON PORTAL_UPLOAD_CONFIG.UP_APPID = PORTAL_APPS.APP_ID "
		  sSql = sSql & "WHERE (((PORTAL_UPLOAD_CONFIG.ID)=" & cLng(request.QueryString("loc")) & "));"
		else
		  sSql = "SELECT PORTAL_UPLOAD_CONFIG.* FROM PORTAL_UPLOAD_CONFIG "
		  sSql = sSql & "WHERE (((PORTAL_UPLOAD_CONFIG.ID)=" & cLng(request.QueryString("loc")) & "));"
		end if
		set rsUp = my_Conn.execute(sSql)
		if not rsUp.eof then
			showResize = false
			strID = cLng(rsUp("ID"))

			strLoc = chkString(rsUp("UP_LOCATION"),"display")
			strActive = cLng(rsUp("UP_ACTIVE"))
			strExt = chkString(rsUp("UP_ALLOWEDEXT"),"display")
			  strGrps = rsUp("UP_ALLOWEDGROUPS")
			if request.QueryString("loc") > 2 then
			strGrpsW = rsUp("APP_GROUPS_USERS")
			end if
			strFile = chkString(rsUp("UP_LOGFILE"),"display")
			strLog = cLng(rsUp("UP_LOGUSERS"))
			strSize = cLng(rsUp("UP_SIZELIMIT"))
			
			strAppID = cLng(rsUp("UP_APPID"))
			strMaxTW = cLng(rsUp("UP_THUMB_MAX_W"))
			strMaxTH = cLng(rsUp("UP_THUMB_MAX_H"))
			strMaxW = cLng(rsUp("UP_NORM_MAX_W"))
			strMaxH = cLng(rsUp("UP_NORM_MAX_H"))
			intResize = cLng(rsUp("UP_RESIZE"))
			strDoThumb = cLng(rsUp("UP_CREATE_THUMB"))
			strFolder = rsUp("UP_FOLDER")
			if instr(strExt,"gif") > 0 or instr(strExt,"jpg") > 0 or instr(strExt,"png") > 0 then
			  select case strImgComp
			    case "aspnet"
			  	  det = checkForDotNet("includes/scripts/checkfordotnet.aspx")
			  	  if det <> "" then
			    	showResize = true
			  	  end if
				case "aspjpeg"
			      showResize = true
			  end select
			end if
		  set rsUp = nothing
		else
			closeAndGo("admin_config_uploads.asp?error=eof")
		end if
	end if %>
<form action="admin_config_uploads.asp" method="post" id="Form" name="Form">
<input type="hidden" name="Method_Type" value="upload_config">
<input type="hidden" name="upLocation" value="<%= strLoc %>">
<input type="hidden" name="strAppID" value="<%= strAppID %>">
<input type="hidden" name="strID" value="<%= strID %>">
  <table style="width:570px" border="0" cellspacing="0" class="tCellAlt2" cellpadding="0" align=center>
    <tr> 
      <td> 
        <table border="0" cellspacing="1" cellpadding="1" width="100%">
          <tr valign="middle"> 
            <td class="tTitle" colspan="2">
			<span class="fTitle"><%= txtCU11 %></span></td>
          </tr>
          <tr valign="middle"> 
            <td class="tSubTitle" align="center" colspan="2">
			<span class="fSubTitle"><%= ucase(strLoc) %></span>
            </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU06 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="upAllow" value="1" <% if strActive = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="upAllow" value="0" <% if strActive = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#uallow')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right" valign="top" class="tCellAlt0"><b><%= txtCU14 %>:&nbsp; </b><br /><br />
			<a href="JavaScript:allowgroups('Form','upGrps','<%= strGrpsW %>');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a><br />
				<a href="JavaScript:removeGroup('Form','upGrps');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>
			</td>
            <td class="tCellAlt0"> 
            <select size="5" name="upGrps" style="width:170;" multiple>
			  <% if strGrps <> "" then
			  		getGroups(strGrps)
				 end if %>
			  <option value="0"></option>
            </select>
              </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU12 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="upSize" size="25" value="<% if strSize <> "" then Response.Write(strSize) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#usize')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU07 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="upExt" size="25" value="<% if strExt <> "" then Response.Write(strExt) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#uextentions')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU13 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="upFile" size="25" value="<% if strFile <> "" then Response.Write(strFile) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#ulogfile')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU09 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="upLog" value="1" <% if strLog = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="upLog" value="0" <% if strLog = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#ulog')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <!-- <tr valign="middle"> 
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU14 %>:&nbsp; </b></td>
            <td class="tCellAlt0"> 
              <select name="upLvl" id="upLvl">
                <option value="0"<% if strLvl = "0" then Response.Write(" selected") %>> 
                <%= txtEveryone %> </option>
                <option value="1"<% if strLvl = "1" then Response.Write(" selected") %>> 
                <%= txtAllMem %> </option>
                <option value="3"<% if strLvl = "3" then Response.Write(" selected") %>> 
                <%= txtModAdmin %> </option>
                <option value="4"<% if strLvl = "4" then Response.Write(" selected") %>> 
                <%= txtAdmin %> </option>
                <option value="5"<% if strLvl = "5" then Response.Write(" selected") %>> 
                <%= txtSuprAdmin %> </option>
              </select>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#uwho')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
            </td>
          </tr> -->
          <tr valign="middle"> 
            <td class="tCellAlt0" align="right"><b><%= txtCU15 %>:</b>&nbsp;</td>
            <td class="tCellAlt0"> 
              <input type="text" class="textbox" name="strFolder" size="25" value="<% if strFolder <> "" then Response.Write(strFolder) %>">
              </td>
          </tr>
			<% if showResize then %>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU16 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <%= txtYes %>: 
              <input type="radio" class="radio" name="intResize" value="1" <% if intResize = 1 then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="intResize" value="0" <% if intResize = 0 then Response.Write("checked") %>>
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU17 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <input type="text" class="textbox" name="strMaxH" size="25" value="<% if strMaxH <> "" then Response.Write(strMaxH) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU18 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <input type="text" class="textbox" name="strMaxW" size="25" value="<% if strMaxW <> "" then Response.Write(strMaxW) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU19 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <%= txtYes %>: 
              <input type="radio" class="radio" name="strDoThumb" value="1" <% if strDoThumb = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="strDoThumb" value="0" <% if strDoThumb = "0" then Response.Write("checked") %>>
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU20 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <input type="text" class="textbox" name="strMaxTH" size="25" value="<% if strMaxTH <> "" then Response.Write(strMaxTH) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0"><b><%= txtCU21 %>:</b>&nbsp;</td>
            <td class="tCellAlt0">
              <input type="text" class="textbox" name="strMaxTW" size="25" value="<% if strMaxTW <> "" then Response.Write(strMaxTW) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0">&nbsp;</td>
            <td class="tCellAlt0">&nbsp;</td>
          </tr>
			<% else %>
          <tr valign="top">
            <td align="right" valign="middle" class="tCellAlt0">&nbsp;</td>
            <td class="tCellAlt0">&nbsp;
			<input type="hidden" name="intResize" value="0">
			<input type="hidden" name="strDoThumb" value="0">
			<input type="hidden" name="strMaxH" value="0">
			<input type="hidden" name="strMaxW" value="0">
			<input type="hidden" name="strMaxTH" value="0">
			<input type="hidden" name="strMaxTW" value="0">
			</td>
          </tr>
			<% end if %>
          <tr valign="middle"> 
            <td class="tCellAlt0" colspan="2" align="center"> 
              <input type="submit" value="<%= txtSubmit %>" onclick="selectAll('Form','upGrps');" id="submit1" name="submit1" class="button">
              <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form><br />
<% End Sub

sub getGroups(gid)
  if gid <> "" then
	arrTemp = split(gid,",")
	for xp = 0 to ubound(arrTemp)
	  sSQL = "select * from " & strTablePrefix & "GROUPS where G_ID = " & arrTemp(xp) & " ORDER BY G_ACTIVE, G_NAME"
	  'response.Write(sSQL & "<br />")
	  set rsGrp = my_Conn.execute(sSQL)
	  if not rsGrp.eof then
	    'do until rsGrp.eof
		  if lcase(rsGrp("G_INAME")) = "administrator" then
		    'rsGrp.movenext
	        response.Write("<option value=""" & rsGrp("G_ID") & """>" & rsGrp("G_NAME") & "</option>" & vbnewline)
		  else
	        response.Write("<option value=""" & rsGrp("G_ID") & """>" & rsGrp("G_NAME") & "</option>" & vbnewline)
		  end if
	      'rsGrp.movenext
	    'loop
	  end if
	next
	set rsGrp = nothing
  end if
end sub %>