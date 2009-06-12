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
dim grpRead,grpWrite,grpFull,upl
upl = 0
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<%	
  Err_Msg = ""
  strCMResult = ""
  pgMode = cInt(request.QueryString("mode"))
	if Request.Form("Method_Type") = "modify_config" then 
		sAdmin = false
		app_id = cInt(Request.Form("APP_ID"))
		app_active = cInt(Request.Form("APP_ACTIVE"))
		
		app_users = chkGrpAdmin(Request.Form("g_read"))
		app_write = chkGrpAdmin(Request.Form("g_write"))
		app_full = chkGrpAdmin(Request.Form("g_full"))

		if Err_Msg = "" then

			strSql = "UPDATE " & strTablePrefix & "APPS"
			strSql = strSql & " SET APP_ACTIVE = " & app_active
			strSql = strSql & ", APP_GROUPS_USERS = '" & app_users & "'"
			strSql = strSql & ", APP_GROUPS_WRITE = '" & app_write & "'"
			strSql = strSql & ", APP_GROUPS_FULL = '" & app_full & "'"
			if trim(request.Form("APP_SUBSCRIPTIONS")) <> "" then
			  strSql = strSql & ", APP_SUBSCRIPTIONS = " & cLng(request.Form("APP_SUBSCRIPTIONS"))
			end if
			if trim(request.Form("APP_BOOKMARKS")) <> "" then
			  strSql = strSql & ", APP_BOOKMARKS = " & cLng(request.Form("APP_BOOKMARKS"))
			end if
			if trim(request.Form("APP_SUBSEC")) = "1" then
			  strSql = strSql & ", APP_SUBSEC = 1"
			else
			  strSql = strSql & ", APP_SUBSEC = 0"
			end if
			if trim(request.Form("iDATA1")) <> "" then
			  strSql = strSql & ", APP_iDATA1 = " & cLng(request.Form("iDATA1"))
			end if
			if trim(request.Form("iDATA2")) <> "" then
			  strSql = strSql & ", APP_iDATA2 = " & cLng(request.Form("iDATA2"))
			end if
			if trim(request.Form("iDATA3")) <> "" then
			  strSql = strSql & ", APP_iDATA3 = " & cLng(request.Form("iDATA3"))
			end if
			if trim(request.Form("iDATA4")) <> "" then
			  strSql = strSql & ", APP_iDATA4 = " & cLng(request.Form("iDATA4"))
			end if
			if trim(request.Form("iDATA5")) <> "" then
			  strSql = strSql & ", APP_iDATA5 = " & cLng(request.Form("iDATA5"))
			end if
			if trim(request.Form("iDATA6")) <> "" then
			  strSql = strSql & ", APP_iDATA6 = " & cLng(request.Form("iDATA6"))
			end if
			if trim(request.Form("iDATA7")) <> "" then
			  strSql = strSql & ", APP_iDATA7 = " & cLng(request.Form("iDATA7"))
			end if
			if trim(request.Form("iDATA8")) <> "" then
			  strSql = strSql & ", APP_iDATA8 = " & cLng(request.Form("iDATA8"))
			end if
			if trim(request.Form("iDATA9")) <> "" then
			  strSql = strSql & ", APP_iDATA9 = " & cLng(request.Form("iDATA9"))
			end if
			if trim(request.Form("iDATA10")) <> "" then
			  strSql = strSql & ", APP_iDATA10 = " & cLng(request.Form("iDATA10"))
			end if
			if trim(request.Form("tDATA1")) <> "" then
			  strSql = strSql & ", APP_tDATA1 = '" & request.Form("tDATA1") & "'"
			end if
			if trim(request.Form("tDATA2")) <> "" then
			  strSql = strSql & ", APP_tDATA2 = '" & request.Form("tDATA2") & "'"
			end if
			if trim(request.Form("tDATA3")) <> "" then
			  strSql = strSql & ", APP_tDATA3 = '" & request.Form("tDATA3") & "'"
			end if
			if trim(request.Form("tDATA4")) <> "" then
			  strSql = strSql & ", APP_tDATA4 = '" & request.Form("tDATA4") & "'"
			end if
			if trim(request.Form("tDATA5")) <> "" then
			  strSql = strSql & ", APP_tDATA5 = '" & request.Form("tDATA5") & "'"
			end if
			strSql = strSql & " WHERE APP_ID = " & app_id
			executeThis(strSql)

			':: Update the PORTAL_UPLOAD_CONFIG table
		  if Request.Form("hasUpload") = "1" then
			strSql = "UPDATE " & strTablePrefix & "UPLOAD_CONFIG"
			strSql = strSql & " SET UP_SIZELIMIT = " & cLng(Request.Form("upSize"))
			strSql = strSql & ", UP_ALLOWEDEXT = '" & ChkString(Request.Form("upExt"),"sqlstring") & "'"
			strSql = strSql & ", UP_LOGFILE = '" & ChkString(Request.Form("upFile"),"") & "'"
			strSql = strSql & ", UP_ACTIVE = " & cLng(Request.Form("upAllow"))
			strSql = strSql & ", UP_ALLOWEDGROUPS = '" & ChkString(chkGrpAdmin(Request.Form("upGrp")),"sqlstring") & "'"
			strSql = strSql & ", UP_LOGUSERS = " & cLng(Request.Form("upLog"))
			
			strSql = strSql & ", UP_RESIZE = " & cLng(Request.Form("intResize"))
			strSql = strSql & ", UP_THUMB_MAX_W = " & cLng(Request.Form("strMaxTW"))
			strSql = strSql & ", UP_THUMB_MAX_H = " & cLng(Request.Form("strMaxTH"))
			strSql = strSql & ", UP_NORM_MAX_W = " & cLng(Request.Form("strMaxW"))
			strSql = strSql & ", UP_NORM_MAX_H = " & cLng(Request.Form("strMaxH"))
			strSql = strSql & ", UP_CREATE_THUMB = " & cLng(Request.Form("strDoThumb"))
			strSql = strSql & ", UP_FOLDER = '" & Request.Form("strFolder") & "'"
			
			strSql = strSql & " WHERE UP_APPID = " & app_id
			'response.Write(strsql)
			'response.End()
			executeThis(strSql)
		  end if
		  if trim(Request.Form("hasBlockedUsers")) = "1" then
			sSQL = "UPDATE " & strTablePrefix & "MEMBERS SET M_PMSTATUS=1 WHERE M_PMSTATUS=0"
			executeThis(sSQL)
		    if trim(Request.Form("BlockedUsers")) <> "0" then
		      arrBanUsers = split(Request.Form("BlockedUsers"),",")
			  for xi = 0 to ubound(arrBanUsers)
			    sSQL = "UPDATE " & strTablePrefix & "MEMBERS SET M_PMSTATUS=0 WHERE MEMBER_ID=" & arrBanUsers(xi) & " AND M_LEVEL < 3"
			    executeThis(sSQL)
			  next
			end if
		  end if
			
			strCMResult = "<span class=""fTitle"">" & txtCM01 & "</span>"
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		   %>
<%		else 
			strCMResult = "<span class=""fTitle"">" & txtThereIsProb & "</span>"
			strCMResult = strCMResult & "<ul>" & Err_Msg & "</ul>"
			strCMResult = strCMResult & "<a href=""JavaScript:history.go(-1)"">" & txtGoBack & "</a>"

		end if %>
<%	end if %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol">
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
  arg2 = txtCM02 & "|admin_config_modules.asp"
  select case pgMode
	case 1 ' edit
      arg3 = txtEdit & "|javascript:;"
	case 2 ' add
      arg3 = txtAdd & "|javascript:;"
	case else ' list all
      arg3 = ""
  end select
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if strCMResult <> "" then
    call showMsgBlock(1,strCMResult)
  end if

spThemeBlock1_open(intSkin)
		select case pgMode
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
<% Response.Redirect "admin_login.asp?target=admin_config_modules.asp" %>
<% End If

Sub listAll() %>
  <table border="0" style="width:500px;" cellspacing="0" cellpadding="0" class="tCellAlt1" align="center">
    <tr> 
      <td> 
        <table border="0" cellspacing="1" cellpadding="3" width="100%">
          <tr valign="top"> 
            <td class="tTitle" colspan="5"><%= txtCM02 %></td>
          </tr>
    <tr align="center" valign="middle" class="tSubTitle"> 
      <td width="60">&nbsp;</td>
      <td height="25"><%= txtCM04 %></td>
      <td width="75" height="25"><%= txtActive %></td>
      <td width="90" height="25"><%= txtUploads %></td>
      <td width="90" height="25"><%= txtVersionInfo %></td>
    </tr>
 <% 
 	sSQL = "SELECT PORTAL_APPS.*, PORTAL_UPLOAD_CONFIG.UP_ACTIVE FROM PORTAL_APPS LEFT JOIN PORTAL_UPLOAD_CONFIG ON PORTAL_APPS.APP_ID = PORTAL_UPLOAD_CONFIG.UP_APPID"
	set rsUp = my_Conn.execute(sSQL)
	if not rsUp.eof then
		do until rsUp.eof
		  hasUploads = "--"
		  if rsUp("UP_ACTIVE") = 1 then
		    hasUploads = "On"
		  elseif rsUp("UP_ACTIVE") = 0 then
		    hasUploads = "Off"
		  end if %>
    <tr align="center">
      <td><a href="admin_config_modules.asp?mode=1&id=<%= cInt(rsUp("APP_ID")) %>"><%= icon(icnEdit,txtEdit,"","","") %></a></td>
      <td><%= chkString(rsUp("APP_NAME"),"display") %></td>
      <td><% if cInt(rsUp("APP_ACTIVE")) = 1 then Response.Write("Yes") else Response.Write("No") %></td>
      <td><%= hasUploads %></td>
      <td>v
	  <% on error resume next
	   response.Write(rsUp("APP_VERSION"))
	   on error goto 0 %></td>
    </tr>
	<% 		rsUp.movenext
		loop 
		set rsUp = nothing %>
 <% Else %>
    <tr align="center"> 
      <td colspan="5"><b><%= txtCM03 %></b></td>
    </tr>
 <% End If %>
  </table></td></tr></table>
<%
End Sub 

sub displayForm(typ)
	if typ = "edit" then
 	  sSQL = "SELECT * FROM PORTAL_APPS WHERE APP_ID = " & cInt(request.QueryString("id"))
	  set rsUp = my_Conn.execute(sSQL)
		if not rsUp.eof then
			strID = cInt(rsUp("APP_ID"))
			strName = chkString(rsUp("APP_NAME"),"display")
			strActive = cInt(rsUp("APP_ACTIVE"))
			grpRead = rsUp("APP_GROUPS_USERS")
			grpWrite = rsUp("APP_GROUPS_WRITE")
			grpFull = rsUp("APP_GROUPS_FULL")
			iSubscriptions = rsUp("APP_SUBSCRIPTIONS")
			iBookMarks = rsUp("APP_BOOKMARKS")
			iSecImg = rsUp("APP_SUBSEC")
			strFunction = rsUp("APP_CONFIG")
			tDATA1 = rsUp("APP_tDATA1")
			iDATA1 = rsUp("APP_iDATA1")
			tDATA2 = rsUp("APP_tDATA2")
			iDATA2 = rsUp("APP_iDATA2")
			tDATA3 = rsUp("APP_tDATA3")
			iDATA3 = rsUp("APP_iDATA3")
			tDATA4 = rsUp("APP_tDATA4")
			iDATA4 = rsUp("APP_iDATA4")
			tDATA5 = rsUp("APP_tDATA5")
			iDATA5 = rsUp("APP_iDATA5")
			'tDATA6 = rsUp("APP_tDATA6")
			iDATA6 = rsUp("APP_iDATA6")
			'tDATA7 = rsUp("APP_tDATA7")
			iDATA7 = rsUp("APP_iDATA7")
			'tDATA8 = rsUp("APP_tDATA8")
			iDATA8 = rsUp("APP_iDATA8")
			'tDATA9 = rsUp("APP_tDATA9")
			iDATA9 = rsUp("APP_iDATA9")
			'tDATA10 = rsUp("APP_tDATA10")
			iDATA10 = rsUp("APP_iDATA10")
		    'set rsUp = nothing
			showUpload = false
 	  		sSQL = "SELECT * FROM PORTAL_UPLOAD_CONFIG WHERE UP_APPID = " & strID
		    set rsUpl = my_Conn.execute(sSQL)
		    if not rsUpl.eof then
			  showUpload = true
			  showResize = false
			  strULoc = chkString(rsUpl("UP_LOCATION"),"display")
			  strUActive = cLng(rsUpl("UP_ACTIVE"))
			  strExt = chkString(rsUpl("UP_ALLOWEDEXT"),"display")
			  strUpGrps = rsUpl("UP_ALLOWEDGROUPS")
			  strFile = chkString(rsUpl("UP_LOGFILE"),"display")
			  strLog = cLng(rsUpl("UP_LOGUSERS"))
			  strSize = cLng(rsUpl("UP_SIZELIMIT"))
			  strMaxTW = cLng(rsUpl("UP_THUMB_MAX_W"))
			  strMaxTH = cLng(rsUpl("UP_THUMB_MAX_H"))
			  strMaxW = cLng(rsUpl("UP_NORM_MAX_W"))
			  strMaxH = cLng(rsUpl("UP_NORM_MAX_H"))
			  intResize = cLng(rsUpl("UP_RESIZE"))
			  strDoThumb = cLng(rsUpl("UP_CREATE_THUMB"))
			  strFolder = rsUpl("UP_FOLDER")
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
			end if
		    set rsUpl = nothing
		else
		    set rsUpl = nothing
		    set rsUp = nothing
			closeAndGo("admin_config_Modules.asp?error=eof")
		end if
	end if %>
<script type="text/javascript">

function selectModUsers(frm,up)
{ 
	//alert(up);
	selectUsers(frm);
	if (up == 1) {
	//alert(up);
	 selectAll(frm,'upGrp')
	 //return;
	}
}

function DeleteSelection()
{
	var user,mText;
	var count,finished;

		finished = false;
		count = 0;
		count = document.PostTopic.AuthUsers.length - 1;
		if (count<1) {
			return;
		}
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
}

function allowmembersX(fm,ob) { 
  var whereto = "pop_portal.asp?cmd=5&mode=1&frm=" + fm + "&sel=" + ob;
  var MainWindow = window.open (whereto, "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=330,top=100,left=100,status=no"); }
//-->
</script>
<form action="admin_config_modules.asp" method="post" id="PostTopic" name="PostTopic">
<input type="hidden" name="Method_Type" value="modify_config">
<input type="hidden" name="APP_ID" value="<%= strID %>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" style="width:500px;">
    <tr> 
      <td class="tCellAlt2" width="500" align="center"> 
        <table border="0" cellspacing="1" cellpadding="1" class="tCellAlt1" width="500">
          <tr valign="middle"> 
            <td class="tTitle" colspan="2">
			<%= txtCM05 %></td>
          </tr>
          <tr valign="middle"> 
            <td class="tSubTitle" align="center" colspan="2">
			<%= ucase(strName) %>
            </td>
          </tr>
          <tr valign="top"> 
            <td align="right" height="25">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#usize')"><%= icon(icnHelp,txtHelp,"","","") %></a>&nbsp;&nbsp;<b><%= txtActive %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="APP_ACTIVE" value="1"<% if strActive = 1 then Response.Write(" checked") %> />
              <%= txtNo %>: 
              <input type="radio" class="radio" name="APP_ACTIVE" value="0"<% if strActive = 0  then Response.Write(" checked") %> />
	<% If iBookmarks <> 3 and intBookmarks = 1 Then %>
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right" height="25">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#usize')"><%= icon(icnHelp,txtHelp,"","","") %></a>&nbsp;&nbsp;<b><%= txtCM06 %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="APP_BOOKMARKS" value="1"<% if iBookmarks = 1 then Response.Write(" checked") %> />
              <%= txtNo %>: 
              <input type="radio" class="radio" name="APP_BOOKMARKS" value="0"<% if iBookmarks = 0  then Response.Write(" checked") %> />
	<% ElseIf iBookmarks = 3 then %>
		  <input type="hidden" name="APP_BOOKMARKS" value="3" />
	<% Else %>
		  <input type="hidden" name="APP_BOOKMARKS" value="0" />
	<% End If %>
	<% If iSubscriptions <> 3 and intSubscriptions = 1 Then %>
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right" height="25">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#usize')"><%= icon(icnHelp,txtHelp,"","","") %></a>&nbsp;&nbsp;<b><%= txtCM07 %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="APP_SUBSCRIPTIONS" value="1"<% if iSubscriptions = 1 then Response.Write(" checked") %> />
              <%= txtNo %>: 
              <input type="radio" class="radio" name="APP_SUBSCRIPTIONS" value="0"<% if iSubscriptions = 0  then Response.Write(" checked") %> />
	<% Else %>
		  <input type="hidden" name="APP_SUBSCRIPTIONS" value="3" />
	<% End If %>
              </td>
          </tr>
          <tr valign="top"> 
            <td align="right" height="25">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#secimage')"><%= icon(icnHelp,txtHelp,"","","") %></a>&nbsp;&nbsp;<b><%= txtSecImg %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="APP_SUBSEC" value="1"<% if iSecImg = 1 then Response.Write(" checked") %> />
              <%= txtNo %>: 
              <input type="radio" class="radio" name="APP_SUBSEC" value="0"<% if iSecImg = 0  then Response.Write(" checked") %> />
              </td>
          </tr>
		  
          <tr valign="middle"> 
            <td class="tSubTitle" align="center" colspan="2"><%= txtCM08 %></td>
          </tr>
        <tr><td align="center" valign="middle" colspan="2">
		<% 
		  Call shoGroupAccess("PostTopic",grpRead,grpWrite,grpFull,"")
		  %></td></tr><%
  		  if strFunction <> "" then
    		execute("call " & strFunction)
  		  end if %>
		<% if showUpload and bFso then
			 upl = 1 %>
          <tr valign="middle"> 
            <td class="tSubTitle" align="center" colspan="2"><span class="fSubTitle"><%= txtCM14 %></span></td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU06 %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="upAllow" value="1" <% if strUActive = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="upAllow" value="0" <% if strUActive = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#uallow')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU12 %>:</b>&nbsp;</td>
            <td> 
              <input type="text" class="textbox" name="upSize" size="25" value="<% if strSize <> "" then Response.Write(strSize) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#usize')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU07 %>:</b>&nbsp;</td>
            <td> 
              <input type="text" class="textbox" name="upExt" size="25" value="<% if strExt <> "" then Response.Write(strExt) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#uextentions')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
            </td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU15 %>:</b>&nbsp;</td>
            <td> 
              <input type="text" class="textbox" name="strFolder" size="25" value="<% if strFolder <> "" then Response.Write(strFolder) %>">
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU13 %>:</b>&nbsp;</td>
            <td> 
              <input type="text" class="textbox" name="upFile" size="25" value="<% if strFile <> "" then Response.Write(strFile) %>">
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#ulogfile')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right"><b><%= txtCU09 %>:</b>&nbsp;</td>
            <td> 
              <%= txtYes %>: 
              <input type="radio" class="radio" name="upLog" value="1" <% if strLog = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="upLog" value="0" <% if strLog = "0" then Response.Write("checked") %>>
              <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#ulog')"><%= icon(icnHelp,txtHelp,"","","") %></a> 
              </td>
          </tr>
          <tr valign="middle"> 
            <td align="right" valign="top"><b><%= txtCU14 %>:&nbsp; </b><br /><br />
			<a href="JavaScript:allowgroups('PostTopic','upGrp','0');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a><br />
				<a href="JavaScript:removeGroup('PostTopic','upGrp');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>
			</td>
            <td> 
            <select size="5" name="upGrp" id="upGrp" style="width:170;" multiple>
			  <% if cInt(request.QueryString("mode")) = 1 and strUpGrps <> "" then
			  		getGroups(strUpGrps)
				 end if %>
			  <option value="0"></option>
            </select>
            </td>
          </tr>
			<% if showResize then %>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU16 %>:</b>&nbsp;</td>
            <td>
              <%= txtYes %>: 
              <input type="radio" class="radio" name="intResize" value="1" <% if intResize = 1 then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="intResize" value="0" <% if intResize = 0 then Response.Write("checked") %>>
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU17 %>:</b>&nbsp;</td>
            <td>
              <input type="text" class="textbox" name="strMaxH" size="25" value="<% if strMaxH <> "" then Response.Write(strMaxH) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU18 %>:</b>&nbsp;</td>
            <td>
              <input type="text" class="textbox" name="strMaxW" size="25" value="<% if strMaxW <> "" then Response.Write(strMaxW) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU19 %>:</b>&nbsp;</td>
            <td>
              <%= txtYes %>: 
              <input type="radio" class="radio" name="strDoThumb" value="1" <% if strDoThumb = "1" then Response.Write("checked") %>>
              <%= txtNo %>: 
              <input type="radio" class="radio" name="strDoThumb" value="0" <% if strDoThumb = "0" then Response.Write("checked") %>>
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU20 %>:</b>&nbsp;</td>
            <td>
              <input type="text" class="textbox" name="strMaxTH" size="25" value="<% if strMaxTH <> "" then Response.Write(strMaxTH) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle"><b><%= txtCU21 %>:</b>&nbsp;</td>
            <td>
              <input type="text" class="textbox" name="strMaxTW" size="25" value="<% if strMaxTW <> "" then Response.Write(strMaxTW) %>">
			  </td>
          </tr>
          <tr valign="top">
            <td align="right" valign="middle">&nbsp;</td>
            <td>&nbsp;
			<input type="hidden" name="hasUpload" value="1"></td>
          </tr>
			<% else %>
          <tr valign="top">
            <td align="right" valign="middle">&nbsp;</td>
            <td>&nbsp;
			<input type="hidden" name="intResize" value="0">
			<input type="hidden" name="strDoThumb" value="0">
			<input type="hidden" name="strMaxH" value="0">
			<input type="hidden" name="strMaxW" value="0">
			<input type="hidden" name="strMaxTH" value="0">
			<input type="hidden" name="strMaxTW" value="0">
			<input type="hidden" name="hasUpload" value="1">
			</td>
          </tr>
			<% end if %>
		<% else %>
          <tr valign="top">
            <td align="right" valign="middle">&nbsp;</td>
            <td>&nbsp;
			<input type="hidden" name="hasUpload" value="0">
			</td>
          </tr>
		<% end if %>
          <tr valign="top">
            <td align="right" valign="middle">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr valign="middle"> 
            <td colspan="2" align="center"> 
              <input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" onclick="selectModUsers('PostTopic',<%= upl %>)<% If cInt(request.QueryString("id"))=1 Then response.Write(";selectAll('PostTopic','tDATA1');") %>" class="button" style="width:120px;">
			  &nbsp;&nbsp;
              <input type="button" value="<%= txtGoBack %>" onclick="JavaScript:history.go(-1)" id="cancel" name="cancel" class="button" style="width:120px;">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form><br />
<% 
  set rsUp = nothing
End Sub

Sub ListGroups() 
 response.Write("<select size=""12"" name=""listGroups"" style=""width:200; font-size:10pt;"">" & vbNewLine)
 set rsGrp = my_Conn.execute("select G_ID, G_NAME from " & strTablePrefix & "GROUPS ORDER BY G_ACTIVE, G_NAME")
	if not rsGrp.eof then
	  do until rsGrp.eof
		response.Write("<option value=""" & rsGrp("G_ID") & """>" & rsGrp("G_Name") & "</option>" & vbNewLine)
		rsGrp.movenext
	  loop
	end if
	response.Write("</select>" & vbNewLine)
  set rsGrp = nothing
end Sub

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
end sub

sub config_downloads() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_downloadsX() %>
        <tr> <td align="center" valign="middle"><br />
				
				<a href="javascript:InsertSelection2('Add');" title="Add Group"><b>Add Group(s)</b></a><br />
				<a href="JavaScript:delLeader();" title="Remove selected Group(s)"><b>Remove Group(s)</b></a></td>
          <td align=center> <br />
            <br />
            <b>NOT USED</b><br />
            <select size="5" name="grpLeader" style="width:150;" multiple>
			  <% if cInt(request.QueryString("mode")) = 1 then
			  		'getGroupLeaders(gid)
				 end if %>
			  <option value="0"></option>
            </select>
            <input name="groupLeader" type=hidden value="" size=20 readOnly>
          </td>
        </tr>
<%
end sub

sub config_articles() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_links() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_forums() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_events() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_pictures() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub

sub config_classifieds() %>
        <tr> <td align="center" valign="middle" colspan="2"><br /><br />
            <b>&nbsp;</b><br />
          </td>
        </tr>
<%
end sub
 %>