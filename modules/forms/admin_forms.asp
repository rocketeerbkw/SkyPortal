<!--#include file="config.asp" --><%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2007 Dogg Software All Rights Reserved
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Classic ASP Form Creator v1.0
' Copyright David Angell, http://www.angells.com/FormCreator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/**
' * SkyPortal Forms Module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://www.opensource.org/licenses/mit-license.php MIT License
' */

pgType = "manager"
  modPgType = "addForm"
  uploadPg = false
  hasEditor = true
  strEditorElements = "fldIntroText, fldThankYou, fldInactiveText"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<!--#include file="inc_functions.asp" -->
<%
If len(Request.querystring("ajax")) > 0 then
  Select Case Request.querystring("ajax")
    case "AddFormField"
      tmpGetID = trim(request.querystring("Form")&" ")
      if len(tmpGetID) > 0 then if not isnumeric(tmpGetID) then closeAndGo("stop")
      formID = tmpGetID

      strSql = "INSERT INTO " & strTablePrefix & "FORMFIELDS (FLDLINKFORMID, FLDCAPTION, FLDFIELDTYPE, FLDVALIDATION, FLDREQUIRED, FLDWIDTH, FLDHEIGHT, FLDORDER, FLDDEFAULT, FLDOPTIONS) VALUES (" & formID & ", ' ', ' ', ' ', 'N', 0, 0, 0, ' ', ' ');"
      executeThis(strSQL)
      strSQL = "SELECT MAX(ID) AS ID FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID
      Set dbtable = my_Conn.Execute(strSql)

      response.clear
      response.write dbtable.fields("ID")
      response.end

    case "DeleteFormField"
      tmpGetID = trim(request.querystring("field")&" ")
      if len(tmpGetID) > 0 then if not isnumeric(tmpGetID) then closeAndGo("stop")
      fieldID = tmpGetID

      strSql = "DELETE FROM " & strTablePrefix & "FORMFIELDS WHERE ID=" & FieldID & ";"
      executeThis(strSQL)

      response.clear
      response.write "ffielddeleted"
      response.end

    case "DeleteForm"
      tmpGetID = trim(Request.Querystring("Form") & " ")
      if len(tmpGetID) > 0 then if not isNumeric(tmpGetID) then closeAndGo("stop")
      formID = tmpGetID

      strSQL = "DELETE FROM " & strTablePrefix & "FORMHEADER WHERE ID=" & FormID & ";"
      executeThis(strSQL)
      strSQL = "DELETE FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID=" & FormID & ";"
      executeThis(strSQL)

      response.clear
      response.write "fdeleted"
      response.end

    case else
      response.clear
      response.write "no match"
      response.end

  End select
End if

%>
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<!--#include file="modules/forms/form_functions.asp" -->
<% 
iPgType = 0
sMode = 0
a_id = 0
if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		sMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("w_id") <> "" or  Request("w_id") <> " " then
	if IsNumeric(Request("w_id")) = True then
		w_id = cLng(Request("w_id"))
	else
		closeAndGo("default.asp")
	end if
end if


%>
<style type="text/css">
@import "includes/css/form.import.css";
</style>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td width="190" class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
'getting rid of menu for now
''  if bFSO then
''  	mnu.menuName = "b_forms"
''    mnu.template = 4
''    mnu.thmBlk = 0
''    mnu.title = ""
''    mnu.shoExpanded = 1
''    mnu.canMinMax = 0
''    mnu.keepOpen = 1
''    mnu.GetMenu()
''  else
''    response.write("this module does not support servers w/o FSO yet.  You are not able to see this menu.")
''  end if
''  	response.Write("<hr />")
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<% 
	  intSkin = getSkin(intSubSkin,2)
	  'breadcrumb here
  	  arg1 = txtAdminHome & "|admin_home.asp"
  	  arg2 = "Forms Manager|admin_forms.asp"
  	  select case Request.querystring("action")
      	case "NewForm"
      		arg3 = "Add New Form"
      	case "EditForm"
      		arg3 = "Edit Form"
      	case "DefineFields"
      		arg3 = "Edit Form Fields"
      	case "DeleteForm"
      		arg3 = "Delete Form"
      	case "AddFormField"
      		'InsertBlankFormField
      		'Response.Redirect "admin_forms.asp?action=DefineFields&form=" & formID
      		'DefineFields
      	case else
      		arg3 = ""
      end select
  	  arg4 = ""
  	  arg5 = ""
  	  arg6 = ""

  
  	  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
spThemeTitle = ""
spThemeBlock1_open(intSkin)
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr><td width="100%">
<%
formID = GetID("form")
'select case iPgType
select case Request.querystring("action")
	case "NewForm"
		NewForm
	case "EditForm"
		EditForm
	case "DefineFields"
		DefineFields
	case "DeleteForm"
		DeleteForm
	case "CopyForm"
		CopyForm
	case "AddFormField"
		InsertBlankFormField
		Response.Redirect "admin_forms.asp?action=DefineFields&form=" & formID
		DefineFields
	case else
	  If Request.Querystring("next") <> "" then
  		ShowForms
		else
  		FormSummary
		end if
end select
%>
</td></tr>
</table>
<%
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp?target=admin_forms.asp" %><% end if %>
