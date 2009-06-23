<!-- #include file="config.asp" --><%
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
CurPageType = "downloads"
sDLpage = "admin_dl_admin.asp"
sMDLpage = "dl.asp"

%>
<!-- #include file="lang/en/core_admin.asp" -->
<!-- #include file="includes/inc_admin_functions.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then

dim iPgType, iMode, cid, sid, strMsg, nPageTo, nPageCnt
strMsg = ""
fsoMsg = ""
iPgType = 22
iMode = 0
cid = 0
sid = 0
nPageTo = ""
isOK = false

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("cid") <> "" or  Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cid = cLng(Request("cid"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("sid") <> "" or  Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sid = cLng(Request("sid"))
	else
		closeAndGo("default.asp")
	end if
end if
if Request("item") <> "" or  Request("item") <> " " then
	if IsNumeric(Request("item")) = True then
		intItemID = cLng(Request("item"))
	else
		closeAndGo("default.asp")
	end if
end if

intDLID = intItemID

if iPgType = 0 then iPgType = 22

if iPgType = 20 or iPgType = 21 or iPgType = 22 then
  sMode = iMode
end if
if iPgType = 22 and sMode = 321 then
  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true
  bShoRight = false
end if

 %>
<!-- #include file="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="inc_top.asp" -->
<!-- #include file="modules/downloads/dl_functions.asp" -->
<%

	':: set default module permissions
	setAppPerms CurPageType,"iName"
	
	sSQL = "select UP_FOLDER from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
	set rsU = my_Conn.execute(sSQL)
	  downloadDir = rsU("UP_FOLDER")
	set rsU = nothing

Dim cat_id,sub_id,strPrice,strBanner,strPoster,strDLname,strDescription
Dim intDLID,intHit,intShow,strOwner,strPostDate,dateSince,expired

if iPgType = 20 or iPgType = 21 then
 Select case iMode
  case 1  'add category
    mod_addCategory()
  case 2  'add subcategory
    mod_addSubCategory(cid)
  case 3  'rename category
    mod_renameCategory(cid)
  case 4  'delete category
    deleteDlCategory(cat_id)
  case 5  'rename subcategory
    mod_renameSubCategory cid,sid
  case 6  'delete subcategory
    deleteDlSubCategory(sid)
  case 9  'update category order
    mod_updateCatOrder cid,sid
  case 10 'update subcategory order
    mod_updateSubCatOrder cid,sid
 end select
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign=top class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	downloadConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	spThemeBlock1_close(intSkin) %>
		</td>
		<td class="mainPgCol">
<%
  intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = "Admin Area|admin_home.asp"
  arg2 = "Downloads Admin|admin_dl_admin.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  select case iPgType
	case 20 'select category to edit
  	  arg3 = "Category Manager"
	  spThemeTitle = "Downloads - Category Manager"
	case 21
  	  arg3 = "Subcategory Manager"
	  spThemeTitle = "Downloads - Subcategory Manager"
	case 22
  	  arg3 = "Attention Manager"
	  spThemeTitle = "Downloads - Attention Manager"
  end select
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  spThemeBlock1_open(intSkin)
  select case iPgType				
	case 20 'Category Manager
	    mod_showCategoryManager()
	case 21 'Subcategory Manager
	    mod_showSubCategoryManager()
	case 22 'View items that need attention
	    showAttentionSubCat(sid)
	case 23 'edit item
	    if sMode = 322 then
		  processEditItemForm()
		else
	      mod_edit_Item(intDLID)
		end if
	case else 'View items that need attention
	    showAttentionSubCat(sid)
  end select
  response.Write("<br />&nbsp;")
  spThemeBlock1_close(intSkin) %>
		</td>
	</tr>
</table>
<!-- #include file="inc_footer.asp" -->
<% 
Else 
 Response.Redirect "admin_login.asp?target=admin_dl_main.asp" 
End If
%>