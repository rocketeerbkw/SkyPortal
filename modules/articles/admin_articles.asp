<!-- #INCLUDE FILE="config.asp" --><%
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
CurPageType = "article"
%>
<!-- #INCLUDE file="lang/en/core_admin.asp" -->
<!-- #INCLUDE FILE="lang/en/article_lang.asp" -->
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/inc_admin_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="modules/articles/article_functions.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<% 
If Session(strCookieURL & "Approval") = "256697926329" Then

dim iPgType, sMode, cid, sid, strMsg, nPageTo, nPageCnt
strMsg = ""
iPgType = 0
sMode = 0
cid = 0
sid = 0
nPageTo = ""
isOK = false

setAppPerms "article","iName"

mod_setPageAppVars()

cat_id = cid
sub_id = sid

if iPgType = 20 or iPgType = 21 then
 Select case sMode
  case 1  'add category
    mod_addCategory()
  case 2  'add subcategory
    mod_addSubCategory(cid)
  case 3  'rename category
    mod_renameCategory(cid)
  case 4  'delete category
    deleteArtCategory(cid)
  case 5  'rename subcategory
    mod_renameSubCategory cid,sid
  case 6  'delete subcategory
    deleteArtSubCategory(sid)
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
	articleConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	spThemeBlock1_close(intSkin) %>
		</td>
		<td class="mainPgCol">
<%
  intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = "Admin Area|admin_home.asp"
  arg2 = "Article Admin|admin_articles.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  select case iPgType
	case 20
  	  arg3 = "Category Manager"
	  spThemeTitle = "Articles - Category Manager"
	case 21
  	  arg3 = "Subcategory Manager"
	  spThemeTitle = "Articles - Subcategory Manager"
	case 22
  	  arg3 = "Attention Manager"
	  spThemeTitle = "Articles - Attention Manager"
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
	      mod_edit_Item(intItemID)
		end if
	case else
		showAttentionSubCat(sid)
  end select
  response.Write("<br>&nbsp;")
  spThemeBlock1_close(intSkin) %>
		</td>
	</tr>
</table>
<!-- #INCLUDE FILE="inc_footer.asp" -->
<% 
Else
	Response.Redirect "admin_login.asp?target=admin_article.asp"
End If
%>