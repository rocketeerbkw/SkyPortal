<!-- #INCLUDE FILE="config.asp" -->
<!-- #INCLUDE FILE="lang/en/article_lang.asp" --><%
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
bShoRight = true
Dim iPgType, cat_id, sub_id, cid, sid, intItemID
Dim intDir, intShow, intLen, search, sMode
%>
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/core_module_functions.asp" -->
<!-- #INCLUDE FILE="Modules/articles/article_functions.asp" -->
<!-- #INCLUDE FILE="Modules/articles/article_custom.asp" -->
<%
iPageSize = art_PageSize
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = "?" & ChkActUsrUrl(Request.QueryString)
	if len(strOnlineQueryString) = 1 then
	  strOnlineQueryString = ""
	end if
	PageName = "Articles"
	PageAction = "Viewing<br>" 
	PageLocation = sScript & strOnlineQueryString
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function

  uploadPg = false
  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true
  
  iPgType = 0
  cat_id = 0
  sub_id = 0
  sMode = 0	
  cid = 0
  sid = 0
  intItemID = 0
  Comments = 0
		'strContent = ""
		'strArticleTitle = ""
		'strSummary = ""
		'intHit = 0		
		'strPostDate = ""
		'dateSince=""
		'strPoster = ""
		'intRating = 0
		'intVotes = 0

  mod_setPageAppVars()
  
  cat_id = cid
  sub_id = sid

if iPgType = 23 or (iPgType = 22 and sMode = 321) then
  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true
  bShoRight = false
end if

if iPgType = 7 then 'we are showing form
  modPgType = "addForm"
  uploadPg = false
  hasEditor = true
  strEditorElements = "Message"
  'strEditorType = ""
end if

  'get the default layout 
  cpSQL = "select * from PORTAL_PAGES where p_iname = '"&skyPage_iName&"'"
  set rsCPs = my_Conn.execute(cpSQL)
  if not rsCPs.eof then
  	  left_Col = rsCPs("p_leftcol")
  	  maint_Col = rsCPs("p_maintop")
	  mainb_Col = rsCPs("p_mainbottom")
  	  right_Col = rsCPs("p_rightcol")
  else
    set rsCPs = nothing
    closeAndGo("default.asp")
  end if
  set rsCPs = nothing
%>
<!-- #INCLUDE FILE="inc_top.asp" -->
<% 
setAppPerms "article","iName"
	
arg1 = txtArticles & "|" & app_page 'this is for the page breadcrumb

if iPgType = 20 or iPgType = 21 then
  Select case sMode
    case 1  'add category
	  if bAppFull then
        mod_addCategory()
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 2  'add subcategory
	  if mod_chkCatFull(cid) then
        mod_addSubCategory(cid)
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 3  'rename category
	  if mod_chkCatFull(cid) then
        mod_renameCategory(cid)
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 4  'delete category
	  if bAppFull then
        deleteArtCategory(cat_id)
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 5  'rename subcategory
	  if mod_chkSubCatFull(sid) then
        mod_renameSubCategory cid,sid
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 6  'delete subcategory
	  if mod_chkCatFull(cid) then
        deleteArtSubCategory(sid)
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 9  'update category order
	  if bAppFull then
        mod_updateCatOrder cid,sid
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
    case 10 'update subcategory order
	  if mod_chkCatFull(cid) then
        mod_updateSubCatOrder cid,sid
	  else
	    closeAndGo("error.asp?type=nopermtask")
	  end if
  end select
end if

  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if trim(left_Col) <> "" then
    cont = cont + 1
    response.Write("<td class=""leftPgCol"" valign=""top"">")
	intSkin = getSkin(intSubSkin,1)
	shoColumnBlocks(left_Col)
    response.Write("</td>")
  end if

    response.Write("<td class=""mainPgCol"" valign=""top"">")  
	intSkin = getSkin(intSubSkin,2)
    cont = cont + 1
  
  	'shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

  select case iPgType
	case 0
	  showall()
	case 1
	  showcat(cat_id)
	case 2
	  showsub()
	case 3, 4, 5
	  showAllSummaries()
	case 6
	  doSearch()
	case 7
	  addArticleForm()
	  bShoRight = false
	case 10
	  closeAndGo(app_rpage & "item=" & sub_id)
	case 20 'Category Manager
	  if strUserMemberID > 0 then
	    mod_showCategoryManager()
		bShoRight = false
	  else
	    showall()
	  end if
	case 21 'Subcategory Manager
	  if strUserMemberID > 0 then
	    mod_showSubCategoryManager()
		bShoRight = false
	  else
	    showall()
	  end if
	case 22 'View items that need attention
	  if strUserMemberID > 0 then
	    showAttentionSubCat(sub_id)
		bShoRight = false
	  else
	    showall()
	  end if
	case 23 'edit item
	  if strUserMemberID > 0 then
	    if sMode = 322 then
		  processEditItemForm()
		else
	      mod_edit_Item(intItemID)
		end if
	    'EditItemForm(intDLID)
	  else
	    showall()
	  end if
	case 24 'delete item
	  deleteItem item_tbl,item_fld,intItemID
	case else
	  'showcats()
  end select

  if trim(mainb_Col) <> "" then
	 shoColumnBlocks(mainb_Col)
  end if
    response.Write("</td>")
  
  if trim(right_Col) <> "" and bShoRight then
    if cont = 2 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	shoColumnBlocks(right_Col)
    response.Write("</td>")
  end if
  response.Write("</tr></table>")
%>
<!-- #INCLUDE FILE="inc_footer.asp" -->