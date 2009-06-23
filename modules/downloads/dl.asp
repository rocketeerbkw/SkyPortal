<!-- #include file="config.asp" -->
<!-- #include file="Lang/en/downloads_lang.asp" --><%
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
curIP = request.ServerVariables("REMOTE_ADDR")
canDL = true

sPage_iName = CurPageType
sPage_id = 0
sDLpage = "dl.asp"

':: Breadcrumb values
    arg1 = txtDownloads & "|" & sDLpage
  	arg2 = ""
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
':::::::::::::::::::::::::::::::::::::::::::::::::

pgname = "ERROR!"
CurPageInfoChk = "1"

Dim iPgType, cat_id, sub_id, intDir, intShow, intLen, ord1, ord2
dim dateSince, intHit, intDLID, sTxt, hp, search
dim strDescription, strBannerURL, strDLName, strPostDate
dim strDLEMAIL, strDLFILESIZE, strDLLICENSE, strDLLANGUAGE, strDLPLATFORM
dim strDLPUBLISHER, strDLPUBLISHER_URL, strDLUPLOADER, strPoster

  uploadPg = false
  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true
		
iPgType = 0
sMode = 0
cat_id = 0
sub_id = 0
cid = 0
sid = 0
intDLID = 0
intHit = 0
hp = 0
ord1 = "d"
ord2 = "Desc"
bShoRight = true
  if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp?cmd")
	end if
  end if
  if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		sMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp?mode")
	end if
  end if
  if Request("cid") <> "" or  Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cat_id = cLng(Request("cid"))
	else
		closeAndGo("default.asp?cid")
	end if
  end if
  if Request("sid") <> "" or  Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sub_id = cLng(Request("sid"))
	else
		closeAndGo("default.asp?sid")
	end if
  end if
  if Request("item") <> "" or  Request("item") <> " " then
	if IsNumeric(Request("item")) = True then
		intDLID = cLng(Request("item"))
	else
		closeAndGo("default.asp?item")
	end if
  end if
  
if iPgType = 20 or iPgType = 21 or iPgType = 22 then
  cid = cat_id
  sid = sub_id
end if
if iPgType = 23 or (iPgType = 22 and sMode = 321) then
  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true
  bShoRight = false
end if
%>

<!-- #include file="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="modules/downloads/dl_functions.asp" -->
<%
'closeAndGo("stop")
  'get the default layout 
  if sPage_id = 0 then
    cpSQL = "select * from PORTAL_PAGES where p_iname = '" & sPage_iName & "'"
  else
    cpSQL = "select * from PORTAL_PAGES where P_ID = " & sPage_id & ""
  end if
  set rsCPs = my_Conn.execute(cpSQL)
  if not rsCPs.eof then
	  pgtitle = rsCPs("p_title")
	  pgname = rsCPs("p_name")
	  if rsCPs("p_acontent") <> "" then
	    pgbody = replace(rsCPs("p_acontent"),"''","'")
	  else
	    if rsCPs("p_content") <> "" then
	      pgbody = replace(rsCPs("p_content"),"''","'")
	    end if
	  end if
  	  left_Col = rsCPs("p_leftcol")
  	  maint_Col = rsCPs("p_maintop")
	  mainb_Col = rsCPs("p_mainbottom")
  	  right_Col = rsCPs("p_rightcol")
	  
	  m_title = rsCPs("P_META_TITLE")
	  addToMeta "NAME","Description",rsCPs("P_META_DESC")
	  addToMeta "NAME","Keywords",rsCPs("P_META_KEY")
	  addToMeta "HTTP-EQUIV","Expires",rsCPs("P_META_EXPIRES")
	  addToMeta "NAME","Rating",rsCPs("P_META_RATING")
	  addToMeta "NAME","Distribution",rsCPs("P_META_DIST")
	  addToMeta "NAME","Robots",rsCPs("P_META_ROBOTS")
  end if
  set rsCPs = nothing

PageTitle = m_title

function CurPageInfo () 
	PageName = pgname 
	PageAction = txtBrows & "<br />" 
	PageLocation = request.ServerVariables("URL")
	if request.QueryString() <> "" then
	  PageLocation = PageLocation & "?" & chkString(request.QueryString(),"sqlstring")
	end if
	CurPageInfo = PageAction & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function 
%>
<!-- #include file="inc_top.asp" -->
<% 
':: set default module permissions
setAppPerms CurPageType,"iName"

  sScript = request.ServerVariables("SCRIPT_NAME")
  if instr(sScript,"/") > 0 then
    sScript = mid(sScript,instrrev(sScript,"/")+1,len(sScript))
  end if
  
  cont = 0
  bLeft = false
  bMaint = false
  bMainb = false
  bRight = false
  
  if trim(left_Col) <> "" then
    if right(left_Col,1) = "," then
      left_Col = left(left_Col,len(left_Col)-1)
    end if
    if instr(left_Col,",") > 0 then
      l_col = split(left_Col,",")
	else
	  dim l_col(0)
      l_col(0) = left_Col
	end if
    bLeft = true
    cont = cont + 1
  end if
  if trim(maint_Col) <> "" then
    if right(maint_Col,1) = "," then
      maint_Col = left(maint_Col,len(maint_Col)-1)
    end if
    if instr(maint_Col,",") > 0 then
      mt_col = split(maint_Col,",")
	else
	  dim mt_col(0)
      mt_col(0) = maint_Col
	end if
    bMaint = true
    cont = cont + 1
  end if
  if trim(mainb_Col) <> "" then
    if right(mainb_Col,1) = "," then
      mainb_Col = left(mainb_Col,len(mainb_Col)-1)
    end if
    if instr(mainb_Col,",") > 0 then
      mb_col = split(mainb_Col,",")
	else
	  dim mb_col(0)
      mb_col(0) = mainb_Col
	end if
    bMainb = true
    cont = cont + 1
  end if
  if trim(right_Col) <> "" then
    if right(right_Col,1) = "," then
      right_Col = left(right_Col,len(right_Col)-1)
    end if
    if instr(right_Col,",") > 0 then
      r_col = split(right_Col,",")
	else
	  dim r_col(0)
      r_col(0) = right_Col
	end if
    bRight = true
    cont = cont + 1
  end if

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
        deleteDlCategory(cat_id)
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
        deleteDlSubCategory(sid)
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
	case else
	  'closeAndGo("error.asp?type=nopermtask")
  end select
end if
  
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if bLeft then
    response.Write("<td class=""leftPgCol"" valign=""top"">")
	intSkin = getSkin(intSubSkin,1)
	  cStart = timer
	 shoBlocks(l_col)
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
  end if

    response.Write("<td class=""mainPgCol"" valign=""top"">")  
	intSkin = getSkin(intSubSkin,2)
	  cStart = timer
  
  if bMaint then
	 'shoBlocks(mt_col)
  end if
  
':: start main content
  select case iPgType
	case 0
	  showall()
	case 1
	  showcat(cat_id)
	case 2
	  showsub()
	case 3
	  shownew()
	case 4
	  showpopular()
	case 5
	  showtoprated()
	case 6
	  if sMode = 99 then
    	if cat_id = 0 then cat_id = intDLID
	    Call mod_deleteComment("DL","DL_ID")
		closeAndGo(sDLpage & "?cmd=6&cid=" & cat_id)
	  else
	    showItem()
	  end if
	case 7
	  doSearch()
	case 8
	  addDownload()
	case 10
	  closeAndGo(sDLpage & "?cmd=6&cid=" & sub_id)
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
		  'Response.Write "Old Process"
		  processEditItemForm()
		else
	      mod_edit_Item(intDLID)
		end if
	    'EditItemForm(intDLID)
	  else
	    showall()
	  end if
	case 24 'delete item
	  deleteItem(intDLID)
	
	case else
	  showall()
  end select
':: end main content

  if bMainb then
	 shoBlocks(mb_col)
  end if
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
	
  if bRight and bShoRight then
    if cont = 3 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	  cStart = timer
	shoBlocks(r_col)
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
  end if
  response.Write("</tr></table>")
  app_Footer()
 %>
<!--#include file="inc_footer.asp" -->