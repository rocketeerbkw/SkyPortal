<%
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

item_tbl = "ARTICLE"
item_fld = "ARTICLE_ID"
app_pop = "article_pop.asp"
app_page = "article.asp"
app_rpage = "article_read.asp?"
app_admin = "admin_articles.asp"
app_addForm = "article.asp?cmd=7&"
skyPage_iName = "article"
%>
<!-- #include file="article_admin.asp" -->
<!-- #include file="article_custom.asp" -->
<%
sub showAllSummaries()
  dim objPagingRS
  
  arg2 = "All Articles|"& app_page &"?cmd="& iPgType &"&amp;mode="& sMode
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  sSql = mod_singleItemSql(item_tbl)
  sSql = sSql & "WHERE ACTIVE = 1"
  if request("initial") <> "" then
    initial = left(chkString(request("initial"),"sqlstring"),1)
    sSql = sSql & " AND TITLE LIKE '" & initial & "%'"
  end if
  ord1 = chkString(request("ord1"),"sqlstring")
  ord2 = chkString(request("ord2"),"sqlstring")
  ord = ord1 & ord2
  select case ord
    case "hDesc"
      sSQL = sSQL & " ORDER BY " & item_tbl & ".HIT DESC;"
    case "hAsc"
  	  ord1 = "h"
  	  ord2 = "Asc"
      sSQL = sSQL & " ORDER BY " & item_tbl & ".HIT;"
    case "dDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".POST_DATE DESC;"
    case "dAsc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".POST_DATE;"
    case "rDesc"
	  sSQL = sSQL & " ORDER BY ROUND(RATING/VOTES, 0) DESC, VOTES DESC;"
    case "rAsc"
	  sSQL = sSQL & " ORDER BY ROUND(RATING/VOTES, 0), VOTES DESC;"
    case "tDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".TITLE DESC;"
    case "tAsc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".TITLE;"
    case else
	  ord = "dDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & "." & item_fld & " DESC;"
  end select
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  'app_MainColumn_top()
  
	iPageCurrent = 1
    If Request("page") = "" Then
	  iPageCurrent = 1
    Else
	  iPageCurrent = cLng(Request("page"))
    End If
  
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.PageSize = iPageSize
	objPagingRS.CacheSize = iPageSize
	objPagingRS.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	reccount = objPagingRS.recordcount
	iPageCount = objPagingRS.PageCount

	If iPageCurrent > iPageCount Then
	  iPageCurrent = iPageCount 
	end if
	If iPageCurrent < 1 Then
	  iPageCurrent = 1
	end if
	
	if iPageCount > 1 then
  	  tPgCnt = " - " & iPageCount & " pages"
	end if

    spThemeTitle= "<b>All Articles</b>"
	spThemeTitle = spThemeTitle & tPgCnt
	spThemeBlock1_open(intSkin)
	response.Write("<form method=""post"" action=""" & sScript & """>")
	response.Write("<input name=""cmd"" type=""hidden"" value=""" & iPgType & """>")
	response.Write("<input name=""cid"" type=""hidden"" value=""" & cat_id & """>")
	response.Write("<input name=""sid"" type=""hidden"" value=""" & sub_id & """>")
	response.Write("<input name=""initial"" type=""hidden"" value=""" & initial & """>")
	response.Write("Sort by:&nbsp;")
	response.Write("<select name=""ord1"" id=""ord1"" style=""margin-top:2px;"">")
	response.Write("<option value=""t""" & chkSelect(left(ord,1),"t") & ">Title</option>")
	response.Write("<option value=""h""" & chkSelect(left(ord,1),"h") & ">Hits</option>")
	response.Write("<option value=""r""" & chkSelect(left(ord,1),"r") & ">Rating</option>")
	response.Write("<option value=""d""" & chkSelect(left(ord,1),"d") & ">Post Date</option>")
	response.Write("</select>&nbsp;")
	response.Write("<select name=""ord2"">")
	response.Write("<option value=""Asc""" & chkSelect(right(ord,3),"Asc") & ">Asc</option>")
	response.Write("<option value=""Desc""" & chkSelect(right(ord,4),"Desc") & ">Desc</option>")
	response.Write("</select>&nbsp;")
	response.Write("&nbsp;<input name=""sub1"" type=""submit"" value="" Go "">")
	response.Write("</form>")
	
	arrAlpha = split(txtAlphabet,",")
	response.Write("<a href=""" & sScript & "?cmd=" & iPgType & "&amp;mode=" & sMode & "&amp;ord1=" & ord1 & "&amp;ord2=" & ord2 & """>"& txtAll &"</a>&nbsp;")
	for xa = 0 to ubound(arrAlpha)
	response.Write("&nbsp;<a href="""& sScript &"?cmd=" & iPgType & "&amp;mode=4&amp;initial="& arrAlpha(xa) &"&amp;ord1=" & ord1 & "&amp;ord2=" & ord2 & """>" & arrAlpha(xa) & "</a>")
	next
	Response.Write("<hr />")
	
 	If iPageCount = 0 Then
	  Response.Write("<div class=""fTitle"" style=""width:100%;"">")
	  Response.Write("<center><span class=""fAlert"" style=""text-align:center;"">")
	  Response.Write("<br /><b>No items found!</b><br /><br />")
	  Response.Write("</span></center></div>")
 	Else
	
	  objPagingRS.AbsolutePage = iPageCurrent

	  if iPageCount > 1 then
	    showDaPaging iPageCurrent,iPageCount,0
	  end if
	  iRecordsShown = 0
	  Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		Call DisplayArticle(objPagingRS)
	    iRecordsShown = iRecordsShown + 1
		objPagingRS.MoveNext
	  Loop
	  objPagingRS.Close
	  Set objPagingRS = Nothing

	  if iPageCount > 1 then
  	    showDaPaging iPageCurrent,iPageCount,2
	  end if
	end if
	
	Response.Write("<center><hr /><br />")
	Response.Write("<a href=""" & app_addForm & """>")
	Response.Write("Add an Article</a></center><br />")
	spThemeBlock1_close(intSkin)
end sub

sub showall()

  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  'app_MainColumn_top()
  'mod_displayIntro(intAppID)
  if trim(maint_Col) <> "" then
	 shoColumnBlocks(maint_Col)
  end if
  
  spThemeTitle = spThemeTitle & mod_iconSubscribe(0,0)
  
  If bAppFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit(app_pop,14,0,0,"right",2)
	spThemeTitle = spThemeTitle & "<a href="""&sScript&"?cmd=20&cid="&cid&""" title=""Category Manager"">"
	spThemeTitle = spThemeTitle & icon(icnToolbox,"Category Manager","display:inline","","align=""right""") & "</a>"
	if artAttention > 0 then
	  spThemeTitle = spThemeTitle & "<a href="""&sScript&"?cmd=22&amp;sid=" & sub_id & """>"
	  spThemeTitle = spThemeTitle & icon(icnAttention,"Items need attention","","","align=""right""") & "</a>"
	end if
  end if
  
  spThemeTitle = spThemeTitle & txtArticles
spThemeBlock1_open(intSkin)
  'strSql = "SELECT * FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID & " ORDER BY C_ORDER, CAT_NAME"
  'set rsCategories = server.CreateObject("adodb.recordset")
  'rsCategories.Open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
  dim rsAll
  sSql = mod_CatSubCatsql(0,0,intAppID)
  Set rsAll = my_Conn.execute(sSql)
  'response.Write(sSql)
  response.Write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" width=""100%"">"
  if rsAll.eof then
    ':: no records found
  else
   Do until rsAll.EOF
    response.Write "<tr>"
	ColNum = 1 
	Do while ColNum < 3
	  blkTimer = timer
	  if not rsAll.EOF then
	    'curCat = rsAll(strTablePrefix & "M_CATEGORIES.CAT_ID")
	    curCat = rsAll(sMCPre & "CAT_ID")
		if hasAccess(trim(rsAll("CG_READ"))) then
		  Response.Write "<td align=""left"" valign=""top"" width=""50%"">"
  		  If hasAccess(trim(rsAll("CG_FULL"))) or bAppFull Then 
			sTo = ""&sScript&"?cmd=21&amp;cid=" & curCat
			Response.Write "<a href="""&sTo&""" title=""Category Manager"">"
			Response.Write icon(icnToolbox,"Category Manager","","","")
			Response.Write "</a>"
			Response.Write modGrpEdit(app_pop,14,curCat,0,"bottom",rsAll("CG_INHERIT"))
		  else
		    Response.Write icon("images/icons/icon_folder_new_topic.gif","","","","")
  		  end if
		  Response.Write "<a href=""" & sScript & "?cmd=1&amp;cid=" & curCat & """>"
		  Response.Write "<b><span class=""fTitle"">"
		  Response.Write ChkString(rsAll("cat_name"),"display")
		  Response.Write "</span></b></a><br />"
		  'shoCatSubcats rsAll("cat_id"),rsAll("CG_FULL")
		  shoCatSubcats(rsAll)
	  
	  	  if shoBlkTimer then
	  		blkLoadTime = formatnumber((timer - blkTimer),3)
	  		response.Write(blkLoadTime)
	  	  end if
  		  response.Write("</td>")
		  ColNum = ColNum + 1 
		else
		  rsAll.movenext
		end if
	  else
		response.Write("<td>&nbsp;</td>")
		ColNum = ColNum + 1 
	  end if 
	Loop
    response.Write("</tr>")
   Loop
  end if
  response.Write("</table>")
  Call mod_shoLegend("main",art_chkNew,art_chkUpdated)
	  
  spThemeBlock1_close(intSkin)
  rsAll.close
  set rsAll = nothing
end sub

sub shoCatSubcats(ob)
  parent_id = ob(sMCPre & "CAT_ID")
  c = ob(sMCPre & "CAT_ID")
  do while not ob.EOF
   if parent_id <> ob(sMSPre & "CAT_ID") then exit sub
	if hasAccess(trim(ob("SG_READ"))) then
      subcatID = ob("SUBCAT_ID")
      bCatFull = false
      bSCatFull = false
	  if bAppFull then
        bCatFull = true
        bSCatFull = true
	  else
	    if hasAccess(ob("CG_FULL")) then
      	  bCatFull = true
          bSCatFull = true
		else
	      if hasAccess(ob("SG_FULL")) then
            bSCatFull = true
		  end if
		end if
	  end if
		  
	  rcounts = ob("ITEM_CNT")
	 if rcounts > 0 or (rcounts = 0 and (art_ShowEmptySubs or bSCatFull)) then
	  Response.Write icon(icnBar,"","","","")
 	  If bSCatFull Then 
		'Response.Write("&nbsp;" & modGrpEdit(app_pop,14,parent_id,rsSubcat("subcat_id"),"middle",rsSubcat("SG_INHERIT")))
		sTo = ""&sScript&"?cmd=21&amp;cid=" & parent_id
		response.Write(modGrpEdit(sTo,,,,"middle",ob("SG_INHERIT")))
		if rcounts >= 0 then
		  chkSubCatAttention(subcatID)
		end if
	  else
	    Response.Write(icon(icnArticle,"","","","align=""bottom"""))
  	  end if
	  %>
	  <a href="<%= sScript %>?cmd=2&amp;cid=<%=c%>&amp;sid=<%= subcatID %>"><span class="fNorm"><%= ob("SUBCAT_NAME") %>&nbsp;(<%= rcounts %>)</span></a>
	  <%
	  if rcounts > 0 then
	    Call mod_chkNewSubCatItems(subcatID,art_chkNew,art_chkUpdated)
	  end if
	  Response.Write "<br />"
	 end if
	end if
	ob.movenext
  loop
end sub

sub showcat(cid)
  sSql = mod_CatSubCatsql(cid,0,intAppID)
  Set rsC = my_Conn.execute(sSql)
	cat_name = rsC("CAT_NAME")
	inherit = rsC("CG_INHERIT")
	call setPermVars(rsC,1)
  
   if not bCatRead then
     closeandgo(app_page)
   end if
  
  arg2 = cat_name
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  if trim(maint_Col) <> "" then
	 shoColumnBlocks(maint_Col)
  end if
  
  spThemeTitle = spThemeTitle & mod_iconSubscribe(cid,0)
  spThemeTitle = spThemeTitle & mod_iconBookmark(cid,0,0)
  If bCatFull Then
	spThemeTitle = spThemeTitle & "<a href="""&sScript&"?cmd=21&cid="&cid&""" title=""Category Manager"">"
	spThemeTitle = spThemeTitle & icon(icnToolbox,"Category Manager","display:inline","","align=""middle""") & "</a>"
	spThemeTitle = spThemeTitle & modGrpEdit(app_pop,14,cid,0,"middle",inherit)
  else
    spThemeTitle = spThemeTitle & icon(icnNewFolder,txtPrint,"","","align=""middle"" hspace=""4""")
  end if
  spThemeTitle = spThemeTitle & "&nbsp;" & cat_name
  spThemeBlock1_open(intSkin)
  Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""6"" width=""100%"">")
  Do while NOT rsC.EOF
	sSCatRead = rsC("SG_READ")
	sSCatWrite = rsC("SG_WRITE")
	sSCatFull = rsC("SG_FULL")
	if bCatFull then
	  bSCatFull = true
	else
	  bSCatFull = hasAccess(sSCatFull)
	end if
	if hasAccess(sSCatRead) then
	  sSQL = "SELECT count(" & item_fld & ") FROM " & item_tbl & " where category=" & rsC("SUBCAT_ID") & " and ACTIVE=1"
	  Set RScount = my_Conn.Execute(sSQL)
	  rCount = RScount(0)
  	  Set RScount = nothing
	 if rcounts > 0 or (rcounts = 0 and (art_ShowEmptySubs or bSCatFull)) then
	  Response.Write "<tr>"
      Response.Write "<td align=""left"" class=""fNorm"">"
	  Response.Write icon(icnSpacer,"","","","width=""15""")
	  Response.Write icon(icnBar,"","","","hspace=""3""")
	  If bSCatFull Then
	    response.Write(modGrpEdit(app_pop,14,cid,rsC("SUBCAT_ID"),"middle",rsC("SG_INHERIT")))
		chkSubCatAttention(rsC("SUBCAT_ID"))
	  else
	    Response.Write icon(icnNewFolder,"","","","align=""middle"" hspace=""2""")
	  end if
      Response.Write("&nbsp;<a href=""" & sScript & "?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & rsC("SUBCAT_ID") & """><b>")
	  'Response.Write("<span class=""fSubTitle"">")
	  Response.Write(ChkString(rsC("subcat_name"), "display"))
	  'Response.Write("</span>")
	  Response.Write "</a> (" & rCount & ")</b>&nbsp;&nbsp;"
	  
	  if rCount > 0 then
	    call mod_chkNewSubCatItems(rsC("SUBCAT_ID"),dl_chkNew,dl_chkUpdated)
	  end if
      Response.Write("</td></tr>")
	 end if
	end if
	rsC.MoveNext
  Loop
  Response.Write("</table>")
  Call mod_shoLegend("main",dl_chkNew,dl_chkUpdated)
  spThemeBlock1_close(intSkin)
  set rsC = nothing
end sub

function showsub()  
  Dim iPageSize       
  Dim iPageCount      
  Dim iPageCurrent    
  Dim strOrderBy      
  Dim ssSQL   
  Dim iRecordsShown   
  Dim I      
  Dim cat_name
  Dim sub_name   
  Dim objPagingRS

  iPageSize = 6
  iPageCurrent = 1
  'set page size

  If Request("page") = "" Then
	iPageCurrent = 1
  Else
	iPageCurrent = cLng(Request("page"))
  End If

sSQL = "SELECT CAT_ID, SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & sub_id & " AND APP_ID = " & intAppID
set rsT = my_Conn.execute(sSQL)
  cat_id = rsT(0)
  sub_name = rsT(1)
set rsT = nothing
sSQL = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & cat_id & " AND APP_ID = " & intAppID
set rsT = my_Conn.execute(sSQL)
  cat_name = rsT(0)
set rsT = nothing

sSQL = "SELECT " & strTablePrefix & "M_CATEGORIES.*, " & strTablePrefix & "M_SUBCATEGORIES.* "
sSQL = sSQL & "FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID "
sSQL = sSQL & "WHERE (((" & strTablePrefix & "M_CATEGORIES.CAT_ID)=" & cat_id & ") AND ((" & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID)=" & sub_id & ") AND ((" & strTablePrefix & "M_CATEGORIES.APP_ID) = " & intAppID & "));"
	
  set rsT = my_Conn.execute(sSQL)
  cat_name = rsT("CAT_NAME")
  sub_name = rsT("SUBCAT_NAME")
  inherit = rsT("SG_INHERIT")
  call setPermVars(rsT,2)
  set rsT = nothing
  
  'shoDebugVars()

if bSCatRead then
  arg2 = cat_name & "|" & sScript & "?cmd=1&amp;cid=" & cat_id
  arg3 = sub_name & "|" & sScript & "?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & sub_id
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  if trim(maint_Col) <> "" then
	 shoColumnBlocks(maint_Col)
  end if

  sSQL = "SELECT * FROM " & item_tbl & " where CATEGORY=" & sub_id & " and ACTIVE = 1"
  if request("initial") <> "" then
    initial = left(chkString(request("initial"),"sqlstring"),1)
	sSQL = sSQL & " AND TITLE LIKE '" & initial & "%'"
  end if
  ord1 = chkString(request("ord1"),"sqlstring")
  ord2 = chkString(request("ord2"),"sqlstring")
  ord = ord1 & ord2
  select case ord
    case "hDesc"
      sSQL = sSQL & " ORDER BY " & item_tbl & ".HIT DESC;"
    case "hAsc"
  	  ord1 = "h"
  	  ord2 = "Asc"
      sSQL = sSQL & " ORDER BY " & item_tbl & ".HIT;"
    case "dDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".POST_DATE DESC;"
    case "dAsc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".POST_DATE;"
    case "rDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".RATING DESC;"
    case "rAsc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".RATING;"
    case "tDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".TITLE DESC;"
    case "tAsc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & ".TITLE;"
    case else
	  ord = "dDesc"
	  sSQL = sSQL & " ORDER BY " & item_tbl & "." & item_fld & " DESC;"
  end select

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.PageSize = iPageSize
objPagingRS.CacheSize = iPageSize
objPagingRS.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = objPagingRS.recordcount
iPageCount = objPagingRS.PageCount

If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1

if iPageCount = 1 then
  tPgCnt = " - " & iPageCount & " page"
else
  tPgCnt = " - " & iPageCount & " pages"
end if
if reccount = 1 then
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " items"&tPgCnt&")"
else
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " items"&tPgCnt&")"
end if
  
  spThemeTitle = spThemeTitle & mod_iconSubscribe(0,sub_id)
  spThemeTitle = spThemeTitle & mod_iconBookmark(0,sub_id,0)
	  sNewsFeedUrl = ""
      'spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:popUpWind('article_pop.asp?mode=xml&amp;sid=" & sub_id & "','xmlPub','540','620','yes','yes');""><img src=""" & icnPlus & """ title=""Publish to XML"" alt=""Publish"" border=""0"" align=""right"" hspace=""10"" style=""display:inline;"" hspace=""2""></a>"
  If bSCatFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit(app_pop,14,cat_id,sub_id,"right",inherit)
  end if
  spThemeTitle = spThemeTitle & "&nbsp;" & tTitle
  
    spThemeBlock1_open(intSkin)
  'response.write(ssSQL & "<br />" & objPagingRS("TITLE") & "<br />" & iPageCount & "<br />")
	response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	response.Write("<tr><td align=""center"" valign=""top"" class=""fNorm"">")
	response.Write("<form method=""post"" action=""" & sScript & """>")
	response.Write("<input name=""cmd"" type=""hidden"" value=""" & iPgType & """>")
	response.Write("<input name=""cid"" type=""hidden"" value=""" & cat_id & """>")
	response.Write("<input name=""sid"" type=""hidden"" value=""" & sub_id & """>")
	'response.Write("<input name=""initial"" type=""hidden"" value=""" & initial & """>")
	response.Write("Sort by:&nbsp;")
	response.Write("<select name=""ord1"" id=""ord1"" style=""margin-top:2px;"">")
	response.Write("<option value=""t""" & chkSelect(left(ord,1),"t") & ">Title</option>")
	response.Write("<option value=""h""" & chkSelect(left(ord,1),"h") & ">Hits</option>")
	response.Write("<option value=""r""" & chkSelect(left(ord,1),"r") & ">Rating</option>")
	response.Write("<option value=""d""" & chkSelect(left(ord,1),"d") & ">Post Date</option>")
	response.Write("</select>&nbsp;")
	response.Write("<select name=""ord2"">")
	response.Write("<option value=""Asc""" & chkSelect(right(ord,3),"Asc") & ">Asc</option>")
	response.Write("<option value=""Desc""" & chkSelect(right(ord,4),"Desc") & ">Desc</option>")
	response.Write("</select>&nbsp;")
	response.Write("&nbsp;<input name=""sub1"" type=""submit"" class=""button"" value="" Go "">")
	response.Write("</form>")
	arrAlpha = split(txtAlphabet,",")
	response.Write("<a href=""" & sScript & "?cmd=" & iPgType & "&amp;cid=" & cid & "&amp;sid=" & sid & "&amp;ord1=" & ord1 & "&amp;ord2=" & ord2 & """>"& txtAll &"</a>&nbsp;")
	for xa = 0 to ubound(arrAlpha)
	response.Write("&nbsp;<a href="""& sScript &"?cmd=" & iPgType & "&amp;cid=" & cid & "&amp;sid=" & sid & "&amp;initial="& arrAlpha(xa) &""">" & arrAlpha(xa) & "</a>")
	'response.Write("&nbsp;<a href="""& sScript &"?cmd=" & iPgType & "&amp;cid=" & cid & "&amp;sid=" & sid & "&amp;initial="& arrAlpha(xa) &"&amp;ord1=" & ord1 & "&amp;ord2=" & ord2 & """>" & arrAlpha(xa) & "</a>")
	next
	Response.Write("<hr />")
	response.Write("</td></tr>")
	response.Write("<tr>")
	response.Write("<tr><td align=""center"" valign=""top"">")
  If iPageCount = 0 Then
    sG = "<span class=""fSubTitle"">No items found!</span>"
    sG = sG & "<br /><br /><hr /><center>"
	sG = sG & "<a href=""" & app_addForm & "&amp;sid=" &sub_id& "&amp;cat_name=" &sub_name& "&amp;cid=" &cat_id& "&amp;parent_name=" &cat_name& """>"
	sG = sG & "Submit an Article</a></center><br />"
    response.Write(sG) 
  Else
	
	objPagingRS.AbsolutePage = iPageCurrent
	if iPageCount > 1 then
	  showDaPaging iPageCurrent,iPageCount,0
	end if
	iRecordsShown = 0
	rCount = 0
	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	
		call displayArticle(objPagingRS)
	    iRecordsShown = iRecordsShown + 1
	    objPagingRS.MoveNext
	Loop

    objPagingRS.Close
    Set objPagingRS = Nothing
    if iPageCount > 1 then
      showDaPaging iPageCurrent,iPageCount,2
    end if
    If bSCatWrite Then
    %>
    <center>
    <hr /><a href="<%= app_addForm %>&amp;sid=<%=sub_id%>&amp;cat_name=<%=sub_name%>&amp;cid=<%=cat_id%>&amp;parent_name=<%=cat_name%>">
  <span class="fNorm"><b>Submit an Article</b></span></a>
    </center>
    <br />
    <%
	end if
  End If
	response.Write("</td></tr></table>")
    spThemeBlock1_close(intSkin)
else ':: no access so redirect
  closeandgo(sScript)
end if
end function

function doSearch()
  search = ChkString(Request("search"), "SQLString")
  show = 8
  if request("num") <> "" then
    show = clng(Request("num"))
  end if
  if show <> "" then
	Dim iPageCount      
	Dim iPageCurrent    
	Dim strOrderBy      
	Dim strSQL          
	Dim objPagingConn   
	Dim objPagingRS     
	Dim iRecordsShown   
	Dim I

	iPageSize = show

	If Request("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = cLng(Request("page"))
	End If

	'::::: variable search reutine :::::::::
	if sMode = 0 then 'search all articles
      strSQL = "select * from ARTICLE where (KEYWORD like'%" & search & "%' or SUMMARY like '%" & search & "%' or TITLE like '%" & search & "%' or CONTENT like '%" & search & "%') and ACTIVE=1 order by HIT DESC, ARTICLE_ID DESC"
	  strSrchTxt = "Search results for"
	elseif sMode = 1 then ':: search member submitted articles
      strSQL = "select * from ARTICLE where (POSTER = '" & search & "' or POSTER like '%" & search & "%') and ACTIVE=1 order by POSTER, HIT DESC, ARTICLE_ID DESC"
	  strSrchTxt = "Items submitted by"
	elseif sMode = 2 then ':: search by Author
      strSQL = "select * from ARTICLE where (TDATA2 = '" & search & "' or TDATA2 like '%" & search & "%') and ACTIVE=1 order by HIT DESC, ARTICLE_ID DESC"
	  strSrchTxt = "Search results for Author"
	end if
'Response.Write strSQL
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.PageSize = iPageSize
	objPagingRS.CacheSize = iPageSize
	objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	reccount = objPagingRS.recordcount
	iPageCount = objPagingRS.PageCount

	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
	If iPageCurrent < 1 Then iPageCurrent = 1

  	arg1 = "Articles|" & app_page
  	arg2 = strSrchTxt & ": " & search & ""
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
    if trim(maint_Col) <> "" then
	 shoColumnBlocks(maint_Col)
    end if

	spThemeBlock1_open(intSkin)

	If iPageCount = 0 Then
		Response.Write "<div class=""fTitle"" class=""text-align:center;"">"
		Response.Write "<b>No Articles found!</b></div><br><br>"
	Else
		objPagingRS.AbsolutePage = iPageCurrent %>
		<center><div class="fTitle"><b><%= strSrchTxt %> :&nbsp;</b><span class="fAlert"><b><%=search%></b></span></div>
		<span class="fAlert"> found <%=reccount%><% if reccount = 1 then %> article<% else %> articles<% end if %></span></center><%
		if iPageCount > 1 then
	  	showDaPaging iPageCurrent,iPageCount,0
		end if
		curPoster = search
		iRecordsShown = 0
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		  if sMode = 1 then
		    pClass = "fAlert"
		  else
		    pClass = "fSmall"
		  end if
			dagar=DateDiff("d", Date, strtodate(objPagingRS("POST_DATE")))+7
			spThemeBlock4_open() %>
			<table border="0" width="95%" cellspacing="1" cellpadding="6" align="center">
  			<tr><td width="100%">
      		&#149; <b><a href="<%= app_rpage %>item=<%=objPagingRS("ARTICLE_ID")%>"><%=objPagingRS("TITLE")%></a></b><%if dagar >= 0 then response.write "<img src=""themes/" &  strTheme & "/new.gif"" alt=""new"" />"%><br>
			<span class="fSmall">Posted by: <span class="<%= pClass %>"><b><%=objPagingRS("POSTER")%></b></span></span><br />
      		(Added : <%=formatdatetime(strtodate(objPagingRS("POST_DATE")), 2)%> Hits : <%=objPagingRS("HIT")%>)<br />
      		<%=objPagingRS("Summary")%>
      		<br />
    		</td></tr></table>
			<br /><%
			spThemeBlock4_close()
			iRecordsShown = iRecordsShown + 1
			objPagingRS.MoveNext
		Loop
		%><%
	End If

	objPagingRS.Close
	Set objPagingRS = Nothing
	if iPageCount > 1 then
  	  showDaPaging iPageCurrent,iPageCount,2
	end if
  	spThemeBlock1_close(intSkin)
  else
	' hmmmm
  end if
end function

sub showDaPaging(nPageTo,nPageCnt,nPaging)
	'Display Paging Buttons
	sHidden = vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""mode"" value=""" & sMode & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""cid"" value=""" & cid & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sid & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""search"" value=""" & search & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""initial"" value=""" & initial & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""ord1"" value=""" & ord1 & """ />"
	sHidden = sHidden & vbCrLf & "<input type=""hidden"" name=""ord2"" value=""" & ord2 & """ />"
	
  Response.Write("<center><table border=""0"" cellpadding=""4"" cellspacing=""4"">")
	if (nPageCnt > totSho) and nPaging = 1 then
	  Response.Write("<tr>")
	  Response.Write("<td colspan=""5"" align=""center""><span class=""fSmall""><b>Page <span class=""fAlert"">" &  nPageTo & "</span> of <span class=""fAlert"">" & nPageCnt & "</span></b></span>")
	  Response.Write("</td>")
	  Response.Write("</tr>")
	end if
	' Display <<
	Response.Write(vbCrLf & "<tr><td align=""center"">")
	Response.Write(vbCrLf & "<form action=""" & sScript & """ method=""post"" name=""formP"&nPaging&"01"" id=""formP"&nPaging&"01"">")
	If int(nPageTo) = 1 Then 
	  Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold}"" disabled=""disabled"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2"" /><input type=""hidden"" name=""page"" value=""1"" />")
	Else
	  Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold;cursor:pointer;}"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2""><input type=""hidden"" name=""page"" value=""1"" />")
	End IF
	Response.Write(sHidden)
	Response.Write(vbCrLf & "</form>")
	Response.Write(vbCrLf & "</td>")
	' Display <
	Response.Write(vbCrLf & "<td align=""center"">")
	Response.Write(vbCrLf & "<form action=""" & sScript & """ method=""post"" name=""formP"&nPaging&"02"" id=""formP"&nPaging&"02"">")
	If int(nPageTo) = 1 Then 
	  Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""1"" />")
	Else
	  Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold;cursor:pointer;}"" />")
	  Response.Write(vbCrLf & "<input type=""hidden"" name=""page"" value=""" & nPageTo-1 & """ />")
	End If
	Response.Write(sHidden)
	Response.Write(vbCrLf & "</form>")
	Response.Write(vbCrLf & "</td>")
	' Display >
	strQryStr = ""
	if sMode <> "" then
	  strQryStr = strQryStr & "&amp;mode=" & sMode
	end if
	if initial <> "" then
	  strQryStr = strQryStr & "&amp;initial=" & initial
	end if
	if search <> "" then
	  strQryStr = strQryStr & "&amp;search=" & search
	end if
	if ord1 <> "" and ord2 <> "" then
	  strQryStr = strQryStr & "&amp;ord1=" & ord1
	  strQryStr = strQryStr & "&amp;ord2=" & ord2
	end if
	if nPageCnt > 1 then
	  Response.Write("<td align=""center"">")
	  totSho = 5
	  b4 = cint((totSho-1)/2)
	  pgS = nPageTo-b4
	  if pgS < 1 then
		pgS = 1
	  end if 
	  pgE = pgS+(totSho-1)
	  if pgE > nPageCnt then
		pgE = nPageCnt
		pgS = pgE-(totSho-1)
	  end if
	  if pgS < 1 then
		pgS = 1
	  end if 
	  for pgc = pgS to pgE
		if nPageTo = pgc then
		  Response.Write("<span class=""fAlert"">")
		  Response.Write("&nbsp;[" & pgc & "]</span>")
		else
		  Response.Write("&nbsp;<a href=""" & sScript & "?cmd=" & iPgType & "&amp;cid=" & cat_id & "&amp;sid=" & sub_id & "&amp;page=" & pgc & strQryStr & """>")
		  Response.Write("<span class=""fBold"">" & pgc & "</span></a>")
		end if
	  next
	  Response.Write("&nbsp;</td>")
	end if
						
	Response.Write(vbCrLf & "<td align=""center"">")
	Response.Write(vbCrLf & "<form action=""" & sScript & """ method=""post"" id=""formP"&nPaging&"03"" name=""formP"&nPaging&"03"">")
	If int(nPageTo) = nPageCnt Then 
	  Response.Write(vbCrLf & "<input type=""submit"" value='  Next &gt;  ' id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""" & nPageTo & """ />")
	Else
	  Response.Write(vbCrLf & "<input type=""submit"" value=""  Next &gt;  "" id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold;cursor:pointer;}"" />")
	  Response.Write(vbCrLf & "<input type=""hidden"" name=""page"" value=""" & nPageTo+1 & """ />")
	End IF
	Response.Write(sHidden)
	Response.Write(vbCrLf & "</form>")
	Response.Write(vbCrLf & "</td>")
	' Display >>
	Response.Write(vbCrLf & "<td align=""center"">")
	Response.Write(vbCrLf & "<form action=""" & sScript & """ method=""post"" id=""formP"&nPaging&"04"" name=""formP"&nPaging&"04"">")
	If int(nPageTo) = nPageCnt Then 
	  Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""" & nPageTo & """ />")
	Else
	  Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold;cursor:pointer;}"" /><input type=""hidden"" name=""page"" value=""" & nPageCnt & """ />")
	End IF
	Response.Write(sHidden)
	Response.Write(vbCrLf & "</form>")
	Response.Write(vbCrLf & "</td>")
	Response.Write("</tr>")
	if (nPageCnt > totSho) and nPaging = 2 then
	  Response.Write("<tr>")
	  Response.Write("<td colspan=""5"" align=""center""><span class=""fSmall""><b>Page <span class=""fAlert"">" &  nPageTo & "</span> of <span class=""fAlert"">" & nPageCnt & "</span></b></span>")
	  Response.Write("</td>")
	  Response.Write("</tr>")
	end if
	Response.Write("</table></center>")
end sub

sub addArticleForm()
  	arg1 = "Articles|" & app_page
  	arg2 = "Add Article|" & app_addForm
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
	intDir = 0
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  	'app_MainColumn_top()
	
	sMode = 0
	if Request.form("mode") <> "" or  Request.form("mode") <> " " then
	  if IsNumeric(Request.form("mode")) = True then
		sMode = cLng(Request.form("mode"))
	  else
		closeAndGo("default.asp")
	  end if
    end if
	
	spThemeTitle = "Add an Article"
	spThemeBlock1_open(intSkin)%>
<%  
if sMode = 22 then ' a form has been submitted
	sString = ""
  
  bSecCodeMatch = true
  if intSecCode <> 0 then
    fSecCode = ChkString(request.form("secCode"),"sqlstring")
    if DoSecImage(fSecCode) then
      'Image matched their input 
      bSecCodeMatch = true
    else
      'Image did not match their input
      bSecCodeMatch = false
	  sString = sString & "<li>Your Security Code didn't match.</li>"
    end if
  end if
  
	cat = cLng(Request.Form("subcat"))
	title = replace(ChkString(Request.Form("title"),"sqlstring"), "'","''", 1, -1, 1)
	content = ChkString(ChkBadWords(Request.Form("Message")),"message")
	summary = ChkString(ChkBadWords(Request.Form("summary")),"message")
	key = replace(ChkString(Request.Form("key"),"sqlstring"), "'","''", 1, -1, 1)
	poster = strDBNTUserName
	posteremail = replace(ChkString(Request.Form("posteremail"),"sqlstring"),"'","",1,-1,1)
	today = strCurDateString
	
	if len(trim(title)) = 0 then
	  sString = sString & "<li>Please enter article title.</li>"
	else
	  strSql="Select TITLE from ARTICLE where TITLE='" & title & "'"
	  Set rs = my_Conn.execute(strSql)
	  if not rs.eof then
		sString = sString & "<li>This article already exists in our database.</li>"
	  end if
	  Set rs = nothing
	end if
	if cat = 0 then
	  sString = sString & "<li>Please select a category that matches your article.</li>"
	end if
	if len(trim(content)) = 0 then
	  sString = sString & "<li>Please enter the article content.</li>"
	end if
	if len(trim(summary)) = 0 then
	  sString = sString & "<li>Please enter the article summary.</li>"
	else
	  if len(trim(summary)) > 400 then
	  	sString = sString & "<li>" &len(trim(summary))&" characters."
		sString = sString & " Your summary is too long!<br />400 characters max.</li>"
	  end if
	end if
	if len(trim(posteremail)) = 0 then 
	  sString = sString & "<li>You must give your email address.</li>"
	else
  	  if EmailField(posteremail) = 0 then 
		sString = sString & "<li>You must enter a valid email address.</li>"
  	  end if
	end if
	
  if sString = "" then ' no error, add to the db
	
	sT1Sql = ""
	sT2Sql = ""

  if txtArtLabel1 <> "" then
    sTDATA1 = ChkString(Request.Form("TDATA1"),"sqlstring")
	sT1Sql = ", TDATA1"
	sT2Sql = ", '" & sTDATA1 & "'"
  end if
  if txtArtLabel2 <> "" then
    sTDATA2 = ChkString(Request.Form("TDATA2"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2"
	sT2Sql = sT2Sql & ", '" & sTDATA2 & "'"
  end if
  if txtArtLabel3 <> "" then
    sTDATA3 = ChkString(Request.Form("TDATA3"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3"
	sT2Sql = sT2Sql & ", '" & sTDATA3 & "'"
  end if
  if txtArtLabel4 <> "" then
    sTDATA4 = ChkString(Request.Form("TDATA4"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4"
	sT2Sql = sT2Sql & ", '" & sTDATA4 & "'"
  end if
  if txtArtLabel5 <> "" then
    sTDATA5 = ChkString(Request.Form("TDATA5"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5"
	sT2Sql = sT2Sql & ", '" & sTDATA5 & "'"
  end if
  if txtArtLabel6 <> "" then
    sTDATA6 = ChkString(Request.Form("TDATA6"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6"
	sT2Sql = sT2Sql & ", '" & sTDATA6 & "'"
  end if
  if txtArtLabel7 <> "" then
    sTDATA7 = ChkString(Request.Form("TDATA7"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7"
	sT2Sql = sT2Sql & ", '" & sTDATA7 & "'"
  end if
  if txtArtLabel8 <> "" then
    sTDATA8 = ChkString(Request.Form("TDATA8"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8"
	sT2Sql = sT2Sql & ", '" & sTDATA8 & "'"
  end if
  if txtArtLabel9 <> "" then
    sTDATA9 = ChkString(Request.Form("TDATA9"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9"
	sT2Sql = sT2Sql & ", '" & sTDATA9 & "'"
  end if
  if txtArtLabel10 <> "" then
    sTDATA10 = ChkString(Request.Form("TDATA10"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10"
	sT2Sql = sT2Sql & ", '" & sTDATA10 & "'"
  end if

	strSQL = "SELECT CAT_ID FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & cat & " AND APP_ID=" & intAppID & ""
	dim rsCategories
	set rsCategories = my_Conn.execute(strSQL)
	parent = rsCategories("CAT_ID")
	set rsCategories = nothing

	strSql = "INSERT INTO ARTICLE"
	strSql = strSql & " (TITLE"
	strSql = strSql & ", KEYWORD"
	strSql = strSql & ", CATEGORY"
	strSql = strSql & ", CONTENT"
	strSql = strSql & ", SUMMARY"
	strSql = strSql & ", POST_DATE"
	strSql = strSql & ", PARENT_ID"
	strSql = strSql & ", ACTIVE"
	strSql = strSql & ", HIT"
	strSql = strSql & ", POSTER"
	strSql = strSql & ", POSTER_EMAIL"
	strSql = strSql & sT1Sql
	strSql = strSql & ") VALUES ("
	strSql = strSql & "'" & title & "'"
	strSql = strSql & ", '" & key & "'"
	strSql = strSql & ", '" & cat & "'"
	'strSql = strSql & ", '" & replace(content,"'","''") & "'"
	'strSql = strSql & ", '" & replace(summary,"'","''") & "'"
	strSql = strSql & ", '" & content & "'"
	strSql = strSql & ", '" & summary & "'"
	strSql = strSql & ", '" & today & "'"
	strSql = strSql & ", '" & parent & "'"
	if bAppFull or s_full or c_full then
	  strSql = strSql & ", 1"
	else
	  strSql = strSql & ", 0"
	end if 
	strSql = strSql & ", 0"
	strSql = strSql & ", '" & poster & "'"
	strSql = strSql & ", '" & posteremail & "'"
	strSql = strSql & sT2Sql
	strSql = strSql & ")"
      executeThis(strSQL)
	  
	  if (bAppFull or s_full or c_full) then
	    mod_increaseSubcatCount(cat)
	    if intSubscriptions = 1 and strEmail = 1 then
	      'send subscriptions emails
	      eSubject = strSiteTitle & " - New Article"
		  eMsg = "A new article has been submitted at " & strSiteTitle & vbCrLf
		  eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		  eMsg = eMsg & "You can view the new articles by visiting " & strHomeUrl & app_page & "?cmd=3" & vbCrLf
	      sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
		  'response.Write("<br>Email sent<br>" )
		end if
	  end if
	%>
		  <p align="center"><span class="fSubTitle">
		  <% if hasAccess(1) then%>
			Your article has been posted.
		  <%else%>
			Your article has been accepted for review.<br />
			Please wait 1-3 days for your article to be reviewed.
		  <%end if%></span></p>
				<table border="0" cellpadding="4" cellspacing="0" width="75%" align="center">
					<tr><td width="30%" class="fNorm">
							<b>Title:</b>&nbsp;
						</td><td class="fNorm" align="left" width="70%">
							<%= replace(replace(title,"''","'", 1,-1,1),"''","'",1,-1,1)%>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Summary:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(summary, "''","'", 1,-1,1) %>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Keywords:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(key, "''","'", 1, -1, 1) %>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Posted by:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(poster, "''","'", 1, -1, 1) %>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Poster's Email:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(posteremail, "''","'", 1, -1, 1) %>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Author:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(author, "''","'", 1, -1, 1) %>
						</td>
					</tr>
					<tr><td class="fNorm">
							<b>Author's email/website:</b>&nbsp;
						</td><td class="fNorm" align="left">
							<%= replace(authoremail, "''","'", 1, -1, 1)%>
						</td>
					</tr>
				</table>
		<meta http-equiv="Refresh" content="6; URL=<%= app_addForm %>?cid=<%= parent %>&amp;sid=<%= cat %>">
<%else%>
		<center><div style="width:400px;">
		<p align="center" class="fTitle">There Was A Problem.</p>
		<center><span><ul style="text-align:left;"><% =sString %></ul></span></center>
		<p align="center"><a href="JavaScript:history.go(-1)">
		Go Back To Enter Data</a></p><br></div></center>
<%
  end if
else 'show form
	%>
<form method="post" action="<%= app_addForm %>" id="formArt" name="formArt">
  <table><tr>
    <td width="30%" align="right" class="fNorm" nowrap="nowrap"> 
      Subcategory:&nbsp;</td><td>
	  <% mod_selectCatSubcat sub_id,"WRITE" %>
	</td>
  </tr>
    <tr>
    <td colspan="2" class="fNorm">&nbsp;</td>
  </tr>
  <tr>
    <td align="right" class="fNorm"> 
      <span class="fAlert">*</span><%= txtArtTitle %>:&nbsp; 
       </td>
	<td><input type="text" name="title" size="40" maxlength="90" /></td>
  </tr>
  <% customFormElements(oz) %>
    <tr>
    <td colspan="2" class="fNorm">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" align="right" class="fNorm">
	<span class="fAlert">*</span>Summary:&nbsp;<br />(400 characters max.)&nbsp;<br><br><span id="charLeft1">400 characters left&nbsp;</span></td>
    <td align="left"><textarea rows="10" name="summary" id="summary" cols="45" onKeyUp="cntChar('summary','charLeft1','{CHAR} characters left&nbsp;',400);"></textarea></td>
  </tr>
    <tr>
    <td colspan="2" class="fNorm">&nbsp;</td>
  </tr><!-- insert editor here --> 
  <% 
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<span class=""fAlert"">*</span> Article ","Article Content"
  else
  	displayPLAINeditor 1,""
  end if
  if intSecCode <> 0 then
  %>
  <tr>
    <td colspan="2" align="center"><% shoSecurityImg %></td>
  </tr>
  <% 
  End If
  %>
  <tr>
    <td colspan="2" height="1">
    <%if strDBNTUserName = "" then%>
    <input type="hidden" name="poster" value="Anonymous" />
    <% else %>
    <input type="hidden" name="poster" value="<%=strDBNTUserName %>" />
      <input type="hidden" name="mode"  value="22" />
    <% end if%>
    </td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="submit" value="Submit Article" name="B1" accesskey="s" title="Shortcut Key: Alt+S" class="button" />&nbsp;&nbsp;<% If strAllowHtml <> 1 Then %><input name="Preview" type="button" class="Button" value=" Preview " onclick="OpenPreview()" />&nbsp;&nbsp;<input type="reset" value="Reset" name="B2" class="button" /><% End If %></td>
  </tr></table></form><hr />
<center>
  <p><span class="fAlert">*</span> = required field<br />
    <br />
    If you did not see a category that fit your article's content,<br />
    <a href="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getMemberID(split(strwebmaster,",")(0)) %>');"><u><span class="fAlert">contact 
    us</span></u></a> and we'll be happy to consider it.<br />
    <br />
    <%	if lcase(strEmail) = "1" then%>
    We will notify you by Email when the article gets added to our database. 
    <%  end if%>
    <br />
    <br />
    <a href="<%= app_page %>">Back</a> 
  </p>
</center>
<%
end if 'show or process form
%>
	<%spThemeBlock1_close(intSkin)%>
	<p align="center"><a href="<%= app_page %>">Back to main Categories</a></p>
<%
end sub

function GetCategories_old(ii)
  ':: ii = subcat_ID
  sSQL = "SELECT " & strTablePrefix & "M_CATEGORIES.*, " & strTablePrefix & "M_SUBCATEGORIES.*"
  sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
  sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_SUBCATEGORIES.APP_ID)=" & intAppID & "))"
  sSQL = sSQL & " ORDER BY " & strTablePrefix & "M_CATEGORIES.CAT_NAME, " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_NAME;"

	dim rsC
	set rsC = my_Conn.execute(sSql)
	
    Response.Write "<td>"
    Response.Write "<select name=""cat"">"
	curCat = ""
	do while not rsC.EOF
	  if curCat <> rsC("" & strTablePrefix & "M_CATEGORIES.CAT_ID") then
	    curCat = rsC("" & strTablePrefix & "M_CATEGORIES.CAT_ID")
	    Response.Write "<optgroup label=""" & rsC("CAT_NAME") & """>"
	  end if
	  Response.Write "<option value="""&rsC("SUBCAT_ID")&""""
	  Response.Write chkSelect(ii,rsC("SUBCAT_ID")) & ">"
	  'Response.Write rsC("CAT_NAME")&" / "
	  Response.Write "- " & rsC("SUBCAT_NAME")
	  Response.Write "</option>"
	  rsC.MoveNext
	  if rsC.eof then
	    Response.Write "</optgroup>"
		Response.Write "<optgroup title=""Spacer""></optgroup>"
	  else
	   if curCat <> rsC(""& strTablePrefix &"M_CATEGORIES.CAT_ID") then
	    Response.Write "</optgroup>"
		Response.Write "<optgroup title=""Spacer""></optgroup>"
	   end if
	  end if
	loop
    Response.Write "</select>"
    Response.Write "</td>"
	set rsC = nothing
end function

sub menu_article()
	spThemeTitle= txtMenu
	spThemeBlock1_open(intSkin)
 if bFSO then
    mnu.menuName = "m_article"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
 else
 %>
	<div class="menu">
      <a href="<%= app_page %>?cmd=3">- <%= txtNewArts %><br /></a>
      <a href="<%= app_page %>?cmd=4">- <%= txtPopArts %><br /></a>
      <a href="<%= app_page %>?cmd=5">- <%= txtTopArts %><br /></a>
	<%if hasAccess("1,2") then%>
      <a href="<%= app_page %>?cmd=7">- <%= txtSubArt %><br /></a><% End If %>
      <a href="javascript:openWindow3('<%= app_pop %>?mode=10')">- <%= txtArtFAQ %><br /></a>
	</div>
<% End If %>
 <br />
<script type="text/javascript">
function chkSrchForm1() {
mt=document.formS1.search.value;
if (mt.length < 3) {
alert("Search word must be more than 3 characters");
return false;
}
else { return true; }
}
</script>
	<form method="get" action="<%= app_page %>" id="formS1" name="formS1" onsubmit="return chkSrchForm1()">
	<% 
	spThemeTitle = txtSearch & ":"
	spThemeBlock3_open() %>
    <div class="tPlain" style="text-align:center;">
	<input type="hidden" name="cmd" value="6" />
	<input type="text" name="search" size="15" style="margin-top:5px;margin-bottom:5px;" />
  <select name="mode" id="mode">
    <option value="0" selected="selected">All Articles</option>
    <option value="1">By Submitter</option>
    <option value="2">By Author</option>
  </select>
      <div class="fNorm" style="margin-bottom:3px;text-align:center;">
      <input type="submit" value=" <%= txtSearch %> " id="searchA" name="searchA" class="button" /></div></div><% spThemeBlock3_close() %></form>
<%spThemeBlock1_close(intSkin)
end sub
%>