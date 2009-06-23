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

sub template()
  if instr(sScript,app_page) > 0 then
    arg2 = ""
    shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	
	spThemeTitle = txtDownloads & " - " & txtCatMgr
	spThemeBlock1_open(intSkin)
  end if
  chkSessionMsg()
  
  if instr(sScript,app_page) > 0 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub deleteArticle(itm)
  bOk = false
  sSql = "SELECT " & strTablePrefix & "M_CATEGORIES.CG_FULL, " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL, ARTICLE.CATEGORY"
  sSql = sSql & " FROM ARTICLE INNER JOIN (" & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID) ON ARTICLE.CATEGORY = " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID"
  sSql = sSql & " WHERE (((ARTICLE.ARTICLE_ID)=" & itm & "));"
  set rsT = my_Conn.execute(sSql)
  if rsT.eof then
    ':: item not found "ARTICLE","ARTICLE_ID",
  else
    if bAppFull or hasAccess(rsT("CG_FULL")) or hasAccess(rsT("SG_FULL")) then
	  bOk = true
	  sucat = rsT("CATEGORY")
	end if
  end if
  set rsT = nothing
  
  if bOk then
    sSql = "DELETE FROM ARTICLE WHERE ARTICLE_ID = " & itm & ""
	executeThis(sSql)
    sSql = "DELETE FROM " & strTablePrefix & "M_RATING WHERE ITEM_ID = " & itm & " AND APP_ID = " & intAppID
	executeThis(sSql)
	mod_decreaseSubcatCount(sucat)
    sMsg = "Item, ratings and comments have been deleted"
  else
    sMsg = "Item was not found or you do not have the correct <br>permissions to perform the specified task"
  end if
	Call setSession("sMsg",sMsg)
end sub

sub chkSubCatAttention(i)
  tAtt = getCount("ARTICLE_ID","ARTICLE","ACTIVE = 0 AND CATEGORY=" & i)
  if tAtt > 0 then
	Response.Write "<a href="""&sScript&"?cmd=22&amp;sid=" & i & """>"
	Response.Write icon(icnAttention,"Items need approval","","","align=""middle"" hspace=""2""")
	Response.Write "</a>"
  end if
end sub

sub deleteArtCategory(c)
  sSql = mod_CatSubCatsql(c,0,intAppID)
  'sSql = sSql & "WHERE " & strTablePrefix & "M_CATEGORIES.CAT_ID=" & c & " AND " & strTablePrefix & "M_CATEGORIES.APP_ID = " & intAppID & ""
  
  set rsA = my_Conn.execute(sSQL)
  if not rsA.eof then
    cat= rsA("CAT_NAME")
	do until rsA.eof
  
      ':: delete ratings
      sSQL = "SELECT ARTICLE_ID FROM ARTICLE WHERE CATEGORY=" & rsA("SUBCAT_ID")
      set rsDel = my_Conn.execute(sSQL)
      do until rsDel.eof
  	    executeThis("DELETE from " & strTablePrefix & "M_RATING where ITEM_ID=" & rsDel("ARTICLE_ID") & " AND APP_ID = " & intAppID & "")
	    rsDel.movenext
	  loop
      set rsDel = nothing
	
      executeThis("DELETE FROM ARTICLE WHERE CATEGORY=" & rsA("SUBCAT_ID"))
	  rsA.movenext
	loop
    strMsg = strMsg & "Category (<span class=""fAlert"">" & cat & "</span>) along with all Sub-Categories"
    strMsg = strMsg & "<br />and associated data have been deleted.<br /><br />"
  else
    strMsg = strMsg & "<span class=""fAlert"">Category Not Found (" & cat & ")</span>"
  end if
  set rsA = nothing
  
  executeThis("DELETE FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & c & " AND APP_ID = " & intAppID & "")
  executeThis("DELETE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE CAT_ID=" & c & " AND APP_ID = " & intAppID & "")
  Call setSession("sMsg",strMsg)
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "")
end sub

sub deleteArtSubCategory(sc)
  sSql2 = "SELECT CAT_ID, SUBCAT_NAME From " & strTablePrefix & "M_SUBCATEGORIES where SUBCAT_ID=" & sc
  set rsDel = my_Conn.execute(sSQL2)
    c = rsDel("CAT_ID")
	scn = rsDel("SUBCAT_NAME")
  set rsDel = nothing
    sSQL = "SELECT * FROM ARTICLE WHERE CATEGORY=" & sc
    set rsDel = my_Conn.execute(sSQL)
     if not rsDel.eof then
      do until rsDel.eof
	    'delete ratings
  	    executeThis("DELETE from " & strTablePrefix & "M_RATING where ITEM_ID=" & rsDel("ARTICLE_ID") & " AND APP_ID = " & intAppID & "")
	    rsDel.movenext
	  loop
     end if
    set rsDel = nothing
    executeThis("delete From " & strTablePrefix & "M_SUBCATEGORIES where SUBCAT_ID=" & sc)
    executeThis("delete From ARTICLE where CATEGORY=" & sc)
    strMsg = strMsg & "Subcategory: <b>" & scn & "</b><br />"
    strMsg = strMsg & "and all its contents have been deleted."
	Call setSession("sMsg",strMsg)
	resetCoreConfig()
    closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & cid & "")
end sub

sub showAttentionSubCat(i)
  if instr(sScript,app_page) > 0 then
    arg2 = ""
    shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	
	spThemeTitle = "Attention Items"
    if sMode = 321 then
	  spThemeTitle = "Edit Item"
	end if
	spThemeBlock1_open(intSkin)
  end if
  chkSessionMsg()
 if sMode = 321 then
   mod_EditItemForm(intItemID)
 elseif sMode = 322 then
   processEditItemForm()
 else
  
  sSql = "SELECT PORTAL_M_CATEGORIES.*, PORTAL_M_SUBCATEGORIES.*"
  sSql = sSql & " FROM PORTAL_M_CATEGORIES INNER JOIN PORTAL_M_SUBCATEGORIES ON PORTAL_M_CATEGORIES.CAT_ID = PORTAL_M_SUBCATEGORIES.CAT_ID"
  sSql = sSql & " WHERE PORTAL_M_SUBCATEGORIES.SUBCAT_ID=" & i & ""
  sSql = sSql & " AND PORTAL_M_SUBCATEGORIES.APP_ID = " & intAppID & ";"
  set rsA = my_Conn.execute(sSql)
  if rsA.eof and not bAppFull then
    ':: no items found
    Response.Write "<p>&nbsp;</p>"
    Response.Write "<p class=""fTitle"" align=""center"">"
    Response.Write "Subcategory not found"
    Response.Write "</p>"
    Response.Write "<p>&nbsp;</p>"
  else
    hasFull = false
	if bAppFull then
	  hasFull = true
	else
     if hasAccess(rsA("CG_FULL")) OR hasAccess(rsA("SG_FULL")) then
	  hasFull = true
	 end if
	end if
    if hasFull then
	  select case sMode
	    case 122
		  Response.Write("Mode 122<br><br>")
		  sSql = "UPDATE " & item_tbl & " SET"
		  sSql = sSql & " ACTIVE=1"
		  sSql = sSql & ", POST_DATE='" & strCurDateString & "'"
		  'sSql = sSql & ", BADLINK=0"
		  sSql = sSql & " WHERE " & item_fld & "=" & cid
		  sSql = sSql & " AND CATEGORY=" & sid
		  'response.Write(sSql)
		  'response.end
		  executeThis(sSql)
		  Call setSession("sMsg","Item Approved")
		  closeAndGo(sScript & "?cmd=" & iPgType & "&sid=" & sid)
		  
		case 123 ':: delete new item or bad link
		  Response.Write("Mode 123")
		  sSql = "SELECT " & item_fld & " FROM " & item_tbl & " WHERE " & item_fld & " = " & cid
		  set rsA = my_Conn.execute(sSql)
		  if not rsA.eof then
		    isOK = true
		  else
		    isOK = false
			Call setSession("sMsg","Item not found")
		  end if
		  set rsA = nothing
		  if isOK then
		    sSql = "DELETE FROM " & item_tbl & ""
		    sSql = sSql & " WHERE " & item_fld & "=" & cid
		    sSql = sSql & " AND CATEGORY=" & sid
		    executeThis(sSql)
			Call setSession("sMsg","Item successfully deleted")
		  end if
		  closeAndGo(sScript & "?cmd=" & iPgType & "&sid=" & sid)
		  
		case 133
		  Response.Write("Mode 133")
		case else
	  end select
	  mod_writeApprovalJS()
  	  Response.Write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">"
  	  Response.Write "<tr><td width=""30%"" align=""right"">"
  	  Response.Write icon(imgAttention,"Attention","","","")
  	  Response.Write "</td><td>"
  	  Response.Write "Items that need approved are listed here"
  	  Response.Write "</td></tr>"
  	  Response.Write "</table>"
	  
  	  Response.Write "<hr/><br/>"
	  
  	  'Response.Write "Category: " & rsA("CAT_NAME") & "<br>"
  	  'Response.Write "SubCategory: " & rsA("SUBCAT_NAME") & "<br>"
	  
  	  Response.Write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"" class=""grid"">"
  	  Response.Write "<tr><td align=""center"" colspan=""4"" class=""tTitle"">"
  	  Response.Write "Items needing Approval"
  	  Response.Write "</td></tr>"
  	  Response.Write "<tr><td align=""center"" class=""tSubTitle"" width=""100"">Options"
  	  Response.Write "</td><td class=""tSubTitle"">Name"
  	  Response.Write "</td><td align=""center"" class=""tSubTitle"" width=""100"">Date"
  	  Response.Write "</td><td align=""center"" class=""tSubTitle"" width=""100"">By"
  	  Response.Write "</td></tr>"
	  sSql = mod_singleItemSql(item_tbl)
	  sSql = sSql & " WHERE " & item_tbl & ".ACTIVE=0"
	  sSql = sSql & " ORDER BY PORTAL_M_CATEGORIES.CAT_NAME, PORTAL_M_SUBCATEGORIES.SUBCAT_NAME, " & item_tbl & ".TITLE;"
	  'response.Write(sSql)
	  'closeAndGo("stop")
	  set rsN = my_Conn.execute(sSql)
	  if rsN.eof then
  	    Response.Write "<tr><td align=""center"" colspan=""4"" class=""fSubTitle"">"
  	    Response.Write "<br/>No items to approve<br/><br/>"
  	    Response.Write "</td></tr>"
	  else
	    do until rsN.eof
		 if hasAccess(rsN("CG_FULL")) OR hasAccess(rsN("SG_FULL")) OR bAppFull then
  	      Response.Write "<tr><td align=""center"">"
  	      Response.Write icon(icnCheck,"Approve","display:inline;cursor:pointer;","jsApprDl('" & rsN("TITLE") & "','" & rsN(item_fld) & "','" & rsN("SUBCAT_ID") & "',1)","")
  	      Response.Write icon(icnDelete,txtDel,"display:inline;cursor:pointer;","jsDelDl('" & rsN("TITLE") & "','" & rsN(item_fld) & "','" & rsN("SUBCAT_ID") & "')","")
  	      Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","jsEditDL('" & rsN(item_fld) & "','" & rsN("SUBCAT_ID") & "')","")
		  if isMac then
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","jsAttnDL('view" & rsN(item_fld) & "')","")
		  else
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","openJsLayer('view" & rsN(item_fld) & "','550','450')","")
		  end if
  	      Response.Write "</td><td>"
  	      Response.Write rsN("TITLE")
		  Call mod_writeViewItem(rsN,"showArticle")
  	      Response.Write "</td><td align=""center"">" & chkDate2(rsN("POST_DATE"))
  	      Response.Write "</td><td align=""center"">" & rsN("POSTER")
  	      Response.Write "</td></tr>"
		 end if
		 rsN.movenext
		loop
	  end if
	  set rsN = nothing
  	  Response.Write "</table>"
	  
	  Response.Write "<div id=""view_pane"">"
	  Response.Write "</div><br/>"
	  Call mod_shoLegend("admin",art_chkNew,art_chkUpdated)
	  'sSql = singleDLsql()
	  'sSql = sSql & " WHERE (((DL.DL_ID)=" & tmpID & "));"
	  'set rsV = my_Conn.execute(sSql)
	  'if rsV.eof then
	  'else
	  'end if
	  'set rsV = nothing
	  
	else
	  ':: NO ACCESS
      Response.Write "<p>&nbsp;</p>"
      Response.Write "<p class=""fTitle"" align=""center"">No Access</p>"
      Response.Write "<p>&nbsp;</p>"
	end if
  end if
  set rsA = nothing
 end if ':: if sMode = 321
  
  if instr(sScript,app_page) > 0 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub mainArticle() %>
  <%
      spThemeTitle = spThemeTitle & "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow5('" & app_pop & "?mode=print&amp;cid=" & articleID & "')"">" & icon(icnPrint,txtPrint,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
      if strUserMemberID > 0 and strEmail = 1 then
	    spThemeTitle = spThemeTitle & "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=emailitem&amp;cid=" & articleID & "')"">" & icon(icnEmail,"Email this to a friend","display:inline;","","align=""right"" hspace=""4""") & "</a>"
      end if
  
  	spThemeTitle = spThemeTitle & mod_iconBookmark(0,0,intItemID)
	If bAppFull Then 
  	  spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=" & hp & "&amp;cid=" & articleID & "')"">" & icon(sImg,sTxt,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
	End If

	spThemeTitle = spThemeTitle & "&nbsp;" & title & ""
	spThemeBlock1_open(intSkin)
	chkSessionMsg()

	showArticle(rs)
  %>
  <table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td colspan="2" align="center" class="fSubTitle">
  <% 
  If art_Comments or art_Rate Then
    pop = app_pop & "?mode=rate&amp;cid=" & articleid
    'popUpWind(pop,"spRate",400,500,"yes","yes")
	
    Response.Write("<a href=""javascript:;"" title=""Add Comment"" onClick=""popUpWind('" & pop & "','spRate',400,530,'yes','yes');"">" & icon(icnComment,"Comment","","","align=""bottom"" hspace=""5""") & "<b>Add&nbsp;")
    'Response.Write("<a href=""javascript:;"" title=""Add Comment"" onClick=""openWindow4('" & app_pop & "?mode=rate&amp;cid=" & articleid & "');"">" & icon(icnComment,"Comment","","","align=""bottom"" hspace=""5""") & "Add&nbsp;")
   If art_Comments Then
    response.Write("Comment")
	if art_Rate then
      response.Write("/")
	end if
   end if
   if art_Rate then
    response.Write("Rating")
   end if
   response.Write("</b></a>")
  end if
  %>
</td></tr>
  <% If art_Comments Then %>
<tr><td width="100%" colspan="2">
<% Call mod_GetComments(articleid,intAppID,app_rpage & "cmd=25") %>
</td></tr>
  <% End If %>
<tr><td width="100%" align="center" colspan="2" class="fNorm"><hr />
<br />
<center>
<a href="<%= app_page %>?cmd=6&amp;mode=2&amp;search=<%= server.URLEncode(strAuthor)%>">Other Articles by this Author</a><br /><br />
<a href="<%= chkString(Request.ServerVariables("HTTP_REFERER"), "refer") %>"><%= txtBack %></a>
</center>
</td></tr></table>
<%
end sub

sub showArticle(ob)
  isOwner = false
  bFull = false
  if bAppFull or hasAccess(ob("CG_FULL")) or hasAccess(ob("CG_FULL")) then
    bFull = true
  end if
  if bFull or (strDBNTUserName = ob("POSTER")) then
    isOwner = true
  end if
  mainContent = formatStr(ob("CONTENT"))
  'mainContent = chkString(mainContent, "display")
  
  %><hr />
<table border="0" class="tBorder" width="100%" cellspacing="1" cellpadding="6" align="center">
  <tr>
    <td class="tTitle">
	  <%
	  if isOwner then
	   if iPgType = 22 then
	    Response.Write "<a href=""" & app_page & "?cmd=22&amp;mode=321&amp;item="& ob("ARTICLE_ID") &""">"
	   else
	    Response.Write "<a href=""" & app_page & "?cmd=23&amp;item="& ob("ARTICLE_ID") &""">"
	   end if
	    Response.Write icon(icnEdit,"Edit Item","display:inline;","","align=""right""")
	    Response.Write "</a>"
	    Response.Write "<a href=""" & app_page & "?cmd=24&amp;item="& ob("ARTICLE_ID") &""">"
	    Response.Write "</a>"
	    Response.Write icon(icnDelete,"Delete Item","display:inline;cursor:pointer;","jsDelArt('" & title & "','" & ob("ARTICLE_ID") & "')","align=""right""")
	  end if
      Response.Write "<span class=""fTitle"">" & title & "</span>"

	  call chkNewItem(ob("POST_DATE"),art_chkNew,ob("UPDATED"),art_chkUpdated) %>
    </td>
  </tr>
  <tr>
    <td>
	  <% showInfoBlock(ob) %>
      <p><%= mainContent %></p>
    </td>
  </tr>
</table><hr />
  <%
end sub

sub showEditForm(oo)
  Response.Write "<form method=""post"" action=""" & sScript & """>"
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""4"" width=""100%"" align=""center"">"
  Response.Write "<tr><td align=""right"" width=""30%"">"
  Response.Write "Subcategory:&nbsp;</td>"
  Response.Write "<td>"
  mod_selectCatSubcat oo("SUBCAT_ID"),"WRITE"
  Response.Write "</td></tr>"
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<span class=""fAlert"">*</span>"
  Response.Write "Title:</td>"
  Response.Write "<td>"
  Response.Write "<input type=""text"" name=""title"" size=""40"" maxlength=""90"" value=""" & oo("TITLE") & """ />"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "Submitted by:</td>"
  Response.Write "<td>"
  Response.Write "<b>" & oo("POSTER") & "</b>"
  Response.Write "</td></tr>"
  customFormElements(oo)
  'Response.Write "<tr><td colspan=""2""><hr></td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<span class=""fAlert"">* </span>"
  Response.Write "Summary: <br /><br />"
  Response.Write "<span id=""charLeft"">" & 400-len(oo("SUMMARY")) & " characters left&nbsp;"
  Response.Write "</span></td>"
  Response.Write "<td>"
  Response.Write "<textarea rows=""9"" name=""sdes"" id=""sdes"" cols=""50"" wrap=""virtual"" onKeyUp=""cntChar('sdes','charLeft','{CHAR} characters left.',400);"">"
  Response.Write oo("SUMMARY")
  Response.Write "</textarea>"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<span class=""fAlert"">*</span> Article: ",oo("CONTENT")
  else
  	displayPLAINeditor 1,oo("CONTENT")
  end if
  
  Response.Write "<tr><td colspan=""2""><hr/>"
  Response.Write "</td></tr>"
  
  if oo("ACTIVE") = 0 then
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""approve"" value=""1"" checked=""checked"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Approve Item</b>"
    Response.Write "</td></tr>"
  else
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""marknew"" value=""1"" checked=""checked"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Mark as Updated</b>"
    Response.Write "</td></tr>"
  end if
  
  Response.Write "<tr><td colspan=""2""><hr/>"
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">&nbsp;</td>"
  Response.Write "<td><br />"
  Response.Write "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />"
  Response.Write "<input type=""hidden"" name=""mode"" value=""322"" />"
  Response.Write "<input type=""hidden"" name=""itemID"" value=""" & oo(item_fld) & """ />"
  Response.Write "<input type=""hidden"" name=""orig_subcat"" value=""" & oo("SUBCAT_ID") & """ />"
  Response.Write "<input id=""button"" class=""button"" type=""submit"" value="" Update "" style=""width:150px;height:25px;"" name=""B1"" accesskey=""s"" title=""Shortcut Key: Alt+S"" />"
  Response.Write "</td></tr>"
  Response.Write ""
  
  Response.Write "</table>"
  Response.Write "</form>"
end sub

sub processEditItemForm()
	cat = cLng(Request.Form("subcat"))
    itemID = clng(Request.Form("itemID"))
	title = replace(ChkString(Request.Form("title"),"sqlstring"), "'","''", 1, -1, 1)
	content = ChkString(ChkBadWords(Request.Form("Message")),"message")
	summary = ChkString(ChkBadWords(Request.Form("sdes")),"message")
	key = replace(ChkString(Request.Form("key"),"sqlstring"), "'","''", 1, -1, 1)
	poster = strDBNTUserName
	posteremail = replace(ChkString(Request.Form("posteremail"),"sqlstring"),"'","",1,-1,1)
	
	sT1Sql = ""

  if txtArtLabel1 <> "" then
    sTDATA1 = ChkString(Request.Form("TDATA1"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA1='" & sTDATA1 & "'"
  end if
  if txtArtLabel2 <> "" then
    sTDATA2 = ChkString(Request.Form("TDATA2"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2='" & sTDATA2 & "'"
  end if
  if txtArtLabel3 <> "" then
    sTDATA3 = ChkString(Request.Form("TDATA3"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3='" & sTDATA3 & "'"
  end if
  if txtArtLabel4 <> "" then
    sTDATA4 = ChkString(Request.Form("TDATA4"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4='" & sTDATA4 & "'"
  end if
  if txtArtLabel5 <> "" then
    sTDATA5 = ChkString(Request.Form("TDATA5"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5='" & sTDATA5 & "'"
  end if
  if txtArtLabel6 <> "" then
    sTDATA6 = ChkString(Request.Form("TDATA6"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6='" & sTDATA6 & "'"
  end if
  if txtArtLabel7 <> "" then
    sTDATA7 = ChkString(Request.Form("TDATA7"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7='" & sTDATA7 & "'"
  end if
  if txtArtLabel8 <> "" then
    sTDATA8 = ChkString(Request.Form("TDATA8"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8='" & sTDATA8 & "'"
  end if
  if txtArtLabel9 <> "" then
    sTDATA9 = ChkString(Request.Form("TDATA9"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9='" & sTDATA9 & "'"
  end if
  if txtArtLabel10 <> "" then
    sTDATA10 = ChkString(Request.Form("TDATA10"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10='" & sTDATA10 & "'"
  end if
    'marknew = Request.Form("marknew")
    approve = Request.Form("approve")
	if Request.Form("marknew") = 1 then
	  marknew = true
	else
	  marknew = false
	end if
	
	sString = ""
	if len(trim(title)) = 0 then
	  sString = sString & "<li>Please enter article title.</li>"
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
	  if len(trim(summary)) => 400 then
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
	if len(trim(author)) = 0 then 
	  'sString = sString & "<li>You must supply the Author's name.</li>"
	else
	end if
	if sString = "" then
	
      sSql = "UPDATE " & item_tbl & " set "
      sSql = sSql & "TITLE='" & title & "'"
      sSql = sSql & ",CATEGORY =" & cat & ""
      sSql = sSql & ",KEYWORD ='" & key & "'"
      sSql = sSql & ",SUMMARY ='" & summary & "'"
      sSql = sSql & ",CONTENT ='" & content & "'"
      sSql = sSql & ",POSTER_EMAIL ='" & posteremail & "'"
	  sSql = sSql & sT1Sql
      if marknew then
        sSql = sSql & ",UPDATED ='" & strCurDateString & "'"
      end if
      if approve = "1" then
        sSql = sSql & ",ACTIVE = 1"
      end if
      sSql = sSql & " where " & item_fld & " =" & itemID
      executeThis(sSql)
  
      Call setSession("sMsg","Item successfully updated")
	  if approve = "1" then
	    mod_increaseSubcatCount(cat)
	  end if
	  if approve = "1" and intSubscriptions = 1 and strEmail = 1 then
	    'send subscriptions emails
		sSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & cat
		set rsA = my_Conn.execute(sSql)
		  parent = rsA(0)
		set rsA = nothing
	    eSubject = strSiteTitle & " - New Article"
		eMsg = "A new article has been submitted at " & strSiteTitle & vbCrLf
		eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		eMsg = eMsg & "You can view the new articles by visiting " & strHomeUrl & app_page & "?cmd=3" & vbCrLf
	    sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
		'response.Write("<br>Email sent<br>" )
	  end if
	  
  	  if iPgType = 23 then
    	closeAndGo(app_rpage & "item=" & itemID)
  	  else
    	closeAndGo(sScript & "?cmd=" & iPgType & "&sid=" & subcat)
  	  end if
	else
	  ':: required fields not filled in
	  Response.Write("Required fields not filled in<br><br>")
	  Response.Write(sString)
	end if
end sub
%>
