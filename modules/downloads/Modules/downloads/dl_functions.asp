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

item_tbl = "DL"
item_fld = "DL_ID"
app_pop = "dl_pop.asp"
app_page = "dl.asp"
app_rpage = "dl.asp"
app_admin = "admin_dl_admin.asp"
app_addForm = "dl_add_form.asp?"
skyPage_iName = "downloads"

%>
<!-- #include file="dl_admin.asp" -->
<!-- #include file="dl_custom.asp" -->
<% 
if len(dl_InvalidIPs & "x") > 0 then
  if not isValidIP(dl_InvalidIPs) then
    closeAndGo("error.asp?type=noperm")
  end if
end if

function singleDLsql()
  tSql = "SELECT " & strTablePrefix & "M_CATEGORIES.*"
  tSql = tSql & ", " & strTablePrefix & "M_SUBCATEGORIES.*"
  tSql = tSql & ", DL.*"
  tSql = tSql & " FROM (" & strTablePrefix & "M_CATEGORIES"
  tSql = tSql & " INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES"
  tSql = tSql & " ON " & strTablePrefix & "M_CATEGORIES.CAT_ID"
  tSql = tSql & " = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID)"
  tSql = tSql & " INNER JOIN DL"
  tSql = tSql & " ON " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID"
  tSql = tSql & " = DL.CATEGORY "
  singleDLsql = tSql
end function

function GetNewDL(daysShown)
	dim i
	for i = 0 to daysShown - 1
		curDate = dateadd("d",-i,strCurDateAdjust)
		strSQL = "SELECT count(DL_ID) as DLCOUNT FROM DL WHERE POST_DATE LIKE '" & left(DateToStr(curDate),8) & "%' AND ACTIVE = 1"
		set rsDay = server.CreateObject("adodb.recordset")
		rsDay.Open strSQL, my_Conn
		%>
		  <div class="tPlain" style="padding: 4px;">
		    <span style="width: 50px; text-align: right;">&#149;</span>
		    <span style="width: 240px;">&nbsp;<a href="dl.asp?cmd=3&amp;daysago=<%= i %>"><span class="fNorm"><%= formatdatetime(curDate,1) %></span></a></span>
		    <span style="width: 50px; text-align: right;">&nbsp;(<%=rsDay("DLCOUNT")%>)</span>
		  </div>
		<%
		rsDay.Close
		set rsDay = nothing
	next
end function

sub showtoprated()
  dim intTop, cnt
  cnt = 1
  arg1 = txtDownloads & "|" & sScript
  arg2 = txtTopRDLs & "|dl.asp?cmd=5"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intTop = "" or not isnumeric(intTop) then intTop = 10
  spThemeTitle= "<b>" & txtTopRDL &"</b>"
  spThemeBlock1_open(intSkin)
  'response.Write("<table><tr><td>")
  'strSQL = "SELECT * FROM DL WHERE SHOW = 1 AND VOTES > 0 ORDER BY ROUND(RATING/VOTES, 0) DESC, VOTES DESC"
  sSql = getDlSql_sm("rated")

  dim rsPop
  set rsPop = server.CreateObject("adodb.recordset")
  rsPop.Open sSql, my_Conn
  If rsPop.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No Downloads have been rated!</span>"
  else
    Do While Not rsPop.EOF and cnt <= intTop
	  if hasAccess(rsPop("SG_READ")) or hasAccess(rsPop("CG_FULL")) or bAppFull then
	    Call DisplayDL(rsPop)
	    cnt = cnt + 1
	  end if
	  rsPop.MoveNext
    Loop
  end if

  rsPop.Close
  Set rsPop = Nothing
  'response.Write("</td></tr></table>")
  spThemeBlock1_close(intSkin)
end sub

sub showpopular()
  dim intPopular, cnt
  cnt = 1
  intPopular = 10
  arg2 = txtPopDL & "|dl.asp?cmd=4"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intPopular = "" or not isnumeric(intpopular) then intPopular = 10
  spThemeTitle= "<b>Popular Downloads by hit count</b>"
  spThemeBlock1_open(intSkin)

  'strSQL = "SELECT * FROM DL WHERE SHOW = 1 and HIT > 0 ORDER BY HIT DESC"
  sSql = getDlSql_sm("top")

  dim rsPop
  set rsPop = my_Conn.execute(sSql)
  If rsPop.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No Downloads have been downloaded!</span>"
  else
    Do While Not rsPop.EOF and cnt <= intPopular
	  if hasAccess(rsPop("SG_READ")) or hasAccess(rsPop("CG_FULL")) or bAppFull then
	    Call DisplayDL(rsPop)
	    cnt = cnt + 1
	  end if
	  rsPop.MoveNext
    Loop
  end if
  Set rsPop = Nothing
  spThemeBlock1_close(intSkin)
end sub

sub shownew()
  dim intPopular, cnt
  cnt = 1
  intPopular = 10
  arg2 = txtNewDL & "|dl.asp?cmd=3"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intPopular = "" or not isnumeric(intpopular) then intPopular = 10
  sSql = getDlSql_sm("new")
  spThemeTitle= "<b>Newest Downloads</b>"
  spThemeBlock1_open(intSkin)

  dim rsPop
  set rsPop = my_Conn.execute(sSql)
  If rsPop.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No Downloads are available!</span>"
  else
    Do While Not rsPop.EOF and cnt <= intPopular
	  if hasAccess(rsPop("SG_READ")) or hasAccess(rsPop("CG_FULL")) or bAppFull then
	    Call DisplayDL(rsPop)
	    cnt = cnt + 1
	  end if
	  rsPop.MoveNext
    Loop
  end if
  Set rsPop = Nothing
  spThemeBlock1_close(intSkin)
end sub

sub shownew2()
  dim intDaysAgo, rsDay
  if Request.QueryString("daysago") <> "" or  Request.QueryString("daysago") <> " " then
	if IsNumeric(Request.QueryString("daysago")) = True then
		intDaysAgo = chkString(Request.QueryString("daysago"),"sqlstring")
	else
		closeAndGo(sScript)
	end if
  end if
  
  arg2 = txtNewDl & "|dl.asp?cmd=3"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()
  
  if intDaysAgo <> "" then
	curDate = dateadd("d",-intDaysAgo,strCurDateAdjust)
	spThemeTitle= "<b>Total new Downloads on " &split(curDate," ")(0)&".</b>"
	spThemeBlock1_open(intSkin)
	strSQL = "SELECT * FROM DL WHERE POST_DATE LIKE '" & left(DateToStr(curDate),8) & "%' AND ACTIVE = 1 ORDER BY CATEGORY"
	set rsDay = server.CreateObject("adodb.recordset")
	rsDay.Open strSQL, my_Conn
	If rsDay.eof Then
		Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No files found!</span>"
	else		
	  Do While Not rsDay.EOF
		Call DisplayDL(rsDay)
		rsDay.MoveNext
	  Loop
	end if
	rsDay.Close
	Set rsDay = Nothing
	spThemeBlock1_close(intSkin)
  else
	dim intDaysShown
	intDaysShown = 7
	spThemeTitle= "<b>Total new Downloads for last " &intDaysShown&" Days.</b>"
	spThemeBlock1_open(intSkin)
 	  GetNewDL(intDaysShown)
  	spThemeBlock1_close(intSkin)
  end if 
end sub

sub showall()

  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  'app_MainColumn_top()
  mod_displayIntro(intAppID)
  spThemeTitle = spThemeTitle & mod_iconSubscribe(0,0)
  
  If bAppFull Then
	spThemeTitle = spThemeTitle & modGrpEdit(app_pop,14,0,0,"right",2)
	spThemeTitle = spThemeTitle & "<a href="""&sDLpage&"?cmd=20&cid="&cid&""" title=""Category Manager"">"
	spThemeTitle = spThemeTitle & icon(icnToolbox,"Category Manager","display:inline","","align=""right""") & "</a>"
	if dlAttention > 0 then
	  spThemeTitle = spThemeTitle & "<a href="""&sDLpage&"?cmd=22&amp;sid=" & sub_id & """>"
	  spThemeTitle = spThemeTitle & icon(icnAttention,"Items need attention","","","align=""right""") & "</a>"
	end if
  end if
  
  spThemeTitle = spThemeTitle & txtDownloads
  spThemeBlock1_open(intSkin)
  'strSql = "SELECT * FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID & " ORDER BY C_ORDER, CAT_NAME"
  'set rsCategories = server.CreateObject("adodb.recordset")
  'rsCategories.Open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
  
  response.Write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" width=""100%"">"
  
  dim rsAll
  sSql = mod_CatSubCatsql(0,0,intAppID)
  'Set rsAll = oSpData.GetRecordset(sSql)
  Set rsAll = my_Conn.execute(sSql)
  if rsAll.eof then
    ':: no records found
  else
   Do until rsAll.EOF
    response.Write "<tr>"
	ColNum = 1 
	Do while ColNum < 3
	  blkTimer = timer
	  if not rsAll.EOF then
	    curCat = rsAll(sMCPre & "CAT_ID")
		if hasAccess(trim(rsAll("CG_READ"))) then
		  Response.Write "<td align=""left"" valign=""top"" width=""50%"">"
  		  If hasAccess(trim(rsAll("CG_FULL"))) or bAppFull Then 
			sTo = sDLpage&"?cmd=21&amp;cid=" & curCat
			Response.Write "<a href="""&sTo&""" title=""Category Manager"">"
			Response.Write icon(icnToolbox,"Category Manager","","","")
			Response.Write "</a>"
			Response.Write modGrpEdit("dl_pop.asp",14,curCat,0,"bottom",rsAll("CG_INHERIT"))
		  else
		    Response.Write icon("images/icons/icon_folder_new_topic.gif","","","","")
  		  end if
		  Response.Write "<a href=""" & sDLpage & "?cmd=1&amp;cid=" & curCat & """>"
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
  Call mod_shoLegend("main",dl_chkNew,dl_chkUpdated)
	  
  spThemeBlock1_close(intSkin)
  rsAll.close
  set rsAll = nothing
end sub

'sub shoCatSubcats(c,cf)
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
	  'rcounts = getCount("DL_ID","DL","CATEGORY=" & subcatID & " AND ACTIVE = 1")
	  isOK = false
	 if rcounts > 0 or (rcounts = 0 and (dl_ShowEmptySubs or bSCatFull)) then
	  isOK = true
	 end if
	 if isOK then
	  Response.Write icon(icnBar,"","","","")
 	  If bSCatFull Then 
		'Response.Write("&nbsp;" & modGrpEdit("dl_pop.asp",14,parent_id,rsSubcat("subcat_id"),"middle",rsSubcat("SG_INHERIT")))
		sTo = ""&sScript&"?cmd=21&amp;cid=" & parent_id
		response.Write(modGrpEdit(sTo,,,,"middle",ob("SG_INHERIT")))
		if rcounts > 0 then
		  chkSubCatAttention(subcatID)
		end if
	  else
	    Response.Write(icon("images/icons/img_dl.gif","","","","align=""middle"""))
  	  end if
	  %>
	  <a href="<%= sScript %>?cmd=2&amp;cid=<%=c%>&amp;sid=<%= subcatID %>"><span class="fNorm"><%= ob("SUBCAT_NAME") %>&nbsp;(<%= rcounts %>)</span></a>
	  <%
	  if rcounts > 0 then
	    call mod_chkNewSubCatItems(subcatID,dl_chkNew,dl_chkUpdated)
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
     closeandgo(sDLpage)
   end if
  
  arg2 = cat_name
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()
  
  spThemeTitle = spThemeTitle & mod_iconSubscribe(cid,0)
  spThemeTitle = spThemeTitle & mod_iconBookmark(cid,0,0)
  If bCatFull Then 
	spThemeTitle = spThemeTitle & "<a href="""&sScript&"?cmd=21&cid="&cid&""" title=""Category Manager"">"
	spThemeTitle = spThemeTitle & icon(icnToolbox,"Category Manager","display:inline","","align=""middle""") & "</a>"
	spThemeTitle = spThemeTitle & modGrpEdit("dl_pop.asp",14,cid,0,"middle",inherit)
  else
    spThemeTitle = spThemeTitle & icon(icnNewFolder,"","","","align=""middle"" hspace=""4""")
  end if
  spThemeTitle = spThemeTitle & "&nbsp;" & cat_name
  spThemeBlock1_open(intSkin)
  Response.Write("<table border=""0"" cellpadding=""4"" cellspacing=""0"" width=""100%"">")
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
	  sSQL = "SELECT count(DL_ID) FROM DL where category=" & rsC("SUBCAT_ID") & " and ACTIVE=1"
	  Set RScount = my_Conn.Execute(sSQL)
	  rCount = RScount(0)
  	  Set RScount = nothing
	 if rCount > 0 or (rCount = 0 and (dl_ShowEmptySubs or bSCatFull)) then
	  Response.Write "<tr>"
      Response.Write "<td align=""left"" valign=""top"">"
	  Response.Write icon(icnSpacer,"","","","width=""15""")
	  Response.Write icon(icnBar,"","","","hspace=""3""")
	  If bSCatFull Then
		chkSubCatAttention(rsC("SUBCAT_ID"))
	    response.Write(modGrpEdit("dl_pop.asp",14,cid,rsC("SUBCAT_ID"),"middle",rsC("SG_INHERIT")))
	  else
	    Response.Write icon(icnNewFolder,"","","","align=""middle"" hspace=""4""")
	  end if
      Response.Write("&nbsp;<a href=""" & sDLpage & "?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & rsC("SUBCAT_ID") & """>")
	  Response.Write("<span class=""fSubTitle"">")
	  Response.Write(ChkString(rsC("subcat_name"), "display"))
	  Response.Write("</span></a>")
	  Response.Write " (" & rCount & ")&nbsp;&nbsp;"
	  
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
  arg2 = cat_name & "|dl.asp?cmd=1&amp;cid=" & cat_id
  arg3 = sub_name & "|dl.asp?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & sub_id
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  sSQL = "SELECT * From DL where Category=" & sub_id & " and ACTIVE = 1"
  ord1 = chkString(request("ord1"),"sqlstring")
  ord2 = chkString(request("ord2"),"sqlstring")
  ord = ord1 & ord2
  select case ord
    case "hDesc"
      sSQL = sSQL & " ORDER BY DL.HIT DESC;"
    case "hAsc"
      sSQL = sSQL & " ORDER BY DL.HIT;"
    case "dDesc"
	  sSQL = sSQL & " ORDER BY DL.O_POST_DATE DESC;"
    case "dAsc"
	  sSQL = sSQL & " ORDER BY DL.O_POST_DATE;"
    case "rDesc"
	  sSQL = sSQL & " ORDER BY DL.RATING DESC;"
    case "rAsc"
	  sSQL = sSQL & " ORDER BY DL.RATING;"
    case "tDesc"
	  sSQL = sSQL & " ORDER BY DL.NAME DESC;"
    case "tAsc"
	  sSQL = sSQL & " ORDER BY DL.NAME;"
    case else
	  ord = "dDesc"
	  sSQL = sSQL & " ORDER BY DL.O_POST_DATE DESC;"
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
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " download"&tPgCnt&")"
else
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " downloads"&tPgCnt&")"
end if

  spThemeTitle = spThemeTitle & mod_iconSubscribe(0,sub_id)
  spThemeTitle = spThemeTitle & mod_iconBookmark(0,sub_id,0)
  
  If bSCatFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit(app_pop,14,cat_id,sub_id,"right",inherit)
  end if
  
  If iPageCount = 0 Then
    call showMsgBlock(1,"No items found!") 
  Else
    spThemeTitle = spThemeTitle & "&nbsp;" & tTitle
    spThemeBlock1_open(intSkin)
	chkSessionMsg()
  'response.write(ssSQL & "<br />" & objPagingRS("TITLE") & "<br />" & iPageCount & "<br />")
	objPagingRS.AbsolutePage = iPageCurrent
	if iPageCount > 1 then
	  showDaPaging iPageCurrent,iPageCount,0
	end if
	iRecordsShown = 0
	rCount = 0
	response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	response.Write("<tr><td align=""center"" valign=""top"">")
	response.Write("<form method=""post"" action=""dl.asp"">")
	response.Write("<input name=""cmd"" type=""hidden"" value=""" & iPgType & """>")
	response.Write("<input name=""cid"" type=""hidden"" value=""" & cat_id & """>")
	response.Write("<input name=""sid"" type=""hidden"" value=""" & sub_id & """>")
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
	response.Write("</td></tr>")
	response.Write("<tr>")
	response.Write("<tr><td align=""center"" valign=""top""><hr />")
	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
	
		call displayDL(objPagingRS)
	    iRecordsShown = iRecordsShown + 1
	    objPagingRS.MoveNext
	Loop
	response.Write("</td></tr></table>")

    objPagingRS.Close
    Set objPagingRS = Nothing
    if iPageCount > 1 then
      showDaPaging iPageCurrent,iPageCount,2
    end if
    If bSCatWrite Then
    %>
    <center>
    <hr /><a href="dl_add_form.asp?cat_id=<%=sub_id%>&amp;cat_name=<%=sub_name%>&amp;parent_id=<%=cat_id%>&amp;parent_name=<%=cat_name%>">
  Submit a Download</a>
    </center>
    <br />
    <%
	end if
    spThemeBlock1_close(intSkin)
  End If
else ':: no access so redirect
  closeandgo("dl.asp")
end if
end function

function doSearch()
  search = ChkString(Request("search"), "SQLString")
  show = 10
  if request("num") <> "" then
    show = clng(Request("num"))
  end if
  if show <> "" then
	Dim iPageSize       
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
	if sMode <> 1 then 'search all
  	  strSQL = "select * from DL where KEYWORD like '%" & search & "%' or DESCRIPTION like '%" & search & "%' or TDATA4 like '%" & search & "%' or NAME like '%" & search & "%' or TDATA2 like '%" & search & "%' or TDATA3 like '%" & search & "%' or TDATA1 like '%" & search & "%' or FILESIZE like '%" & search & "%' and ACTIVE=1 order by HIT DESC, DL_ID DESC"
	  strSrchTxt = "Search results for"
	else ':: search member submitted downloads
	  'srchMemberID = getMemberId(search)
      strSQL = "select * from DL where UPLOADER = '" & search & "' or UPLOADER like '%" & search & "%' and ACTIVE=1 order by UPLOADER, HIT DESC, DL_ID DESC"
	  strSrchTxt = "Items submitted by"
	end if
	

	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.PageSize = iPageSize
	objPagingRS.CacheSize = iPageSize
	objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	reccount = objPagingRS.recordcount
	iPageCount = objPagingRS.PageCount

	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
	If iPageCurrent < 1 Then iPageCurrent = 1

  	arg2 = strSrchTxt & ": " & search & "|javascript:;"
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  	app_MainColumn_top()

	spThemeBlock1_open(intSkin)

	If iPageCount = 0 Then
		Response.Write "<br /><div class=""fTitle"" class=""text-align:center;"">"
		Response.Write "<b>" & strSrchTxt & ": """ & search & """<br />returned no results!</b></div><br />"
	Else
		objPagingRS.AbsolutePage = iPageCurrent %>
		<center><div class="fTitle"><b><%= strSrchTxt %> :&nbsp;</b><span class="fAlert"><b><%=search%></b></span></div>
		<!-- <span class="fAlert"> found <% 'reccount%> item(s)</span> --></center><%
		if iPageCount > 1 then
	  	showDaPaging iPageCurrent,iPageCount,0
		end if
		iRecordsShown = 0
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		  'if hasAccess(objPagingRS("SG_READ")) then
			call displayDL(objPagingRS)
			iRecordsShown = iRecordsShown + 1
		  'end if
		  objPagingRS.MoveNext
		Loop
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

sub showItem()
  'set page size
  iPageSize = 1
  iPageCurrent = 1

  If Request("page") = "" Then
	iPageCurrent = 1
	bFirst = true
  Else
	iPageCurrent = cLng(Request("page"))
	bFirst = false
  End If
  
  if bFirst then
	strLSQL = "SELECT * FROM DL WHERE DL_ID = " & cat_id
	set rsT = my_Conn.execute(strLSQL)
	if not rsT.eof then
	  dl_id = cat_id
	  'cat_id = rsT("PARENT_ID")
	  sub_id = rsT("CATEGORY")
	else
	  dl_id = 0
	end if
	set rsT = nothing
  end if
  
	sSql = "SELECT " & strTablePrefix & "M_CATEGORIES.*, " & strTablePrefix & "M_SUBCATEGORIES.*, DL.* "
	sSql = sSql & " FROM (" & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID) INNER JOIN DL ON " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID = DL.CATEGORY"
	sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID)=" & sub_id & ") AND ((" & strTablePrefix & "M_SUBCATEGORIES.APP_ID) = " & intAppID & "));"
	'sSQL = sSQL & " WHERE ((" & strTablePrefix & "M_CATEGORIES.APP_ID) = " & intAppID & ");"
	
	'strLinkSQL = "SELECT * FROM DL WHERE DL_ID = " & cat_id
	
	Set rsItem = Server.CreateObject("ADODB.Recordset")
  	rsItem.PageSize = iPageSize
  	rsItem.CacheSize = iPageSize
	rsItem.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  	reccount = rsItem.recordcount
  	iPageCount = rsItem.PageCount

  	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  	If iPageCurrent < 1 Then iPageCurrent = 1
	
	if rsItem.EOF then
	  call showMsgBlock(1,"Download does not exist.")
	else	
	  
		  if bFirst then
		    pc = 0
		    foundit = false
		    do until foundit = true
			  pc = pc + 1
		      if rsItem("DL_ID") = dl_id then
			    foundit = true
				iPageCurrent = pc
		      end if
			  rsItem.movenext
		    loop
		  else
		  end if
		    rsItem.AbsolutePage = iPageCurrent
		  
	  cat_name = rsItem("CAT_NAME")
	  sub_name = rsItem("SUBCAT_NAME")
		s_id = rsItem("CATEGORY")
		c_id = rsItem(sMSPre & "CAT_ID")
	  call setPermVars(rsItem,2)
	  
	  if not bSCatRead then
	    closeandgo("dl.asp")
	  else

  		arg2 = cat_name & "|dl.asp?cmd=1&amp;cid=" & c_id
  		arg3 = sub_name & "|dl.asp?cmd=2&amp;cid=" & c_id & "&amp;sid=" & s_id
  		'arg4 = rsItem("NAME") & "|dl.asp?cmd=6&amp;cid=" & dl_id
  		arg5 = ""
  		arg6 = ""
  
  		shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  		app_MainColumn_top()
		
		'strBannerURL = rsLink("BANNER_URL")
		hp = rsItem("FEATURED")
		if hp = 0 then
	  	  hp = 1
	  	  sTxt = "Make this a Featured Download'"
	  	  sImg = icnFeature
		else
	  	  hp = 2
	  	  sTxt = "Remove from Featured Downloads"
	  	  sImg = icnUnfeature
		end if
		
		'spThemeTitle = spThemeTitle & "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow5('dl_pop.asp?mode=9&amp;cid=" & intDLID & "')"">" & icon(icnPrint,txtPrint,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
		if strUserMemberID > 0 and strEmail = 1 then
		spThemeTitle = "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow('dl_pop.asp?mode=8&amp;cid=" & rsItem("DL_ID") & "')"">" & icon(icnEmail,"Email this Download to a friend","display:inline;","","align=""right"" hspace=""4""") & "</a>"
		end if
  If strUserMemberID > 0 and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,"0","0",rsItem("DL_ID"),strUserMemberID)
	  if bookmark_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('dl_pop.asp?mode=10&amp;cid=" & bookmark_id & "');"">" & icon(icnUnBookmark,"Remove bookmark for this item","","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('dl_pop.asp?mode=3&amp;cmd=3&amp;cid=" & rsItem("DL_ID") & "');"">" & icon(icnBookmark,"Bookmark this item","","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>" 
	  end if
  end if
		'If bSCatFull Then 
		  'spThemeTitle = spThemeTitle & "<a href=""admin_dl_editurl.asp?id=" & rsItem("DL_ID") & """><img border=""0"" src=""images/icons/icon_edit_topic.gif"" align=""right"" title=""Edit Download"" alt=""Edit Download"" style=""display:inline;"" hspace=""4"" /></a>"
		'end if
	  
		If bAppFull Then 
		  spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('dl_pop.asp?mode=" & hp & "&amp;cid=" & rsItem("DL_ID") & "')"">" & icon(sImg,sTxt,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
		End If
		spThemeTitle = spThemeTitle & "&nbsp;" & rsItem("NAME")
		spThemeBlock1_open(intSkin)
		chkSessionMsg()
		  if iPageCount > 1 then
	  	    showDaPaging iPageCurrent,iPageCount,0
		  end if
		'displayDL()
		showInfo(rsItem)
		Call mod_GetComments(rsItem("DL_ID"),intAppID,"dl.asp?cmd=6&amp;mode=99")
		spThemeBlock1_close(intSkin)
	  end if
	end if ':: rsItem.eof check
	set rsItem = nothing
end sub

sub showInfo(ob)
  isOwner = false
  bFull = false
  if bAppFull or hasAccess(ob("CG_FULL")) or hasAccess(ob("CG_FULL")) then
    bFull = true
  end if
  if bFull or (strDBNTUserName = ob("UPLOADER")) then
    isOwner = true
  end if
  temp1 = "n/a"
  strDescription = ob("CONTENT")
  strDLName = ob("NAME")
  intDLID = ob("DL_ID")
  %><hr />
<table border="0" class="tBorder" width="100%" cellspacing="1" cellpadding="6" align="center">
  <tr>
    <td class="tTitle">
	  <%
	  if isOwner then
	   if iPgType = 22 then
	    Response.Write "<a href=""" & sDLpage & "?cmd=22&amp;mode=321&amp;item="& intDLID &""">"
	   else
	    Response.Write "<a href=""" & sDLpage & "?cmd=23&amp;item="& intDLID &""">"
	   end if
	    Response.Write icon(icnEdit,"Edit Item","display:inline;","","align=""right""")
	    Response.Write "</a>"
	    Response.Write "<a href=""" & sDLpage & "?cmd=24&amp;item="& intDLID &""">"
	    Response.Write icon(icnDelete,"Delete Item","display:inline;","","align=""right""")
	    Response.Write "</a>"
	  end if
		
	  chkItemAttention(ob)
	
	   if not strUserMemberID > 0 then %>
      <%=strDLName%>
	<% else %>
      <b><a href="javascript:openWindow5('dl_pop.asp?mode=4&amp;cid=<%=intDLID%>');"><span class="fTitle"><%=strDLName%></span></a></b>
	<% end if
	   call chkNewItem(ob("POST_DATE"),dl_chkNew,ob("UPDATED"),dl_chkUpdated) %>
    </td>
  </tr>
  <tr>
    <td>
	  <% showInfoBlock(ob) %>
      <p><%=strDescription%></p><br />
    </td>
  </tr>
</table><hr />

<table class="grid" width="100%" border="0" cellspacing="0" cellpadding="5">
<% shoDLGridInfo(ob) %>
<% If strDBNTUsername <> "" or dl_GuestsCanDL Then %>
<tr>
<td colspan="2" align="center" class="fSubTitle">
<% If canDL Then %>
	<br /><a href="javascript:;" onClick="popUpWind('dl_pop.asp?mode=4&amp;cid=<%=intDLID%>','spRate',200,200,'yes','yes');"><%= icon(icnDL,"Download Now!","","","align=""bottom""") %>&nbsp;<span class="fTitle"><b>Click here to download</b></span></a><br /><br />
<% End If %>
</td></tr>
<tr><td colspan="2" align="center" class="fNorm">
  <% 
  If (dl_Comments or dl_Rate) and strDBNTUsername <> "" Then
    Response.Write("<a href=""javascript:;"" title=""Add Comment"" onClick=""popUpWind('dl_pop.asp?mode=rate&amp;cid=" & intDLID & "','spRate',400,530,'yes','yes');"">" & icon(icnComment,"Comment","","","align=""bottom"" hspace=""5""") & "<b>Add&nbsp;")
   If dl_Comments Then
    response.Write("Comment")
	if dl_Rate then
      response.Write("/")
	end if
   end if
   if dl_Rate then
    response.Write("Rating")
   end if
   response.Write("</b></a>")
  end if
  %>&nbsp;|&nbsp;<a href="javascript:;" title="Report bad URL" onClick="openWindow4('dl_pop.asp?mode=6&amp;cid=<%=intDLID%>');"><b><%= icon(icnAttention,"Report bad URL","","","align=""bottom"" hspace=""5""") %>Report bad URL</b></a>
</td></tr>
<% else %>
<tr><td colspan="2" align="center" class="fNorm">
<span class="fAlert">You must be <a href="policy.asp"><span class="fAlert"><u>registered</u></span></a> and logged in to download this item.</span>
</td></tr>
<% end if %>
</table>
<%
end sub

sub showDaPaging(nPageTo,nPageCnt,nPaging)
	'Display Paging Buttons
				Response.Write("<center><table border=""0"" cellpadding=""4"" cellspacing=""4"">")
					if (nPageCnt > totSho) and nPaging = 1 then
					  Response.Write("<tr>")
						Response.Write("<td colspan=""5"" align=""center""><span class=""fSmall""><b>Page <span class=""fAlert"">" &  nPageTo & "</span> of <span class=""fAlert"">" & nPageCnt & "</span></b></span>")
						Response.Write("</td>")
					  Response.Write("</tr>")
					end if
					' Display <<
						Response.Write(vbCrLf & "<tr><td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" name=""formP"&nPaging&"01"" id=""formP"&nPaging&"01"">")
						If int(nPageTo) = 1 Then 
							Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold}"" disabled=""disabled"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2"" /><input type=""hidden"" name=""page"" value=""1"" />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold;cursor:pointer;}"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2""><input type=""hidden"" name=""page"" value=""1"" />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""mode"" value=""" & sMode & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cid"" value=""" & cat_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sub_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""search"" value=""" & search & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord1"" value=""" & chkString(request("ord1"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord2"" value=""" & chkString(request("ord2"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display <
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" name=""formP"&nPaging&"02"" id=""formP"&nPaging&"02"">")
						If int(nPageTo) = 1 Then 
							Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""1"" />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold;cursor:pointer;}"" />")
							Response.Write(vbCrLf & "<input type=""hidden"" name=""page"" value=""" & nPageTo-1 & """ />")
						End If
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""mode"" value=""" & sMode & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cid"" value=""" & cat_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sub_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""search"" value=""" & search & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord1"" value=""" & chkString(request("ord1"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord2"" value=""" & chkString(request("ord2"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display >
					      strQryStr = ""
						  if sMode <> "" then
						    strQryStr = strQryStr & "&amp;mode=" & sMode
						    strMode = "&amp;mode=" & sMode
						  end if
						  if request("ord1") <> "" and request("ord2") <> "" then
						    strQryStr = strQryStr & "&amp;ord1=" & chkString(request("ord1"),"sqlstring")
						    strQryStr = strQryStr & "&amp;ord2=" & chkString(request("ord2"),"sqlstring")
						  end if
						  if search <> "" then
						    strQryStr = strQryStr & "&amp;search=" & search
						  end if
						if nPageCnt > 1 then
						  Response.Write("<td align=""center"" class=""fNorm"">")
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
							  Response.Write("&nbsp;<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?cmd=" & iPgType & "&amp;cid=" & cat_id & "&amp;sid=" & sub_id & "&amp;page=" & pgc & strQryStr & """>")
						      Response.Write("<span class=""fBold"">" & pgc & "</span></a>")
							end if
						  next
						  Response.Write("&nbsp;</td>")
						end if
						
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" id=""formP"&nPaging&"03"" name=""formP"&nPaging&"03"">")
						If int(nPageTo) = nPageCnt Then 
							Response.Write(vbCrLf & "<input type=""submit"" value='  Next &gt;  ' id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""" & nPageTo & """ />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value=""  Next &gt;  "" id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold;cursor:pointer;}"" />")
							Response.Write(vbCrLf & "<input type=""hidden"" name=""page"" value=""" & nPageTo+1 & """ />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""mode"" value=""" & sMode & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cid"" value=""" & cat_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sub_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""search"" value=""" & search & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord1"" value=""" & chkString(request("ord1"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord2"" value=""" & chkString(request("ord2"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display >>
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" id=""formP"&nPaging&"04"" name=""formP"&nPaging&"04"">")
						If int(nPageTo) = nPageCnt Then 
							Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""page"" value=""" & nPageTo & """ />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold;cursor:pointer;}"" /><input type=""hidden"" name=""page"" value=""" & nPageCnt & """ />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""mode"" value=""" & sMode & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cid"" value=""" & cat_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sub_id & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""search"" value=""" & search & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord1"" value=""" & chkString(request("ord1"),"sqlstring") & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""ord2"" value=""" & chkString(request("ord2"),"sqlstring") & """ />")
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

sub RecordsetPaging(items_per_page,rsp_sql,disp_function,paging_position)
  'set default page size
  iPageSize = 10
  if items_per_page > 0 then
    iPageSize = items_per_page
  end if
  
  iRecordsShown = 0
  iPageCurrent = 1
  If Request("page") <> "" and isnumeric(Request("page")) Then
	iPageCurrent = cLng(Request("page"))
  End If

  Set objPagingRS = Server.CreateObject("ADODB.Recordset")
  objPagingRS.PageSize = iPageSize
  objPagingRS.CacheSize = iPageSize
  objPagingRS.Open rsp_sql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  reccount = objPagingRS.recordcount
  iPageCount = objPagingRS.PageCount
  
  If iPageCount = 0 Then
    call showMsgBlock(1,"No items found!")
  Else
    If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
    If iPageCurrent < 1 Then iPageCurrent = 1

	if iPageCount = 1 then
  	  tPgCnt = " - " & iPageCount & " page"
	else
  	  tPgCnt = " - " & iPageCount & " pages"
	end if
	if reccount = 1 then
  	  tTitle = sub_name & " ( " & reccount & " download"&tPgCnt&")"
	else
  	  tTitle = sub_name & " ( " & reccount & " downloads"&tPgCnt&")"
	end if
	
    spThemeTitle = spThemeTitle & "&nbsp;" & tTitle
    spThemeBlock1_open(intSkin)
	
	objPagingRS.AbsolutePage = iPageCurrent
	if iPageCount > 1 then
	  showDaPaging iPageCurrent,iPageCount,0
	end if
	
	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		
	    rspFunct = disp_function & "(" & objPagingRS & ")"
		execute(rspFunct)
		
	    iRecordsShown = iRecordsShown + 1
	    objPagingRS.MoveNext
	Loop
    spThemeBlock1_close(intSkin)
  End If
  
  if iPageCount > 1 then
    showDaPaging iPageCurrent,iPageCount,2
  end if

  objPagingRS.Close
  Set objPagingRS = Nothing
end sub

sub menu_downloads()
	spThemeTitle= txtMenu
	spThemeBlock1_open(intSkin)
 if bFso then
    mnu.menuName = "m_downloads"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
 else %>
	<div class="menu">
      <a href="dl.asp?cmd=3">- <%= txtNewDL %><br /></a>
      <a href="dl.asp?cmd=4">- <%= txtPopDL %><br /></a>
      <a href="dl.asp?cmd=5">- <%= txtTopDL %><br /></a>
	<%if not strDBNTUserName = "" then%>
      <a href="dl_add_form.asp">- <%= txtSubDL %><br /></a><% End If %>
      <a href="javascript:openWindow3('dl_pop.asp?mode=12')">- <%= txtDLFAQ %><br /></a>
	</div>
<% End If %>
<br />
<SCRIPT LANGUAGE="JavaScript">
function chkSrchForm1() {
mt=document.formS1.search.value;
if (mt.length<3) {
alert("Search word must be more than 3 characters");
return false;
}
else { return true; }
}
</SCRIPT>
	<form method="get" action="dl.asp" id="formS1" name="formS1" onSubmit="return chkSrchForm1()">
	<% 
	spThemeTitle = txtSearch & ":"
	spThemeBlock3_open() %>
    <div class="tPlain" style="text-align:center;">
	<input type="text" name="search" size="15" style="margin-top:5px;margin-bottom:5px;" />
  <select name="mode" id="mode">
    <option value="0" selected>All Downloads</option>
    <option value="1">By Submitter</option>
  </select></div>
      <div class="fNorm" style="margin-bottom:3px;text-align:center;">
      <input type="submit" value=" <%= txtSearch %> " id="searchA" name="searchA" class="button" /><input type="hidden" name="cmd" value="7" /></div><% spThemeBlock3_close() %></form>
<%spThemeBlock1_close(intSkin)
end sub
%>