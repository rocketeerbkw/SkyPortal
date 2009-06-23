<!-- #include file="article_config.asp" --><%
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

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Articles - small
' :::::::::::::::::::::::::::::::::::::::::::::::
Function article_sm(Byval astrType)
  if chkApp("article","USERS") then
	  blkStart = timer	
	Dim intDescLen
	Dim numToDisp
	Dim sSql
	Dim lrsData
	
	intDir = cLng(intDir)
	intCount = 1
	
	numToDisp = cLng(intShow)
	if isnumeric(numToDisp) then
	  if numToDisp = 0 then
	    numToDisp = 5
	  end if
	else
	  numToDisp = 5
	end if
	
	sSql = getArtSql_sm(astrType)

	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin)
	'Set lrsData = Server.CreateObject("adodb.recordset")
	'lrsData.Open sSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	Set lrsData = my_Conn.execute(sSql)
	
	If lrsData.EOF Then
%>
	<br /><table border="0"><tr><td width="100%" valign="top" align="center" class="fSubTitle"><b>No Items Found!</b></td></tr></table><br /><br />
<%	
	Else
		reccount = lrsData.recordcount
		if intDir = 1 then
	  	  intWid = Int(100 / reccount) & "%"
		else
	  	  intWid = "100%"
		end if
	response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td width=""" & intWid & """ valign=""top"">")
		myCurCnt = 0
		Do While Not lrsData.Eof and myCurCnt < numToDisp
		  if not hasAccess(lrsData("SG_READ")) then
			lrsData.MoveNext
		  else
		    myCurCnt = myCurCnt + 1
			article_DisplaySmall(lrsData)
			lrsData.MoveNext
			if intDir = 1 and intCount < 3 and not lrsData.eof then
			  intCount = intCount + 1
			  Response.write("</td><td style=""background:url(" & icnLine & ");width:1px;"">" & icon(icnSpacer,"","","","width=""1""") & "</td>") 
			  response.write("<td width=""" & intWid & """ valign=""top"">")
			end if
			
			if intDir <> 1 and not lrsData.eof Then
				Response.write "</td></tr><tr><td align=""center"" style=""height:1px;"">" & icon(icnLine,"","","","height=""1"" width=""98%""") & "</td></tr><tr><td>"
			end if
		  end if
		Loop
	response.Write("</td></tr></table>")
	End If
	
	lrsData.Close
	Set lrsData = Nothing
	spThemeBlock1_close(intSkin)
	intShow = 0
	intLen = 0
	intDir = 0
	if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - blkStart),3)
	  response.Write(blkLoadTime)
	end if
  end if
End Function
incArtFp = true
' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Articles - large
' :::::::::::::::::::::::::::::::::::::::::::::::

Function article_lg(Byval astrType)
  if chkApp("article","USERS") then
	  blkStart = timer	
	Dim numToDisp
	Dim sSql
	Dim lrsData
	Dim lintWidth
	Dim lintColumns, lintMaxColumns
	Dim numToDispIndex
	
	lintMaxColumns = 2
	numToDisp = cLng(intShow)'This is how many items to display
	if numToDisp = 0 then
	  numToDisp = 6  
	end if
	
	sSql = getArtSql_sm(astrType)
	
	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin)
	response.Write("<table border=""0"" cellspacing=""3"" cellpadding=""0"" width=""100%""><tr>")
	
	Set lrsData = Server.CreateObject("adodb.recordset")
	lrsData.Open sSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	If lrsData.EOF Then
%>
	<td width="100%" valign="top" class="fTitle">No Articles Found!</td>
<%	
	Else
		'lintWidth = Int(100 / (numToDisp/2))
		lintWidth = Int(100 / 2)
		lcnt = 1
		myCurCnt = 0
		Do While Not lrsData.Eof and myCurCnt < numToDisp
		  if not hasAccess(lrsData("SG_READ")) then
			lrsData.MoveNext
		  else
		    myCurCnt = myCurCnt + 1
		    Response.Write("<td width="""&lintWidth&"%"" valign=""top"">")
			
			article_DisplayLarge(lrsData)
			
			Response.Write("</td>")
			lrsData.MoveNext
			  If lcnt = 2 and not lrsData.eof Then
				lcnt = 0
				Response.write "</tr><tr><td colspan=""3"" align=""center"" style=""height:1px;"">" & icon(icnLine,"","","","height=""1"" width=""98%""") & "</td></tr><tr>"
			  elseif lcnt = 1 and not lrsData.eof then
					Response.write "<td style=""background:url(" & icnLine & ");width:1px;"">" & icon(icnSpacer,"","","","width=""1""") & "</td>"
			  elseif lcnt = 1 and lrsData.eof then
					Response.write "<td style=""background:url(" & icnLine & ");width:1px;"">" & icon(icnSpacer,"","","","width=""1""") & "</td>"
					Response.write "<td>&nbsp;</td>"
			  End If
			lcnt = lcnt + 1
		  end if
		Loop
	End If
	
	lrsData.Close
	Set lrsData = Nothing
	response.Write("</tr></table>")
	spThemeBlock1_close(intSkin)
	intShow = 0
	intLen = 0
	intDir = 0
	if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - blkStart),3)
	  response.Write(blkLoadTime)
	end if
  end if
End Function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Articles - common
' :::::::::::::::::::::::::::::::::::::::::::::::
	
function getArtSql_sm(typ)
    dim tSQL
	
	tSQL = sql_selectArtSm()
	Select Case LCase(typ)
		Case "top"
			spThemeTitle= "Popular Articles"
			tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
		Case "new"
			spThemeTitle= "Newest Articles"
			tSQL = tSQL & " ORDER BY POST_DATE DESC"
		Case "rated"
			spThemeTitle= "Toprated Articles"
			tSQL = tSQL & " ORDER BY RATING DESC, POST_DATE DESC"
		Case "featured"
			spThemeTitle= "Featured Articles"
			tSQL = tSQL & " AND ((FEATURED) <> 0) ORDER BY POST_DATE DESC"
		Case "rand"
			spThemeTitle= "Random Articles"
			Select Case strDBType
				Case "mysql"
					tSQL = tSQL & " ORDER BY Rand()"
				Case "sqlserver"
					tSQL = tSQL & " ORDER BY NEWID()"
				Case Else
					Randomize()
					lintRandomNumber = Int (1000*Rnd)+1
					tSQL = tSQL & " ORDER BY Rnd(" & -1*(lintRandomNumber) & "*ARTICLE_ID)"
			End Select
	  	Case Else
			spThemeTitle= "Top Downloads"
			tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
	End Select
	getArtSql_sm = tSQL
end function

function sql_selectArtSm()
  tS = "SELECT ARTICLE.*, " & strTablePrefix & "M_CATEGORIES.CG_READ, " & strTablePrefix & "M_CATEGORIES.CG_FULL, " & strTablePrefix & "M_SUBCATEGORIES.SG_READ, " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL "
  tS = tS & "FROM (ARTICLE INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON ARTICLE.CATEGORY = " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID) INNER JOIN " & strTablePrefix & "M_CATEGORIES ON " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID = " & strTablePrefix & "M_CATEGORIES.CAT_ID "
  tS = tS & "WHERE ((ARTICLE.ACTIVE)=1) "
  sql_selectArtSm = tS
end function

Sub article_DisplaySmall(ob)
  if len(ob("SUMMARY")) > 100 then
    tSummary = left(ob("SUMMARY"),100) & "..."
  end if
%>
	<table width="100%" cellspacing="0" cellpadding="0"><tr><td width="100%"><p>&nbsp;<b><a href="article_read.asp?title=<%=server.URLEncode(ob("TITLE"))%>&amp;item=<%= ob("article_id") %>"><span class="fSubTitle"><%=ob("TITLE")%></span></a></b> 
        <% call chkNewArtItem(ob("POST_DATE"),art_chkNew,ob("UPDATED"),art_chkUpdated) %><br /><i>(Hits: <%= ob("HIT") %>)</i>
      </p></td></tr>
	<tr><td width="100%"><p><%= tSummary %><br />
        <a href="article_read.asp?title=<%=server.URLEncode(ob("Title"))%>&amp;item=<%= ob("article_id") %>"><span class="fNorm"><b>read 
        more...</b></span></a></p></td></tr></table>
<%
End Sub

Sub article_DisplayLarge(ob)
%>
	<table width="100%"><tr><td width="100%">&nbsp;<b><a href="article_read.asp?title=<%=server.URLEncode(ob("TITLE"))%>&amp;item=<%=ob("article_id")%>"><span class="fSubTitle"><%=ob("TITLE")%></span></a></b> 
        <% call chkNewArtItem(ob("POST_DATE"),art_chkNew,ob("UPDATED"),art_chkUpdated) %><br /><span class="fSmall"><i>(Hits: <%=ob("HIT")%>)</i></span>
      </td></tr>
	<tr><td width="100%"><p>
        <%=ob("SUMMARY")%>
        <a href="article_read.asp?title=<%=server.URLEncode(ob("Title"))%>&amp;item=<%= ob("article_id") %>"><br/><span class="fNorm"><b>read 
        more...</b></span></a></p></td></tr></table>
<%
End Sub

function chkNewArtItem(p,bn,u,bu)
  ':: checks if item passes is new or updated and writes an icon to the browser
  bTF = false
  lastVisit = Session(strUniqueID & "last_here_date")
 if bn then
  if len(p) = 14 then
    tdtSince = getDateDiff(strCurDateString,p)
    if lastVisit <= p then
	  response.Write icon(icnNew1,"New since last visit","","","hspace=""4"" align=""middle""")
	  bTF = true
    elseif tdtSince < 7 then
	  response.Write icon(icnNew2,"New in last 7 days","","","hspace=""4"" align=""middle""")
	  bTF = true
      elseif tdtSince < 14 then
	  response.Write icon(icnNew3,"New in last 14 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    end if
  end if
 end if
 if bu then
  if len(u) = 14 then
    tdtSince = getDateDiff(strCurDateString,u)
    if lastVisit <= u then
	  response.Write icon(icnUpdate1,"Updated since last visit","","","hspace=""4"" align=""middle""")
	  bTF = true
	elseif tdtSince < 7 then
	  response.Write icon(icnUpdate2,"Updated in last 7 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    elseif tdtSince < 14 then
	  response.Write icon(icnUpdate3,"Updated in last 14 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    end if
  end if
 end if
  chkNewArtItem = bTF
end function

'::::::: NEW FOR V0.10 ::::::::::::::::::::::::::::::::::

artAttention = getCount("ARTICLE_ID","ARTICLE","ACTIVE = 0")

newArtcnt = chkNewArt()

function cntNewArticles()
  If newArtcnt = "" Then
    aImg = ""
  else
    aImg = "&nbsp;" & newArtcnt
  end if
  cntNewArticles = aImg
end function

function chkNewArt()
  tStr = ""
  lastVisit = getCount("ARTICLE_ID","ARTICLE","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND ACTIVE = 1")
  if lastVisit > 0 then
	tStr = icon(icnNew1,"New since last visit","","","align=""middle""")
  else
	d7 = DateToStr(dateAdd("d",-7,date()))
	last7 = getCount("ARTICLE_ID","ARTICLE","POST_DATE >= '" & d7 & "' AND ACTIVE = 1")
	if last7 > 0 then
	  tStr = icon(icnNew2,"New in last 7 days","","","align=""middle""")
	else
	  d14 = DateToStr(dateAdd("d",-14,date()))
	  last14 = getCount("ARTICLE_ID","ARTICLE","POST_DATE >= '" & d14 & "' AND ACTIVE = 1")
	  if last14 > 0 then
		tStr = icon(icnNew3,"New in last 14 days","","","align=""middle""")
	  end if
	end if
  end if
  chkNewArt = tStr
end function

sub doCatCountUpdate()
  art_CatCountUpdate()
  Call setSession("sMsg","Counts successfully updated")
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "&cid=0")
end sub

sub art_CatCountUpdate()
  ':: reset counts to zero
  sSql = "UPDATE PORTAL_M_SUBCATEGORIES SET ITEM_CNT=0"
  sSql = sSql & " WHERE APP_ID=" & intAppID
  executeThis(sSql)
  
  ':: reset counts
  sSql = "SELECT SUBCAT_ID FROM PORTAL_M_SUBCATEGORIES"
  sSql = sSql & " WHERE APP_ID=" & intAppID
  set rsA = my_Conn.execute(sSql)
  if not rsA.eof then
    do until rsA.eof
	  iSid = rsA("SUBCAT_ID")
      rsA.movenext
	  iCnt = getCount("POST_DATE","ARTICLE","CATEGORY=" & iSid & "")
	  sSql = "UPDATE PORTAL_M_SUBCATEGORIES SET ITEM_CNT=" & iCnt
	  sSql = sSql & " WHERE SUBCAT_ID=" & iSid
	  executeThis(sSql)
    loop
  end if
  set rsA = nothing
end sub

sub article_PendTaskCnt()
  ' Pending Articles count
  PTcnt = PTcnt + getCount("ARTICLE_ID","ARTICLE","ACTIVE=0")
end sub

sub article_adminPndLink()
  if chkApp("article","USERS") then
    cntPI = getCount("ARTICLE_ID","ARTICLE","ACTIVE=0")
	If cntPI <> 0 then
	  Response.Write "<li><a href=""admin_articles.asp?cmd=22""><b>"
	  Response.Write cntPI & "&nbsp;" & txtNwArtclsApprv
	  Response.Write "</b></a></li>"
	End IF
  End IF
end sub

sub article_SiteSearch()
  If chkApp("article","USERS") Then
    strSQL = "SELECT * from ARTICLE WHERE KEYWORD LIKE '%" & search & "%' OR SUMMARY LIKE '%" & search & "%' OR CONTENT LIKE '%" & search & "%' AND ACTIVE=1 ORDER BY ARTICLE_ID DESC"
	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	reccount = objPagingRS.recordcount
	response.Write "<center><span class=""fSubTitle"">"
	Response.Write txtArticles & " - " & txtFound & "&nbsp;" & reccount & "&nbsp;" & txtSitems
	Response.Write "</span></center><br />" 
 	If reccount > 0 Then
	  Response.Write "<center><a href=""article.asp?cmd=6&search=" & search & "&submit1=Search&num=" & show & """>"
	  Response.Write txtVSrchRslts & "</a></center><br />"
	End If
	objPagingRS.Close
	Set objPagingRS = Nothing
	response.Write "<hr />"
  end if 
end sub

sub menu_art()
  blkStart = timer
  getMenu(intAppID)
  if shoBlkTimer then
	blkLoadTime = formatnumber((timer - blkStart),3)
	response.Write(blkLoadTime)
  end if
end sub

':: article admin menu
sub articleConfigMenu(typ)
 if bFSO then
    mnu.menuName = "articles_admin"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = typ
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
 else
	if typ = 1 then
	  cls = "block"
	  icn = "min1"
	  alt = "Collapse"
	else
	  cls = "none"
	  icn = "max1"
	  alt = "Expand"
	end if %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block6<%= typ %>','2');"><span style="margin: 2px;"><img name="block6<%= typ %>Img" id="block6<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="middle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b>Articles</b></div>
      <div class="menu" id="block6<%= typ %>" style="display: <%= cls %>; text-align:left;">
	  	<a href="admin_article_main.asp"><%= icn_bar %>Approve Articles (<%= getCount("ARTICLE_ID","ARTICLE","ACTIVE=0") %>)<br></a>
		<a href="admin_articles.asp"><%= icn_bar %>Create Category<br></a>
		<a href="admin_articles.asp?cmd=2"><%= icn_bar %>Edit Category<br></a>
		<a href="admin_articles.asp?cmd=4"><%= icn_bar %>Delete Category<br></a>
		<a href="admin_articles.asp?cmd=1"><%= icn_bar %>Create SubCategory<br></a>
		<a href="admin_articles.asp?cmd=5"><%= icn_bar %>Edit SubCategory<br></a>
		<a href="admin_articles.asp?cmd=8"><%= icn_bar %>Delete SubCategory<br></a>
		<a href="admin_articles.asp?cmd=10"><%= icn_bar %>Edit Article<br></a>
		<a href="admin_articles.asp?cmd=20"><%= icn_bar %>Delete Article<br></a>
		<a href="admin_articles.asp?cmd=30"><%= icn_bar %>Browse Articles<br></a>
		   </div>
  <%
 end if
end sub
':: end article admin menu

%>
