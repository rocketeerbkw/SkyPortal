<!-- #include file="dl_config.asp" --><%
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
' :::		downloads - small
' :::::::::::::::::::::::::::::::::::::::::::::::
Function dl_small(Byval astrType)
  if chkApp("downloads","USERS") then
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
	
	sSql = getDlSql_sm(astrType)

	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin)
	Set lrsData = Server.CreateObject("adodb.recordset")
	lrsData.Open sSql, my_Conn  ', adOpenStatic, adLockReadOnly, adCmdText
	'Set lrsData = oSpData.GetRecordset(sSql)
	'Set lrsData = my_Conn.execute(sSql)
	
	If lrsData.EOF Then
%>
	<br /><table border="0"><tr><td width="100%" valign="top" align="center" class="fTitle"><b>No Downloads Found!</b></td></tr></table><br /><br />
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

			dl_DisplaySmall(lrsData)
			lrsData.MoveNext
			if intDir = 1 and intCount < 3 and not lrsData.eof then
			  intCount = intCount + 1
			  Response.write("</td><td style=""background:url(themes/" & strTheme & "/line.gif);width:1px;"">" & icon(icnSpacer,"","","","width=""1""") & "</td>") 
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
incDlFp = true
' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		downloads - large
' :::::::::::::::::::::::::::::::::::::::::::::::

Function dl_large(Byval astrType)
  if chkApp("downloads","USERS") then
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
	
	sSql = getDlSql_sm(astrType)
	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin)
	response.Write("<table border=""0"" cellspacing=""3"" cellpadding=""0"" width=""100%""><tr>")
	
	Set lrsData = Server.CreateObject("adodb.recordset")
	lrsData.Open sSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	'Set lrsData = oSpData.GetRecordset(sSql)

	If lrsData.EOF Then
%>
	<td width="100%" valign="top" class="fTitle">No Downloads Found!</td>
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
			
			dl_DisplayLarge(lrsData)
			
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
' :::		downloads - common
' :::::::::::::::::::::::::::::::::::::::::::::::
	
function getDlSql_sm(typ)
    dim tSQL
	
	tSQL = sql_selectDlSm()
	Select Case LCase(typ)
		Case "top"
			spThemeTitle= "Popular Downloads"
			tSQL = tSQL & " ORDER BY HIT DESC, O_POST_DATE DESC"
		Case "new"
			spThemeTitle= "Newest Downloads"
			tSQL = tSQL & " ORDER BY O_POST_DATE DESC"
		Case "rated"
			spThemeTitle= "Toprated Downloads"
			tSQL = tSQL & " ORDER BY RATING DESC, O_POST_DATE DESC"
		Case "featured"
			spThemeTitle= "Featured Downloads"
			tSQL = tSQL & " AND ((FEATURED) <> 0) ORDER BY O_POST_DATE DESC"
		Case "rand"
			spThemeTitle= "Random Downloads"
			Select Case strDBType
				Case "mysql"
					tSQL = tSQL & " ORDER BY Rand()"
				Case "sqlserver"
					tSQL = tSQL & " ORDER BY NEWID()"
				Case Else
					Randomize()
					lintRandomNumber = Int (1000*Rnd)+1
					tSQL = tSQL & " ORDER BY Rnd(" & -1*(lintRandomNumber) & "*DL_ID)"
  					'tSQL = "SELECT DL.DL_ID, DL.NAME, DL.DESCRIPTION, DL.POST_DATE, DL.HIT, DL.SHOW, Rnd(" & -1 * (lintRandomNumber) & " * DL.DL_ID), DL_CATEGORIES.CG_READ, DL_CATEGORIES.CG_FULL, " & strTablePrefix & "M_SUBCATEGORIES.SG_READ, " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL "
  					'tSQL = tSQL & "FROM (DL INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON PIC.CATEGORY = " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID) INNER JOIN DL_CATEGORIES ON " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID = DL_CATEGORIES.CAT_ID "
  					'tSQL = tSQL & "WHERE ((DL.SHOW)=1) "
					'tSQL = tSQL = "ORDER BY 7"
			End Select
	  	Case Else
			spThemeTitle= "Top Downloads"
			tSQL = tSQL & " ORDER BY HIT DESC, O_POST_DATE DESC"
	End Select
	getDlSql_sm = tSQL
end function

function sql_selectDlSm()
  tS = "SELECT DL.*, " & strTablePrefix & "M_CATEGORIES.CG_READ, " & strTablePrefix & "M_CATEGORIES.CG_FULL, " & strTablePrefix & "M_SUBCATEGORIES.SG_READ, " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL "
  tS = tS & "FROM (DL INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON DL.CATEGORY = " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID) INNER JOIN " & strTablePrefix & "M_CATEGORIES ON " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID = " & strTablePrefix & "M_CATEGORIES.CAT_ID "
  tS = tS & "WHERE ((DL.ACTIVE)=1) "
  sql_selectDlSm = tS
end function

Sub dl_DisplaySmall(ob)
  intId = ob("DL_ID")
  strTitle = ob("NAME")
  strDescription = ob("DESCRIPTION")
	
  intDescLen = cLng(intLen)
  if intDescLen = 0 then
	intDescLen = 50
  end if
  If len(strDescription) > intDescLen then
    strDescription = Left(strDescription , intDescLen) & "..."
  End If
  intHit = ob("Hit")
  dtPostDate = ChkDate(ob("POST_DATE"))
%>
	<table width="100%" cellspacing="3" cellpadding="0" border="0">
	<tr><td width="100%"><p>
        <% chkItemAttention(ob) %>
        <a href="dl.asp?title=<%= strTitle %>&amp;cmd=6&amp;cid=<%=intId%>"><span class="fNorm"><b><%=strTitle%></b></span></a>
        <% 
		n = ob("POST_DATE")
		u = ob("UPDATED")
		if u <> "0" then
		  n = "00"
		end if
		call chkNewItem(n,dl_chkNew,u,dl_chkUpdated) %>
        <br />
        <!-- <i>(Downloaded:&nbsp;<%=intHit%>)</i> --></p></td></tr>
	<tr><td width="100%"><p><%=strDescription%></p></td></tr></table>
<%
End Sub

Sub dl_DisplayLarge(ob)
  lstrDescription = ob("DESCRIPTION")
  If len(lstrDescription) > intDescLen then
    'lstrDescription = Left(lstrDescription , intDescLen) & "..."
  End If
  ldtPostDate = ChkDate(ob("POST_DATE"))
%>	<table width="100%" border="0" cellspacing="0" cellpadding="2">
	<tr><td width="100%"><p>
        <% chkItemAttention(ob) %>
        <b><a href="dl.asp?title=<%= ob("NAME") %>&amp;cmd=6&amp;cid=<%= ob("DL_ID") %>"><span class="fSubTitle"><%= ob("NAME") %></span></a></b> 
        <%
		n = ob("POST_DATE")
		u = ob("UPDATED")
		if u <> "0" then
		  n = "00"
		end if
		call chkNewItem(n,dl_chkNew,u,dl_chkUpdated)
		 %><br /><i>(Downloaded:&nbsp;<%= ob("Hit") %>)</i>
      </p></td></tr>
	<tr><td width="100%"><p>
        <%= lstrDescription %></p></td></tr></table>
<%
End Sub

'::::::: NEW FOR V0.11 ::::::::::::::::::::::::::::::::::
if curPageType = "downloads" then 
dlAttention = getCount("DL_ID","DL","ACTIVE = 0 OR BADLINK <> 0")
end if

newDLcnt = chkNewDL()

function cntNewDL()
  If newDLcnt = "" Then
    aImg = ""
  else
    aImg = "&nbsp;" & newDLcnt
  end if
  cntNewDL = aImg
end function

function chkNewDL()
  tStr = ""
  lastVisit = getCount("DL_ID","DL","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND ACTIVE = 1")
  if lastVisit > 0 then
	tStr = icon(icnNew1,"New since last visit","","","align=""middle""")
  else
	d7 = DateToStr(dateAdd("d",-7,date()))
	last7 = getCount("DL_ID","DL","POST_DATE >= '" & d7 & "' AND ACTIVE = 1")
	if last7 > 0 then
	  tStr = icon(icnNew2,"New in last 7 days","","","align=""middle""")
	else
	  d14 = DateToStr(dateAdd("d",-14,date()))
	  last14 = getCount("DL_ID","DL","POST_DATE >= '" & d14 & "' AND ACTIVE = 1")
	  if last14 > 0 then
		tStr = icon(icnNew3,"New in last 14 days","","","align=""middle""")
	  end if
	end if
  end if
  chkNewDL = tStr
end function

sub doCatCountUpdate()
  dl_CatCountUpdate()
  Call setSession("sMsg","Counts successfully updated")
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "&cid=0")
end sub

sub dl_CatCountUpdate()
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
	  iCnt = getCount("POST_DATE","DL","CATEGORY=" & iSid & "")
	  sSql = "UPDATE PORTAL_M_SUBCATEGORIES SET ITEM_CNT=" & iCnt
	  sSql = sSql & " WHERE SUBCAT_ID=" & iSid
	  executeThis(sSql)
    loop
  end if
  set rsA = nothing
end sub

sub dl_PendTaskCnt()
  ' Pending Downloads count
  PTcnt = PTcnt + getCount("DL_ID","DL","ACTIVE=0 OR BADLINK<>0")
end sub

sub dl_adminPndLink()
  if chkApp("downloads","USERS") then
    cntDL = getCount("DL_ID","DL","ACTIVE=0")
	cntDLB = getCount("DL_ID","DL","BADLINK <> 0")
	If cntDL <> 0 or cntDLB <> 0 then
	  Response.Write "<li><a href=""dl.asp?cmd=22&sid=0""><b>"
	  If cntDL <> 0 then
	    response.Write cntDL & "&nbsp;" & txtNwDLsApprv
	  End IF
	  If cntDLB <> 0 then
	    response.Write cntDLB & "&nbsp;" & txtBdDLs
	  End IF
	  Response.Write "</b></a></li>"
	End IF
  End IF
end sub

sub dl_SiteSearch()
  '############# Download Search ####################
  If chkApp("downloads","USERS") Then
    strSQL = "SELECT * FROM DL WHERE KEYWORD LIKE '%" & search & "%' OR CONTENT LIKE '%" & search & "%' OR DESCRIPTION LIKE '%" & search & "%' or NAME like '%" & search & "%' AND ACTIVE=1 ORDER BY DL_ID"

    Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.Open strSQL,my_Conn,adOpenStatic,adLockReadOnly,adCmdText
	reccount = objPagingRS.recordcount
	%>
	<center><span class="fSubTitle"><b><%= txtDownloads %> - <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br />
	<% 
    If reccount > 0 Then
	  %> 	
	  <center>
	  <a href="dl.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
	  <%= txtVSrchRslts %></a></center>
	  <br />
	  <%
    End If

    objPagingRS.Close
    Set objPagingRS = Nothing
	response.Write("<hr />")
    response.Flush()
  end if
end sub

sub menu_dl()
	  blkStart = timer
  getMenu(intAppID)
	if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - blkStart),3)
	  response.Write(blkLoadTime)
	end if
end sub

':: download admin menu
sub downloadConfigMenu(typ)
 if bFSO then
    mnu.menuName = "downloads_admin"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
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
    <div class="tCellAlt1" onMouseOver="this.className='tCellHover';" onMouseOut="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block8<%= typ %>','2');"><span style="margin: 2px;"><img name="block8<%= typ %>Img" id="block8<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b>Downloads</b></div>
      <div class="menu" id="block8<%= typ %>" style="display: <%= cls %>; text-align:left;">
	 <a href="admin_dl_main.asp"><%= icn_bar %>Approve New (<%= getCount("DL_ID","DL","ACTIVE=0") %>)<br /></a>
	 <a href="admin_dl_admin.asp?cmd=40"><%= icn_bar %>Bad Downloads (<%= getCount("DL_ID","DL","BADLINK <> 0") %>)<br /></a>
	 <a href="admin_dl_admin.asp"><%= icn_bar %>Create Category<br /></a>
	 <a href="admin_dl_admin.asp?cmd=2"><%= icn_bar %>Edit Category<br /></a>
	 <a href="admin_dl_admin.asp?cmd=4"><%= icn_bar %>Delete Category<br /></a>
	 <a href="admin_dl_admin.asp?cmd=1"><%= icn_bar %>Create SubCategory<br /></a>
	 <a href="admin_dl_admin.asp?cmd=5"><%= icn_bar %>Edit SubCategory<br /></a>
	 <a href="admin_dl_admin.asp?cmd=8"><%= icn_bar %>Delete SubCategory<br /></a>
	 <a href="admin_dl_admin.asp?cmd=10"><%= icn_bar %>Edit Download<br /></a>
	 <a href="admin_dl_admin.asp?cmd=20"><%= icn_bar %>Delete Download<br /></a>
	 <a href="admin_dl_admin.asp?cmd=30"><%= icn_bar %>Browse Downloads<br /></a>
		   </div>
<%
 end if
end sub

%>
