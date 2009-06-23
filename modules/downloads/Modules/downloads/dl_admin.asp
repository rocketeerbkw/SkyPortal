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

sub showAttentionSubCat(i)
  if instr(sScript,"dl.asp") > 0 then
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
   mod_EditItemForm(intDLID)
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
	    case 121 ':: unmark bad item
		  Response.Write("Mode 121<br><br>")
		  sSql = "UPDATE DL SET"
		  sSql = sSql & " ACTIVE=1"
		  sSql = sSql & ", BADLINK=0"
		  sSql = sSql & " WHERE DL_ID=" & cid
		  'sSql = sSql & " AND CATEGORY=" & sid
		  Response.Write(sSql)
		  executeThis(sSql)
		  Call setSession("sMsg","Item successfully updated")
		  resetCoreConfig() 
		  closeAndGo(sDLpage & "?cmd=" & iPgType & "&sid=" & i)
	    case 122 ':: approve item
		  Response.Write("Mode 122<br><br>")
		  sSql = "UPDATE DL SET"
		  sSql = sSql & " ACTIVE=1"
		  sSql = sSql & ", BADLINK=0"
		  sSql = sSql & " WHERE DL_ID=" & cid
		  'sSql = sSql & " AND CATEGORY=" & sid
		  Response.Write(sSql)
		  executeThis(sSql)
		  Call setSession("sMsg","Item successfully updated")
		  mod_increaseSubcatCount(i)
		  resetCoreConfig() 
		  closeAndGo(sDLpage & "?cmd=" & iPgType & "&sid=" & i)
		  
		case 123 ':: delete new item or bad link
		  Response.Write("Mode 123")
		  sSql = "SELECT URL FROM DL WHERE DL_ID = " & cid
		  set rsA = my_Conn.execute(sSql)
		  if not rsA.eof then
		    sFile = rsA("URL")
			if left(sFile,16) = "files/downloads/" then
		      pFile = server.MapPath(sFile)
			  on error resume next
			  set fso = Server.CreateObject("Scripting.FileSystemObject")
			  if fso.FileExists(pFile) = true then
			    fso.DeleteFile(pFile)
			  end if
			  set fso = nothing
			  on error goto 0			  
			end if
		    sSql = "DELETE FROM DL"
		    sSql = sSql & " WHERE DL_ID=" & cid
		    sSql = sSql & " AND CATEGORY=" & i
		    executeThis(sSql)
			Call setSession("sMsg","Item successfully deleted")
			mod_decreaseSubcatCount(i)
		  else
			Call setSession("sMsg","Item not found")
		  end if
		  resetCoreConfig()
		  closeAndGo(sDLpage & "?cmd=" & iPgType & "&sid=" & i)
		  
		case 133
		  Response.Write("Mode 133")
		case else
	  end select
      'tNew = getCount("DL_ID","DL","ACTIVE = 0 AND CATEGORY=" & i)
      'tBad = getCount("DL_ID","DL","BADLINK <> 0 AND CATEGORY=" & i)
	  mod_writeApprovalJS()
  	  Response.Write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">"
  	  Response.Write "<tr><td width=""30%"" align=""right"">"
  	  Response.Write icon(imgAttention,"Attention","","","")
  	  Response.Write "</td><td>"
  	  Response.Write "Items that need attention are listed here"
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
	  sSql = singleDLsql()
	  sSql = sSql & "WHERE (((DL.ACTIVE)=0))"
	  sSql = sSql & " ORDER BY PORTAL_M_CATEGORIES.CAT_NAME, PORTAL_M_SUBCATEGORIES.SUBCAT_NAME, DL.NAME;"
	  
	  'sSql = "SELECT * FROM DL WHERE ACTIVE = 1"
	  set rsN = my_Conn.execute(sSql)
	  if rsN.eof then
  	    Response.Write "<tr><td align=""center"" colspan=""4"" class=""fSubTitle"">"
  	    Response.Write "<br/>No items to approve<br/><br/>"
  	    Response.Write "</td></tr>"
	  else
	    do until rsN.eof
		 if hasAccess(rsN("CG_FULL")) OR hasAccess(rsN("SG_FULL")) OR bAppFull then
  	      Response.Write "<tr><td align=""center"">"
  	      Response.Write icon(icnCheck,"Approve","display:inline;cursor:pointer;","jsApprDl('" & rsN("NAME") & "','" & rsN("DL_ID") & "','" & i & "',1)","")
  	      Response.Write icon(icnDelete,txtDel,"display:inline;cursor:pointer;","jsDelDl('" & rsN("NAME") & "','" & rsN("DL_ID") & "','" & i & "')","")
  	      Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","jsEditDL('" & rsN("DL_ID") & "','" & i & "')","")
		  if isMac then
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","jsAttnDL('view" & rsN("DL_ID") & "')","")
		  else
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","openJsLayer('view" & rsN("DL_ID") & "','550','450')","")
		  end if
  	      Response.Write "</td><td>"
  	      Response.Write rsN("NAME")
		  Call mod_writeViewItem(rsN,"showInfo")
  	      Response.Write "</td><td align=""center"">" & chkDate2(rsN("POST_DATE"))
  	      Response.Write "</td><td align=""center"">" & rsN("UPLOADER")
  	      Response.Write "</td></tr>"
		 end if
		 rsN.movenext
		loop
	  end if
	  set rsN = nothing
  	  Response.Write "</table>"
	  
      Response.Write "<br><br>"
	  
  	  Response.Write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"" class=""grid"">"
  	  Response.Write "<tr><td align=""center"" colspan=""4"" class=""tTitle"">"
  	  Response.Write "Items marked as Bad"
  	  Response.Write "</td></tr>"
  	  Response.Write "<tr><td align=""center"" class=""tSubTitle"" width=""100"">Options"
  	  Response.Write "</td><td class=""tSubTitle"">Name"
  	  Response.Write "</td><td align=""center"" class=""tSubTitle"" width=""100"">Date"
  	  Response.Write "</td><td align=""center"" class=""tSubTitle"" width=""100"">By"
  	  Response.Write "</td></tr>"
	  'sSql = "SELECT * FROM DL WHERE BADLINK <> 0"' AND CATEGORY = " & i
	  sSql = singleDLsql()
	  sSql = sSql & "WHERE (((DL.BADLINK)<>0))"
	  sSql = sSql & " ORDER BY PORTAL_M_CATEGORIES.CAT_NAME, PORTAL_M_SUBCATEGORIES.SUBCAT_NAME, DL.NAME;"
	  set rsB = my_Conn.execute(sSql)
	  if rsB.eof then
  	    Response.Write "<tr><td align=""center"" colspan=""4"" class=""fSubTitle"">"
  	    Response.Write "<br/>No items to correct<br/><br/>"
  	    Response.Write "</td></tr>"
	  else
	    tmpID = rsB("DL_ID")
	    do until rsB.eof
		 if hasAccess(rsB("CG_FULL")) OR hasAccess(rsB("SG_FULL")) OR bAppFull then
  	      Response.Write "<tr><td align=""center"">"
  	      Response.Write icon(icnCheck,"Unmark","display:inline;cursor:pointer;","jsApprDl('" & rsB("NAME") & "','" & rsB("DL_ID") & "','" & rsB("CATEGORY") & "',0)","")
  	      Response.Write icon(icnDelete,txtDel,"display:inline;cursor:pointer;","jsDelDl('" & rsB("NAME") & "','" & rsB("DL_ID") & "','" & i & "')","")
  	      Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","jsEditDL('" & rsB("DL_ID") & "','" & i & "')","")
		  if isMac then
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","jsAttnDL('view" & rsB("DL_ID") & "')","")
		  else
  	        Response.Write icon(icnBinoc,txtView,"display:inline;cursor:pointer;","openJsLayer('view" & rsB("DL_ID") & "','550','450')","")
		  end if
  	      Response.Write "</td><td>"
  	      Response.Write rsB("NAME")
		  Call mod_writeViewItem(rsB,"showInfo")
		  'writeEditItem(rsB)
  	      Response.Write "</td><td align=""center"">" & chkDate2(rsB("POST_DATE"))
  	      Response.Write "</td><td align=""center"">" & getMemberName(rsB("BADLINK"))
  	      Response.Write "</td></tr>"
		 end if
		 rsB.movenext
		loop
	  end if
	  set rsB = nothing
  	  Response.Write "</table><br/>"
	  Response.Write "<div id=""view_pane"">"
	  Response.Write "</div><br/>"
	  call mod_shoLegend("admin",dl_chkNew,dl_chkUpdated)
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
  
  if instr(sScript,"dl.asp") > 0 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub procNewEditItemForm()
if not isObject(objUpload) then
	sString = sString & "<li>Your session has expired.</li>"
	sString = sString & "<li>You will need to refresh the submission page<br />to get the session back.</li>"
	response.Write(sString)
else
	sT1Sql = ""
	sT2Sql = ""
  response.Write("Process form<br>")
  uLoad = filename
  itemID = clng(objUpload.Fields("itemID").Value)
  subcat = clng(objUpload.Fields("subcat").Value)
  orig_subcat = clng(objUpload.Fields("orig_subcat").Value)
  iname = ChkString(objUpload.Fields("name").Value,"sqlstring")
  URL = trim(ChkString(objUpload.Fields("URL").Value,"url"))
  filesize = formatSize(size)
  sdesc = ChkString(objUpload.Fields("sdes").Value,"sqlstring")
  ldesc = ChkString(objUpload.Fields("Message").Value,"message")
  email = ChkString(objUpload.Fields("email").Value,"sqlstring")
  key = ChkString(objUpload.Fields("key").Value,"sqlstring")

  if txtDlLabel1 <> "" then
    sTDATA1 = ChkString(objUpload.Fields("TDATA1").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA1='" & sTDATA1 & "'"
  end if
  if txtDlLabel2 <> "" then
    sTDATA2 = ChkString(objUpload.Fields("TDATA2").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2='" & sTDATA2 & "'"
  end if
  if txtDlLabel3 <> "" then
    sTDATA3 = ChkString(objUpload.Fields("TDATA3").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3='" & sTDATA3 & "'"
  end if
  if txtDlLabel4 <> "" then
    sTDATA4 = ChkString(objUpload.Fields("TDATA4").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4='" & sTDATA4 & "'"
  end if
  if txtDlLabel5 <> "" then
    sTDATA5 = ChkString(objUpload.Fields("TDATA5").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5='" & sTDATA5 & "'"
  end if
  if txtDlLabel6 <> "" then
    sTDATA6 = ChkString(objUpload.Fields("TDATA6").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6='" & sTDATA6 & "'"
  end if
  if txtDlLabel7 <> "" then
    sTDATA7 = ChkString(objUpload.Fields("TDATA7").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7='" & sTDATA7 & "'"
  end if
  if txtDlLabel8 <> "" then
    sTDATA8 = ChkString(objUpload.Fields("TDATA8").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8='" & sTDATA8 & "'"
  end if
  if txtDlLabel9 <> "" then
    sTDATA9 = ChkString(objUpload.Fields("TDATA9").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9='" & sTDATA9 & "'"
  end if
  if txtDlLabel10 <> "" then
    sTDATA10 = ChkString(objUpload.Fields("TDATA10").Value,"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10='" & sTDATA10 & "'"
  end if
  
  unmarkbad = objUpload.Fields("unmarkbad").Value
  marknew = objUpload.Fields("marknew").Value
  markupd = objUpload.Fields("markupd").Value
  approve = objUpload.Fields("approve").Value
  today = strCurDateString
  
  Set objUpload = Nothing

  dim rsA
  strSQL = "SELECT CAT_ID, SG_FULL, SG_WRITE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & subcat
  set rsA = my_Conn.execute(strSql)
  if rsA.eof then
    closeAndGo("dl.asp?bad=scatID")
  else
    parent = rsA("CAT_ID")
    s_full = hasAccess(rsA("SG_FULL"))
    s_write = hasAccess(rsA("SG_WRITE"))
  end if
  set rsA = nothing

  if not s_write then
    closeAndGo("error.asp?type=no_access")
  end if

  strSQL = "SELECT CG_FULL FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & parent
  set rsA = my_Conn.execute(strSql)
  if rsA.eof then
    closeAndGo("dl.asp?bad=catID")
  else
    c_full = hasAccess(rsA("CG_FULL"))
  end if
  set rsA = nothing

  strSQL = "select UP_ACTIVE, UP_ALLOWEDGROUPS, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'download'"
  set rsA = server.CreateObject("adodb.recordset")
  rsA.Open strSQL, my_Conn
  uActive = rsA("UP_ACTIVE")
  uAllowed = hasAccess(rsA("UP_ALLOWEDGROUPS"))
  extAllowed = rsA("UP_ALLOWEDEXT")
  rsA.Close
  set rsA = nothing

  if trim(uLoad) <> "" and bFSO = true and strAllowUploads = 1 and uActive = 1 and uAllowed then
	banner = remotePath  & orig_subcat & "/"  & uLoad
	URL = banner
	on error resume next
	set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(remotePath) & "\"
		if fso.FolderExists(dirPath & orig_subcat) = false then
			fso.CreateFolder(dirPath & orig_subcat)
		end if
		if fso.FileExists(dirPath & uLoad) = true then
			fso.MoveFile dirPath & uLoad, dirPath & orig_subcat & "\" & uLoad
		end if
	set fso = nothing
	on error goto 0
  end if
  
    if approve = "1" then
	  mod_increaseSubcatCount(subcat)
    end if
  
  ':: check for subcat change
  if orig_subcat <> subcat then
    if approve <> "1" then
	  mod_increaseSubcatCount(subcat)
	  mod_decreaseSubcatCount(orig_subcat)
    end if
    if left(URL,16) = "files/downloads/" then
	  if instr(URL,"/" & orig_subcat & "/") > 0 then
	    ':: edit URL
  	    tUrl = replace(URL,"/" & orig_subcat & "/","/" & subcat & "/")
	    ':: move file to new subcat folder
		call moveFile(server.MapPath(URL),server.MapPath(tUrl))
	    URL = tUrl
	  end if
	end if
  end if
  
  sSql = "UPDATE DL SET "
  sSql = sSql & "NAME='" & iname & "'"
  sSql = sSql & ",CATEGORY =" & subcat & ""
  sSql = sSql & ",EMAIL ='" & email & "'"
  sSql = sSql & ",DESCRIPTION ='" & sdesc & "'"
  sSql = sSql & ",CONTENT ='" & ldesc & "'"
  sSql = sSql & ",URL ='" & URL & "'"
  sSql = sSql & ",FILESIZE ='" & filesize & "'"
  sSql = sSql & ",KEYWORD ='" & key & " '"
  sSql = sSql & sT1Sql
  if marknew = "1" then
    sSql = sSql & ",UPDATED ='0'"
    sSql = sSql & ",O_POST_DATE ='" & today & "'"
    sSql = sSql & ",POST_DATE ='" & today & "'"
  end if
  if markupd = "1" then
    sSql = sSql & ",UPDATED ='" & today & "'"
    sSql = sSql & ",POST_DATE ='" & today & "'"
  end if
  if approve = "1" then
    sSql = sSql & ",ACTIVE = 1"
  end if
  if unmarkbad = "1" then
    sSql = sSql & ",BADLINK = 0"
  end if
  sSql = sSql & " WHERE DL_ID =" & itemID
  'response.Write(email)
  'response.End()
  executeThis(sSql)
  Response.Write("<br>" & sSql)
  'Call setSession("sMsg","Item successfully updated")
  resetCoreConfig()
  closeAndGo("stop")
  'session.Contents("uploadType") = ""
  'closeAndGo(sDLpage & "?cmd=6&cid=" & itemID)
end if
end sub

sub processEditItemForm()
  response.Write("Process form<br>")
  itemID = clng(Request.Form("itemID"))
  subcat = clng(Request.Form("subcat"))
  orig_subcat = clng(Request.Form("orig_subcat"))
  iname = ChkString(Request.Form("name"),"sqlstring")
  URL = trim(ChkString(Request.Form("URL"),"url"))
  filesize = ChkString(Request.Form("size"),"sqlstring")
  sdesc = ChkString(Request.Form("sdes"),"sqlstring")
  ldesc = ChkString(Request.Form("Message"),"message")
  email = ChkString(Request.Form("email"),"sqlstring")
  key = ChkString(Request.Form("key"),"sqlstring")
  
  'license = ChkString(Request.Form("license"),"sqlstring")
  'language = ChkString(Request.Form("language"),"sqlstring")
  'platform = ChkString(Request.Form("platform"),"sqlstring")
  'publisher = ChkString(Request.Form("publisher"),"sqlstring")
  'publisherURL = ChkString(Request.Form("publisherURL"),"url")
  'uploader = ChkString(Request.Form("uploader"),"title")
	
	sT1Sql = ""
	sT2Sql = ""

  if txtDlLabel1 <> "" then
    sTDATA1 = ChkString(Request.Form("TDATA1"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA1='" & sTDATA1 & "'"
  end if
  if txtDlLabel2 <> "" then
    sTDATA2 = ChkString(Request.Form("TDATA2"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA2='" & sTDATA2 & "'"
  end if
  if txtDlLabel3 <> "" then
    sTDATA3 = ChkString(Request.Form("TDATA3"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA3='" & sTDATA3 & "'"
  end if
  if txtDlLabel4 <> "" then
    sTDATA4 = ChkString(Request.Form("TDATA4"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA4='" & sTDATA4 & "'"
  end if
  if txtDlLabel5 <> "" then
    sTDATA5 = ChkString(Request.Form("TDATA5"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA5='" & sTDATA5 & "'"
  end if
  if txtDlLabel6 <> "" then
    sTDATA6 = ChkString(Request.Form("TDATA6"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA6='" & sTDATA6 & "'"
  end if
  if txtDlLabel7 <> "" then
    sTDATA7 = ChkString(Request.Form("TDATA7"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA7='" & sTDATA7 & "'"
  end if
  if txtDlLabel8 <> "" then
    sTDATA8 = ChkString(Request.Form("TDATA8"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA8='" & sTDATA8 & "'"
  end if
  if txtDlLabel9 <> "" then
    sTDATA9 = ChkString(Request.Form("TDATA9"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA9='" & sTDATA9 & "'"
  end if
  if txtDlLabel10 <> "" then
    sTDATA10 = ChkString(Request.Form("TDATA10"),"sqlstring")
	sT1Sql = sT1Sql & ", TDATA10='" & sTDATA10 & "'"
  end if
  
  unmarkbad = Request.Form("unmarkbad")
  marknew = Request.Form("marknew")
  markupd = Request.Form("markupd")
  approve = Request.Form("approve")
  today = strCurDateString
  
    if approve = "1" then
	  mod_increaseSubcatCount(subcat)
    end if
  
  ':: check for subcat change
  if orig_subcat <> subcat then
    if approve <> "1" then
	  mod_increaseSubcatCount(subcat)
	  mod_decreaseSubcatCount(orig_subcat)
    end if
    if left(URL,16) = "files/downloads/" then
	  if instr(URL,"/" & orig_subcat & "/") > 0 then
	    ':: edit URL
  	    tUrl = replace(URL,"/" & orig_subcat & "/","/" & subcat & "/")
	    ':: move file to new subcat folder
		call moveFile(server.MapPath(URL),server.MapPath(tUrl))
	    URL = tUrl
	  end if
	end if
  end if
  
  sSql = "UPDATE DL SET "
  sSql = sSql & "NAME='" & iname & "'"
  sSql = sSql & ",CATEGORY =" & subcat & ""
  sSql = sSql & ",EMAIL ='" & email & "'"
  sSql = sSql & ",DESCRIPTION ='" & sdesc & "'"
  sSql = sSql & ",CONTENT ='" & ldesc & "'"
  sSql = sSql & ",URL ='" & URL & "'"
  sSql = sSql & ",FILESIZE ='" & filesize & "'"
  sSql = sSql & ",KEYWORD ='" & key & " '"
  sSql = sSql & sT1Sql
  if marknew = "1" then
    sSql = sSql & ",UPDATED ='0'"
    sSql = sSql & ",O_POST_DATE ='" & today & "'"
    sSql = sSql & ",POST_DATE ='" & today & "'"
  end if
  if markupd = "1" then
    sSql = sSql & ",UPDATED ='" & today & "'"
    sSql = sSql & ",O_POST_DATE ='" & today & "'"
  end if
  if approve = "1" then
    sSql = sSql & ",ACTIVE = 1"
  end if
  if unmarkbad = "1" then
    sSql = sSql & ",BADLINK = 0"
  end if
  sSql = sSql & " WHERE DL_ID =" & itemID
  'response.Write(email)
  'response.End()
  my_Conn.Execute(sSql)
  
	  if approve = "1" and intSubscriptions = 1 and strEmail = 1 then
	    'send subscriptions emails
		sSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & subcat
		set rsA = my_Conn.execute(sSql)
		  parent = rsA(0)
		set rsA = nothing
	    eSubject = strSiteTitle & " - New Download"
		eMsg = "A new download has been submitted at " & strSiteTitle & vbCrLf
		eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		eMsg = eMsg & "You can view the new downloads by visiting " & strHomeUrl & "dl.asp?cmd=3" & vbCrLf
	    sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
		'response.Write("<br />Email sent<br />" )
	  end if
  
  Call setSession("sMsg","Item successfully updated")
  resetCoreConfig()
  if iPgType = 23 then
    'closeAndGo(sDLpage & "?cmd=6&cid=" & itemID)
  else
    'closeAndGo(sDLpage & "?cmd=" & iPgType & "&sid=" & subcat)
  end if
  
end sub

sub showEditForm(oo)
  if bAppFull or bCatFull or bSubcatFull or (strDBNTUserName = oo("UPLOADER")) then
  Response.Write "<form method=""post"" action=""dl_edit_url.asp"" enctype=""multipart/form-data"">"
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""4"" width=""100%"" align=""center"">"
  Response.Write "<tr><td align=""right"" width=""30%"">"
  Response.Write "<b>Subcategory:&nbsp;</b></td>"
  Response.Write "<td>" '& 
  'GetCategories(oo("SUBCAT_ID"))
  mod_selectCatSubcat oo("SUBCAT_ID"),"WRITE"
  Response.Write "</td></tr>"
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b><span class=""fAlert"">*</span>"
  Response.Write "File Name:</b></td>"
  Response.Write "<td>"
  Response.Write "<input type=""text"" name=""name"" size=""40"" maxlength=""90"" value=""" & oo("NAME") & """ />"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b><span class=""fAlert"">*</span>"
  Response.Write "URL of the file:</b></td>"
  Response.Write "<td>"
  Response.Write "<input type=""text"" name=""url"" size=""40"" maxlength=""190"" value=""" & oo("URL") & """ />"
  Response.Write "</td></tr>"
  Response.Write "<tr><td colspan=""2"" align=""center"">"
  Response.Write "<a href="""& oo("URL") &""">"& oo("URL") &"</a>"
  Response.Write "<br>&nbsp;</td></tr>"
  
  ':: Start Upload File code
		  	strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'download'"
			set rsUload = my_Conn.execute(strSQL)
			uActive = rsUload("UP_ACTIVE")
			uUpGrps = rsUload("UP_ALLOWEDGROUPS")
			uSize = rsUload("UP_SIZELIMIT")
			uExt = rsUload("UP_ALLOWEDEXT")
			uID = rsUload("ID")
			set rsUload = nothing
		  	session.Contents("uploadType") = uID
		  	session.Contents("loggedUser") = strdbntusername
		  If bFSO = true and strAllowUploads = 1 and uActive = 1 and hasAccess(uUpGrps) Then
		    ast = "**"
			btxt = "<span class=""fAlert"">**</span> = " & txtLnkOrUpld & "<br />" %>
          <tr>
            <td align="right">&nbsp;</td>
            <td align="left" valign="top">
			  <br /><%= txtMaxUpldSize %> <b><%= uSize %> kb</b><br />
			  <%= txtAllowExt %> <b><%= uExt %></b><br />
            </td>
          </tr>
          <tr>
            <td align="right">
			  <span class="fAlert">**</span> Upload new file:&nbsp; </td>
            <td><input type="hidden" name="max" value="1" />
              <input class="textbox" name="file1" id="file1" type="file" size="30" />
            </td>
          </tr>
		  <% Else
		  	   ast = "*"
			   btxt = "" %>
		  		<input class="textbox" name="file1" id="file1" type="hidden" value="" />
		  <% End If
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  if oo("BADLINK") <> 0 then
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>Bad link reported by:</b>"
  Response.Write "&nbsp;</td><td>"
  Response.Write "<b>" & getMemberName(oo("BADLINK")) & "</b>"
  Response.Write "</td></tr>"
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  end if
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>Submitted by:</b></td>"
  Response.Write "<td>"
  Response.Write "<b>" & oo("UPLOADER") & "</b>"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>Submitter Email:</b></td>"
  Response.Write "<td>"
  Response.Write "<input type=""text"" name=""email"" size=""40"" maxlength=""90"" value=""" & oo("EMAIL") & """ />"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>File Size:</b></td>"
  Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""size"" size=""40"" maxlength=""90"" value=""" & oo("FILESIZE") & """ />"
  if left(oo("URL"),4) = "http" then
  Response.Write "<input type=""text"" name=""size"" size=""40"" maxlength=""90"" value=""" & oo("FILESIZE") & """ />"
  else
  Response.Write "<input type=""text"" name=""size"" size=""40"" maxlength=""90"" value=""" & mod_getFileInfo(server.MapPath(oo("URL")),"Size") & """ />"
  end if
  Response.Write "</td></tr>"
  
  customEditFormElements(oo)
  
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>License:</b></td>"
  'Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""license"" size=""40"" maxlength=""90"" value=""" & oo("LICENSE") & """ />"
  'Response.Write "</td></tr>"
  
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>Language:</b></td>"
  'Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""language"" size=""40"" maxlength=""90"" value=""" & oo("LANG") & """ />"
  'Response.Write "</td></tr>"
  
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>Platform:</b></td>"
  'Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""platform"" size=""40"" maxlength=""90"" value=""" & oo("PLATFORM") & """ />"
  'Response.Write "</td></tr>"
  
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>Publisher:</b></td>"
  'Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""publisher"" size=""40"" maxlength=""90"" value=""" & oo("PUBLISHER") & """ />"
  'Response.Write "</td></tr>"
  
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>Publisher URL:</b></td>"
  'Response.Write "<td>"
  'Response.Write "<input type=""text"" name=""publisherURL"" size=""40"" maxlength=""90"" value=""" & oo("PUBLISHER_URL") & """ />"
  'Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>Keywords:</b></td>"
  Response.Write "<td>"
  Response.Write "<input type=""text"" name=""key"" size=""40"" maxlength=""240"" value=""" & oo("KEYWORD") & """ />"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b><span class=""fAlert"">* </span>"
  Response.Write "Short Description: </b><br /><br />"
  Response.Write "<span id=""charLeft"">" & 250-len(oo("DESCRIPTION")) & " characters left&nbsp;"
  Response.Write "</span></td>"
  Response.Write "<td>"
  Response.Write "<textarea rows=""7"" name=""sdes"" id=""sdes"" cols=""50"" wrap=""virtual"" onKeyUp=""cntChar('sdes','charLeft','{CHAR} characters left.',250);"">"
  Response.Write oo("DESCRIPTION")
  Response.Write "</textarea>"
  Response.Write "</td></tr>"
  
  Response.Write "<tr><td colspan=""2"">&nbsp;"
  Response.Write "</td></tr>"
  
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b><span class=""fAlert"">*</span> Long Description:</b> ",oo("CONTENT")
  else
  	displayPLAINeditor 1,oo("CONTENT")
  end if
  
  Response.Write "<tr><td colspan=""2""><hr/>"
  Response.Write "</td></tr>"
  
  if oo("BADLINK") <> 0 and (bAppFull or hasAccess(oo("CG_FULL")) or hasAccess(oo("SG_FULL")) or strDBNTUserName = oo("UPLOADER")) then
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""unmarkbad"" value=""1"" checked=""checked"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Unmark as bad</b>"
    Response.Write "</td></tr>"
  end if
  
  if oo("ACTIVE") = 0 then
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""approve"" value=""1"" checked=""checked"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Approve Item</b>"
    Response.Write "</td></tr>"
  else
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""markupd"" value=""1"" checked=""checked"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Mark as Updated</b>"
    Response.Write "</td></tr>"
    Response.Write "<tr><td align=""right"">"
    Response.Write "<input type=""checkbox"" name=""marknew"" value=""1"" />"
    Response.Write "&nbsp;</td><td>"
    Response.Write "<b>Mark as New</b>"
    Response.Write "</td></tr>"
  end if
  
  Response.Write "<tr><td colspan=""2""><hr/>"
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">&nbsp;</td>"
  Response.Write "<td><br />"
  'Response.Write "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />"
  Response.Write "<input type=""hidden"" name=""mode"" value=""322"" />"
  Response.Write "<input type=""hidden"" name=""itemID"" value=""" & oo("DL_ID") & """ />"
  Response.Write "<input type=""hidden"" name=""orig_subcat"" value=""" & oo("SUBCAT_ID") & """ />"
  Response.Write "<input id=""button"" class=""button"" type=""submit"" value="" Update "" style=""width:150px;height:25px;"" name=""B1"" accesskey=""s"" title=""Shortcut Key: Alt+S"" />"
  Response.Write "</td></tr>"
  Response.Write ""
  
  Response.Write "</table>"
  Response.Write "</form>"
  else
    closeAndGo("error.asp?type=noperm")
  end if
end sub

sub deleteItem(idd)
  sSql = singleDLsql()
  sSql = sSql & "WHERE (((DL.DL_ID)=" & idd & "))"
  set rsB = my_Conn.execute(sSql)
  if rsB.eof then
	'Response.Write("No items found")
  else
    isOwner = false
    bFull = false
    if bAppFull or hasAccess(rsB("CG_FULL")) or hasAccess(rsB("CG_FULL")) then
      bFull = true
    end if
    if bFull or (strDBNTUserName = rsB("UPLOADER")) then
      isOwner = true
    end if
  end if
  
  if bFull or isOwner then
    doDelete(rsB)
  end if
  whereTo = sDLpage & "?cmd=2&cid=" & rsB(sMCPre & "CAT_ID") & "&sid=" & rsB("SUBCAT_ID")
  itemID = rsB("DL_ID")
  set rsB = nothing
  executeThis("DELETE from DL where DL_ID=" & itemID)
  resetCoreConfig()
  closeAndGo(whereTo)
end sub

sub doDelete(ob)
  if bFSO = true then
    sSQL = "select UP_FOLDER from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
    set rsU = my_Conn.execute(sSQL)
	  downloadDir = rsU("UP_FOLDER")
    set rsU = nothing
	
      tmpBanner = ob("URL")
	  banner = right(tmpBanner, len(tmpBanner) - instrrev(tmpBanner,"/"))
	  set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirFPath = server.MapPath(downloadDir) & "\" & ob("CATEGORY") & "\" & replace(banner,"_rs.",".")
		if fso.FileExists(dirFPath) = true then
			fso.DeleteFile dirFPath
		end if
		dirPath = server.MapPath(downloadDir) & "\" & ob("CATEGORY") & "\" & banner
		if fso.FileExists(dirPath) = true then
			fso.DeleteFile dirPath
		end if
		if fso.FileExists(dirPath) = true then
			fsoMsg = "The file was not removed from the server"
		end if
	  set fso = nothing
  end if
  executeThis("DELETE from " & strTablePrefix & "M_RATING where ITEM_ID=" & ob("DL_ID") & " AND APP_ID = " & intAppID & "")
  mod_decreaseSubcatCount(ob("SUBCAT_ID"))
  strMsg = strMsg & "The Download and all of its data has been deleted"
  if fsoMsg <> "" then
    strMsg = strMsg & "<br />" & fsoMsg
  end if
  Call setSession("sMsg",strMsg)
end sub

sub chkSubCatAttention(i)
  tAtt = getCount("DL_ID","DL","(ACTIVE = 0 OR BADLINK <> 0) AND CATEGORY=" & i)
  if tAtt > 0 then
	Response.Write "<a href="""&sDLpage&"?cmd=22&amp;sid=" & i & """>"
	Response.Write icon(icnAttention,"Items need attention","","","align=""middle""")
	Response.Write "</a>"
  end if
end sub

sub deleteDlCategory(c)  
  sSQL = "select UP_FOLDER from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
  set rsU = my_Conn.execute(sSQL)
	downloadDir = rsU("UP_FOLDER")
  set rsU = nothing
  
  sSql = mod_CatSubCatsql(c,0,intAppID)
  'sSql = sSql & "WHERE " & strTablePrefix & "M_CATEGORIES.CAT_ID=" & c & " AND " & strTablePrefix & "M_CATEGORIES.APP_ID = " & intAppID & ""
  
  set rsA = my_Conn.execute(sSQL)
  if not rsA.eof then
    cat= rsA("CAT_NAME")
	do until rsA.eof
  
      ':: delete ratings
      sSQL = "SELECT DL_ID FROM DL WHERE CATEGORY=" & rsA("SUBCAT_ID")
      set rsDel = my_Conn.execute(sSQL)
      do until rsDel.eof
  	    executeThis("DELETE from " & strTablePrefix & "M_RATING where ITEM_ID=" & rsDel("DL_ID") & " AND APP_ID = " & intAppID & "")
	    rsDel.movenext
	  loop
      set rsDel = nothing
	
	  ':: 
  	  if bFSO = true then
	    set fso = Server.CreateObject("Scripting.FileSystemObject")
	    dirFPath = server.MapPath(downloadDir & rsA("SUBCAT_ID"))
	    if fso.FolderExists(dirFPath) = true then
		  set objF = fso.getfolder(dirFPath)
		  objF.Delete
		  set objF = nothing
	    end if
	    if fso.FolderExists(dirFPath) = true then
		  strMsg = strMsg & "<h4>Category Folder /" & downloadDir & rsA("SUBCAT_ID") & " could not be deleted</h4><br />"
	    else
		  'strMsg = strMsg & "<h3>Category Folder successfully deleted</h3><br />"
	    end if
	    set fso = nothing
  	  end if
      executeThis("DELETE FROM DL WHERE CATEGORY=" & rsA("SUBCAT_ID"))
	  rsA.movenext
	loop
    strMsg = strMsg & "Category (<span class=""fAlert"">" & c & "</span>) along with all Sub-Categories"
    strMsg = strMsg & "<br />and associated data have been deleted.<br /><br />"
  else
    strMsg = strMsg & "Category Deleted!<br>(<span class=""fAlert"">" & c & "</span>)"
  end if
  set rsA = nothing
  
  executeThis("DELETE FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & c & " AND APP_ID = " & intAppID & "")
  executeThis("DELETE FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE CAT_ID=" & c & " AND APP_ID = " & intAppID & "")
  Call setSession("sMsg",strMsg)
  resetCoreConfig()
  closeAndGo(sDLpage & "?cmd=" & iPgType & "")
end sub

sub deleteDlSubCategory(sc)
  'cat = trim(chkString(Request.Form("cat"), "sqlstring"))
  'scat = trim(chkString(Request.Form("scat"), "sqlstring"))
	sSQL = "select UP_FOLDER from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
	set rsU = my_Conn.execute(sSQL)
	  downloadDir = rsU("UP_FOLDER")
	set rsU = nothing
	
  sSql2 = "SELECT CAT_ID, SUBCAT_NAME From " & strTablePrefix & "M_SUBCATEGORIES where SUBCAT_ID=" & sc
  set rsDel = my_Conn.execute(sSQL2)
    c = rsDel("CAT_ID")
	scn = rsDel("SUBCAT_NAME")
  set rsDel = nothing
    sSQL = "SELECT * FROM DL WHERE CATEGORY=" & sc
    set rsDel = my_Conn.execute(sSQL)
     if not rsDel.eof then
      do until rsDel.eof
	    'delete ratings
  	    executeThis("DELETE from " & strTablePrefix & "M_RATING where ITEM_ID=" & rsDel("DL_ID") & " AND APP_ID = " & intAppID & "")
	    rsDel.movenext
	  loop
     end if
    set rsDel = nothing
	  'if uploaded files, lets delete the subcat folder
  	  if bFSO = true then
	  	set fso = Server.CreateObject("Scripting.FileSystemObject")
		  dirFPath = server.MapPath(downloadDir & sc)
		  if fso.FolderExists(dirFPath) = true then
		    set objF = fso.getfolder(dirFPath)
			objF.Delete
			set objF = nothing
		  end if
		  if fso.FolderExists(dirFPath) = true then
			strMsg = strMsg & "<h4>Folder could not be deleted</h4>"
			strMsg = strMsg & "<b>" & dirFPath & "</b>"
		  else
			'strMsg = strMsg & "<h3>SubFolder successfully deleted</h3>"
		  end if
	  	set fso = nothing
  	  end if
    executeThis("delete From " & strTablePrefix & "M_SUBCATEGORIES where SUBCAT_ID=" & sc)
    executeThis("delete From DL where CATEGORY=" & sc)
    strMsg = strMsg & "Subcategory: <b>" & scn & "</b><br />"
    strMsg = strMsg & "and all its contents have been deleted."
	Call setSession("sMsg",strMsg)
    resetCoreConfig()
    closeAndGo(sDLpage & "?cmd=" & iPgType & "&cid=" & cid & "")
end sub
%>
