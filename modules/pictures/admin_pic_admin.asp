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
CurPageType = "pictures"
%>
<!-- #INCLUDE file="lang/en/core_admin.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %> 
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/inc_admin_functions.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<%
dim iPgType, iMode, cid, sid, strMsg, nPageTo, nPageCnt
strMsg = ""
iPgType = 0
iMode = 0
cid = 0
sid = 0
nPageTo = ""

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

	' get module config info
call setAppPerms("pictures","iName")
	
	sSQL = "select UP_FOLDER from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
	set rsU = my_Conn.execute(sSQL)
	  galleryDir = rsU("UP_FOLDER")
	set rsU = nothing

Dim cat_id,sub_id,strURL,strTURL,strPoster,strpicTitle,strSummary
Dim intpicID,intHit,intShow,strOwner,strPostDate,dateSince

if iMode = 1 then 'add category
  newcat = trim(chkString(Request.Form("newcat"), "sqlstring"))
  if newcat = "" then
	strMsg = strMsg & "<b>Please enter category name</b>"
  else
	Set RS=Server.CreateObject("ADODB.Recordset")
	strSql="Select CAT_NAME from PIC_CATEGORIES where CAT_NAME='" & newcat & "'"
	RS.Open strSql,my_Conn , 2, 2
	if rs.eof then
		been_here_before="No"
	end if
	RS.close
	set RS = nothing

	if been_here_before="No" then 
		sSql = "insert into PIC_CATEGORIES ("
		sSql = sSql & "CAT_NAME, CG_READ, CG_WRITE, CG_FULL, CG_INHERIT, CG_PROPAGATE"
		sSql = sSql & ") values ("
		sSql = sSql & "'" & newcat & "','" & sAppRead & "'"
		sSql = sSql & ",'" & sAppWrite & "','" & sAppFull & "'"
		sSql = sSql & ",1,1"
		sSql = sSql & ");"
		set rsinsert = my_Conn.Execute (sSql)
		strMsg = strMsg & "New Category <b>" & newcat & "</b> Added"
	else
		strMsg = strMsg & "<b>This category name already exists, please enter a different category name</b>"
	end if
	'iPgType = 1
  end if
  
elseif iMode = 2 then 'add subcategory
  newsub = trim(chkString(Request.Form("newsub"), "sqlstring"))
  if newsub = "" then
	strMsg = strMsg & "Please enter a subcategory name"
  else
	if cid = 0 then
		strMsg = strMsg & "Please select a category for the new subcategory"
	else
		strSql="Select CAT_NAME, CG_READ, CG_WRITE, CG_FULL from PIC_CATEGORIES where CAT_ID=" & cid
		set rsC = my_Conn.execute(strSql)
		   CAT_NAME = rsC("CAT_NAME")
		   cR = rsC("CG_READ")
		   cW = rsC("CG_WRITE")
		   cF = rsC("CG_FULL")
		set rsC = nothing
		sSql = "insert into PIC_SUBCATEGORIES ("
		sSql = sSql & "SUBCAT_NAME, CAT_ID, SG_READ, SG_WRITE, SG_FULL, SG_INHERIT"
		sSql = sSql & ") values ("
		sSql = sSql & "'" & newsub & "'," & cid & ""
		sSql = sSql & ",'" & cR & "','" & cW & "','" & cF & "'"
		sSql = sSql & ",1);"
		sSql = sSql & ""
		executeThis(sSql)
		strMsg = strMsg & "New subcategory <span class=""fAlert""><b>" & newsub & "</b></span> has been<br>"
		strMsg = strMsg & "added to category <b>" & CAT_NAME & "</b>"
	end if
  end if
  
elseif iMode = 3 then 'rename category
  cat= ChkString(Request.Form("newcat"),"SQLString")
  if trim(cat) <> "" then
    sSql = "UPDATE PIC_CATEGORIES SET CAT_NAME='" & cat & "' WHERE CAT_ID=" & cid
    executeThis(sSql)
    strMsg = strMsg & "Category renamed to <b>" & cat & "</b>"
	iPgType = 2
  else
    strMsg = strMsg & "You must chose a name in<br>order to rename the category"
  end if
  
elseif iMode = 4 then 'delete category
  cat= ChkString(Request.Form("cat"),"SQLString")
  sSQL = "SELECT PIC_ID FROM PIC WHERE PARENT_ID=" & cid
  set rsDel = my_Conn.execute(sSQL)
    do until rsDel.eof
	  'delete picture ratings
  	  executeThis("DELETE from PIC_RATING where PIC=" & rsDel("PIC_ID"))
	  rsDel.movenext
	loop
  set rsDel = nothing
	  'if uploaded pictures, lets delete the cat folder
  	  if bFso then
	  	set fso = Server.CreateObject("Scripting.FileSystemObject")
		  dirFPath = server.MapPath(galleryDir & cid)
		  if fso.FolderExists(dirFPath) = true then
		    set objF = fso.getfolder(dirFPath)
			for each p in objF.SubFolders
		      set objSF = fso.getfolder(p.path)
			  for each sf in objF.Files
			    fso.DeleteFile sf.path
			  next
			  set objSF = nothing
			  fso.DeleteFolder p.path
			next
			for each pf in objF.Files
			  fso.DeleteFile pf.path
			next
			set objF = nothing
			fso.DeleteFolder dirFPath
		  end if
		  if fso.FolderExists(dirFPath) = true then
			strMsg = strMsg & "<h4>Category Folder /" & galleryDir & cid & " could not be deleted</h4><br>"
		  else
			'strMsg = strMsg & "<h3>Category Folder successfully deleted</h3><br>"
		  end if
	  	set fso = nothing
  	  end if
  executeThis("delete From PIC_CATEGORIES where CAT_ID=" & cid)
  executeThis("delete From PIC_SUBCATEGORIES where CAT_ID=" & cid )
  executeThis("delete From PIC where PARENT_ID=" & cid )
  strMsg = strMsg & "Category (<span class=""fAlert"">" & cat & "</span>) along with all Sub-Categories"
  strMsg = strMsg & "<br>and associated pictures have been deleted.<br><br>"
  
elseif iMode = 5 then
  cat = trim(chkString(Request.Form("newcat"), "sqlstring"))
  executeThis("UPDATE PIC_SUBCATEGORIES SET SUBCAT_NAME='" & cat & "' WHERE SUBCAT_ID=" & sid)
  strMsg = strMsg & "Sub-Category name changed!"
  
elseif iMode = 6 then 'delete subcategory
  cat = trim(chkString(Request.Form("cat"), "sqlstring"))
  scat = trim(chkString(Request.Form("scat"), "sqlstring"))
  sSQL = "SELECT PIC_ID, URL, TURL, PARENT_ID FROM PIC WHERE CATEGORY=" & sid
  set rsDel = my_Conn.execute(sSQL)
    if rsDel.eof and rsDel.bof then
   	  parentID = 0
    else
      parentID = rsDel("PARENT_ID")
	end if

    do until rsDel.eof
	  'delete picture rating
  	  executeThis("DELETE from PIC_RATING where PIC=" & rsDel("PIC_ID"))
	  rsDel.movenext
	loop
	  'if uploaded pictures, lets delete the subcat folder
	if parentID <> 0 then
  	  if bFso then
	  	set fso = Server.CreateObject("Scripting.FileSystemObject")
		  dirFPath = server.MapPath(galleryDir & parentID & "/" & sid)
		  'response.Write("Path: " & dirFPath & "<br>")
		  'response.End()
		  if fso.FolderExists(dirFPath) = true then
		    set objF = fso.getfolder(dirFPath)
			for each p in objF.Files
			  fso.DeleteFile p.path
			next
			set objF = nothing
		  end if
		  fso.DeleteFolder dirFPath
		  if fso.FolderExists(dirFPath) = true then
			strMsg = strMsg & "<h4>Folder could not be deleted</h4>"
		  else
			'strMsg = strMsg & "<h3>SubFolder successfully deleted</h3>"
		  end if
	  	set fso = nothing
  	  end if
    end if
  set rsDel = nothing
  executeThis("delete From PIC_SUBCATEGORIES where SUBCAT_ID=" & sid)
  executeThis("delete From PIC where CATEGORY=" & sid)
  strMsg = strMsg & "<b>" & cat & "</b> subcategory: <b>" & scat & "</b><br>"
  strMsg = strMsg & "and all its contents have been deleted."
  
elseif iMode = 7 then 'delete picture
  if bFso then
	
    sSQL = "select CATEGORY, PARENT_ID, URL, TURL FROM PIC where PIC_ID = " & cid
    set rsUp = my_Conn.execute(sSQL)
      tmpBanner = rsUp("URL")
      tmpThumb = rsUp("TURL")
	  banner = right(tmpBanner, len(tmpBanner) - instrrev(tmpBanner,"/"))
	  Tbanner = right(tmpThumb, len(tmpThumb) - instrrev(tmpThumb,"/"))
	  set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirFPath = server.MapPath(galleryDir) & "\" & rsUp("PARENT_ID") & "\" & rsUp("CATEGORY") & "\" & replace(banner,"_rs.",".")
		if fso.FileExists(dirFPath) = true then
			fso.DeleteFile dirFPath
		end if
		dirPath = server.MapPath(galleryDir) & "\" & rsUp("PARENT_ID") & "\" & rsUp("CATEGORY") & "\" & banner
		if fso.FileExists(dirPath) = true then
			fso.DeleteFile dirPath
		end if
		dirTPath = server.MapPath(galleryDir) & "\" & rsUp("PARENT_ID") & "\" & rsUp("CATEGORY") & "\" & Tbanner
		if fso.FileExists(dirTPath) = true then
			fso.DeleteFile dirTPath
		end if
	  set fso = nothing
  end if
  executeThis("DELETE from PIC where PIC_ID=" & cid)
  executeThis("DELETE from PIC_RATING where PIC=" & cid)
  strMsg = strMsg & "The picture and all of its data has been deleted"
  
elseif iMode = 8 then 
  LocFile = false
  TLocFile = false
  webid = cLng(Request("wid"))
  
  set rsUp = my_Conn.execute("select URL, TURL, CATEGORY, PARENT_ID from PIC where PIC_ID = " & webid)
    frm = rsUp("PARENT_ID") & "\" & rsUp("CATEGORY")
    tmpBanner = rsUp("URL")
    tmpTBanner = rsUp("TURL")
	banner = right(tmpBanner, len(tmpBanner) - instrrev(tmpBanner,"/"))
	Tbanner = right(tmpTBanner, len(tmpTBanner) - instrrev(tmpTBanner,"/"))
  set rsUp = nothing
	
	if bFso then	
	    set fso = Server.CreateObject("Scripting.FileSystemObject")
		dirPath = server.MapPath(galleryDir) & "\"
		if fso.FolderExists(dirPath & cid) = false then
			fso.CreateFolder(dirPath & cid)
		end if
		if fso.FolderExists(dirPath & cid & "\" & sid) = false then
			fso.CreateFolder(dirPath & cid & "\" & sid)
		end if
		if fso.FileExists(dirPath & frm & "\" & replace(banner,"_rs","")) = true then
			fso.MoveFile dirPath & frm & "\" & replace(banner,"_rs",""), dirPath & cid & "\" & sid & "\" & replace(banner,"_rs","")
			LocFile = true
		end if
		if fso.FileExists(dirPath & frm & "\" & banner) = true then
			fso.MoveFile dirPath & frm & "\" & banner, dirPath & cid & "\" & sid & "\" & banner
			LocFile = true
		end if
		if fso.FileExists(dirPath & frm & "\" & Tbanner) = true then
			fso.MoveFile dirPath & frm & "\" & Tbanner, dirPath & cid & "\" & sid & "\" & Tbanner
			TLocFile = true
		end if
	    set fso = nothing
	end if
	
	nBanner = strHomeUrl & galleryDir & cid & "/"  & sid & "/"  & banner
	nTBanner = strHomeUrl & galleryDir & cid & "/"  & sid & "/"  & Tbanner
	
	sSQL = "UPDATE PIC set Category=" & sid & ", PARENT_ID=" & cid
	if LocFile = true then
		sSQL = sSQL & ", URL='" & nBanner & "'"
	end if
	if TLocFile = true then
		sSQL = sSQL & ", TURL='" & nTBanner & "'"	
	end if
	sSQL = sSQL & " where PIC_ID =" & webid
 	executeThis(sSQL)
	
    strMsg = strMsg & "The picture has been moved to a new category/subcategory"
  
elseif iMode = 9 then
  sSql = "UPDATE PIC_CATEGORIES SET C_ORDER=" & sid & " WHERE CAT_ID=" & cid
  executeThis(sSql)
  strMsg = strMsg & "Picture category order updated"
  
elseif iMode = 10 then
  ord = clng(request("ord"))
  sSql = "UPDATE PIC_SUBCATEGORIES SET C_ORDER=" & ord & " WHERE SUBCAT_ID=" & sid
  executeThis(sSql)
  strMsg = strMsg & "Picture sub-category order updated"
  
elseif iMode = 11 then 'unmark bad link
  ord = clng(request("ord"))
  sSql = "UPDATE PIC SET BADLINK=0 WHERE PIC_ID=" & cid
  executeThis(sSql)
  strMsg = strMsg & "Picture updated"
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign=top class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	pictureConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	spThemeBlock1_close(intSkin) %>
		</td>
		<td class="mainPgCol">
<%
  intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = "Admin Area|admin_home.asp"
  arg2 = "Picture Admin|admin_pic_main.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

	  	   spThemeBlock1_open(intSkin) 
		   If strMsg <> "" Then %>
			  <span class="fTitle"><%= strMsg %></span>
		<% End If %>
		<% select case iPgType
			  case 1
			    addSub()
			  case 2
				iCmd = 3
				strTxt = "Edit Category...<br>"
				strTxt = strTxt & "Click a category to rename it or<br>"
				strTxt = strTxt & "use the drop down boxes to re-order them"
				showCats() 'order
			  case 3
			    editCat()
			  case 4
			    delCat()
			  case 5
				iCmd = 6
				strTxt = "Edit Subcategory...<br>"
				strTxt = strTxt & "Click the category that the subcategory is in"
				showCats()
			  case 6
				iCmd = 6
				iLink = "admin_pic_admin.asp?cmd=7&sid="
				strTxt = "Edit Subcategory...<br>"
				strTxt = strTxt & "<b>Click the subcategory to edit."
			    showSubs() 'order
			  case 7
			    editSub()
			  case 8
				iCmd = 9
				strTxt = "Delete Subcategory...<br>"
				strTxt = strTxt & "Click the category that the subcategory is in"
				showCats()
			  case 9
			    delPickSub()
			  case 10
				iCmd = 11
				strTxt = "Edit Picture...<br>"
				strTxt = strTxt & "Click the category that the Picture is in."
				showCats()
			  case 11
				'iCmd = 12
				iLink = "admin_pic_admin.asp?cmd=12&sid="
				strTxt = "Edit Picture...<br>"
				strTxt = strTxt & "Click the subcategory that the Picture is in."
			    showSubs()
			  case 12 'pick Picture to edit
				'iCmd = 22
				iLink = "admin_pic_editpic.asp?id="
				strTxt = "Edit Picture...<br>"
				strTxt = strTxt & "Click the Picture Title that you want to edit."
				showArticles()
			  case 13
				iCmd = 14
				webid = cid
				'iLink = "admin_pic_admin.asp?cmd=33&cid="
				strTxt = strTxt & "First Step Complete<br>"
				strTxt = strTxt & "Picture updated!."
				strTxt = strTxt & "<br><br>"
				strTxt = strTxt & "Second Step : Optional<br>"
				strTxt = strTxt & "Move the Picture to a new category by clicking a category below.<br>"
				strTxt = strTxt & "Do nothing and no further changes will be made."
				strTxt = strTxt & "<br><br>"
				showCats()
			    'picUpdate_S1()
			  case 14
				'iCmd = 22
				webid = cLng(request("wid"))
				iLink = "admin_pic_admin.asp?cmd=10&mode=8&wid=" & webid & "&cid=" & cid & "&sid="
				strTxt = strTxt & "<b>First Step Complete</b><br>"
				strTxt = strTxt & "Picture updated!."
				strTxt = strTxt & "<br><br>"
				strTxt = strTxt & "<b>Second Step Complete</b><br>"
				strTxt = strTxt & "Category selected"
				strTxt = strTxt & "<br><br>"
				strTxt = strTxt & "<b>Third Step</b><br>"
				strTxt = strTxt & "Move the Picture to a new subcategory by clicking a subcategory below.<br>"
				strTxt = strTxt & "Do nothing and no changes will be made."
				strTxt = strTxt & "<br><br>"
			    showSubs()
			  case 20
				iCmd = 21
				strTxt = "Delete Picture...<br>"
				strTxt = strTxt & "Click the category that the Picture is in."
				showCats()
			  case 21
				'iCmd = 22
				iLink = "admin_pic_admin.asp?cmd=22&sid="
				strTxt = "Delete Picture...<br>"
				strTxt = strTxt & "Click the subcategory that the Picture is in."
			    showSubs()
			  case 22
				iLink = "admin_pic_admin.asp?cmd=20&mode=7&cid="
				strTxt = "Delete Picture...<br>"
				strTxt = strTxt & "Click the Picture Title that you want to delete."
				strTxt = strTxt & "<br><br><b>THIS CANNOT BE UNDONE</b>."
				showArticles()
			  case 30
			    'dim strSummary,strARTICLETitle,strPostDate
				'dim dateSince,intHit,intARTICLEID
				strTxt = "Browse All Pictures..."
				browseArticles()
			  case 40 ' check bad links
				iLink = "admin_pic_admin.asp?cmd=40&mode=7&cid="
				strTxt = "Bad Links...<br>"
				strTxt = strTxt & "These pictures have been reported by users:"
				strTxt = strTxt & "<br>Click each picture link to correct the error.."
				showBadLinks()
			  case else
			    addCat()
		   end select
		   response.Write("<br>&nbsp;")
	  	   spThemeBlock1_close(intSkin) %>
		</td>
	</tr>
</table>
<!-- #INCLUDE FILE="inc_footer.asp" -->
<% Else %><% Response.Redirect "admin_login.asp?target=admin_pic_main.asp" %><% End If %>
<%

sub displayArticle
%><hr />
		<center>
          <table border="0" width="500" cellspacing="1" cellpadding="2" class="tCellAlt2">
          <tr>
            <td rowspan="2" align="center" valign="middle" width="150"><%
			if instr(strTURL,"_sm") > 0 then
			  stImg = "<img src=""" & strTURL & """ border=""0"" alt=""Image"" title=""Image"" />"
			else
			  stImg = "<img src=""" & strURL & """ border=""0"" width=""120"" alt=""image"" title=""Image"" />"
			end if 
  			response.Write(stImg)%>	
		    </td>
          <td class="tSubTitle" height="20" width="300" valign="top">
	  <a href="admin_pic_editpic.asp?id=<%=intpicID%>"><%= strpicTitle %></a><br><span class="fAlert"><% if intShow = 0 then%> [not approved]<%else%> [approved]<%end if%><% if strOwner = "0" then%> [public]<%else%> [private]<%end if%></span>
	 </td><td height="20" class="tTitle" width="100" valign="top"><span class="fAlert">Hit : <%=intHit%> </span></td>
	 </tr>
	 <tr>
	 <td class="tCellAlt1" valign="top">
	 <%= strSummary %></td>
	</tr>
	</table>
	</center><hr />
<%
end sub

function GetRating(ArticleID)
	strSQL = "SELECT VOTES, RATING FROM PIC WHERE PIC_ID = " & intArticleID
	set rsArticleRating = server.CreateObject("adodb.recordset")
	rsArticleRating.Open strSQL, my_Conn
	dim intVotes
	dim intRating
	intVotes = rsArticleRating("VOTES")
	intRating = rsArticleRating("RATING")
	rsArticleRating.Close
	set rsArticleRating = nothing
	if intVotes > 0 then
		intRating = Round(intRating/intVotes)
		Response.Write(" Rating: " & intRating & " Votes: " & intVotes)
	end if
end function

sub browseArticles()
  iPageSize = 10
  nPageTo = 1
  Set RS=Server.CreateObject("ADODB.Recordset")
  If Request("PageTo") <> "" Then
	strPageTo = Request("PageTo")
    If strPageTo <> "" Then
        nPageTo = int(strPageTo)
        If nPageTo <  1 Then 
			nPageTo = 1 
        End If
    Else
        nPageTo = 1
    End If
  else
  end if
	strSql="select * from PIC order by TITLE"
	RS.PageSize = iPageSize
	RS.CacheSize = iPageSize
	RS.Open strSql, my_Conn , 1, 1
	If not (RS.BOF and RS.EOF) Then
		RS.AbsolutePage = nPageTo
	End If 

	reccount = RS.recordcount
	iPageCount = RS.PageCount

    If nPageTo > iPageCount Then nPageTo = iPageCount
	If nPageTo < 1 Then nPageTo = 1
  
  
  Response.Write("<center><br><br><b>" & strTxt & "</b><br></center>")
  If RS.EOF Then  
    Response.Write("<center><br><br><span class=""fAlert""><b>No Pictures were found!</b></span><br></center>")
  Else
	nRowCount = 0
	showDaPaging nPageTo,iPageCount,0
	Do While not RS.EOF and nRowCount < iPageSize
		  cat_id = rs("PARENT_ID")
		  sub_id = rs("CATEGORY")
		  strURL = rs("URL")
		  strTURL = rs("TURL")
		  strPoster = rs("POSTER")
		  strpicTitle = rs("TITLE")
		  strSummary = rs("DESCRIPTION")
		  intpicID = rs("PIC_ID")
		  intHit = rs("HIT")
		  intShow = rs("ACTIVE")
		  strOwner = rs("OWNER")		
		  strPostDate = strtodate(left(rs("POST_DATE"),8) & "000000")
		  dateSince=DateDiff("d", strCurDateAdjust, strPostDate)+7
		  if len(trim(rs("TURL"))) = 7 or rs("TURL") = "" then
                strTURL = rs("URL")
 	  	  end if
		Call DisplayARTICLE()
		rs.MoveNext
		nRowCount = nRowCount + 1
	loop
	'Display Paging Buttons
	nPageCnt = iPageCount
	showDaPaging nPageTo,iPageCount,2
  end if
	rs.close
	set rs = nothing
end sub

sub showArticles()
  Response.Write "<center>" & strTxt & "</center><br>"
  iPageSize = 10
  nPageTo = 1

  If Request("PageTo") <> "" Then
	strPageTo = Request("PageTo")
    If strPageTo <> "" Then
        nPageTo = int(strPageTo)
        If nPageTo < 1 Then 
			nPageTo = 1 
        End If
    Else
        nPageTo = 1
    End If
	'Set RS = Session("ListingRS")
  else
  end if
	Set RS=Server.CreateObject("ADODB.Recordset")
	strSql="SELECT * From PIC where CATEGORY=" & sid & " order by PIC_ID DESC"
	RS.Open strSql, my_Conn, 1, 1
	RS.PageSize = iPageSize
	RS.CacheSize = iPageSize
	'Set Session("ListingRS") = RS
	If not (RS.BOF and RS.EOF) Then
		RS.AbsolutePage = nPageTo
	End If 

	reccount = RS.recordcount
	iPageCount = RS.PageCount

    If nPageTo > iPageCount Then nPageTo = iPageCount
	If nPageTo < 1 Then nPageTo = 1
  
  If RS.EOF Then
	Response.Write(" <center><br><br><span class=""fAlert""><b>No Picture found for you to edit</b></span><br></center>")
  Else
	nRowCount = 0
	showDaPaging nPageTo,iPageCount,0
	Do While not RS.EOF and nRowCount < iPageSize
		%>
		<center>
          <table border="0" width="500" cellspacing="1" cellpadding="2" class="tCellAlt2">
          <tr>
            <td rowspan="2" align="center" valign="middle" width="150"><%
			if instr(rs("TURL"),"_sm") > 0 then
			  stImg = "<img src=""" & rs("TURL") & """ border=""0"" alt=""Image"" title=""Image"" />"
			else
			  stImg = "<img src=""" & rs("URL") & """ border=""0"" width=""120"" alt=""image"" title=""Image"" />"
			end if 
  			response.Write(stImg)%>	
		    </td>
          <td class="tSubTitle" height="20" width="300" valign="top">
	  <a href="<%=iLink & RS("pic_ID")%>"><%= ChkString(rs("Title"), "display") %></a><br><span class="fAlert"><% if rs("ACTIVE") = 0 then%> [not approved]<%else%> [approved]<%end if%><% if rs("OWNER") = "0" then%> [public]<%else%> [private]<%end if%></span>
	 </td><td height="20" class="tTitle" width="100" valign="top"><span class="fAlert">Hit : <%=RS("Hit")%> </span></td>
	 </tr>
	 <tr>
	 <td class="tCellAlt1" colspan="2" valign="top">
	 <%= ChkString(rs("Description"), "display") %></td>
	</tr>
	</table>
	</center><br>
    <%  nRowCount = nRowCount + 1
		RS.MoveNext
	loop
	nPageCnt = iPageCount
	showDaPaging nPageTo,iPageCount,2
	end if 'end eof check
	rs.close
	set rs = nothing
end sub

sub showBadLinks()
  Response.Write "<center>" & strTxt & "</center><br>"
  iPageSize = 10
  nPageTo = 1

  If Request("PageTo") <> "" Then
	strPageTo = Request("PageTo")
    If strPageTo <> "" Then
        nPageTo = int(strPageTo)
        If nPageTo < 1 Then 
			nPageTo = 1 
        End If
    Else
        nPageTo = 1
    End If
	'Set RS = Session("ListingRS")
  else
  end if
	Set RS=Server.CreateObject("ADODB.Recordset")
	strSql="SELECT * From PIC where BADLINK <> 0 order by PIC_ID DESC"
	RS.Open strSql, my_Conn, 1, 1
	RS.PageSize = iPageSize
	RS.CacheSize = iPageSize
	'Set Session("ListingRS") = RS
	If not (RS.BOF and RS.EOF) Then
		RS.AbsolutePage = nPageTo
	End If 

	reccount = RS.recordcount
	iPageCount = RS.PageCount

    If nPageTo > iPageCount Then nPageTo = iPageCount
	If nPageTo < 1 Then nPageTo = 1
  
  If RS.EOF Then
	Response.Write(" <center><br><br><span class=""fAlert""><b>No Picture found for you to edit</b></span><br></center>")
  Else
	nRowCount = 0
	showDaPaging nPageTo,iPageCount,0
	Do While not RS.EOF and nRowCount < iPageSize
		%>
		<center>
        <table border="0" width="500" cellspacing="1" cellpadding="2" class="tCellAlt0">
          <tr>
            <td rowspan="2" align="center" valign="middle" width="150"><%
			if instr(rs("TURL"),"_sm") > 0 then
			  stImg = "<img src=""" & rs("TURL") & """ border=""0"" alt=""Image"" title=""Image"" />"
			else
			  stImg = "<img src=""" & rs("URL") & """ border=""0"" width=""120"" alt=""image"" title=""Image"" />"
			end if 
  			response.Write(stImg)%>	
		    </td>
          <td class="tCellAlt0" height="20" width="300" valign="top">
	  <a href="<%=iLink & RS("pic_ID")%>" title="Click to edit"><%= ChkString(rs("Title"), "display") %></a><br><span class="fAlert"><% if rs("ACTIVE") = 0 then%> [not approved]<%else%> [approved]<%end if%><% if rs("OWNER") = "0" then%> [public]<%else%> [private]<%end if%></span>
	 </td><td height="20" class="tCellAlt0" width="100" valign="top"><span class="fAlert">Hit : <%=RS("Hit")%> </span></td>
	 </tr>
	 <tr>
	 <td class="tCellAlt0" colspan="2" valign="top">
	 <%= ChkString(rs("Description"), "display") %></td>
	</tr>
	 <tr>
	 <td class="tCellAlt2" colspan="3" valign="middle" align="center">
		<a href="admin_pic_editpic.asp?ID=<%=rs("pic_id")%>">Edit this picture</a> |
		<a href="admin_pic_admin.asp?cmd=40&mode=7&cid=<%=rs("pic_id")%>">Delete this picture</a> |
		<a href="admin_pic_admin.asp?cmd=40&mode=11&cid=<%=rs("pic_id")%>">Unmark bad picture</a>
	 </td>
	</tr>
	</table>
	</center><br>
    <%  nRowCount = nRowCount + 1
		RS.MoveNext
	loop
	nPageCnt = iPageCount
	showDaPaging nPageTo,iPageCount,2
	end if 'end eof check
	rs.close
	set rs = nothing
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
							Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold}"" disabled=""disabled"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2"" /><input type=""hidden"" name=""PageTo"" value=""1"" />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value="" &lt;&lt; First "" style=""{font-weight:bold;cursor:pointer;}"" id=""submit"&nPaging&"2"" name=""submit"&nPaging&"2""><input type=""hidden"" name=""PageTo"" value=""1"" />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sid & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display <
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" name=""formP"&nPaging&"02"" id=""formP"&nPaging&"02"">")
						If int(nPageTo) = 1 Then 
							Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""PageTo"" value=""1"" />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value=""&lt; Previous "" id=""submit"&nPaging&"3"" name=""submit"&nPaging&"3"" style=""{font-weight:bold;cursor:pointer;}"" />")
							Response.Write(vbCrLf & "<input type=""hidden"" name=""PageTo"" value=""" & nPageTo-1 & """ />")
						End If
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sid & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display >
					      strQryStr = ""
						  if sMode <> "" then
						    strQryStr = strQryStr & "&amp;mode=" & sMode
						    strMode = "&amp;mode=" & sMode
						  end if
						  if search <> "" then
						    strQryStr = strQryStr & "&amp;search=" & search
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
							  Response.Write("&nbsp;<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?cmd=" & iPgType & "&amp;sid=" & sid & "&amp;PageTo=" & pgc & strQryStr & """>")
						      Response.Write("<span class=""fBold"">" & pgc & "</span></a>")
							end if
						  next
						  Response.Write("&nbsp;</td>")
						end if
						
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" id=""formP"&nPaging&"03"" name=""formP"&nPaging&"03"">")
						If int(nPageTo) = nPageCnt Then 
							Response.Write(vbCrLf & "<input type=""submit"" value='  Next &gt;  ' id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""PageTo"" value=""" & nPageTo & """ />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value=""  Next &gt;  "" id=""submit"&nPaging&"4"" name=""submit"&nPaging&"4"" style=""{font-weight:bold;cursor:pointer;}"" />")
							Response.Write(vbCrLf & "<input type=""hidden"" name=""PageTo"" value=""" & nPageTo+1 & """ />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sid & """ />")
						Response.Write(vbCrLf & "</form>")
						Response.Write(vbCrLf & "</td>")
					' Display >>
						Response.Write(vbCrLf & "<td align=""center"">")
						Response.Write(vbCrLf & "<form action=""" & Request.ServerVariables("SCRIPT_NAME") & """ method=""post"" id=""formP"&nPaging&"04"" name=""formP"&nPaging&"04"">")
						If int(nPageTo) = nPageCnt Then 
							Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold}"" disabled=""disabled"" /><input type=""hidden"" name=""PageTo"" value=""" & nPageTo & """ />")
						Else
							Response.Write(vbCrLf & "<input type=""submit"" value="" Last &gt;&gt; "" id=""submit"&nPaging&"5"" name=""submit"&nPaging&"5"" style=""{font-weight:bold;cursor:pointer;}"" /><input type=""hidden"" name=""PageTo"" value=""" & nPageCnt & """ />")
						End IF
						Response.Write(vbCrLf & "<input type=""hidden"" name=""cmd"" value=""" & iPgType & """ />")
						Response.Write(vbCrLf & "<input type=""hidden"" name=""sid"" value=""" & sid & """ />")
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

sub showSubs() 'list of subcategories for selection 
	If iCmd = 6 Then
	  sSQL = "SELECT count(CAT_ID) FROM PIC_SUBCATEGORIES where CAT_ID=" & cid
	  Set RScount = my_Conn.Execute(sSQL)
	    scount = RScount(0)
	  set RScount = nothing
	end if
  sql = "SELECT CAT_NAME From PIC_CATEGORIES where CAT_ID=" & cid
  set rsC = my_Conn.execute(sql)
    catname = rsC(0)
  set rsC = nothing %>
	Category: <b><%=catname%></b><br><br>
	<%= strTxt %><br><br>
 <%
	sql = "SELECT * From PIC_SUBCATEGORIES where CAT_ID=" & cid & " order by C_ORDER, SUBCAT_ID"
	set rs = my_Conn.Execute (sql)
	if rs.eof then
	  Response.Write "No subcategories found in this category"
	else %>
	  <table border=1 width=500 cellspacing=0 cellpadding="2" class="grid" style="border-collapse: collapse;">
	  <tr class="tSubTitle">
		<% If iCmd = 6 Then %>
		<td width="100" align="center">
		<b>Order</b></td>
		<% else
			  colspn = " colspan=2" %>
		<% End If %>
		<td<%= colspn %>>
		&nbsp;&nbsp;<b>Category Names</b></td>
	  </tr>
	  <% 
	  Do while NOT rs.EOF 
	    response.Write("<tr>")
	    If iCmd = 6 Then 
	  	  ColNum = 2 
	    else
	  	  ColNum = 1 
	    end if
	 	Do while ColNum < 3
	   	  if NOT rs.EOF then 
		  
		 	If iCmd = 6 Then %>
	    	<form action="admin_pic_admin.asp" method="post" name="dorder<%= rs("SUBCAT_ID") %>" id="dorder<%= rs("SUBCAT_ID") %>">
			<td width="100" align="center">
        	<input name="mode" type="hidden" id="mode" value="10">
  			<input name="sid" type="hidden" id="sid" value="<%= rs("SUBCAT_ID") %>">
  			<input name="cid" type="hidden" id="cid" value="<%= cid %>">
  			<input name="cmd" type="hidden" id="cmd" value="6">
	    	<select name="ord" onChange="submit()"><% for xc = 1 to scount %>
    	  		<option value="<%= xc %>"<% If rs("C_ORDER") = xc Then response.Write(" selected") %>><%= xc %></option><% next %>
			</select>
			</td></form><% 
			End If 
			
		    rcount = 0
		    sSQL = "SELECT count(PIC_ID) FROM PIC where category=" & rs("SUBCAT_ID")
			Set RScount = my_Conn.Execute(sSQL)
			  rcount = RScount(0)
			set RScount = nothing %>
		    <td width="400">
			
			<a href="<%= iLink & rs("SUBCAT_ID")%>">
			<%= ChkString(rs("SUBCAT_NAME"), "display") %> <small>(<%= rcount %>)</small></a><br></td><% 
	  	  else %>
			<td width="50%">&nbsp;</td><%
	  	  end if
	  	  if ColNum = 1 then
 	  	  rs.MoveNext
	  	  end if
 	  	  ColNum = ColNum + 1
		Loop %>
		</TR>
		<% 
		if NOT rs.EOF then  
		rs.MoveNext  
		end if 
 	  Loop %>
	  </TABLE><%
	end if 
  	rs.close
  	set rs = nothing
end sub

sub showCats()
	If iCmd = 3 Then
	  sSQL = "SELECT count(CAT_ID) FROM PIC_CATEGORIES"
	  Set RScount = my_Conn.Execute(sSQL)
	    rcount = RScount(0)
	  set RScount = nothing
	end if
	sql = "select * from PIC_CATEGORIES order by C_ORDER, CAT_NAME"
	set rs = my_Conn.Execute (sql)
	%>
	<p><b><%= strTxt %></b></p><br>
	<table border=1 width=500 cellspacing=0 cellpadding="2" class="grid" style="border-collapse: collapse;">
	  <tr class="spThemeBlock_subTitleCell">
		<% If iCmd = 3 Then %>
		<td width="100" align="center">
		<b>Order</b></td>
		<% else
			  colspn = " colspan=2" %>
		<% End If %>
		<td<%= colspn %>>
		<span class="fTitle">&nbsp;&nbsp;Category Names</span></td>
	  </tr>
	<% 
	response.Write("<tr>")
	Do while NOT rs.EOF 
	If iCmd = 3 Then 
	  ColNum = 2 
	else
	  ColNum = 1 
	end if
	Do while ColNum < 3
	  if NOT rs.EOF then  
	  %>
		<% If iCmd = 3 Then %>
	    <form action="admin_pic_admin.asp" method="post" name="dorder<%= rs("CAT_ID") %>" id="dorder<%= rs("CAT_ID") %>">
		<td width="100" align="center">
        <input name="mode" type="hidden" id="mode" value="9">
  		<input name="cid" type="hidden" id="cid" value="<%= rs("CAT_ID") %>">
  		<input name="cmd" type="hidden" id="cmd" value="2">
	    <select name="sid" onChange="submit()">
		<% for xc = 1 to rcount %>
    	  	<option value="<%= xc %>"<% If rs("C_ORDER") = xc Then response.Write(" selected") %>><%= xc %></option>
		<% next %>
		</select>
		</td></form>
		<% End If %>
		<td>
		
		<a href="admin_pic_admin.asp?cmd=<%= iCmd %>&cid=<%=rs("CAT_ID")%>&wid=<%= webid %>">&nbsp;<%= ChkString(rs("CAT_NAME"), "display") %> </a></td>
	  	<% 
	  else %>
		<td width="50%">&nbsp;</td>
		<%
	  end if
	  if ColNum = 1 then
 	  rs.MoveNext
	  end if
 	  ColNum = ColNum + 1
	Loop 
	  %>
	</tr>
	  <% 
	  if NOT rs.EOF then  
	  rs.MoveNext  
	  end if 
	Loop 
	%>
	</table>
	<% 
	'end if
	rs.close 
	set rs = nothing
end sub

sub delPickSub()
	strSql="Select CAT_NAME from PIC_CATEGORIES where CAT_ID=" & cid
	set rsC = my_Conn.execute(strSql)
	  CAT_NAME = rsC(0)
	set rsC = nothing
		
	sql = "select * from PIC_SUBCATEGORIES where CAT_ID=" & cid
	set rs = my_Conn.Execute (sql)
	if rs.eof then
	Response.Write "No subcat found in this category"
	else
%>
 <b>Click button to delete subcategory from "<%= CAT_NAME %>"</b><br>
 This will delete all pictures in this subcategory.<br><br><b>Remember, this cannot be undone!</b><br><br>
<table border=1  width=500 cellspacing=1 cellpadding="2" class="grid" style="border-collapse: collapse;">
	<% 
	Do while NOT rs.EOF 
	ColNum = 1 
	 Do while ColNum < 3
	%>
<tr>
<td ALIGN=center VALIGN=middle >	

<%= ChkString(rs("SUBCAT_NAME"), "display")%><br>
<form action="admin_pic_admin.asp" method="post" id=form2 name=form2>
<input type="submit" value="Delete Subcat" id=submit2 name=submit2 class="button">
<input type="hidden" value="<%= CAT_NAME %>" name="cat">
<input type="hidden" value="<%= rs("SUBCAT_NAME") %>" name="scat">
<input type="hidden" value="<%= rs("SUBCAT_ID") %>" name="sid">
<input type="hidden" value="8" name="cmd">
<input type="hidden" value="6" name="mode">
</form>
</TD>
<% 
if NOT rs.EOF then 
 rs.MoveNext 
end if 
 ColNum = ColNum + 1 
if NOT rs.EOF then 
%>
<td ALIGN=center VALIGN=middle >	

<%= ChkString(rs("SUBCAT_NAME"), "display")%><br>
<form action="admin_pic_admin.asp" method="post"  id=form2 name=form2>
<input type="submit" value="Delete subcat"  id=submit2 name=submit2 class="button">
<input type="hidden" value="<%= CAT_NAME %>" name="cat">
<input type="hidden" value="<%= rs("SUBCAT_NAME") %>" name="scat">
<input type="hidden" value="<%= rs("SUBCAT_ID") %>" name="sid">
<input type="hidden" value="8" name="cmd">
<input type="hidden" value="6" name="mode">
</form>
</TD>
<% 
end if
ColNum = ColNum + 1 
Loop 
%>
	</TR>
<% 
if NOT rs.EOF then  
rs.MoveNext  
end if 
 Loop 
 %>
	</TABLE>
	<% 
	end if
	rs.close 
	set rs = nothing
end sub

sub editSub() 'list of subcategories for selection %>
  <br><br><b>Select a new name for this subcategory.</b><br><%
  sql="SELECT SUBCAT_NAME FROM PIC_SUBCATEGORIES WHERE SUBCAT_ID=" & sid 
  set rs = my_Conn.Execute (sql) %>
  <form action="admin_pic_admin.asp" method="post">
  <input type="text" value="<%=chkString(rs("SUBCAT_NAME"), "display")%>" name="newcat" size=30>
  <input type="submit" value="update" class="button">
  <input type="hidden" value="<%= sid %>" name="sid">
  <input type="hidden" value="5" name="mode">
  <input type="hidden" value="5" name="cmd">
  </form><% 
  rs.close
  set rs = nothing
end sub

sub addCat() %>
  <p align="center"><span class="fTitle">Add new category..</span><br><br>
<b>Enter a category name and click the &quot;Create Now&quot; button<br>and the new category will be created.<br><br>
 Do not forget to create at least one subcategory for this new category.</b>
  <br><br>
  <form method="post" action="admin_pic_admin.asp">
  <input type="hidden" value="1" name="mode">
  <p>New Category name : <input type="text" name="newcat" size="30" class="textbox"><br><input type="submit" value="Create Now" name="B1" class="button"></p>
  </form>
<%
end sub

sub addSub()
  %>
  <p align="center"><span class="fTitle">Add new subcategory..</span>
  <p>&nbsp;</p>
  <form method="POST" action="admin_pic_admin.asp">
  <input type="hidden" value="2" name="mode">
  <input type="hidden" value="1" name="cmd">
  New SubCategory name : <input type="text" name="newsub" size="30">
  <br><br>put it in category : 
  <select name="cid">
    <option selected value="0">--select category--</option>
	<%
	set rscat = my_Conn.Execute("SELECT CAT_ID, CAT_NAME FROM PIC_CATEGORIES ORDER BY CAT_NAME") 
	do while not rscat.eof %>
	<option value="<%=rscat("CAT_ID")%>"><%=rscat("CAT_NAME") %></option>
	<%
	rscat.movenext
	loop
	rscat.close 
	set rscat = nothing %>
	</select>
  <br><br>
  <input type="submit" value="Create Now" name="B1" class="button">
  </form>
  <%
end sub

sub editCat()
  sql="SELECT * FROM PIC_CATEGORIES WHERE CAT_ID=" & cid 
  set rs = my_Conn.Execute (sql)
  %> 
  <form action="admin_pic_admin.asp" method="post">  
  <input type="hidden" value="<%=rs("CAT_ID")%>" name="cid">
  <input type="hidden" value="3" name="mode">
  <input type="hidden" value="3" name="cmd">
  <b>Change the name of this category:</b><br><br>
  <input type="text" value="<%=chkString(rs("CAT_NAME"), "edit")%>" name="newcat" size=30>	
  <input type="submit" value="update" class="button">
  </form>
  <%
end sub

sub delCat()
	sql = "SELECT CAT_ID, CAT_NAME FROM PIC_CATEGORIES ORDER BY CAT_NAME"
	set rs = my_Conn.Execute (sql) %>
      <span class="fAltSubTitle"><b>Click button to delete category</b></span><br>
	  This will delete all subcategories and associated pictures.<br><br>
	  <b>Remember, this cannot be undone!</b><br><br>
  <table border=1  width=500 cellspacing=1 cellpadding="2" class="grid" style="border-collapse: collapse;">
  <% 
	Do while NOT rs.EOF 
		ColNum = 1 
		Do while ColNum < 3 %>	
  <tr>
    <td ALIGN=center VALIGN=middle >			
      <%= ChkString(rs("CAT_NAME"), "display") %>
      <form action="admin_pic_admin.asp" method="post"  id=form1 name=form1>
      <input type="submit" value="Delete Now"  id=submit1 name=submit1 class="button">
      <input type="hidden" value="<%=rs("CAT_ID")%>" name="cid">
      <input type="hidden" value="<%=rs("CAT_NAME")%>" name="cat">
      <input type="hidden" value="4" name="cmd">
      <input type="hidden" value="4" name="mode">
      </form>
      
    </td>
			<% 
			if NOT rs.EOF then 
			  rs.MoveNext 
			end if 
			ColNum = ColNum + 1 
			if NOT rs.EOF then 
			%>
    <td ALIGN=center VALIGN=middle >
      <%= ChkString(rs("CAT_NAME"), "display") %>
      <form action="admin_pic_admin.asp" method="post"  id=form1 name=form1>
      <input type="submit" value="Delete Now"  id=submit1 name=submit1 class="button">
      <input type="hidden" value="<%=rs("CAT_ID")%>" name="cid">
      <input type="hidden" value="<%=rs("CAT_NAME")%>" name="cat">
      <input type="hidden" value="4" name="cmd">
      <input type="hidden" value="4" name="mode">
      </form>
      
    </td></tr>
			<%
			else
			%>
    <td ALIGN=center VALIGN=middle>&nbsp;</td></tr>
			<%
			end if
			ColNum = ColNum + 1 
		Loop 
		if NOT rs.EOF then  
		rs.MoveNext  
		end if 
	Loop 
	%>
  </table>
  <%
end sub
%>