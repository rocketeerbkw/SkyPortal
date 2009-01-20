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

if strDBType = "sqlserver" then
  sMPPre = ""
  sMCPre = ""
  sMSPre = ""
else
  sMPPre = strTablePrefix & "M_PARENT."
  sMCPre = strTablePrefix & "M_CATEGORIES."
  sMSPre = strTablePrefix & "M_SUBCATEGORIES."
end if

if item_scat_fld = "" then
  item_scat_fld = "category"
end if
'::::::::::::::::::::::::::::::
'::
'::   COMMON MODULE FUNCTIONS
'::
':::::::::::::::::::::::::::::::

sub mod_increaseSubcatCount(s)
  sSql = "UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET"
  sSql = sSql & " " & strTablePrefix & "M_SUBCATEGORIES.ITEM_CNT=" & strTablePrefix & "M_SUBCATEGORIES.ITEM_CNT+1"
  sSql = sSql & " WHERE SUBCAT_ID=" & s
  executeThis(sSql)
  'Response.Write sSql & "<br>"
end sub

sub mod_decreaseSubcatCount(s)
  sSql = "UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET"
  sSql = sSql & " " & strTablePrefix & "M_SUBCATEGORIES.ITEM_CNT=" & strTablePrefix & "M_SUBCATEGORIES.ITEM_CNT-1"
  sSql = sSql & " WHERE SUBCAT_ID=" & s
  executeThis(sSql)
  'Response.Write sSql & "<br>"
end sub

sub mod_decreaseMultiSubcatCount(s,c)
  sSql = "UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET"
  sSql = sSql & " ITEM_CNT=ITEM_CNT - " & c
  sSql = sSql & " WHERE SUBCAT_ID=" & s
  executeThis(sSql)
end sub

sub mod_updateCatCounts() %>
  <div style="width:250px;">
  <fieldset style="margin:5px;padding:10px;">
  <p align="center"><span class="fTitle">Update Category counts</span><br><br>
  <form method="post" action="<%= sScript %>">
  <input type="hidden" value="<%= iPgType %>" name="cmd">
  <input type="hidden" value="23" name="mode">
  <input type="submit" value="Update Counts Now!" class="button">
  </form><br></p></fieldset></div>
  <%
end sub

function mod_getFileInfo(u,a)
  Dim tf, tr
  tr = ""
  if bFso then
    set oFs = new clsSFSO
    set tf = oFs.GetFileInformation(u)
    set oFs = nothing
    select case a
      case "Size"
	    tr = mod_formatSize(tf.Size)
      case "FileType"
	    tr = tf.FileType
    end select
  end if
  mod_getFileInfo = tr
end function

function mod_formatSize(s)
  dim fSize
  if s > 1024 then
    fSize = round(s/1024)
	if fSize > 1024 then
      fSize = round(fSize/1024)
	  if fSize > 1024 then
	    fSize = round(fSize/1024)
	    fSize = fSize & " gb"
	  else
	    fSize = fSize & " mb"
	  end if
	else
	  fSize = fSize & " kb"
	end if
  else
    fSize = round(s) & " bytes"
  end if
  mod_formatSize = fSize
end function

sub mod_selectCatSubcat(ii,mp)
  ':: ii = subcat_ID
  sgp = "SG_" & mp
  sSQL = "SELECT " & strTablePrefix & "M_CATEGORIES.*, " & strTablePrefix & "M_SUBCATEGORIES.*"
  sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
  sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_SUBCATEGORIES.APP_ID)=" & intAppID & "))"
  sSQL = sSQL & " ORDER BY " & strTablePrefix & "M_CATEGORIES.CAT_NAME, " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_NAME;"

	dim rsC
	set rsC = my_Conn.execute(sSql)
	
    'Response.Write "<td>"
    Response.Write "<select id=""subcat"" name=""subcat"">"
	curCat = ""
	if ii = 0 then
	  Response.Write "<option value=""0""" & chkSelect(ii,0) & ">"
	  Response.Write "[Select One]"
	  Response.Write "</option>"
	end if
	do while not rsC.EOF
	 if hasAccess(rsC("CG_READ")) or bAppFull then
	  if curCat <> rsC(sMCPre & "CAT_ID") then
	    curCat = rsC(sMCPre & "CAT_ID")
	    Response.Write "<optgroup label=""" & rsC("CAT_NAME") & """>"
	  end if
	  if hasAccess(rsC("CG_FULL")) or hasAccess(rsC(sgp)) or bAppFull then
	  Response.Write "<option value="""&rsC("SUBCAT_ID")&""""
	  Response.Write chkSelect(ii,rsC("SUBCAT_ID")) & ">"
	  'Response.Write rsC("CAT_NAME")&" / "
	  Response.Write "- " & rsC("SUBCAT_NAME")
	  Response.Write "</option>"
	  end if
	  rsC.MoveNext
	  if rsC.eof then
	    Response.Write "</optgroup>"
		Response.Write "<optgroup title=""Spacer""></optgroup>"
	  else
	   if curCat <> rsC(sMCPre & "CAT_ID") then
	    Response.Write "</optgroup>"
		Response.Write "<optgroup title=""Spacer""></optgroup>"
	   end if
	  end if
	 else
	  rsC.MoveNext
	 end if
	loop
    Response.Write "</select>"
    'Response.Write "</td>"
	set rsC = nothing
end sub

function mod_selectCats(ii,p)
  ':: ii = cat_ID
  sSQL = "SELECT " & strTablePrefix & "M_CATEGORIES.*"
  sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES"
  sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_CATEGORIES.APP_ID)=" & intAppID & "))"
  sSQL = sSQL & " ORDER BY " & strTablePrefix & "M_CATEGORIES.CAT_NAME;"

	dim rsC
	set rsC = my_Conn.execute(sSql)
	
    Response.Write "<select name=""cat"">"
	curCat = ""
	do while not rsC.EOF
	  if hasAccess(rsC("CG_" & p & "")) or bAppFull or ii = rsC("CAT_ID") then
	    Response.Write "<option value="""&rsC("CAT_ID")&""""
	    Response.Write chkSelect(ii,rsC("CAT_ID")) & ">"
	    Response.Write "- " & rsC("CAT_NAME")
	    Response.Write "</option>"
	  end if
	  rsC.MoveNext
	loop
    Response.Write "</select>"
	set rsC = nothing
end function

function mod_CatSubCatsql(c,s,a)
  tSql = "SELECT " & strTablePrefix & "M_CATEGORIES.*"
  tSql = tSql & ", " & strTablePrefix & "M_SUBCATEGORIES.*"
  tSql = tSql & " FROM " & strTablePrefix & "M_CATEGORIES"
  tSql = tSql & " INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES"
  tSql = tSql & " ON " & strTablePrefix & "M_CATEGORIES.CAT_ID"
  tSql = tSql & " = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
  tSql = tSql & " WHERE "
  if s > 0 then
    tSql = tSql & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID = " & s
    tSql = tSql & " AND "
  end if
  if c > 0 then
    tSql = tSql & strTablePrefix & "M_CATEGORIES.CAT_ID = " & c
    tSql = tSql & " AND "
  end if
  
  tSql = tSql & strTablePrefix & "M_CATEGORIES.APP_ID = " & a
  
  tSql = tSql & " ORDER BY "
  tSql = tSql & strTablePrefix & "M_CATEGORIES.C_ORDER"
  tSql = tSql & ", " & strTablePrefix & "M_CATEGORIES.CAT_NAME"
  tSql = tSql & "," & strTablePrefix & "M_SUBCATEGORIES.C_ORDER"
  tSql = tSql & "," & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_NAME"
  tSql = tSql & ";"
  mod_CatSubCatsql = tSql
end function

function mod_singleItemSql(tbl)
  tSql = "SELECT " & tbl & ".*"
  tSql = tSql & ", " & strTablePrefix & "M_SUBCATEGORIES.*"
  tSql = tSql & ", " & strTablePrefix & "M_CATEGORIES.*"
  tSql = tSql & " FROM " & tbl
  tSql = tSql & " INNER JOIN (" & strTablePrefix & "M_SUBCATEGORIES"
  tSql = tSql & " INNER JOIN " & strTablePrefix & "M_CATEGORIES ON"
  tSql = tSql & " " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID ="
  tSql = tSql & " " & strTablePrefix & "M_CATEGORIES.CAT_ID)"
  tSql = tSql & " ON " & tbl & "." & item_scat_fld & " ="
  tSql = tSql & " " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID "
  mod_singleItemSql = tSql
end function

sub mod_displayIntro(i)
  'if request.QueryString() <> "" then
  'else
	strSql = "SELECT * FROM " & strTablePrefix & "WELCOME WHERE W_MODULE = " & i & " AND W_ACTIVE=1"
	set rsIntro =  my_Conn.Execute (strSql)
    if not rsIntro.EOF then
	  W_ID = rsIntro("W_ID")
	  W_TITLE = trim(replace(rsIntro("W_TITLE"),"''","'"))
	  W_SUBJECT = trim(replace(rsIntro("W_SUBJECT"),"''","'"))
	  W_SUBJECT	= replace(W_SUBJECT,"[%member%]",strdbntusername)
	  W_MESSAGE = trim(replace(rsIntro("W_MESSAGE"),"''","'"))
	  W_MESSAGE	= replace(W_MESSAGE,"</p><p>","<br /><br />")
	  W_MESSAGE	= replace(W_MESSAGE,"<p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"</p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"[%member%]",strdbntusername)
	  W_MESSAGE = FormatStr2(W_MESSAGE)

	  spThemeMM = "m_intro_" & i
	  'spThemeTitle = txtWelcomeTo & " " & strSiteTitle
	  if bAppFull then
	  	  spThemeTitle = icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","mwpHSs('edIntro" & W_ID & "','1');mwpHSs('shoIntro" & W_ID & "','1');","align=""right""")
	  	  'spThemeTitle = icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","openJsLayer('edIntro"& W_ID &"','350','250');","align=""right""")
	  end if
	  
	  spThemeTitle = spThemeTitle & W_SUBJECT
	  spThemeBlock1_open(intSkin)
	  Response.Write "<div style=""display:block;"" id=""shoIntro"& W_ID &""">"
	  Response.Write "<p style=""text-align:left;"">"
	  Response.Write W_MESSAGE
	  Response.Write "</p></div>"
	  if bAppFull then
	    mod_introEditForm(rsIntro)
	  end if
	  spThemeBlock1_close(intSkin)
	End if 
    set rsIntro = nothing
  'end if
end sub

sub mod_introEditForm(ob)
  if request.Form("update") = "update" then
    a_id = cLng(request.Form("W_ID"))
    'a_title = chkString(request.Form("W_TITLE"),"message")
    a_subject = chkString(request.Form("W_SUBJECT"),"sqlstring")
    'a_active = cint(request.Form("W_ACTIVE"))
    a_message = chkString(request.Form("Message"),"message")
    'a_message = request.Form("W_MESSAGE")
	a_message = replace(a_message,"</p><p>","<br /><br />")
	a_message = replace(a_message,"<p>","")
	a_message = replace(a_message,"</p>","")
	'response.Write("a_message: " & a_message & "<br />")
	sSql = "UPDATE " & strTablePrefix & "WELCOME SET "
	sSql = sSql & "W_SUBJECT='" & a_subject & "'"
	sSql = sSql & ",W_MESSAGE='" & a_message & "'"
	'sSql = sSql & ",W_TITLE='" & a_title & "'"
	'sSql = sSql & ",W_ACTIVE=" & a_active & ""
	sSql = sSql & " WHERE W_ID=" & a_id
	'response.Write(sSql & "<br />")
	executeThis(sSql)
	closeAndGo(sScript)
  end if
  Response.Write "<div style=""display:none;"" id=""edIntro"& ob("W_ID") &""">"
  Response.Write "<form name=""introForm"" id=""introForm"" method=""post"" action="""">"
  Response.Write "<table border=""0"" cellpadding=""3"" cellspacing=""0"" width=""100%"">"
  Response.Write "<tr><td colspan=""2"" class=""tSubTitle"">"
  Response.Write txtEdModIntro
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">"
  Response.Write "&nbsp;</td><td class=""fSubTitle"">"
  Response.Write ob("W_TITLE")
  Response.Write "</td></tr>"
  'Response.Write "<tr><td align=""right"">"
  'Response.Write "<b>" & txtActive & "&nbsp;</b></td><td>"
  'Response.Write "<select name=""W_ACTIVE"" id=""W_ACTIVE"">"
  'Response.Write "<option value=""1"""& chkSelect(ob("W_ACTIVE"),1) &">"
  'Response.Write txtYes
  'Response.Write "</option>"
  'Response.Write "<option value=""0"""& chkSelect(ob("W_ACTIVE"),0) &">"
  'Response.Write txtNo
  'Response.Write "</option>"
  'Response.Write "</select>"
  'Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>" & txtTitle & "&nbsp;</b>"
  Response.Write "</td><td>"
  Response.Write "<input type=""text"" class=""textbox"" size=""40"" maxlength=""200"" name=""W_SUBJECT"" id=""W_SUBJECT"" value="""& ob("W_SUBJECT") &""" />"
  Response.Write "</td></tr>"

  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b>" & txtMsg & "</b>", "" & ob("W_MESSAGE") & ""
  else
  	displayPLAINeditor 1,Trim(CleanCode(ob("W_MESSAGE")))
  end if
  
  Response.Write "<tr><td align=""right"">"
  Response.Write "&nbsp;</td><td>"
  Response.Write "<input type=""submit"" class=""button"" name=""submit"" id=""submit"" value="" "& txtSubmit &" "" />"
  Response.Write "<input type=""hidden"" name=""update"" id=""update"" value=""update"" />"
  Response.Write "<input type=""hidden"" name=""W_ID"" id=""W_ID"" value="""& ob("W_ID") &""" />"
  Response.Write "</td></tr>"
  Response.Write "</table>"
  Response.Write "</form>"
  Response.Write "</div>" 
end sub

function mod_GetComments(itm,app,rURL)
  sSql = "SELECT COMMENTS, RATE_BY, RATE_DATE, RATING_ID, RATING FROM "
  sSql = sSql & strTablePrefix & "M_RATING"
  sSql = sSql & " WHERE COMMENTS NOT LIKE ' '"
  sSql = sSql & " AND ITEM_ID = " & itm & ""
  sSql = sSql & " AND APP_ID=" & app & ""
  set rsC = my_Conn.execute(sSql)
  if rsC.eof then
	'show nothing
  else
	spThemeTitle= txtComments & ":&nbsp;"
	spThemeBlock3_open()
	do while not rsC.eof
	  Response.Write "<hr align=""center""/>"
	  Response.Write "<table border=""0"" width=""100%"" align=""center"" cellspacing=""0"" cellpadding=""0"">"
	  Response.Write "<tr><td style=""padding-left:20px;"" class=""fNorm"">"
	  If bSCatFull Then
	    Response.Write "<a href=""" & rURL & "&amp;item=" & itm & "&amp;by=" & rsC("RATE_BY") & "&amp;com=" & rsC("RATING_ID") & """>"
	    Response.Write icon(icnDelete,txtDel,"","","hspace=""4""")
	    Response.Write "</a>"
	  Else
	    Response.Write "&nbsp;"
	  End If
	    Response.Write txtBy & ":<b> " & getMemberName(rsC("RATE_BY")) & "</b>"
	    Response.Write "&nbsp;" & txtOn & "&nbsp;" & ChkDate2(rsC("RATE_DATE")) & "<br/>"
	    Response.Write formatstr(rsC("Comments"))
    	Response.Write "</td></tr></table>"
		rsC.MoveNext
	loop
	response.Write("<hr align=""center"" />")
	spThemeBlock3_close()
  end if
  set rsC = nothing
end function

sub mod_deleteComment(tbl,fld)
  strMsg = ""
  bOk = false
  itm = clng(request.QueryString("item"))
  by = clng(request.QueryString("by"))
  com = clng(request.QueryString("com"))
  sSql = "SELECT " & strTablePrefix & "M_CATEGORIES.CAT_ID"
  sSql = sSql & ", " & strTablePrefix & "M_CATEGORIES.CG_FULL"
  sSql = sSql & ", " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID"
  sSql = sSql & ", " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL"
  sSql = sSql & ", " & strTablePrefix & "M_RATING.RATING"
  sSql = sSql & ", " & strTablePrefix & "M_RATING.COMMENTS"
  sSql = sSql & ", " & tbl & ".VOTES"
  sSql = sSql & ", " & tbl & ".RATING"
  sSql = sSql & " FROM " & strTablePrefix & "M_RATING INNER JOIN (" & tbl & " INNER JOIN (" & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID) ON " & tbl & ".CATEGORY = " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID) ON " & strTablePrefix & "M_RATING.ITEM_ID = " & tbl & "." & fld & ""
  sSql = sSql & " WHERE ("
  sSql = sSql & "((" & strTablePrefix & "M_RATING.RATING_ID)=" & com & ")"
  sSql = sSql & " AND ((" & strTablePrefix & "M_RATING.ITEM_ID)=" & itm & ")"
  sSql = sSql & " AND ((" & strTablePrefix & "M_RATING.RATE_BY)=" & by & ")"
  sSql = sSql & " AND ((" & strTablePrefix & "M_RATING.APP_ID)=" & intAppID & ")"
  sSql = sSql & ");"
  
  set rsRate = my_Conn.execute(sSql)
  if not rsRate.eof then
	if bAppFull or hasAccess(rsRate("CG_FULL")) or hasAccess(rsRate("SG_FULL")) then
	  bOk = true
	  strRate = rsRate("" & strTablePrefix & "M_RATING.RATING")
	  strComm = rsRate("COMMENTS")
	  totalVotes = rsRate("VOTES")
	  totalRate = rsRate("" & tbl & ".RATING")
	  if strRate <> "" then
	    intRating = totalRate - strRate
	  else
	    intRating = totalRate
	  end if
	  intVotes = totalVotes - 1
	  if intVotes < 0 then intVotes = 0
	  if intRating < 0 then intRating = 0
	else
      strMsg = strMsg & "<li>" & txtNoAccPerformTask & "</li>"
	end if
  else
    strMsg = strMsg & "<li>" & txtCommNoFnd & "</li>"
  end if
  set rsRate = nothing
  
  if bOk then
    sSql = "DELETE FROM " & strTablePrefix & "M_RATING"
  	sSql = sSql & " WHERE RATING_ID = "& com &" AND APP_ID = "& intAppID
	executeThis(sSql)
	  
	sSql = "UPDATE " & tbl
  	sSql = sSql & " SET VOTES = " & intVotes
  	sSql = sSql & ", RATING = " & intRating
  	sSql = sSql & " Where " & fld & " = " & itm
	executeThis(sSql)
    strMsg = strMsg & "<li>" & txtCommDeleted & "</li>"
  end if
  
  Call setSession("sMsg",strMsg)
  'closeAndGo(rUrl)
end sub

sub mod_rateItem(itm,app,tbl,fld,tFld,bc,br)
  dim intFormRating, strComments, sMsg
  dim bCanRate, bCanComment, sError
  dim intVotes, intRating
  intFormRating = -1
  strComments = ""
  sMsg = ""
  sError = ""
  bCanComment = bc
  bCanRate = br
  intVotes = 0
  intRating = 0
  intRatingID = 0
  response.Write("<br />")
  'Check to see if they are a member.
  if strUserMemberID = 0 then
	sMsg = "<li>" & txtCommMembOnly & "</li>"
	sMsg = sMsg & "<li>" & txtLoginRate & "</li>"
   	call showMsgBlock(1,sMsg)
  else
	
    if request("method") = "add" then
	  sError = ""
	' Check to see if they rated or commented already
	sSQL = "SELECT RATING,COMMENTS,RATING_ID FROM " & strTablePrefix & "M_RATING WHERE ITEM_ID = " & itm & " AND APP_ID=" & intAppID & " AND RATE_BY = " & strUserMemberID
	set rsA = my_Conn.execute(sSql)
	if not rsA.EOF then
	  intRatingID = rsA("RATING_ID")
	  if rsA("RATING") > 0 and bCanRate then
  	    bCanRate = false
	  end if
	  if rsA("COMMENTS") <> "" and bCanComment then
  	    bCanComment = false
	  end if
	end if
	set rsA = nothing
  	  bSecCodeMatch = true
  	  if intSecCode <> 0 then
    	fSecCode = ChkString(request.form("secCode"),"sqlstring")
    	if DoSecImage(fSecCode) then
          'Image matched their input 
          bSecCodeMatch = true
    	else
      	  'Image did not match their input
      	  bSecCodeMatch = false
	  	  sError = sError & "<li><span class=""fAlert"">" & txtBadSecCode & "</span></li>"
    	end if
  	  end if
	  if bCanRate or bCanComment then
	    if bCanRate then
	      if isnumeric(Request.Form("rating")) then
	        intFormRating = cint(Request.Form("rating"))
			if intFormRating = 0 then
			  intFormRating = 0
			  'sError = sError & "<li><span class=""fAlert"">You didn't select a rating.</span></li>"
			end if
		  end if
	    end if
	    if bCanComment then
	      if len(Request.Form("comments") & "x") > 1 then
	        strComments = ChkString(Request.Form("comments"),"message")
		  else
			'sError = sError & "<li><span class=""fAlert"">You did not make a comment.</span></li>"
		  end if
	    end if
		if (bCanRate or bCanComment) and (intFormRating = 0 and strComments = "") then
		  if bCanComment then
			sError = sError & "<li><span class=""fAlert"">" & txtNoComment & "</span></li>"
		  end if
		  if bCanRate then
			sError = sError & "<li><span class=""fAlert"">" & txtNoRating & "</span></li>"
		  end if
		end if
		
		if len(sError) = 0 and (intFormRating > 0 or strComments <> "") then
		  ':: update item votes and rating
		  sSQL = "SELECT VOTES, RATING FROM " & tbl & " WHERE " & fld & " = " & itm
		  set rsA = my_Conn.execute(sSql)
		  if rsA.eof then
		    sMsg = txtItemNotFnd
   			call showMsgBlock(1,sMsg)
		    set rsA = nothing
			closeAndGo("stop")
		  else
		    if intFormRating > 0 then
		      if rsA("VOTES") <> "" then
		        intVotes = rsA("VOTES") + 1
			  else
		        intVotes = 1
			  end if
		      if rsA("RATING") <> "" then
		        intRating = rsA("RATING") + intFormRating
			  else
		        intRating = intFormRating
			  end if
		  	  strSQL = "UPDATE " & tbl & " SET VOTES = " & intVotes & " , RATING = " & intRating & " WHERE " & fld & " = " & itm
		      executeThis(strSQL)
			end if
		  end if
		  set rsA = nothing
		  
		sMsg = txtThankYouFor & "&nbsp;"
		sA = ""
		  ':: add rating/comment to ratings table
		if intRatingID > 0 then
		  sSQL = "UPDATE " & strTablePrefix & "M_RATING SET"
		  if intFormRating >= 0 then
		    sSQL = sSQL & " RATING = " & intFormRating & ","
		    sMsg = sMsg & lcase(txtRating)
			sA = "&nbsp;" & txtAnd & "&nbsp;"
		  end if
		  if strComments <> "" then
		    sSQL = sSQL & " COMMENTS = '" & strComments & "',"
		    sMsg = sMsg & sA & txtCommOn
		  end if
		  sSQL = sSQL & " RATE_BY = " & strUserMemberID & ","
		  sSQL = sSQL & " RATE_DATE = '" & strCurDateString & "'"
		  sSQL = sSQL & " WHERE RATING_ID = " & intRatingID
		else
		  sSQL = "INSERT INTO " & strTablePrefix & "M_RATING ("
		  sSQL = sSQL & "ITEM_ID"
		  if intFormRating > 0 then
		    sSQL = sSQL & ",RATING"
		  end if
		  if strComments <> "" then
		    sSQL = sSQL & ",COMMENTS"
		  end if
		  sSQL = sSQL & ",RATE_BY"
		  sSQL = sSQL & ",RATE_DATE"
		  sSQL = sSQL & ",APP_ID"
		  sSQL = sSQL & ") VALUES ("
		  sSQL = sSQL & itm
		  if intFormRating > 0 then
		    sSQL = sSQL & "," & intRating
		    sMsg = sMsg & lcase(txtRating)
			sA = "&nbsp;" & txtAnd & "&nbsp;"
		  end if
		  if strComments <> "" then
		    sSQL = sSQL & ",'" & strComments & "'"
		    sMsg = sMsg & sA & txtCommOn
		  end if
		  sSQL = sSQL & "," & strUserMemberID
		  sSQL = sSQL & ",'" & strCurDateString & "'"
		  sSQL = sSQL & "," & intAppID
		  sSQL = sSQL & ")"
		end if
		  executeThis(sSQL)
		  sMsg = sMsg & "&nbsp;" & txtThisItem
    	  sMsg = jsReloadOpener & icon(imgComment,"","","","align=""left"" hspace=""10"" vspace=""1""") & sMsg
    	  'call showMsgBlock(1,sMsg)
		
		else ':: they can comment or rate, but didnt
		  ':: should display the form again
		  sMsg = sError
	  	  Call mod_shoRateForm(itm,app,tbl,fld,tFld,bCanComment,bCanRate,intFormRating,strComments,sMsg)
		end if
	  else ':: already rated and commented
		sMsg = icon(imgComment,"","","","align=""left"" hspace=""10"" vspace=""1""") & sError
	  end if
	
	  if sMsg <> "" then
   	    call showMsgBlock(1,sMsg)
	  end if
	  
	else ':: method <> add
	  ':: show form here
	  Call mod_shoRateForm(itm,app,tbl,fld,tFld,bCanComment,bCanRate,0,"","")
	  
	end if ':: end method check
	
  end if ':: member check
end sub

sub mod_shoRateForm(itm,app,tbl,fld,tFld,bc,br,frmRate,frmCom,msg)
  dim strLinkSQL, rsLink
  bCanComment = bc
  bCanRate = br
  
	' Check to see if they rated or commented already
	sSQL = "SELECT RATING,COMMENTS,RATING_ID FROM " & strTablePrefix & "M_RATING WHERE ITEM_ID = " & itm & " AND APP_ID=" & intAppID & " AND RATE_BY = " & strUserMemberID
	set rsA = my_Conn.execute(sSql)
	if not rsA.EOF then
	  intRatingID = rsA("RATING_ID")
	  if rsA("RATING") > 0 and bCanRate then
  	    bCanRate = false
	    sError = "<li><span class=""fAlert"">" & txtAlreadyRated & "</span>"
	    sError = sError & "<br />" & txtURating & ": " & rsA("RATING") & "<br/>&nbsp;</li>"
	  end if
	  if rsA("COMMENTS") <> "" and bCanComment then
  	    bCanComment = false
	    sError = sError & "<li><span class=""fAlert"">" & txtAlreadyComm & "</span>"
	    sError = sError & "<br />" & txtUComment & ": <br/>" & rsA("COMMENTS") & "</li>"
	    'sError = sError & "<li>You cannot comment on it again.</li>"
	  end if
	end if
	set rsA = nothing
	
	  sSql = "SELECT * FROM " & tbl & " WHERE " & fld & " = " & itm
	  Set rsA = my_Conn.execute(sSql)
	  if rsA.EOF then
		sMsg = txtItemNotFnd
   	    call showMsgBlock(1,sMsg)
	  else
		strLinkTitle = rsA(tFld)
		intLinkID = rsA(fld)
		intHit = rsA("HIT")		
		strPostDate = strtodate(rsA("POST_DATE"))
		dateSince=DateDiff("d", Date(), strPostDate)+7 %>
		<form method="post" name="rateform" action="<%= sScript %>">
		<%
		spThemeTitle = txtAddCommRating
		spThemeBlock1_open(intSkin)%>
		<table cellspacing="4" cellpadding="0">
		<tr><td class="fNorm">
		<%= icon(imgComment,"","","","align=""left"" hspace=""10"" vspace=""1""") %>
		<div style="marginLeft:4px;">
		<% 
		If msg <> "" Then
		  Response.Write(msg)
		else
		  Response.Write(txtPlaceCommRate)
		end if
		%>
		</div>
		</td></tr>
		<tr><td><hr/></td></tr>
		<tr><td class="fNorm"><b><%=strLinkTitle%></b><br />
		(<%= txtAdded %> : <%=formatdatetime(strPostDate, 2)%>&nbsp;<%= txtHits %>&nbsp;:&nbsp;<%=intHit%>)<br /></td></tr>
		<% If sError <> "" and (not bCanComment or not bCanRate) Then %>
		<tr><td><hr/></td></tr>
		<tr><td class="fNorm"><ul><%= sError %></ul></td></tr>
		<% End If %>
		<% If bCanRate Then %>
		<tr><td><hr/></td></tr>
		<tr><td class="fNorm">
		<b><%= txtRateItem %>:</b>&nbsp;
		<select name="rating">
          <option value="0">-</option>
          <option value="1"<%= chkSelect(frmRate,1) %>>1</option>
          <option value="2"<%= chkSelect(frmRate,2) %>>2</option>
          <option value="3"<%= chkSelect(frmRate,3) %>>3</option>
          <option value="4"<%= chkSelect(frmRate,4) %>>4</option>
          <option value="5"<%= chkSelect(frmRate,5) %>>5</option>
          <option value="6"<%= chkSelect(frmRate,6) %>>6</option>
          <option value="7"<%= chkSelect(frmRate,7) %>>7</option>
          <option value="8"<%= chkSelect(frmRate,8) %>>8</option>
          <option value="9"<%= chkSelect(frmRate,9) %>>9</option>
          <option value="10"<%= chkSelect(frmRate,10) %>>10</option>
        </select>
		</td></tr>
		<tr><td class="fSmall"><b><%= txtRatingScale %>: <span class="fAlert">1</span> = <%= txtWorst %>, <span class="fAlert">10</span> = <%= txtBest %></b></td></tr>
		<% End If %>
		<% If bCanComment Then %>
		<tr><td><hr/></td></tr>
		<tr><td class="fNorm"><b><%= txtCommThisItem %>:</b><br />
		<div class="fNorm"><%= txtBeObjective %></div><br />
		<textarea rows="8" cols="50" name="comments"><%= frmCom %></textarea></td></tr>
		<% End If
		if intSecCode <> 0 and (bCanComment or bCanRate) then
		%>
		<tr><td><hr/></td></tr>
		<tr><td colspan="2" align="center" class="fNorm"><% shoSecurityImg %></td></tr>
		<%
		End If %>
		<% If bCanComment or bCanRate Then %>
		<tr><td><hr/></td></tr>
		<tr><td align=center>
		<input type="submit" value=" Submit " class="button" />
		<input type="hidden" name="cid" value="<%= itm %>" />
		<input type="hidden" name="item" value="<%= itm %>" />
		<input type="hidden" name="method" value="add" />
		<input type="hidden" name="mode" value="<%=sMode%>" /></td></tr>
		<% End If %>
		</table>
		<% spThemeBlock1_close(intSkin)%>
		</form>
<%
	  end if
	  set rsA = nothing
end sub

sub mod_showCategoryManager()
  if sMode = "23" then
    doCatCountUpdate()
  end if
  if instr(sScript,app_page) > 0 then
    arg2 = txtCatMgr
    shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	
	spThemeTitle = txtDownloads & " - " & txtCatMgr
	spThemeBlock1_open(intSkin)
  end if
  If strMsg <> "" Then
	Response.Write("<span class=""fTitle"">")
	Response.Write(strMsg)
	Response.Write("</span><hr />")
	strMsg = ""
  End If
  chkSessionMsg()
	if bAppFull then
	  Response.Write "<table border=""0"" width=""70%"" cellspacing=""0"" cellpadding=""5"">"
	  Response.Write "<tr><td width=""50%"" align=""center"">"
	    mod_addCat()
	  Response.Write "</td><td>"
	    mod_updateCatCounts()
	  Response.Write "</td></tr></table>"
	  Response.Write "<hr/>"
	end if
	  sSQL = "SELECT count(CAT_ID) FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID
	  Set RScount = my_Conn.Execute(sSQL)
	    rcount = RScount(0)
	  set RScount = nothing
	sql = "select * from " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID & " ORDER BY C_ORDER, CAT_NAME"
	set rs = my_Conn.Execute (sql)
	%>
	<script type="text/javascript">
function jsDelCat(nam,rid){
	var stM
	stM = "<%= txtThisWillDel %>:\n\n";
	stM += "<%= txtTheCat %> "+nam+"\n";
	stM += "<%= txtAllSubcats %>\n";
	stM += "<%= txtAllAssocItems %>\n";
	stM += "\n<%= txtRemNoBeUndone %>\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="<%= sScript %>?cmd=<%= iPgType %>&mode=4&cid="+rid+"";
	}else{
	  return;
	}
}
	</script><br />
	<table border="0" width="500" cellspacing="0" cellpadding="5" class="grid">
	  <tr>
		<td width="120" align="center" class="tTitle">
		<%= txtOptions %></td>
		<td class="tTitle">&nbsp;&nbsp;<%= txtCatNam %></td>
		<% if bAppFull then %>
		<td width="100" align="center" class="tTitle">
		<%= txtOrder %></td>
		<% end if %>
	  </tr>
	<% 
	Do while NOT rs.EOF
	  if hasAccess(rs("CG_FULL")) or bAppFull then
	  %>
	    <tr><td align="center" valign="top" nowrap>
	<%
	    response.Write(modGrpEdit(app_pop,14,rs("CAT_ID"),0,"middle",rs("CG_INHERIT")))
			'sTo = "dl.asp?cmd=20&amp;cid=" & rsCategories("cat_id")
			'response.Write("&nbsp;" & modGrpEdit(sTo,,,,"middle",rsCategories("CG_INHERIT")))

	If bAppFull Then
	    Response.Write icon(icnDelete,txtDel,"display:inline;cursor:pointer;","jsDelCat('" & rs("CAT_NAME") & "','" & rs("CAT_ID") & "')","align=""middle""")
	End If
	
	If isMAC Then
	  Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","hide('ren"&rs("CAT_ID")&"');show('asub"&rs("CAT_ID")&"');","align=""middle""")
	  Response.Write icon(icnPlus,txtCrtNewSubCat,"display:inline;cursor:pointer;","hide('asub"&rs("CAT_ID")&"');show('ren"&rs("CAT_ID")&"');","align=""middle""")
	Else 
	  Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","openJsLayer('ren"& rs("CAT_ID") &"','250','150');","align=""middle""")
	  Response.Write icon(icnPlus,txtCrtNewSubCat,"display:inline;cursor:pointer;","openJsLayer('asub"& rs("CAT_ID") &"','260','200');","align=""middle""")
	End If
	  Response.Write "<a href="""& sScript &"?cmd="& (iPgType + 1) &"&cid="&rs("CAT_ID")&""">"
	  Response.Write icon(icnBinoc,txtViewSubcats,"display:inline;","","hspace=""4"" align=""middle""")
	  Response.Write "</a>"
	   %>
    </td>
		<td class="fSubTitle">
		<a href="<%= sScript %>?cmd=<%= (iPgType + 1) %>&cid=<%=rs("CAT_ID")%>">&nbsp;<%= rs("CAT_NAME") %></a>
		<% 
		mod_editCat(rs)
		mod_addSub rs("CAT_ID"),rs("CAT_NAME"),iPgType,"none"
		%></td>
		<% if bAppFull then %>
		<td width="100" align="center" valign="top">
	    <form action="<%= sScript %>" method="post" name="dorder<%= rs("CAT_ID") %>" id="dorder<%= rs("CAT_ID") %>">
        <input name="mode" type="hidden" id="mode" value="9">
  		<input name="cid" type="hidden" id="cid" value="<%= rs("CAT_ID") %>">
  		<input name="cmd" type="hidden" id="cmd" value="<%= iPgType %>">
	    <select name="sid" onChange="submit()">
		<% for xc = 1 to rcount %>
    	  	<option value="<%= xc %>"<% If rs("C_ORDER") = xc Then response.Write(" selected") %>><%= xc %></option>
		<% next %>
		</select></form>
		</td>
		<% end if %>
	</tr>
	  <% 
	  end if
	  rs.MoveNext
	Loop 
	%>
	</table><br />&nbsp;
	<% 
	set rs = nothing
  Call mod_shoLegend("admin",true,true)
  if instr(sScript,app_page) > 0 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub mod_showSubCategoryManager() 'list of subcategories for specific category 
  catFull = false
  if instr(sScript,app_page) > 0 then
    arg2 = txtSCatMgr
    shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
    spThemeTitle = txtDownloads & " - " & txtSCatMgr
    spThemeBlock1_open(intSkin)
  end if
  if cid = 0 and (instr(sScript,app_admin) > 0 or bAppFull) then
    sSql = "SELECT CAT_ID FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID & ""
	'response.Write(sSql)
	set rsT = my_Conn.execute(sSql)
	if not rsT.eof then
	  cid = rsT(0)
	end if
	set rsT = nothing
  end if
  If strMsg <> "" Then
	Response.Write("<span class=""fTitle"">")
	Response.Write(strMsg)
	Response.Write("</span><hr />")
	strMsg = ""
  End If
  chkSessionMsg()
 %>
  <table border="0" width="300" cellspacing="0" cellpadding="5">
	<tr>
	  <td align="center" class="fSubTitle">
		<%= txtSelCat %>: <%= mod_selectCatSubmit(cid) %>
	    <%
		If cid > 0 Then
          sSql = "SELECT CAT_ID, CAT_NAME, CG_FULL, CG_INHERIT FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID="& cid & " AND APP_ID = " & intAppID & ""
	      set rsA = my_Conn.execute(sSql)
	      if not rsA.eof then
	        cn = rsA("CAT_NAME")
			catInh = rsA("CG_INHERIT")
			catFull = hasAccess(rsA("CG_FULL"))
	      end if
		  if catFull then
			mod_editCat(rsA)
		    mod_addSub cid,cn,iPgType,"none"
		  end if
	      set rsA = nothing
		End If %>
	  </td>
	</tr>
  </table>
<% 
  if bAppFull then
	catFull = bAppFull
  end if
 if cid > 0 then
%>
	<hr /><br />
	<script type="text/javascript">
function jsDelSCat(nam,rid,c){
	var stM
	stM = "<%= txtThisWillDel %>:\n\n";
	stM += "<%= txtTheSubcat %> - "+nam+"\n";
	stM += "<%= txtAnd %> "+c+" <%= txtAssocItems %>\n";
	stM += "\n<%= txtRemNoBeUndone %>\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="?cmd=<%= iPgType %>&mode=6&cid=<%= cid %>&sid="+rid ;
	}else{
	  return;
	}
}
function jsDelCat(nam,rid){
	var stM
	stM = "<%= txtThisWillDel %>:\n\n";
	stM += "<%= txtTheCat %> "+nam+"\n";
	stM += "<%= txtAllSubcats %>\n";
	stM += "<%= txtAllAssocItems %>\n";
	stM += "\n<%= txtRemNoBeUndone %>\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="<%= sScript %>?cmd=<%= iPgType %>&mode=4&cid="+rid;
	}else{
	  return;
	}
}
	</script>
	  <table border="0" width="560" cellspacing="0" cellpadding="5" class="grid">
	  <tr>
		<td align="center" class="tTitle">&nbsp;
		<% If bAppFull Then %>
	  <a href="<%= sScript %>?cmd=20&cid=<%=cid%>" title="Category Manager"><%= icon(icnArLeft,txtCatMgr,"","","align=""middle""") %></a>&nbsp;&nbsp;
	    <% End If
		
	If catFull or bAppFull Then
  	  response.Write(modGrpEdit(app_pop,14,cid,0,"middle",catInh))
	  If bAppFull Then
		response.Write icon(icnDelete,txtDel,"cursor:pointer;","jsDelCat('"& cn &"', '"& cid &"');","align=""middle""")
	  End If
	  If isMAC Then
		response.Write icon(icnEdit,txtEdit,"cursor:pointer;","mwpHSs('ren" & cid & "','1');","align=""middle""")
	    Response.Write icon(icnPlus,txtCrtNewSubCat,"display:inline;cursor:pointer;","hide('asub"& cid &"');show('ren"&rs("CAT_ID")&"');","align=""middle""")
	  Else
		response.Write icon(icnEdit,txtEdit,"cursor:pointer;","openJsLayer('ren"& cid &"','250','150');","align=""middle""")
	    Response.Write icon(icnPlus,txtCrtNewSubCat,"display:inline;cursor:pointer;","openJsLayer('asub"& cid &"','260','200');","align=""middle""")
	  End If
    End If %>
		</td>
		<td colspan="2" class="tTitle">
		&nbsp;&nbsp;<%= cn %>
		</td>
	  </tr>
	  <tr>
		<td align="center" valign="top" class="tSubTitle" width="130">
		<b><%= txtOptions %></b></td>
		<td class="tSubTitle">
		&nbsp;&nbsp;<b><%= txtSubCatNam %></b></td>
		<td width="60" align="center" valign="top" class="tSubTitle">
		<% If catFull Then %>
		<b><%= txtOrder %></b>
		<% Else %>
		&nbsp;
		<% End If %></td>
	  </tr>
	  <% 
  sSQL = "SELECT count(CAT_ID) FROM " & strTablePrefix & "M_SUBCATEGORIES where CAT_ID=" & cid & " AND APP_ID = " & intAppID & ""
   Set RScount = my_Conn.Execute(sSQL)
	scount = RScount(0)
   set RScount = nothing

	sql = "SELECT * From " & strTablePrefix & "M_SUBCATEGORIES where CAT_ID=" & cid & " AND APP_ID = " & intAppID & " order by C_ORDER, SUBCAT_ID"
	set rs = my_Conn.Execute (sql)
	if rs.eof then
	  Response.Write "<tr><td colspan=""3"" align=""center"">"
	  Response.Write "<p>&nbsp;</p>"
	  Response.Write "<p>" & txtNoSubcatFnd & "</p>"
	  Response.Write "<p>&nbsp;</p>"
	  Response.Write "</td></tr>"
	else
	  Do while NOT rs.EOF
		subFull = hasAccess(rs("SG_FULL"))
		subWrite = hasAccess(rs("SG_WRITE"))
  		if catFull or bAppFull then
		  subFull = true
		  subWrite = true
  		end if
	    If subFull Then
		    rcount = 0
		    sSQL = "SELECT count(" & item_fld & ") FROM " & item_tbl & " where " & item_scat_fld & "=" & rs("SUBCAT_ID") & " AND ACTIVE=1"
			Set RScount = my_Conn.Execute(sSQL)
			  rcount = RScount(0)
			set RScount = nothing %>
		    <tr>	
			<td align="center" valign="top">
		<% 
		  Response.Write(modGrpEdit(app_pop,14,cid,rs("SUBCAT_ID"),"middle",rs("SG_INHERIT")))
		If catFull Then
		  response.Write icon(icnDelete,txtDel,"display:inline;cursor:pointer;","jsDelSCat('"& rs("SUBCAT_NAME") &"', '"& rs("SUBCAT_ID") &"','"& rcount &"');","align=""middle""")
		End If
		Response.Write icon(icnEdit,txtEdit,"display:inline;cursor:pointer;","mwpHSs('rens" & rs("SUBCAT_ID") & "','1');","align=""middle""")
		'response.Write icon(icnEdit,txtEdit,"cursor:pointer;","openJsLayer('rens"& rs("SUBCAT_ID") &"','250','150');","align=""middle""")
	    Response.Write "<a href=""" & app_addForm & "parent_id="& cid &"&cat_id="& rs("SUBCAT_ID") &""">"
	    Response.Write icon(icnPlus,txtAddItmSub,"display:inline;","","align=""middle""")
	    Response.Write "</a>"
	    Response.Write "<a href=""" & app_page & "?cmd=2&cid="& cid &"&sid="& rs("SUBCAT_ID") &""">"
	    Response.Write icon(icnBinoc,txtViewSubContents,"display:inline;","","hspace=""4"" align=""middle""")
	    Response.Write "</a>"
	    chkSubCatAttention(rs("SUBCAT_ID"))
	    %>
			</td>
			<td class="fSubTitle">
			&nbsp;<a href="<%= app_page %>?cmd=2&cid=<%= cid %>&sid=<%= rs("SUBCAT_ID")%>">
			<%= ChkString(rs("SUBCAT_NAME"), "display") %> <small>(<%= rcount %>)</small></a>
			<%
			Call mod_chkNewSubCatItems(rs("SUBCAT_ID"),true,true)
			  mod_editSub(rs)
			%>
			</td>
			<td align="center" valign="top">
		<% If catFull Then %>
			<form action="<%= sScript %>" method="post" name="dorder<%= rs("SUBCAT_ID") %>" id="dorder<%= rs("SUBCAT_ID") %>">
        	<input name="mode" type="hidden" id="mode" value="10">
  			<input name="sid" type="hidden" id="sid" value="<%= rs("SUBCAT_ID") %>">
  			<input name="cid" type="hidden" id="cid" value="<%= cid %>">
  			<input name="cmd" type="hidden" id="cmd" value="<%= iPgType %>">
	    	<select name="ord" onChange="submit()">
			<% for xc = 1 to scount %>
    	  		<option value="<%= xc %>"<% If rs("C_ORDER") = xc Then response.Write(" selected") %>><%= xc %></option>
			<% next %>
			</select></form>
		<% Else %>
			&nbsp;
		<% End If %>
			</td>
		</tr>
		<% 
		end if
		rs.MoveNext  
 	  Loop
	end if 
  	set rs = nothing 
	%></table><br />&nbsp;<%
 end if
  Call mod_shoLegend("admin",true,true)
  if instr(sScript,app_page) > 0 then
    spThemeBlock1_close(intSkin)
  end if
end sub

function mod_selectCatSubmit(c)
  t = "<form action=""" & sScript & """ method=""get"">"
  t = t & "<input name=""cmd"" type=""hidden"" value=""" & iPgType & """>"
  t = t & "<select name=""cid"" onChange=""submit()"">"
  't = t & "<option value=""0""" & chkSelect(c,0) & ">"
  't = t & "&nbsp;-[" & txtSelCat & "]-&nbsp;</option>"
  sql = "select * FROM " & strTablePrefix & "M_CATEGORIES WHERE APP_ID = " & intAppID & " ORDER BY CAT_NAME"
  set rsA = my_Conn.Execute(sql)
  do while not rsA.eof
    t = t & "<option value=""" & rsA("CAT_ID") & """" & chkSelect(rsA("CAT_ID"),c) & ">"
    t = t & rsA("CAT_NAME") & "</option>"
    rsA.movenext
  loop
  set rsA = nothing
  t = t & "</select></form>"
  mod_selectCatSubmit = t
end function

sub mod_addCat() %>
  <div style="width:250px;">
  <fieldset style="margin:5px;padding:10px;">
  <p align="center"><span class="fTitle"><%= txtCrtNewCat %></span>
  <form method="post" action="<%= sScript %>">
  <input type="hidden" value="<%= iPgType %>" name="cmd">
  <input type="hidden" value="1" name="mode">
  <input type="text" name="newcat" size="30" class="textbox" style="margin:5px;"><br /><input type="submit" value="<%= txtSubmit %>" name="B1" class="button">
  </form></p></fieldset></div>
  <%
end sub

sub mod_editCat(ob)
  %><div style="display:none;" id="ren<%= ob("CAT_ID") %>">
  <fieldset style="margin:10px;padding:5px;">
  <b><%= txtEditCat %>:</b><br>
  <form action="<%= sScript %>" method="post">  
  <input type="hidden" value="<%= iPgType %>" name="cmd">
  <input type="hidden" value="3" name="mode">
  <input type="hidden" value="<%= ob("CAT_ID") %>" name="cid">
  <input type="text" value="<%=ob("CAT_NAME")%>" name="newcat" size="30">	
  <input type="submit" value="<%= txtSubmit %>" class="button">
  </form></fieldset></div>
  <%
end sub

sub mod_addSub(i,n,c,d)
  if (c = 2 or c = 20) and i = 0 then
    d = "none"
  end if
  %><div style="display:<%= d %>;" id="asub<%= i %>">
  <fieldset style="margin:10px;padding:5px;">
  <% 
  if (c = 2 or c = 20) or (not isMac and c = 2) then
    if n = "" then
      sSql = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & i & " AND APP_ID = " & intAppID & ""
	  set rsA = my_Conn.execute(sSql)
	  if not rsA.eof then
	    n = rsA("CAT_NAME")
	  end if
	  set rsA = nothing
    end if
    Response.Write("<span class=""fAlert"">")
    Response.Write("<b>" & n & "</b>")
    Response.Write("</span><br />")
  end if%>
  <b><%= txtCrtNewSubCat %></b>
  <form method="post" action="<%= sScript %>">
  <input type="hidden" value="<%= c %>" name="cmd">
  <input type="hidden" value="2" name="mode">
  <input type="hidden" value="<%= i %>" name="cid">
  <input type="text" name="newsub" size="30"><br />
  <input type="submit" value="<%= txtSubmit %>" name="B1" class="button">
  </form></fieldset></div>
  <%
end sub

sub mod_editSub(ob) 'list of subcategories for selection %>
  <div style="display:none;" id="rens<%= ob("SUBCAT_ID") %>">
  <fieldset style="margin:10px;padding:5px;">
  <form action="<%= sScript %>" method="post">
  <b><%= txtSelCat %></b><br>
  <% mod_selectCats ob("CAT_ID"),"FULL" %><br>
  <b><%= txtEditSub %>:</b><br>
  <input type="text" value="<%= ob("SUBCAT_NAME") %>" name="newcat" size=30><br/>
  <input type="submit" value="<%= txtSubmit %>" class="button">
  <input type="hidden" value="<%= iPgType %>" name="cmd">
  <input type="hidden" value="5" name="mode">
  <input type="hidden" value="<%= cid %>" name="cid">
  <input type="hidden" value="<%= ob("SUBCAT_ID") %>" name="sid">
  </form></fieldset></div>
  <% 
end sub

sub mod_addCategory()
  newcat = trim(chkString(Request.Form("newcat"), "sqlstring"))
  if newcat = "" then
	strMsg = strMsg & "<b>" & txtPlzEnterCatName & "</b>"
  else
	Set RS=Server.CreateObject("ADODB.Recordset")
	strSql="Select CAT_NAME from " & strTablePrefix & "M_CATEGORIES where CAT_NAME='" & newcat & "' AND APP_ID = " & intAppID & ""
	RS.Open strSql,my_Conn , 2, 2
	if rs.eof then
		isOK = true
	else
		isOK = false
	end if
	RS.close
	set RS = nothing

	if isOK then
	  sSql = "INSERT INTO " & strTablePrefix & "M_CATEGORIES ("
	  sSql = sSql & "CAT_NAME,CG_READ,CG_WRITE,CG_FULL,CG_INHERIT,CG_PROPAGATE,APP_ID"
	  sSql = sSql & ") VALUES ("
	  sSql = sSql & "'" & newcat & "','" & sAppRead & "'"
	  sSql = sSql & ",'" & sAppWrite & "','" & sAppFull & "',1,1"
	  sSql = sSql & ", " & intAppID & ")"
		executeThis(sSql)
		strMsg = strMsg & txtNewCat & ": "
		strMsg = strMsg & "<span class=""fAlert""><b>"
		strMsg = strMsg & "" & newcat & "</b></span>"
		strMsg = strMsg & "<br><br>"
		strMsg = strMsg & txtCrNewSubForCat
	else
		strMsg = strMsg & "<b>" & txtCatAlrExist & ":<br>"
		strMsg = strMsg & "<span class=""fAlert""><b>" & newcat & "</b></span>"
		strMsg = strMsg & "<br><br>" & txtPlzNewCatNam & "</b>"
	end if
	'iPgType = 1
  end if
  Call setSession("sMsg",strMsg)
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "")
end sub

sub mod_renameCategory(c)
    cat= ChkString(Request.Form("newcat"),"SQLString")
    if trim(cat) <> "" then
      sSql = "UPDATE " & strTablePrefix & "M_CATEGORIES SET CAT_NAME='" & cat & "' WHERE CAT_ID=" & c & " AND APP_ID = " & intAppID & ""
      executeThis(sSql)
      strMsg = strMsg & txtCatRenamed & ": <b>" & cat & "</b>"
    else
      strMsg = strMsg & txtChooseCatName
    end if
	Call setSession("sMsg",strMsg)
    resetCoreConfig()
    closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & c)
end sub

sub mod_addSubCategory(c)
    newsub = trim(chkString(Request.Form("newsub"), "sqlstring"))
    if newsub = "" then
	  strMsg = strMsg & txtPlzEnterSCatName
    else
	 if cid = 0 then
	  strMsg = strMsg & txtSelCatForSub
	 else
	  strSql="Select * from " & strTablePrefix & "M_CATEGORIES where CAT_ID=" & cid & " AND APP_ID = " & intAppID & ""
	  set rsC = my_Conn.execute(strSql)
		cat_name = rsC("CAT_NAME")
		r = rsC("CG_READ")
		w = rsC("CG_WRITE")
		f = rsC("CG_FULL")
	  set rsC = nothing
	  sSql = "INSERT INTO " & strTablePrefix & "M_SUBCATEGORIES ("
	  sSql = sSql & "SUBCAT_NAME,CAT_ID,SG_READ"
	  sSql = sSql & ",SG_WRITE,SG_FULL,SG_INHERIT,APP_ID"
	  sSql = sSql & ") VALUES ("
	  sSql = sSql & "'" & newsub & "','" & c & "','" & r & "'"
	  sSql = sSql & ",'" & w & "','" & f & "',1," & intAppID & ""
	  sSql = sSql & ")"
	  executeThis(sSql)
	  strMsg = strMsg & txtNewSub & "&nbsp;<span class=""fAlert""><b>" & newsub & "</b></span>&nbsp;" & txtAddedToCat & "&nbsp;<b>" & cat_name & "</b>"
	 end if
    end if
	Call setSession("sMsg",strMsg)
    resetCoreConfig()
    closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & c & "")
end sub

sub mod_renameSubCategory(c,s)
  sName = trim(chkString(Request.Form("newcat"), "sqlstring"))
  sCatID = clng(Request.Form("cat"))
  executeThis("UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET SUBCAT_NAME='" & sName & "', CAT_ID=" & sCatID & " WHERE SUBCAT_ID=" & s)
  strMsg = strMsg & txtSubUpdated
  Call setSession("sMsg",strMsg)
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & c & "")
end sub

sub mod_updateCatOrder(c,o)
  sSql = "UPDATE " & strTablePrefix & "M_CATEGORIES SET C_ORDER=" & o & " WHERE CAT_ID=" & c
  executeThis(sSql)
  strMsg = strMsg & txtOrdUpd
  Call setSession("sMsg",strMsg)
  resetCoreConfig()
  closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & cid & "")
end sub

sub mod_updateSubCatOrder(c,s)
    ord = clng(request("ord"))
    sSql = "UPDATE " & strTablePrefix & "M_SUBCATEGORIES SET C_ORDER=" & ord & " WHERE SUBCAT_ID=" & s
    executeThis(sSql)
    strMsg = strMsg & txtOrdUpd
	Call setSession("sMsg",strMsg)
    resetCoreConfig()
    closeAndGo(sScript & "?cmd=" & iPgType & "&cid=" & c & "")
end sub

function mod_chkCatFull(c)
  tOK = false
  if bAppFull then
    tOK = true
  else
    sSql = "SELECT CG_FULL FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & c
    set rsA = my_Conn.execute(sSql)
    if not rsA.eof then
      tOK = hasAccess(rsA("CG_FULL"))
    end if
    set rsA = nothing
  end if
  mod_chkCatFull = tOK
end function

function mod_chkSubCatFull(s)
  tOK = false
  if bAppFull then
    tOK = true
  else
    sSQL = "SELECT " & strTablePrefix & "M_CATEGORIES.CG_FULL, " & strTablePrefix & "M_SUBCATEGORIES.SG_FULL"
    sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
    sSQL = sSQL & " WHERE " & strTablePrefix & "M_SUBCATEGORIES.APP_ID=" & intAppID & " AND " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID=" & s
  
    set rsA = my_Conn.execute(sSql)
    if not rsA.eof then
      if hasAccess(rsA("CG_FULL")) or hasAccess(rsA("SG_FULL")) then
        tOK = true
	  end if
    end if
    set rsA = nothing
  end if
  mod_chkSubCatFull = tOK
end function

sub mod_chkNewSubCatItems(i,bn,bu)
 if bn then
  lastVisit = getCount(item_fld,item_tbl,"POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND ACTIVE = 1 AND CATEGORY=" & i)
  if lastVisit > 0 then
	Response.Write(icon(icnNew1,txtNewLstVisit,"display:inline;","","align=""middle"" hspace=""2"""))
  else
	d7 = DateToStr(dateAdd("d",-7,date()))
	last7 = getCount(item_fld,item_tbl,"POST_DATE >= '" & d7 & "' AND ACTIVE = 1 AND CATEGORY=" & i)
	if last7 > 0 then
	  Response.Write(icon(icnNew2,txtNewLst7,"display:inline;","","align=""middle"" hspace=""2"""))
	else
	  d14 = DateToStr(dateAdd("d",-14,date()))
	  last14 = getCount(item_fld,item_tbl,"POST_DATE >= '" & d14 & "' AND ACTIVE = 1 AND CATEGORY=" & i)
	  if last14 > 0 then
		Response.Write(icon(icnNew3,txtNewLst14,"display:inline;","","align=""middle"" hspace=""2"""))
	  end if
	end if
  end if
 end if
 if bu then
  mod_chkUpdatedSubCatItems(i)
 end if
end sub

sub mod_chkUpdatedSubCatItems(i)
  lastVisit = getCount(item_fld,item_tbl,"UPDATED >= '" & Session(strUniqueID & "last_here_date") & "' AND ACTIVE = 1 AND CATEGORY=" & i)
  if lastVisit > 0 then
	response.Write icon(icnUpdate1,txtUpdLstVisit,"display:inline;","","hspace=""2"" align=""middle""")
  else
	d7 = DateToStr(dateAdd("d",-7,date()))
	last7 = getCount(item_fld,item_tbl,"UPDATED >= '" & d7 & "' AND ACTIVE = 1 AND CATEGORY=" & i)
	if last7 > 0 then
	  response.Write icon(icnUpdate2,txtUpdLst7,"display:inline;","","hspace=""2"" align=""middle""")
	else
	  d14 = DateToStr(dateAdd("d",-14,date()))
	  last14 = getCount(item_fld,item_tbl,"UPDATED >= '" & d14 & "' AND ACTIVE = 1 AND CATEGORY=" & i)
	  if last14 > 0 then
	    response.Write icon(icnUpdate3,txtUpdLst14,"display:inline;","","hspace=""2"" align=""middle""")
	  end if
	end if
  end if
end sub

sub mod_writeApprovalJS() %>
  <script type="text/javascript">
  function jsDelDl(nam,rid,s){
	var stM
	stM = "<%= txtDelItem %>:\n\n";
	stM += ""+nam+"\n";
	stM += "\n<%= txtRemNoBeUndone %>\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="<%= sScript %>?cmd=<%= iPgType %>&mode=123&cid="+rid+"&sid="+s;
	}else{
	  return;
	}
  }
  function jsApprDl(nam,rid,s,t){
	var stM
	if (t == 1){
	  stM = "<%= txtApprItem %>:";
	  iM = "122";
	}else{
	  stM = "<%= txtUnmkBadItm %>";
	  iM = "121";
	}
	stM += "\n\n"+nam+"\n\n";
	stM += "<%= txtAreYouSure %>\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="<%= sScript %>?cmd=<%= iPgType %>&mode=" + iM + "&cid="+rid+"&sid="+s;
	}else{
	  return;
	}
  }
  
  function jsEditDL(rid,s){
	  window.location="<%= sScript %>?cmd=<%= iPgType %>&mode=321&item="+rid+"&sid="+s;
  }
  
  function jsAttnDL(dl){
    var ele = document.getElementById(dl);
	var elev = document.getElementById('view_pane');
	elev.innerHTML = ele.innerHTML;
  }
  </script>
  <%
end sub

sub mod_writeViewItem(o,f)
  Response.Write "<div style=""display:none;"" id=""view" & o(item_fld) & """>"
  if not isMac then
    Response.Write "<div style=""width:500px;"">"
  end if
  Response.Write "<b>View item below</b><br><hr>"
  Response.Write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" align=""center"">"
  Response.Write "<tr><td align=""right"" width=""40%"">"
  Response.Write "<b>" & txtCat & ":&nbsp;</b>"
  Response.Write "</td><td>"
  Response.Write "<b>" & o("CAT_NAME") & "</b>"
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>" & txtSubCat & ":&nbsp;</b>"
  Response.Write "</td><td>"
  Response.Write "<b>" & o("SUBCAT_NAME") & "</b>"
  Response.Write "</td></tr>"
  Response.Write "</table>"
  fct = f & "(o)"
  execute("Call " & fct)
  'showInfo(o)
  'GetComments(cat_id)
  Response.Write "<br/>"
  if not isMac then
    Response.Write "</div>"
  end if
  Response.Write "</div>"
end sub

sub mod_edit_Item(it)
  arg2 = txtEditItem
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	
  spThemeTitle = "&nbsp;"
  spThemeBlock1_open(intSkin)
  chkSessionMsg()
  mod_EditItemForm(it)
  spThemeBlock1_close(intSkin)
end sub

sub mod_EditItemForm(iid)
 sSql = mod_singleItemSql(item_tbl)
 sSql = sSql & "WHERE (((" & item_tbl & "." & item_fld & ")=" & iid & "))"
 set rsB = my_Conn.execute(sSql)
 if rsB.eof then
   Response.Write(txtItemNotFnd)
 else
  Response.Write "<div>"
  if iPgType <> 11 then
  Response.Write "<b>" & txtEditItem & "</b><br/><hr/>"
  end if
  Response.Write "<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%"" align=""center"">"
  Response.Write "<tr><td align=""right"" width=""40%"">"
  Response.Write "<b>" & txtCat & ":&nbsp;</b>"
  Response.Write "</td><td>"
  Response.Write "<b>" & rsB("CAT_NAME") & "</b>"
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""right"">"
  Response.Write "<b>" & txtSubCat & ":&nbsp;</b>"
  Response.Write "</td><td>"
  Response.Write "<b>" & rsB("SUBCAT_NAME") & "</b>"
  Response.Write "</td></tr>"
  Response.Write "</table><hr/>"
  showEditForm(rsB)
  'GetComments(cat_id)
  Response.Write "<br/>"
  Response.Write "</div>"
 end if
 set rsB = nothing
end sub

sub mod_shoLegend(t,bn,bu)
  Response.Write "<hr/>"
  Response.Write "<fieldset style=""width:450px;padding:10px;"">"
  Response.Write "<legend><b>" & txtIcnKey & ":</b></legend><br/>"
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""425"">"
  Response.Write "<tr><td width=""220"" valign=""top"">"
  
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""2"" width=""100%"">"
 if t = "main" then
  if intSubscriptions = 1 then
    Response.Write "<tr><td align=""center"" class=""fSmall"">"
    Response.Write icon(icnSubscribe,txtSubScr,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtAddSubsc
    Response.Write "</td></tr>"
    Response.Write "<tr><td align=""center"" class=""fSmall"">"
    Response.Write icon(icnUnSubscribe,txtUnSubScr,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtRemSubscr
    Response.Write "</td></tr>"
  end if
  if intBookmarks = 1 then
    Response.Write "<tr><td align=""center"" class=""fSmall"">"
    Response.Write icon(icnBookmark,txtAddToBkmks,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtAddToBkmks
    Response.Write "</td></tr>"
    Response.Write "<tr><td align=""center"" class=""fSmall"">"
    Response.Write icon(icnUnBookmark,txtRemBkmk,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtRemBkmk
    Response.Write "</td></tr>"
  end if
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write "</td></tr>"
  
 elseif t = "admin" then
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write icon(icnCheck,txtApprItem,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtApprItem
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write icon(icnDelete,txtDel,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtDelItem
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write icon(icnEdit,txtEditItem,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtEditItem
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write icon(icnPlus,txtAddChild,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtAddChild
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write icon(icnBinoc,txtViewItem,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtViewItem
  Response.Write "</td></tr>"
  Response.Write "<tr><td align=""center"" class=""fSmall"">"
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write "</td></tr>"
 end if
 
  Response.Write "</table>"
  Response.Write "</td><td valign=""top"">"
 if (t = "main" or t = "admin") and (bn or bu) then
  Response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""2"" width=""100%"">")
  if bn then
    Response.Write "<tr><td align=""center"">"
    Response.Write icon(icnNew1,txtNewLstVisit,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtNewLstVisit
    Response.Write "</td></tr><tr><td align=""center"">"
    Response.Write icon(icnNew2,txtNewLst7,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtNewLst7
    Response.Write "</td></tr><tr><td align=""center"">"
    Response.Write icon(icnNew3,txtNewLst14,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtNewLst14
    Response.Write "</td></tr>"
  end if
  if bu then
    Response.Write "<tr><td align=""center"">"
    Response.Write icon(icnUpdate1,txtUpdLstVisit,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtUpdLstVisit
    Response.Write "</td></tr><tr><td align=""center"">"
    Response.Write icon(icnUpdate2,txtUpdLst7,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtUpdLst7
    Response.Write "</td></tr><tr><td align=""center"">"
    Response.Write icon(icnUpdate3,txtUpdLst14,"","","")
    Response.Write "</td><td valign=""middle"" class=""fSmall"">"
    Response.Write " - " & txtUpdLst14
    Response.Write "</td></tr>"
  end if
  Response.Write "</table>"
 else
 end if
  Response.Write "</td></tr>"
 if bAppFull then
  Response.Write "<tr><td colspan=""2"" class=""fSmall"">"
  Response.Write "<hr/>"
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""2"" width=""100%"">"
  Response.Write "<tr><td width=""30"" align=""center"">"
  Response.Write icon(icnToolbox,txtMgr,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtCatMgr
  Response.Write "</td></tr>"
  Response.Write "<tr><td width=""30"" align=""center"">"
  Response.Write icon(icnAttention,txtItemNeedAtt,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtItemNeedAtt
  Response.Write "</td></tr>"
  Response.Write "<tr><td width=""30"" align=""center"">"
  Response.Write icon(icnAccess,txtGrpMgr,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtGrpMgr
  Response.Write "</td></tr>"
  Response.Write "<tr><td width=""30"" align=""center"">"
  Response.Write icon(icnAccessOn,txtGrpMgr,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtGrpMgr & " - " & txtInhFrmParent
  Response.Write "</td></tr>"
  Response.Write "<tr><td width=""30"" align=""center"">"
  Response.Write icon(icnAccessOff,txtGrpMgr,"","","")
  Response.Write "</td><td valign=""middle"" class=""fSmall"">"
  Response.Write " - " & txtGrpMgr & " - " & txtNoInhFrmParent
  Response.Write "</td></tr>"
  Response.Write "</table>"
  Response.Write "</td></tr>"
 end if
  Response.Write "</table>"
  Response.Write "</fieldset><br/><br/>"
end sub

sub mod_addFeatured(typ,atyp)
  response.Write("<br><br>")
  spThemeBlock1_open(intSkin)
  response.Write("<br><br>")
  if hasAccess(1) then
	strSql = "UPDATE " & typ & " set FEATURED = " & hp
	strSql = strSql & " WHERE " & typ & "_ID = " & cid
	executeThis(strSql)
%>
	<P align=center><b><%= uCase(typ) & " " & atyp %>&nbsp;<%= txtFeatItems %></b><br></P><script type="text/javascript"> opener.document.location.reload();</script>
<%
  else %>
	<p align=center><b><%= txtOnlyAdminAction %></b></p>
<%
  end If
  response.Write("<br><br>")
	spThemeBlock1_close(intSkin)
end sub

sub mod_delBookmark(i)
  bSQL = "DELETE FROM " & strTablePrefix & "BOOKMARKS WHERE M_ID=" & strUserMemberID & " AND BOOKMARK_ID=" & i
  executeThis(bSQL)
  strMsg = txtBkmkRem
  response.Write("<br /><br /><br />")
	  spThemeBlock1_open(intSkin)%>
	  <p align="center"><div class="fTitle">&nbsp;</div></p>
	  <script type="text/javascript">opener.document.location.reload();</script>
	  <table align="center" border="0">
	    <tr>
	      <td align="center">
		  <b><%= strMsg %></b><br /><br /><br /><br />
		  </td>
	    </tr>
	  </table><%
	  spThemeBlock1_close(intSkin)
  response.Write("<br /><br /><br />")
end sub

sub mod_delSubscription(i)
  bSQL = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " AND SUBSCRIPTION_ID=" & i
  executeThis(bSQL)
  strMsg = txtSubscrRem
  response.Write("<br /><br /><br />")
	  spThemeBlock1_open(intSkin)%>
	  <p align="center"><div class="fTitle">&nbsp;</div></p>
	  <script type="text/javascript">opener.document.location.reload();</script>
	  <table align="center" border="0">
	    <tr>
	      <td align="center">
		  <b><%= strMsg %></b><br /><br /><br /><br />
		  </td>
	    </tr>
	  </table><%
	  spThemeBlock1_close(intSkin)
  response.Write("<br /><br /><br />")
end sub

function mod_iconSubscribe(c,s)
  a = ""
  t = 0
  ' values for t
  ' 1 = category subscription
  ' 2 = subcat subscription
  ' 3 = module subscription
  If strUserMemberID > 0 and intSubscriptions = 1 Then
  if c = 0 and s = 0 then
    t = 3
	q = c
	tB = txtSubscMod
	tBU = txtRemModSubsc
  elseif c > 0 and s = 0 then
    t = 1
	q = c
	tB = txtSubscCat
	tBU = txtRemCatSubsc
  elseif c = 0 and s > 0 then
    t = 2
	q = s
	tB = txtSubscSCat
	tBU = txtRemSCatSubsc
  end if
  
	subscription_id = chkIsSubscribed(intAppID,c,s,0,strUserMemberID)
	  if subscription_id <> 0 then
		a = "<a href=""javascript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=delsub&amp;cid=" & subscription_id & "');"">" & icon(icnUnsubscribe,tBU,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
	  else
  		a = "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=addsub&amp;cmd=" & t & "&amp;cid=" & q & "')"">" & icon(icnSubscribe,tB,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>"
	  end if
  end if
  mod_iconSubscribe = a
end function

function mod_iconBookmark(c,s,i)
  a = ""
  t = 0
  ' values for t
  ' 1 = category bookmark
  ' 2 = subcat bookmark
  ' 3 = item bookmark
  if c > 0 and s = 0 and i = 0 then
    t = 1
	q = c
	tB = txtBkmkCat
	tBU = txtRemCatBkmk
  elseif c = 0 and s > 0 and i = 0 then
    t = 2
	q = s
	tB = txtBkmkSCat
	tBU = txtRemSCatBkmk
  elseif c = 0 and s = 0 and i > 0 then
    t = 3
	q = i
	tB = txtBkmkItem
	tBU = txtRemItemBkmk
  end if
  
  If strUserMemberID > 0 and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,c,s,i,strUserMemberID)
	  if bookmark_id <> 0 then
		a = "<a href=""javascript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=delbook&amp;cid=" & bookmark_id & "');"">" & icon(icnUnBookmark,tBU,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>" 
	  else
		a = "<a href=""javascript:;"" onclick=""JavaScript:openWindow('" & app_pop & "?mode=addbook&amp;cmd=" & t & "&amp;cid=" & q & "');"">" & icon(icnBookmark,tB,"","","align=""right"" style=""display:inline;"" hspace=""4""") & "</a>" 
	  end if
  end if
  mod_iconBookmark = a
end function

sub mod_setPageAppVars()
  if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
  end if
  if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		sMode = cLng(Request("mode"))
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
  if Request("item") <> "" or  Request("item") <> " " then
	if IsNumeric(Request("item")) = True then
		intItemID = cLng(Request("item"))
	else
		closeAndGo("default.asp")
	end if
  end if
end sub

%>