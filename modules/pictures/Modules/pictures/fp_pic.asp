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
numInRow = 3 ':: this value is for all the main picture module pages only.

':: strPicTablePrefix is used for multiple db installations of the pictures module
':: Make this unique for each installation in the same database
strPicTablePrefix = "" 

':: hard coded default directory for uploads
':: Make this unique for each installation in the same database
galleryDir = "files/" & strPicTablePrefix & "pictures/"  

':::::::::::::::::::::::::::::::::>
':: DO NOT EDIT BELOW THIS LINE ::>
':::::::::::::::::::::::::::::::::>

':: global variables for the image dimension routines
dim fpp_cid, fpp_sid
fpp_cid = 0
fpp_sid = 0
tImg = ""
displayTxt = ""
incPicFp = true

':: start image size and dimension functions
function GetBytes(flnm, offset, bytes)
     Dim objFSO
     Dim objFTemp
     Dim objTextStream
     Dim lngSize
     on error resume next
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     ' First, we get the filesize
     Set objFTemp = objFSO.GetFile(flnm)
     lngSize = objFTemp.Size
     set objFTemp = nothing

     fsoForReading = 1
     Set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)
     if offset > 0 then
        strBuff = objTextStream.Read(offset - 1)
     end if
     if bytes = -1 then		' Get All!
        GetBytes = objTextStream.Read(lngSize)  'ReadAll
     else
        GetBytes = objTextStream.Read(bytes)
     end if
     objTextStream.Close
     set objTextStream = nothing
     set objFSO = nothing
	 on error goto 0
end function

function lngConvert(strTemp)
     lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
end function

function lngConvert2(strTemp)
     lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
end function

function imgDim(flnm, width, height, depth, strImageType)
     dim strPNG 
     dim strGIF
     dim strBMP
     dim strType
     strType = ""
     strImageType = "(unknown)"
     imgDim = False
     strPNG = chr(137) & chr(80) & chr(78)
     strGIF = "GIF"
     strBMP = chr(66) & chr(77)
     strType = GetBytes(flnm, 0, 3)
     if strType = strGIF then				' is GIF
        strImageType = "GIF"
        Width = lngConvert(GetBytes(flnm, 7, 2))
        Height = lngConvert(GetBytes(flnm, 9, 2))
        Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
        imgDim = True
     elseif left(strType, 2) = strBMP then		' is BMP
        strImageType = "BMP"
        Width = lngConvert(GetBytes(flnm, 19, 2))
        Height = lngConvert(GetBytes(flnm, 23, 2))
        Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
        imgDim = True
     elseif strType = strPNG then			' Is PNG
        strImageType = "PNG"
        Width = lngConvert2(GetBytes(flnm, 19, 2))
        Height = lngConvert2(GetBytes(flnm, 23, 2))
        Depth = getBytes(flnm, 25, 2)
        select case asc(right(Depth,1))
           case 0
              Depth = 2 ^ (asc(left(Depth, 1)))
              imgDim = True
           case 2
              Depth = 2 ^ (asc(left(Depth, 1)) * 3)
              imgDim = True
           case 3
              Depth = 2 ^ (asc(left(Depth, 1)))  '8
              imgDim = True
           case 4
              Depth = 2 ^ (asc(left(Depth, 1)) * 2)
              imgDim = True
           case 6
              Depth = 2 ^ (asc(left(Depth, 1)) * 4)
              imgDim = True
           case else
              Depth = -1
        end select
     else
        strBuff = GetBytes(flnm, 0, -1)		' Get all bytes from file
        lngSize = len(strBuff)
        flgFound = 0
        strTarget = chr(255) & chr(216) & chr(255)
        flgFound = instr(strBuff, strTarget)
        if flgFound = 0 then
           exit function
        end if
        strImageType = "JPG"
        lngPos = flgFound + 2
        ExitLoop = false
		
        do while ExitLoop = False and lngPos < lngSize
           do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
              lngPos = lngPos + 1
           loop
           if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
              lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
              lngPos = lngPos + lngMarkerSize  + 1
           else
              ExitLoop = True
           end if
       loop
       '
       if ExitLoop = False then
          Width = -1
          Height = -1
          Depth = -1
       else
          Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
          Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
          Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
          imgDim = True
       end if
     end if
end function
':: end picture dimension functions :::

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Pic - small
' :::::::::::::::::::::::::::::::::::::::::::::::
Function photos_sm(Byval astrType)	
  if chkApp("pictures","USERS") then
	Dim lintPopular
	Dim lstrSQL
	Dim lrsData
	
	Dim lintId
	Dim lstrTitle
	Dim lstrDescription
	Dim ldtPostDate
	Dim lintHit
	Dim ldtSince
	
	fpp_cid = 0
	fpp_sid = 0
	lintPopular = cLng(intShow)
	intDir = cLng(intDir)
	intCount = 1
	displayTxt = "Hits"
	
	if lintPopular = 0 then
	  lintPopular = 5
	end if
	
	lstrSQL = getPicSQL_Sm(astrType)

	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin)
	'on error resume next
	Set lrsData = Server.CreateObject("adodb.recordset")
	lrsData.Open lstrSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	if err.number <> 0 then%>
	<table class="tPlain">
	<tr><td width="100%" valign="top" class="fTitle">
	<b>DB Error getting pics</b><br><%=err.description%>
	</td></tr></table>	
	<%end if
	'on error goto 0
	If lrsData.EOF Then
%>
	<table cellpadding="0" cellspacing="0">
	<tr><td width="100%" valign="top" class="fTitle">
	<b>No Pictures Found!</b></td></tr></table>
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
		Do While Not lrsData.Eof and myCurCnt < lintPopular
		  if not hasAccess(lrsData("SG_READ")) then
			lrsData.MoveNext
		  else
		    myCurCnt = myCurCnt + 1
			fpp_sid = lrsData("CATEGORY")
			fpp_cid = lrsData("PARENT_ID")
            stTURL = lrsData("TURL")
            stURL = lrsData("URL")
			lintId = lrsData("PIC_ID")
			lstrTitle = lrsData("TITLE")
			lstrDescription = lrsData("DESCRIPTION")
			If len(lstrDescription) > lintDescriptionLen then 
				lstrSummary = Left(lstrDescription, lintDescriptionLen) & "..."
			End If
			ldtPostDate = ChkDate2(lrsData("POST_DATE"))
			ldtSince = getDateDiff(strCurDateString,lrsData("POST_DATE"))
			lintHit = lrsData("HIT")

			pic_DisplaySmall lintID, lstrTitle, lstrDescription, ldtPostDate, lintHit, ldtSince, stURL, stTURL
			
			lrsData.MoveNext
			if intDir = 1 and intCount < 3 and not lrsData.eof then
			  intCount = intCount + 1
			  Response.write("</td><td style=""background:url(themes/" & strTheme & "/line.gif);width:1px;""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td>")
			  response.write("<td width=""" & intWid & """ valign=""top"">")
			end if
			
			if intDir <> 1 and not lrsData.eof Then
				Response.write "</td></tr><tr><td align=""center"" style=""height:1px;""><img src=""themes/" & strTheme & "/line.gif"" height=""1"" width=""98%"" alt="""" /></td></tr><tr><td>"
			end if
		  end if
		Loop
		%></td></tr></table><%
	End If
	lrsData.Close
	Set lrsData = Nothing
	
	spThemeBlock1_close(intSkin)
  end if
	intShow = 0
	intLen = 0
	intDir = 0
End Function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Pic - large
' :::::::::::::::::::::::::::::::::::::::::::::::

Function photos_lg(Byval astrType)	
  if chkApp("pictures","USERS") then	
	Dim lintPopular
	Dim lintWidth
	Dim lintColumns, lintMaxColumns
	Dim lintPopularIndex
	Dim lstrSQL
	Dim lrsData

	Dim lintID
	Dim lstrTitle
	Dim lstrDescription
	Dim lstrDescriptionLen
	Dim ldtPostDate
	Dim ldtSince
	Dim lintHit
	Dim strURL
	Dim strTURL
	
	fpp_cid = 0
	fpp_sid = 0
	lintDescriptionLen = cLng(intLen)
	lintPopular = cLng(intShow)
	lintMaxColumns = cLng(numInRow)
	intCount = 0
	lintColumn = 0
	dispTxt = "Hits"
	
	if lintDescriptionLen = 0 then
	  'lintDescriptionLen = 250
	end if
	if lintMaxColumns = 0 then
	  lintMaxColumns = 3
	end if
	lintPopular = cLng(intShow)'This is how many items to display
	if lintPopular = 0 then
	  lintPopular = 6  
	end if
	lintWidth = Int(100 / lintMaxColumns)
	
	lstrSQL = getPicSQL_Sm(astrType)

	'spThemeMM = "home"	
	spThemeBlock1_open(intSkin) 
	%><table width="100%" border="0" cellspacing="3" cellpadding="0"><tr><%
	Set lrsData = Server.CreateObject("adodb.recordset")
'consistent data.open
	lrsData.Open lstrSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	If lrsData.EOF Then
%>
	<td width="100%" valign="top" class="fTitle">No Pictures Found!</td>
<%	
	Else
	    myCurCnt = 0
		Do While Not lrsData.Eof and myCurCnt < lintPopular
		  if not hasAccess(lrsData("SG_READ")) then
			lrsData.MoveNext
		  else
		    myCurCnt = myCurCnt + 1
			intCount = intCount + 1
			lintColumn = lintColumn + 1
			fpp_sid = lrsData("CATEGORY")
			fpp_cid = lrsData("PARENT_ID")
            stTURL = lrsData("TURL")
            stURL = lrsData("URL")
			lintID = lrsData("PIC_ID")
			lstrTitle = lrsData("TITLE")
			lstrDescription = lrsData("DESCRIPTION")
			ldtPostDate = ChkDate2(lrsData("POST_DATE"))
			ldtSince = getDateDiff(strCurDateString,lrsData("POST_DATE"))
			lintHit = lrsData("HIT")
			%><td width="<%=lintWidth%>%" valign="top"><%
			pic_DisplayLarge lintID, lstrTitle, lstrDescription, ldtPostDate, lintHit, ldtSince, stURL, stTURL
			%></td><%
			lrsData.MoveNext
			  If intCount = lintMaxColumns and not lrsData.eof Then
				intCount = 0
				Response.write "</tr><tr><td colspan=""5"" align=""center"" style=""height:1px;""><img src=""themes/" & strTheme & "/line.gif"" height=""1"" width=""98%"" alt="""" /></td></tr><tr>"
			  elseif intCount < lintMaxColumns and not lrsData.eof then
					Response.write "<td style=""background:url(themes/" & strTheme & "/line.gif);width:1px;""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td>"
			  elseif intCount < lintMaxColumns and lrsData.eof then
			    do until intCount = lintMaxColumns
					Response.write "<td style=""background:url(themes/" & strTheme & "/line.gif);width:1px;""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td>"
					Response.write "<td>&nbsp;</td>"
					intCount = intCount + 1
				loop
			  End If
		  end if
		Loop
			 if intCount < lintMaxColumns and lrsData.eof then
			    do until intCount = lintMaxColumns
					intCount = intCount + 1
					Response.write "<td>&nbsp;</td>"
					Response.write "<td style=""background:url(themes/" & strTheme & "/line.gif);width:1px;""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td>"
				loop
			  End If
	End If

	lrsData.Close
	Set lrsData = Nothing
	response.Write("</tr></table>")
	spThemeBlock1_close(intSkin)
 end if
End Function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Pic - common
' :::::::::::::::::::::::::::::::::::::::::::::::
function getPicSQL_Sm(typ)
    tSQL = ""
	Select Case LCase(typ)
		Case "top"
			spThemeTitle= "Top Viewed Pics"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
					tSQL = tSQL & " LIMIT " & lintPopular
				Case else
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
			End Select
			
		Case "new"
			spThemeTitle= "Newest Pictures"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY POST_DATE DESC LIMIT " & lintPopular
				Case else
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY POST_DATE DESC"
			End Select
			
		Case "rated"
			spThemeTitle= "Toprated Pictures"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY RATING DESC, POST_DATE DESC"
					tSQL = tSQL & " LIMIT " & lintPopular
				Case else
					tSQL = "SELECT PIC.PIC_ID, PIC.TITLE, PIC.DESCRIPTION, PIC.PARENT_ID, PIC.CATEGORY, PIC.URL, PIC.TURL, PIC.POST_DATE, PIC.OWNER, PIC.HIT as HITS, PIC.RATING as HIT, PIC_CATEGORIES.CG_READ, PIC_SUBCATEGORIES.SG_READ "
					tSQL = tSQL & "FROM (PIC INNER JOIN PIC_SUBCATEGORIES ON PIC.CATEGORY = PIC_SUBCATEGORIES.SUBCAT_ID) INNER JOIN PIC_CATEGORIES ON PIC_SUBCATEGORIES.CAT_ID = PIC_CATEGORIES.CAT_ID "
					tSQL = tSQL & "WHERE (((PIC.ACTIVE)=1) AND ((PIC.OWNER)='0')) "
					tSQL = tSQL & " ORDER BY RATING DESC, POST_DATE DESC"
					
					'tSQL = "SELECT TOP 20 PIC_ID, TITLE, DESCRIPTION, PARENT_ID, CATEGORY, URL, TURL, POST_DATE, HIT as HITS, RATING as HIT, ACTIVE FROM PIC WHERE ACTIVE = 1 AND OWNER = '0'"
			End Select
			displayTxt = "Rated"
			
		Case "featured"
			spThemeTitle= "Featured Pictures"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " AND FEATURED = TRUE"
					tSQL = tSQL & " ORDER BY POST_DATE DESC LIMIT " & lintPopular
				Case else
'featured fix
'					tSQL = "SELECT TOP " & lintPopular & " PIC_ID, TITLE, DESCRIPTION, PARENT_ID, CATEGORY, URL, TURL, POST_DATE, HIT, ACTIVE FROM PIC WHERE ACTIVE = 1 AND OWNER = '0' AND FEATURED = TRUE ORDER BY POST_DATE DESC"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " AND FEATURED <> 0"
					tSQL = tSQL & " ORDER BY POST_DATE DESC"
			End Select
			
		Case "rand"
			spThemeTitle= "Random Pics"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY Rand() LIMIT " & lintPopular
' sqlserver random fix
' in SQLServer, on windows2000 or higher, you use:
' SELECT TOP 1 column FROM table ORDER BY NEWID()
				Case "sqlserver"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY NEWID()"
				Case else
					Randomize()
' access random fix (change PintRandomNumber to lintRandomNumber for consistency.
					lintRandomNumber = Int(1000*Rnd) + 1
' access random fix (changed "ORDER BY 7" to 'ORDER BY 11")
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY Rnd(" & -1*(lintRandomNumber) & "*PIC_ID)"
			End Select
			
	  	Case Else
			spThemeTitle = "Top Viewed Pics"
			Select Case strDBType
				Case "mysql"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
					tSQL = tSQL & " LIMIT " & lintPopular
				Case else
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"
			End Select
	End Select
	getPicSQL_Sm = tSQL
end function

function sql_selectPicSm()
  tS = "SELECT PIC.*, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_FULL, PIC_SUBCATEGORIES.SG_READ, PIC_SUBCATEGORIES.SG_FULL "
  tS = tS & "FROM (PIC INNER JOIN PIC_SUBCATEGORIES ON PIC.CATEGORY = PIC_SUBCATEGORIES.SUBCAT_ID) INNER JOIN PIC_CATEGORIES ON PIC_SUBCATEGORIES.CAT_ID = PIC_CATEGORIES.CAT_ID "
  tS = tS & "WHERE (((PIC.ACTIVE)=1) AND ((PIC.OWNER)='0')) "
  'tS = tS & "ORDER BY PIC.HIT DESC;"
  sql_selectPicSm = tS
end function

Sub pic_DisplaySmall(aintID, astrTitle, astrSummary, adtPostDate, aintHit, adtSince, astrURL, astrTURL)
  tImg = ""
  showImgDet = ""
  bLocal = false
  if bFso then
    if instr(astrURL,"_rs") > 0 then
	  tImg = replace(astrURL,"_rs","")
	  bLocal = true
	end if
    if instr(astrURL,"_sm") > 0 then
	  tImg = replace(astrURL,"_sm","")
	  bLocal = true
	end if
    if instr(astrTURL,"_sm") > 0 then
	  tImg = replace(astrTURL,"_sm","")
	  bLocal = true
	end if
    if instr(astrTURL,"_rs") > 0 then
	  tImg = replace(astrTURL,"_rs","")
	  bLocal = true
	end if
	'if lcase(strdbntusername) = "skydogg" then
    '  response.Write("<br>" & bLocal & "<br>")
	'end if
	if bLocal then
	t1 = instrrev(tImg,"/")-1
	t1a = left(tImg,t1)
	t2 = instrrev(t1a,"/")-1
	t2a = left(t1a,t2)
	fpp_cid = right(t2a,len(t2a)-instrrev(t2a,"/"))
	fpp_sid = right(t1a,len(t1a)-instrrev(t1a,"/"))
	
	tImg = right(tImg,len(tImg)-instrrev(tImg,"/"))
	fImgPath = server.MapPath(galleryDir & fpp_cid & "/" & fpp_sid & "/" & tImg)
	
    Set obFSO = CreateObject("Scripting.FileSystemObject")
    'response.Write("<br>" & fImgPath & "<br>")
	if obFSO.FileExists(fImgPath) then
	'if lcase(strdbntusername) = "skydogg" then
	'  response.Write("exists")
	'end if
    Set obF = obFSO.GetFile(fImgPath)
	   if obF.Size < 1000 then
	     iSize = obF.Size & " bytes"
	   else
	     iSize = round(obF.Size/1000,2) & " kb"
	   end if
       if imgDim(obF.Path, w, h, c, strType) = true then
          'response.write w & " x " & h & " " & c & " colors"
          showImgDet = "<br><span class=""fSmall"">(<i>" & w & "</i> x <i>" & h & "</i> - " & iSize & ")</span>"
       end if
	 Set obF = nothing
	 end if
	 Set obFSO = nothing
	 end if
  end if
  
  if trim(astrTURL) = "" then
    astrTURL = astrURL
  end if
  if instr(astrTURL,"_sm") > 0 then
     stImg = "<img src=""" & astrTURL & """ border=""0"" alt=""Image"" title=""Click to view picture"" />"
  else
     stImg = "<img src=""" & astrURL & """ border=""0"" width=""120"" alt=""image"" title=""Click to view picture"" />"
  end if
  if intDir = 1 then
    hgt = " height=""160"""
	wdth = "100%"
  else
    hgt = ""
	wdth = "170"
  end if
			'dispTxt = "Rated"
%>
	<table width="<%= wdth %>" cellspacing="3" cellpadding="3" border="0">
	<tr><td valign="top"><b><a href="pic.asp?cmd=6&amp;cid=<%=aintID%>"><span class="fSubTitle"><%=astrTitle%></span></a></b><% if adtSince <= 7 then response.write icon(icnNew1,"New Item","","","align=""middle""") %><%= showImgDet %></td></tr>
	<tr><td<%= hgt %> align="center"><a href="pic.asp?cmd=6&amp;cid=<%=aintID%>"><%= stImg %></a></td></tr>
	<tr><td class="fNorm"><i>(<%=displayTxt%>: <%=aintHit%>)</i></td></tr></table>
<%
End Sub

Sub pic_DisplayLarge(aintID, astrTitle, astrSummary, adtPostDate, aintHit, adtSince, astrURL, astrTURL)
  tImg = ""
  showImgDet = ""
  bLocal = false
  if bFso then
    if instr(astrURL,"_rs") > 0 then
	  tImg = replace(astrURL,"_rs","")
	  bLocal = true
	end if
    if instr(astrTURL,"_sm") > 0 then
	  tImg = replace(astrTURL,"_sm","")
	  bLocal = true
	end if
    if instr(astrTURL,"_rs") > 0 then
	  tImg = replace(astrTURL,"_rs","")
	  bLocal = true
	end if
	if bLocal then
	t1 = instrrev(tImg,"/")-1
	t1a = left(tImg,t1)
	t2 = instrrev(t1a,"/")-1
	t2a = left(t1a,t2)
	fpp_cid = right(t2a,len(t2a)-instrrev(t2a,"/"))
	fpp_sid = right(t1a,len(t1a)-instrrev(t1a,"/"))
	
	tImg = right(tImg,len(tImg)-instrrev(tImg,"/"))
	fImgPath = server.MapPath(galleryDir & fpp_cid & "/" & fpp_sid & "/" & tImg)
    Set obFSO = CreateObject("Scripting.FileSystemObject")
	if obFSO.FileExists(fImgPath) then
      Set obF = obFSO.GetFile(fImgPath)
      'response.Write("<br>" & obF.Path & "<br>")
	   if obF.Size < 1000 then
	     iSize = obF.Size & " bytes"
	   else
	     iSize = round(obF.Size/1000,2) & " kb"
	   end if
       if imgDim(obF.Path, w, h, c, strType) = true then
          'response.write w & " x " & h & " " & c & " colors"
          showImgDet = "<br><span class=""fSmall"">(<i>" & w & "</i> x <i>" & h & "</i> - " & iSize & ")</span>"
       end if
	  Set obF = nothing
	 end if
	 Set obFSO = nothing
	 end if
  end if
  if not trim(astrTURL) <> "" then
    astrTURL = astrURL
  end if
  if instr(astrTURL,"_sm") > 0 then
     stImg = "<img src=""" & astrTURL & """ border=""0"" alt=""Image"" title=""Click to view picture"" />"
  else
     stImg = "<img src=""" & astrURL & """ border=""0"" width=""120"" alt=""image"" title=""Click to view picture"" />"
  end if
%>
	<table width="100%" cellspacing="3" cellpadding="3" border="0">
	<tr><td width="100%" valign="top"><b><a href="pic.asp?cmd=6&amp;cid=<%=aintID%>"><span class="fSubTitle"><%=astrTitle%></span></a></b><% if adtSince <= 7 then response.write icon(icnNew1,"New Item","","","align=""middle""") %><%= showImgDet %></td></tr>
	<tr><td width="100%" height="160" align="center"><a href="pic.asp?cmd=6&amp;cid=<%=aintID%>"><%= stImg %></a></td></tr>
	<tr><td width="100%" class="fNorm"><i>(Hits: <%=aintHit%>)</i><br><%=astrSummary%></td></tr></table>
<%
End Sub

function cntNewPictures()
  aCnt = getCount("PIC_ID","PIC","POST_DATE >= '" & Session(strUniqueID & "last_here_date") & "' AND ACTIVE = 1")
  If aCnt > 0 Then
    aImg = "&nbsp;" & icon(icnNew1,"New Item","","","align=""middle""")
  else
    aImg = ""
  end if
  cntNewPictures = aImg
end function

sub pictures_PendTaskCnt()
  ' Pending Pictures count
  PTcnt = PTcnt + getCount("PIC_ID","PIC","ACTIVE=0 OR BADLINK<>0") 
end sub

sub pictures_adminPndLink()
  if chkApp("pictures","USERS") then
    cntPI1 = getCount("PIC_ID","PIC","ACTIVE=0")
	cntPI2 = getCount("PIC_ID","PIC","BADLINK <> 0")
	If cntPI1 > 0 or cntPI2 > 0 then
	  Response.Write "<li>"
	  If cntPI1 > 0 then
	    Response.Write "<a href=""admin_pic_main.asp""><b>"
	    Response.Write cntPI1 & "&nbsp;" & txtNwPicApprv
	    Response.Write "</b></a>"
	    If cntPI2 > 0 then
	      Response.Write "<br />"
	    End IF
	  End IF
	  If cntPI2 > 0 then
	    Response.Write "<a href=""admin_pic_main.asp?cmd=40""><b>"
	    response.Write cntPI2 & "&nbsp;" & txtBdPicts
	    Response.Write "</b></a>"
	  End IF
	  Response.Write "</li>"
	End IF
  End IF
end sub

sub pictures_SiteSearch()
  '################# Picture Search Routine #############
  If chkApp("pictures","USERS") Then
    strSQL = "SELECT * FROM PIC WHERE TITLE LIKE '%" & search & "%' OR KEYWORD LIKE '%" & search & "%' OR DESCRIPTION LIKE '%" & search & "%' AND ACTIVE=1 ORDER BY PIC_ID DESC"

	Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	reccount = objPagingRS.recordcount
	%>
	<center><span class="fSubTitle"><b><%= txtPics %> - <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>	
	<br />
	<% 
	If reccount > 0 Then
	  %>
	  <center><a href="pic.asp?cmd=7&search=<%=search%>&submit1=Search&num=<%=show%>">
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

':: picture admin menu
sub pictureConfigMenu(typ)
 if bFso then
    mnu.menuName = "pictures_admin"
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
    <div class="tCellAlt1" onMouseOver="this.className='tCellHover';" onMouseOut="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block10<%= typ %>','2');"><span style="margin: 2px;"><img name="block10<%= typ %>Img" id="block10<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b>Pictures</b></div>
      <div class="menu" id="block10<%= typ %>" style="display: <%= cls %>; text-align:left;">
	  	<a href="admin_pic_main.asp"><%= icn_bar %>Approve Pictures (<%= getCount("pic_ID","pic","ACTIVE=0") %>)<br></a>
		<a href="admin_pic_admin.asp?cmd=40"><%= icn_bar %>Bad Links (<%= getCount("pic_ID","pic","BADLINK <> 0") %>)<br></a>
		<a href="admin_pic_admin.asp"><%= icn_bar %>Create Category<br></a>
		<a href="admin_pic_admin.asp?cmd=2"><%= icn_bar %>Edit Category<br></a>
		<a href="admin_pic_admin.asp?cmd=4"><%= icn_bar %>Delete Category<br></a>
		<a href="admin_pic_admin.asp?cmd=1"><%= icn_bar %>Create SubCategory<br></a>
		<a href="admin_pic_admin.asp?cmd=5"><%= icn_bar %>Edit SubCategory<br></a>
		<a href="admin_pic_admin.asp?cmd=8"><%= icn_bar %>Delete SubCategory<br></a>
		<a href="admin_pic_admin.asp?cmd=10"><%= icn_bar %>Edit Picture<br></a>
		<a href="admin_pic_admin.asp?cmd=20"><%= icn_bar %>Delete Picture<br></a>
		<a href="admin_pic_admin.asp?cmd=30"><%= icn_bar %>Browse Picture<br></a>
		</div>
 <%
 end if
end sub
%>
