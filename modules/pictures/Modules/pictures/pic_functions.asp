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
sub displaypic()
  tImg = ""
  showImgDet = ""
    'response.Write("<br>" & strURL & "<br>")
  if bFso then
    if instr(strURL,"_rs") > 0 then
	tImg = replace(strURL,"_rs","")
	tImg = right(tImg,len(tImg)-instrrev(tImg,"/"))
	fImgPath = server.MapPath(galleryDir & cat_id & "/" & sub_id & "/" & tImg)
    Set obFSO = CreateObject("Scripting.FileSystemObject")
    'response.Write("<br>" & fImgPath & "<br>")
	if obFSO.FileExists(fImgPath) then
    Set obF = obFSO.GetFile(fImgPath)
	   if obF.Size < 1000 then
	     iSize = obF.Size & " bytes"
	   else
	     iSize = round(obF.Size/1000,2) & " kb"
	   end if
       if imgDim(obF.Path, w, h, c, strType) = true then
          'response.write w & " x " & h & " " & c & " colors"
          showImgDet = "<br><span class=""fSmall"">(<i>" & h & "</i> x <i>" & w & "</i> - " & iSize & ")</span>"
       end if
	 Set obF = nothing
	 end if
	 Set obFSO = nothing
	 end if
  end if
	  'response.Write("bFso")
  if instr(strTURL,"_sm") > 0 then
     stImg = "<img src=""" & strTURL & """ border=""0"" alt=""Image"" title=""Click for full size picture"" />"
  elseif len(strTURL & "x") > 1 then
     stImg = "<img src=""" & strTURL & """ border=""0"" width=""120"" alt=""Image"" title=""Click for full size picture"" />"
  else
     stImg = "<img src=""" & strURL & """ border=""0"" width=""120"" alt=""Image"" title=""Image"" />"
  end if
  
  spThemeBlock4_open()
  if sMode = 1 then
	pClass = "fAlert"
  else
	pClass = "fSmall"
  end if
%>
<table border="0" width="100%" cellspacing="1" cellpadding="6" align="center"> 
  <tr>
    <td width="100%">
      <a href="pic.asp?cmd=6&amp;cid=<%=intpicID%>">
	  <span class="fSubTitle"><b><%=strpicTitle%></b></span></a><% if dateSince <= 7 then response.write "&nbsp;" & icon(icnNew1,"New Item","","","align=""middle""") %><%= showImgDet %>
    </td>
  </tr> 
  <tr>
    <td height="130" align="center" valign="middle">
	<a href="pic.asp?cmd=6&amp;cid=<%=intpicID%>"><%= stImg %></a>
	<!-- <a href="<%= strURL %>" rel="lightbox"><%= stImg %></a> -->
    </td>
  </tr>  
  <tr>
    <td>
	  <span class="fSmall">Posted by: </span><span class="<%= pClass %>"><b><%=strPoster%></b></span><br />
	  <span class="fSmall">Added: <%=strPostDate%><br>Hits: <%=intHit%><% GetRating(intpicID) %></span><br>
    </td>
  </tr>  
</table>
<%
  spThemeBlock4_close()
end sub

function GetRating(picID)
	strSQL = "SELECT VOTES, RATING FROM PIC WHERE PIC_ID = " & picID
	set rspicRating = server.CreateObject("adodb.recordset")
	rspicRating.Open strSQL, my_Conn
		
	dim intVotes
	dim intRating
	intVotes = rspicRating("VOTES")
	intRating = rspicRating("RATING")
	rspicRating.Close
	set rspicRating = nothing
												
	if intVotes > 0 then
		'intRating = Round(intRating/intVotes)
		Response.Write("&nbsp;&nbsp;&nbsp;Rating: " & intRating & "&nbsp;&nbsp;&nbsp;Votes: " & intVotes)
	else
		Response.Write("&nbsp;")
	end if

end function

function getCommentCount(DLID)
	strSQL = "SELECT count(*) as Comments FROM " & strPicTablePrefix & "PIC_RATING WHERE COMMENTS NOT LIKE ' ' AND PIC = " & intpicID
	dim rsDLComments
	set rsDLComments = server.CreateObject("adodb.recordset")
	rsDLComments.Open strSQL, my_Conn
		
	dim intVotes
	dim intRating
	if not rsDLComments.eof then
		Comments = rsDLComments("Comments")
	end if
	rsDLComments.Close
	set rsDLComments = nothing
												
	if Comments > 0 then
	%>
      | <a href="pic_comments.asp?id=<%=intpicID%>"> Read Comments (<%=Comments%>)</a> 
	<%
	end if

end function

function GetNewPIC(daysShown)
	dim i
	for i = 0 to daysShown - 1
		curDate = dateadd("d",-i,strCurDateAdjust)
		strSQL = "SELECT count(PIC_ID) as PICCOUNT FROM " & strPicTablePrefix & "PIC WHERE POST_DATE LIKE '" & left(DateToStr(curDate),8) & "%' AND ACTIVE = 1"
		set rsDay = server.CreateObject("adodb.recordset")
		rsDay.Open strSQL, my_Conn
		%>
		  <div class="tPlain" style="padding: 4px;">
		    <span style="width: 50px; text-align: right;">&#149;</span>
		    <span style="width: 300px;">&nbsp;<a href="pic.asp?cmd=3&amp;daysago=<%= i %>"><span class="fNorm"><%= formatdatetime(curDate,1) %></span></a></span>
		    <span style="width: 50px; text-align: right;">&nbsp;(<%=rsDay("PICCOUNT")%>)</span>
		  </div>
		<%
		rsDay.Close
		set rsDay = nothing
	next
end function

sub showtoprated()
  dim intTop, cnt
  cnt = 1
  arg2 = txtTopRPics & "|pic.asp?cmd=5"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intTop = "" or not isnumeric(intTop) then intTop = 10
  spThemeTitle= "<b>" & txtTopRPics &"</b>"
  spThemeBlock1_open(intSkin)
  'response.Write("<table><tr><td>")
  'strSQL = "SELECT * FROM PIC WHERE ACTIVE = 1 AND VOTES <> 0 ORDER BY ROUND(RATING/VOTES, 0) DESC, VOTES DESC"
	tSQL = "SELECT PIC.PIC_ID, PIC.TITLE, PIC.POSTER, PIC.DESCRIPTION, PIC.PARENT_ID, PIC.CATEGORY, PIC.URL, PIC.TURL, PIC.POST_DATE, PIC.OWNER, PIC.HIT as HITS, PIC.RATING as HIT, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_FULL, PIC_SUBCATEGORIES.SG_READ "
	tSQL = tSQL & "FROM (PIC INNER JOIN PIC_SUBCATEGORIES ON PIC.CATEGORY = PIC_SUBCATEGORIES.SUBCAT_ID) INNER JOIN PIC_CATEGORIES ON PIC_SUBCATEGORIES.CAT_ID = PIC_CATEGORIES.CAT_ID "
	tSQL = tSQL & "WHERE (((PIC.ACTIVE)=1) AND ((PIC.OWNER)='0') AND ((PIC.VOTES)<> 0)) "
	tSQL = tSQL & " ORDER BY RATING DESC, POST_DATE DESC"

  dim rsPopular
  set rsPopular = server.CreateObject("adodb.recordset")
  rsPopular.Open tSQL, my_Conn
  If rsPopular.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No items found!</span>"
  else	
	  rCount = 0
	  response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	  response.Write("<tr>")
    Do While Not rsPopular.EOF and cnt <= intTop
	  if not hasAccess(rsPopular("SG_READ")) and not hasAccess(rsPopular("CG_FULL")) and not bAppFull then
		rsPopular.MoveNext
	  else
		ShareCheck = "no"	
		strPicOwner = rsPopular("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  cat_id = rsPopular("PARENT_ID")
		  sub_id = rsPopular("CATEGORY")
		  strURL = rsPopular("URL")
		  strPoster = rsPopular("POSTER")
		  strpicTitle = rsPopular("TITLE")
		  strDESC = rsPopular("DESCRIPTION")
		  intpicID = rsPopular("PIC_ID")
		  intHit = rsPopular("HIT")		
		  strPostDate = ChkDate2(rsPopular("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,rsPopular("POST_DATE"))
		  if not len(trim(rsPopular("TURL"))) = 7 or rsPopular("TURL") = "" then
                strTURL = rsPopular("TURL")
 	  	  end if
		  response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		  call displaypic()
		  response.Write("</td>")
		  rCount = rCount + 1
		  cnt = cnt + 1
		end if
		rsPopular.MoveNext
		  if rsPopular.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""" & 100/numInRow & "%"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or rsPopular.eof then
		    response.Write("</tr>")
		  end if
		  if rCount = numInRow and not rsPopular.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
	  end if
    Loop
	response.Write("</table>")
  end if

  rsPopular.Close
  Set rsPopular = Nothing
  'response.Write("</td></tr></table>")
  spThemeBlock1_close(intSkin)
end sub

sub showpopular()
  dim intPopular, cnt
  cnt = 1
  intPopular = 10
  arg2 = txtPopPics & "|pic.asp?cmd=4"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intPopular = "" or not isnumeric(intpopular) then intPopular = 10
  spThemeTitle= "<b>Popular pictures by hit count</b>"
  spThemeBlock1_open(intSkin)

  'strSQL = "SELECT * FROM PIC WHERE ACTIVE = 1 ORDER BY HIT DESC"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY HIT DESC, POST_DATE DESC"

  dim rsPopular
  'set rsPopular = server.CreateObject("adodb.recordset")
  set rsPopular = my_Conn.execute(tSQL)
  If rsPopular.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No items found!</span>"
  else	
		
	  rCount = 0
	  response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	  response.Write("<tr>")
    Do While Not rsPopular.EOF and cnt <= intPopular
	  if not hasAccess(rsPopular("SG_READ")) and not hasAccess(rsPopular("CG_FULL")) and not bAppFull then
		rsPopular.MoveNext
	  else
		ShareCheck = "no"	
		strPicOwner = rsPopular("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  cat_id = rsPopular("PARENT_ID")
		  sub_id = rsPopular("CATEGORY")
		  strURL = rsPopular("URL")
		  strPoster = rsPopular("POSTER")
		  strpicTitle = rsPopular("TITLE")
		  strDESC = rsPopular("DESCRIPTION")
		  intpicID = rsPopular("PIC_ID")
		  intHit = rsPopular("HIT")		
		  strPostDate = ChkDate2(rsPopular("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,rsPopular("POST_DATE"))
		  if not len(trim(rsPopular("TURL"))) = 7 or rsPopular("TURL") = "" then
                strTURL = rsPopular("TURL")
 	  	  end if
		  response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		  call displaypic()
		  response.Write("</td>")
		  rCount = rCount + 1
		  cnt = cnt + 1
		end if
	rsPopular.MoveNext
		  if rsPopular.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""" & 100/numInRow & "%"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or rsPopular.eof then
		    response.Write("</tr>")
		  end if
		  if rCount = numInRow and not rsPopular.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
	  end if
    Loop
	response.Write("</table>")
  end if

  'rsPopular.Close
  Set rsPopular = Nothing
  spThemeBlock1_close(intSkin)
end sub

sub shownew()
  dim intPopular, cnt
  cnt = 1
  intPopular = 10
  arg2 = txtNewPics & "|pic.asp?cmd=3"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

  if intPopular = "" or not isnumeric(intpopular) then intPopular = 10
  if numInRow = 3 then
    intPopular = 15
  elseif numInRow = 2 then
    intPopular = 10
  else
    intPopular = 10
  end if
  spThemeTitle= "<b>Newest Pictures</b>"
  spThemeBlock1_open(intSkin)

  'strSQL = "SELECT * FROM PIC WHERE ACTIVE = 1 ORDER BY HIT DESC"
					tSQL = sql_selectPicSm()
					tSQL = tSQL & " ORDER BY POST_DATE DESC, TITLE"

  dim rsPopular
  'set rsPopular = server.CreateObject("adodb.recordset")
  set rsPopular = my_Conn.execute(tSQL)
  If rsPopular.eof Then
	Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No items found!</span>"
  else	
		
	  rCount = 0
	  response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	  response.Write("<tr>")
    Do While Not rsPopular.EOF and cnt <= intPopular
	  if not hasAccess(rsPopular("SG_READ")) and not hasAccess(rsPopular("CG_FULL")) and not bAppFull then
		rsPopular.MoveNext
	  else
		ShareCheck = "no"	
		strPicOwner = rsPopular("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  cat_id = rsPopular("PARENT_ID")
		  sub_id = rsPopular("CATEGORY")
		  strURL = rsPopular("URL")
		  strPoster = rsPopular("POSTER")
		  strpicTitle = rsPopular("TITLE")
		  strDESC = rsPopular("DESCRIPTION")
		  intpicID = rsPopular("PIC_ID")
		  intHit = rsPopular("HIT")		
		  strPostDate = ChkDate2(rsPopular("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,rsPopular("POST_DATE"))
		  if not len(trim(rsPopular("TURL"))) = 7 or rsPopular("TURL") = "" then
                strTURL = rsPopular("TURL")
 	  	  end if
		  response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		  call displaypic()
		  response.Write("</td>")
		  rCount = rCount + 1
		  cnt = cnt + 1
		end if
	rsPopular.MoveNext
		  if rsPopular.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""" & 100/numInRow & "%"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or rsPopular.eof then
		    response.Write("</tr>")
		  end if
		  if rCount = numInRow and not rsPopular.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
	  end if
    Loop
	response.Write("</table>")
  end if

  'rsPopular.Close
  Set rsPopular = Nothing
  spThemeBlock1_close(intSkin)
end sub

sub shownew2()
  dim intDaysAgo, rsDay
  if Request.QueryString("daysago") <> "" or  Request.QueryString("daysago") <> " " then
	if IsNumeric(Request.QueryString("daysago")) = True then
		intDaysAgo = chkString(Request.QueryString("daysago"),"sqlstring")
	else
		closeAndGo("pic.asp")
	end if
  end if
  
  arg2 = txtNewPics & "|pic.asp?cmd=3"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()
  
  if intDaysAgo <> "" then
	curDate = dateadd("d",-intDaysAgo,strCurDateAdjust)
	spThemeTitle= "<b>Total new pictures on " &split(curDate," ")(0)&".</b>"
	spThemeBlock1_open(intSkin)
	strSQL = "SELECT * FROM PIC WHERE POST_DATE LIKE '" & left(DateToStr(curDate),8) & "%' AND ACTIVE = 1 ORDER BY POST_DATE DESC"
	set rsDay = server.CreateObject("adodb.recordset")
	rsDay.Open strSQL, my_Conn
	If rsDay.eof Then
		Response.Write "<span class=""fAlert"" style=""font-weight: bold; text-align: center;"">No items found!</span>"
	else
	  rCount = 0
	  response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	  response.Write("<tr>")		
	  Do While Not rsDay.EOF
		ShareCheck = "no"	
		strPicOwner = rsDay("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  cat_id = rsDay("PARENT_ID")
		  sub_id = rsDay("CATEGORY")
		  strURL = rsDay("URL")
		  strPoster = rsDay("POSTER")
		  strpicTitle = rsDay("TITLE")
		  strDESC = rsDay("DESCRIPTION")
		  intpicID = rsDay("PIC_ID")
		  intHit = rsDay("HIT")		
		  strPostDate = ChkDate2(rsDay("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,rsDay("POST_DATE"))
		  if len(trim(rsDay("TURL"))) > 8 then
                strTURL = rsDay("TURL")
 	  	  end if
		  response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		  call displaypic()
		  response.Write("</td>")
		  rCount = rCount + 1
		  cnt = cnt + 1
		end if
		rsDay.MoveNext
		  if rsDay.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""" & 100/numInRow & "%"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or rsDay.eof then
		    response.Write("</tr>")
		  end if
		  if rCount = numInRow and not rsDay.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
	  Loop
	  response.Write("</table>")
	end if
	rsDay.Close
	Set rsDay = Nothing
	spThemeBlock1_close(intSkin)
  else
	dim intDaysShown
	intDaysShown = 7
	spThemeTitle= "<b>Total new pictures for last " &intDaysShown&" Days.</b>"
	spThemeBlock1_open(intSkin)
 	  GetNewPIC(intDaysShown)
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
  app_MainColumn_top()
  
  If hasAccess(2) and intSubscriptions = 1 Then
	subscription_id = chkIsSubscribed(intAppID,"0","0","0",strUserMemberID)
	  if subscription_id <> 0 then
		spThemeTitle = spThemeTitle & " <a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=11&amp;cid=" & subscription_id & "');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""UnSubscribe from Pictures"" alt='unsubscribe' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & " <a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=9&amp;cmd=3&amp;cid=" & intAppID & "');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""Subscribe to Pictures"" alt=""subscribe"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
  If bAppFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit("pic_pop.asp",14,0,0,"right",2)
  end if
  
  spThemeTitle = spThemeTitle & txtPics
spThemeBlock1_open(intSkin)
  dim rsCategories
  strSql = "select * from PIC_CATEGORIES order by C_ORDER, CAT_NAME"
  set rsCategories = my_Conn.execute(strSql)
  'rsCategories.Open strSql, my_Conn %>
      <table border="0" cellpadding="6" cellspacing="0" width="100%">
        
        <% 
		Do until rsCategories.EOF  %>
			<tr><%
			ColNum = 1 
			Do while ColNum < 3
			  if not rsCategories.EOF then
			  'Response.Write(rsCategories("CG_READ") & "<br>")
			  'Response.Write(hasAccess(trim(rsCategories("CG_READ"))) & "<br>")
			    if hasAccess(trim(rsCategories("CG_READ"))) then
		%>	
          <td align="left" valign="top" width="50%"><img src="images/icons/icon_folder_new_topic.gif" style="margin-left:15px;" border="0" alt="" title="" />
		  <%
  			If hasAccess(trim(rsCategories("CG_FULL"))) or bAppFull Then 
				response.Write("&nbsp;" & modGrpEdit("pic_pop.asp",14,rsCategories("cat_id"),0,"middle",rsCategories("CG_INHERIT")))
  			end if
		  %><a href="pic.asp?cmd=1&amp;cid=<%= rsCategories("cat_id") %>"><b><span class="fTitle"><%= ChkString(rsCategories("cat_name"), "display") %></span></b></a>
		  <br />
		  <%'= hasAccess(rsCategories("CG_READ")) & "<br>" %>
		  <%'= rsCategories("cat_id") %>
				<%
				parent_id = rsCategories("cat_id")
				subsql = "SELECT * FROM PIC_SUBCATEGORIES WHERE CAT_ID=" & parent_id & " ORDER BY C_ORDER, SUBCAT_NAME"
				set rssubcat = my_Conn.Execute(subsql)
				'count = rssubcat.RecordCount
				'count = 1
				do while not rssubcat.EOF 'and count<=3
				  if hasAccess(trim(rssubcat("SG_READ"))) then			
					%><img src="images/icons/icon_bar.gif" style="margin-left:15px;" border="0" hspace="2" alt="" title="" /><img src="images/icons/img_pic.gif" border="0" alt="" title="" /><%
  					If hasAccess(trim(rsSubcat("SG_FULL"))) or hasAccess(trim(rsCategories("CG_FULL"))) or bAppFull Then 
					  Response.Write("&nbsp;" & modGrpEdit("pic_pop.asp",14,rsCategories("cat_id"),rsSubcat("subcat_id"),"middle",rssubcat("SG_INHERIT")))
  					end if
		  			%>
			  <a href="pic.asp?cmd=2&amp;cid=<%=rsCategories("cat_id")%>&amp;sid=<%=rssubcat("subcat_ID")%>"><span class="fNorm"><%= ChkString(rssubcat("subcat_name"), "display") %>
					<%
					sqlcount = "SELECT count(PIC_ID) FROM PIC where category =" & rsSubcat("subcat_id") & " and ACTIVE = 1"
					Set RScounts = my_Conn.Execute(sqlcount)

					rcounts = RScounts(0)
					if rcounts <> 0 then
						Response.Write "&nbsp;(" & rcounts & ")"
					end if			
					%></span></a><br /><%
				  end if
				  rssubcat.movenext
				  %>
			  
			<%  loop %>
          </td>
				<% 
				ColNum = ColNum + 1 
			    end if
				rsCategories.MoveNext
			  else
			    response.Write("<td>&nbsp;</td>")
				ColNum = ColNum + 1 
			  end if 
			Loop
			%>
		</tr>
			<%
		Loop 
		%>
      </table>
<% spThemeBlock1_close(intSkin)
   'rsCategories.close
   set rsCategories = nothing
end sub

sub showcat(cid)
    strSql = "select CAT_NAME, CG_READ, CG_WRITE, CG_FULL, CG_INHERIT, CG_PROPAGATE from PIC_CATEGORIES WHERE CAT_ID=" & cid
    set rsC = my_Conn.Execute(strSql)
	'if hasAccess(trim(rsCategories("CG_READ"))) then
	cat_name = rsC("CAT_NAME")
	inherit = rsC("CG_INHERIT")
	call setPermVars(rsC,1)
	set rsC = nothing
	
    'shoDebugVars()
	
  arg2 = cat_name & "|pic.asp?cmd=1&cid=" & cid
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()
 
 if not bCatRead then
   closeandgo("pic.asp")
 else
  If strUserMemberID > 0 and intSubscriptions = 1 Then
	subscription_id = chkIsSubscribed(intAppID,cid,"0","0",strUserMemberID)
	  if subscription_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=11&amp;cid=" & subscription_id & "');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""UnSubscribe from this category"" alt='unsubscribe' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=9&amp;cmd=1&amp;cid=" & cid & "');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""Subscribe to this category"" alt=""subscribe"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
  If strUserMemberID > 0 and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,cid,"0","0",strUserMemberID)
	  if bookmark_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=10&amp;cid=" & bookmark_id & "');""><img src=""themes/" &  strTheme & "/icons/unbookmark.gif"" title=""Remove Bookmark for this Category"" alt='remove bookmark' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=3&amp;cmd=1&amp;cid=" & cid & "');""><img src=""themes/" &  strTheme & "/icons/bookmark.gif"" title=""Bookmark this Category"" alt=""bookmark"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
  If bCatFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit("pic_pop.asp",14,cid,0,"right",inherit)
  end if
  spThemeTitle = spThemeTitle & "<img src=""images/icons/icon_folder_new_topic.gif"" border=""0"" alt="""" align=""middle"" hspace=""4"" />"
  spThemeTitle = spThemeTitle & "&nbsp;" & cat_name
  spThemeBlock1_open(intSkin)
  %>
      <table border="0" cellpadding="4" cellspacing="0" width="100%">
		<% 
  		sql = "SELECT * FROM PIC_SUBCATEGORIES where CAT_ID=" & cid & " order by C_ORDER, SUBCAT_NAME"
  		set rs = my_Conn.Execute(sql)
		Do while NOT rs.EOF 
		  sSCatRead = rs("SG_READ")
		  sSCatWrite = rs("SG_WRITE")
		  sSCatFull = rs("SG_FULL")
		  if hasAccess(sSCatRead) then
		    %>	
        	<tr>
          	<td align="left" valign="top">
			<img src="images/spacer.gif" border="0" width="15" alt="" />
			<img src="images/icons/icon_bar.gif" border="0" hspace="3" alt="" title="" /><img src="images/icons/img_pic.gif" border="0" alt="" title="" />
			<% 
			If hasAccess(sSCatFull) or bCatFull Then
			  
			  response.Write(modGrpEdit("pic_pop.asp",14,cid,rs("SUBCAT_ID"),"",rs("SG_INHERIT")))
			end if %>
            <a href="pic.asp?cmd=2&amp;cid=<%=cat_id%>&amp;sid=<%=rs("SUBCAT_ID")%>">
			<span class="fSubTitle"><%= ChkString(rs("subcat_name"), "display") %></span></a>
			<% 
			sSQL = "SELECT count(PIC_ID) FROM PIC where category=" & rs("SUBCAT_ID") & " and ACTIVE=1"
			Set RScount = my_Conn.Execute(sSQL)

			rcount = RScount(0)
			if rcount <> 0 then
				Response.Write " (" & rcount & ")"
			end if
			%>
          	</td>
	    	</tr>
			<% 
		  end if
		  rs.MoveNext  
		Loop 
%>
      </table>
<%  spThemeBlock1_close(intSkin)
 end if
end sub

function showsub()  
  Dim iPageSize       
  Dim iPageCount      
  Dim iPageCurrent    
  Dim strOrderBy      
  Dim iRecordsShown  
  Dim ssSQL    
  Dim I      
  Dim cat_name
  Dim sub_name   
  Dim objPagingRS

  'set page size
  iPageSize = 12
  iPageCurrent = 1

  If Request("page") = "" Then
	iPageCurrent = 1
  Else
	iPageCurrent = cLng(Request("page"))
  End If


sSQL = "SELECT PIC_CATEGORIES.CAT_ID, PIC_CATEGORIES.CAT_NAME, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_WRITE, PIC_CATEGORIES.CG_FULL, PIC_CATEGORIES.CG_INHERIT, PIC_CATEGORIES.CG_PROPAGATE, PIC_SUBCATEGORIES.SUBCAT_ID, PIC_SUBCATEGORIES.SUBCAT_NAME, PIC_SUBCATEGORIES.SG_READ, PIC_SUBCATEGORIES.SG_WRITE, PIC_SUBCATEGORIES.SG_FULL, PIC_SUBCATEGORIES.SG_INHERIT "
sSQL = sSQL & "FROM PIC_CATEGORIES INNER JOIN PIC_SUBCATEGORIES ON PIC_CATEGORIES.CAT_ID = PIC_SUBCATEGORIES.CAT_ID "
sSQL = sSQL & "WHERE (((PIC_CATEGORIES.CAT_ID)=" & cat_id & ") AND ((PIC_SUBCATEGORIES.SUBCAT_ID)=" & sub_id & "));"
	
  set rsT = my_Conn.execute(sSQL)
  cat_name = rsT("CAT_NAME")
  sub_name = rsT("SUBCAT_NAME")
  inherit = rsT("SG_INHERIT")
  call setPermVars(rsT,2)
  set rsT = nothing
  
  'shoDebugVars()

if bSCatRead then
  ssSQL = "SELECT * From PIC where Category=" & sub_id & " and ACTIVE = 1"
  ord = request("ord1") & request("ord2")
  select case ord
    case "hDesc"
      ssSQL = ssSQL & " ORDER BY PIC.HIT DESC;"
    case "hAsc"
      ssSQL = ssSQL & " ORDER BY PIC.HIT;"
    case "dDesc"
	  ssSQL = ssSQL & " ORDER BY PIC.POST_DATE DESC;"
    case "dAsc"
	  ssSQL = ssSQL & " ORDER BY PIC.POST_DATE;"
    case "rDesc"
	  ssSQL = ssSQL & " ORDER BY PIC.RATING DESC;"
    case "rAsc"
	  ssSQL = ssSQL & " ORDER BY PIC.RATING;"
    case "tDesc"
	  ssSQL = ssSQL & " ORDER BY PIC.TITLE DESC;"
    case "tAsc"
	  ssSQL = ssSQL & " ORDER BY PIC.TITLE;"
    case else
	  ord = "dDesc"
	  ssSQL = ssSQL & " ORDER BY PIC.PIC_ID DESC;"
  end select
  Set objPagingRS = Server.CreateObject("ADODB.Recordset")
  objPagingRS.PageSize = iPageSize
  objPagingRS.CacheSize = iPageSize
  objPagingRS.Open ssSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  reccount = objPagingRS.recordcount
  iPageCount = objPagingRS.PageCount

  If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  If iPageCurrent < 1 Then iPageCurrent = 1

  arg2 = cat_name & "|pic.asp?cmd=1&amp;cid=" & cat_id
  arg3 = sub_name & "|pic.asp?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & sub_id
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()

if iPageCount = 1 then
  tPgCnt = " - " & iPageCount & " page"
else
  tPgCnt = " - " & iPageCount & " pages"
end if
if reccount = 1 then
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " pictures"&tPgCnt&")"
else
  tTitle = cat_name &": " & sub_name & " ( " & reccount & " pictures"&tPgCnt&")"
end if

  If strUserMemberID > 0 and intSubscriptions = 1 Then
	subscription_id = chkIsSubscribed(intAppID,"0",sub_id,"0",strUserMemberID)
	  if subscription_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=11&amp;cid=" & subscription_id & "');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""UnSubscribe from this Subcategory"" alt='unsubscribe' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=9&amp;cmd=2&amp;cid=" & sub_id & "');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""Subscribe to this Subcategory"" alt=""subscribe"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
  If strUserMemberID > 0 and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,"0",sub_id,"0",strUserMemberID)
	  if bookmark_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=10&amp;cid=" & bookmark_id & "');""><img src=""themes/" &  strTheme & "/icons/unbookmark.gif"" title=""Remove Bookmark for this Subcategory"" alt='remove bookmark' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=3&amp;cmd=2&amp;cid=" & sub_id & "');""><img src=""themes/" &  strTheme & "/icons/bookmark.gif"" title=""Bookmark this Subcategory"" alt=""bookmark"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
  If bSCatFull Then 
	spThemeTitle = spThemeTitle & modGrpEdit("pic_pop.asp",14,cat_id,sub_id,"right",inherit)
  end if  
  
  spThemeTitle = spThemeTitle & "&nbsp;" & tTitle
spThemeBlock1_open(intSkin)
 If iPageCount = 0 Then %>
  <span class="tPlain" style="width:100%;">
    <div class="tTitle" style="text-align:left;"><%=cat_name%>: <%=sub_name%> ( 0 items)</div><br /><br />
    <center><div class="fAlert" style="text-align:center;"><b>No items found!</b></div></center>
  </span>
<%
 Else
  'response.write(ssSQL & "<br />" & objPagingRS("TITLE") & "<br />" & iPageCount & "<br />")
	objPagingRS.AbsolutePage = iPageCurrent
	if iPageCount > 1 then
	  showDaPaging iPageCurrent,iPageCount,0
	end if
	iRecordsShown = 0
	rCount = 0
	response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""3"">")
	response.Write("<tr><td colspan=""" & numInRow & """ align=""center"" valign=""top"">")
	response.Write("<form method=""post"" action=""pic.asp"">")
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
	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF		
		ShareCheck = "no"	
		strPicOwner = objPagingRS("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  strPoster = objPagingRS("POSTER")
		  strpicTitle = objPagingRS("TITLE")
		  strDESC = objPagingRS("DESCRIPTION")
		  intpicID = objPagingRS("PIC_ID")
		  intHit = objPagingRS("HIT")		
		  strPostDate = ChkDate2(objPagingRS("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,objPagingRS("POST_DATE"))
		  strTURL = ""
		  strURL = ""
		  strURL = objPagingRS("URL")
		  if not len(trim(objPagingRS("TURL"))) = 7 or trim(objPagingRS("TURL")) = "" then
            strTURL = objPagingRS("TURL")
		  else
            strTURL = objPagingRS("URL")
 	  	  end if
		  response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		  call displaypic()
		  response.Write("</td>")
		  rCount = rCount + 1
	      iRecordsShown = iRecordsShown + 1
		end if
	    objPagingRS.MoveNext
		  if iRecordsShown < iPageSize and objPagingRS.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""155"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or objPagingRS.eof then
		    response.Write("</tr>")
		  end if
		  if iRecordsShown < iPageSize and rCount = numInRow and not objPagingRS.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
	Loop
	response.Write("</table>")
End If

objPagingRS.Close
Set objPagingRS = Nothing
if iPageCount > 1 then
  showDaPaging iPageCurrent,iPageCount,2
end if

response.Write("<p>&nbsp;</p>")
if bSCatWrite then
%>
<center>
 <a href="pic.asp?cmd=8&amp;cid=<%= cat_id %>&amp;sid=<%= sub_id %>">Submit a Picture</a>
</center>
<p>&nbsp;</p>
<%
end if
  spThemeBlock1_close(intSkin)
else ':: no access so redirect
  closeandgo("pic.asp")
end if
end function

sub shoDebugVars()
  Response.Write("<table width=""100%""><tr><td width=""40%"">")
  response.Write("bAppRead: " & bAppRead & "<br>")
  response.Write("bAppWrite: " & bAppWrite & "<br>")
  response.Write("bAppFull: " & bAppFull & "<br>")
  response.Write("bCatRead: " & bCatRead & "<br>")
  response.Write("bCatWrite: " & bCatWrite & "<br>")
  response.Write("bCatFull: " & bCatFull & "<br>")
  response.Write("bSCatRead: " & bSCatRead & "<br>")
  response.Write("bSCatWrite: " & bSCatWrite & "<br>")
  response.Write("bSCatFull: " & bSCatFull & "<br>")
  Response.Write("</td><td>")
  response.Write("sAppRead: " & sAppRead & "<br>")
  response.Write("sAppWrite: " & sAppWrite & "<br>")
  response.Write("sAppFull: " & sAppFull & "<br>")
  response.Write("sCatRead: " & sCatRead & "<br>")
  response.Write("sCatWrite: " & sCatWrite & "<br>")
  response.Write("sCatFull: " & sCatFull & "<br>")
  response.Write("sSCatRead: " & sSCatRead & "<br>")
  response.Write("sSCatWrite: " & sSCatWrite & "<br>")
  response.Write("sSCatFull: " & sSCatFull & "<br>")
  Response.Write("</td></tr></table>")
end sub

function doSearch()
  search = ChkString(Request("search"), "SQLString")
  show = 10
  if request("num") <> "" then
    show = clng(Request("num"))
  end if
  if show > 0 then
	Dim iPageSize       
	Dim iPageCount      
	Dim iPageCurrent    
	Dim strOrderBy      
	Dim strSQL          
	Dim objPagingConn   
	Dim objPagingRS     
	Dim iRecordsShown   
	Dim I

	intSkin = 1
	iPageSize = show

	If Request("page") = "" Then
		iPageCurrent = 1
	Else
		iPageCurrent = cLng(Request("page"))
	End If
	
	'::::: variable search reutine :::::::::
	if sMode <> 1 then 'search all
  	  strSQL = "select * from PIC where TITLE like '%" & search & "%' or KEYWORD like'%" & search & "%' or DESCRIPTION like '%" & search & "%' and ACTIVE=1 order by HIT DESC, PIC_ID DESC"
	  strSrchTxt = "Search results for"
	else ':: search member submitted
	  'srchMemberID = getMemberId(search)
      strSQL = "select * from PIC where POSTER = '" & search & "' or POSTER like '%" & search & "%' and ACTIVE=1 order by POSTER, HIT DESC, PIC_ID DESC"
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
		Response.Write "<br><div class=""fTitle"" class=""text-align:center;"">"
		Response.Write "<b>Your search for: """ & search & """<br />returned no results!</b></div><br>"
	Else
		objPagingRS.AbsolutePage = iPageCurrent %>
		<center><div class="fTitle"><b>Search results for :&nbsp;</b><span class="fAlert"><b><%=search%></b></span></div>
		<span class="fAlert"> found <%=reccount%><% if reccount = 1 then %> picture<% else %> pictures<% end if %></span></center><%
		if iPageCount > 1 then
	  	showDaPaging iPageCurrent,iPageCount,0
		end if
		iRecordsShown = 0
		rCount = 0
		response.Write("<table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""2"">")
		response.Write("<tr>")
		Do While iRecordsShown < iPageSize And Not objPagingRS.EOF
		  ShareCheck = "no"	
		  strPicOwner = objPagingRS("OWNER")
		  if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then

		  	if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		      ShareCheck = "ok"
		  	end if
		
		    cat_id = objPagingRS("PARENT_ID")
		    sub_id = objPagingRS("CATEGORY")
		    strURL = objPagingRS("URL")
		    strPoster = objPagingRS("POSTER")
	      	'dagar=DateDiff("d", Date, strtodate(objPagingRS("POST_DATE")))+7
		  	strpicTitle = objPagingRS("TITLE")
		  	strDESC = objPagingRS("DESCRIPTION")
		  	intpicID = objPagingRS("PIC_ID")
		  	intHit = objPagingRS("Hit")		
		  	strPostDate = ChkDate2(objPagingRS("POST_DATE"))
		    dateSince = getDateDiff(strCurDateString,objPagingRS("POST_DATE"))
		  	if not len(trim(objPagingRS("TURL"))) = 7 or objPagingRS("TURL") = "" then
                strTURL = objPagingRS("TURL")
 	  	  	end if
		    response.Write("<td width=""" & 100/numInRow & "%"" valign=""top"" align=""center"">")
		    call displaypic()
		    response.Write("</td>")
		    rCount = rCount + 1
		  	'cnt = cnt + 1
		    iRecordsShown = iRecordsShown + 1
		  end if
		  objPagingRS.MoveNext
		  if iRecordsShown < iPageSize and objPagingRS.eof then
		    if rCount < numInRow then
			  for xp = rCount to numInRow
		        response.Write("<td width=""155"">&nbsp;</td>")
			  next
			end if
		  end if
		  if rCount = numInRow or objPagingRS.eof then
		    response.Write("</tr>")
		  end if
		  if iRecordsShown < iPageSize and rCount = numInRow and not objPagingRS.eof then
		    response.Write("<tr>")
			rCount = 0
		  end if
		Loop
	  response.Write("</table>")
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
		'Response.Write "hello world. " & cat_id
  dim strLinkSQL,rsLink

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
	strLSQL = "SELECT PARENT_ID, CATEGORY FROM PIC WHERE PIC_ID = " & cat_id
	set rsT = my_Conn.execute(strLSQL)
	if not rsT.eof then
	  pic_id = cat_id
	  cat_id = rsT("PARENT_ID")
	  sub_id = rsT("CATEGORY")
	else
	  pic_id = 0
	end if
	set rsT = nothing
  else
  end if
  
	sSql = "SELECT PIC_CATEGORIES.CAT_ID, PIC_CATEGORIES.CAT_NAME, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_WRITE, PIC_CATEGORIES.CG_FULL, PIC_CATEGORIES.CG_INHERIT, PIC_CATEGORIES.CG_PROPAGATE, PIC_SUBCATEGORIES.SUBCAT_ID, PIC_SUBCATEGORIES.SUBCAT_NAME, PIC_SUBCATEGORIES.SG_READ, PIC_SUBCATEGORIES.SG_WRITE, PIC_SUBCATEGORIES.SG_FULL, PIC_SUBCATEGORIES.SG_INHERIT, PIC.* "
	sSql = sSql & " FROM (PIC_CATEGORIES INNER JOIN PIC_SUBCATEGORIES ON PIC_CATEGORIES.CAT_ID = PIC_SUBCATEGORIES.CAT_ID) INNER JOIN PIC ON PIC_SUBCATEGORIES.SUBCAT_ID = PIC.CATEGORY"
	sSQL = sSQL & " WHERE (((PIC_CATEGORIES.CAT_ID)=" & cat_id & ") AND ((PIC_SUBCATEGORIES.SUBCAT_ID)=" & sub_id & "));"
	
	'sSQL = sSQL & "FROM PIC_CATEGORIES INNER JOIN PIC_SUBCATEGORIES ON PIC_CATEGORIES.CAT_ID = PIC_SUBCATEGORIES.CAT_ID "
	'sSql = sSql & " WHERE (((PIC.PIC_ID)=" & pic_id & "));"
	
	Set rsItem = Server.CreateObject("ADODB.Recordset")
  	rsItem.PageSize = iPageSize
  	rsItem.CacheSize = iPageSize
  	rsItem.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  	reccount = rsItem.recordcount
  	iPageCount = rsItem.PageCount

  	If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  	If iPageCurrent < 1 Then iPageCurrent = 1
	
	if rsItem.EOF then
	  call showMsgBlock(1,"Picture does not exist.")
	else	
	  cat_name = rsItem("CAT_NAME")
	  sub_name = rsItem("SUBCAT_NAME")
	  c_id = rsItem("CAT_ID")
	  s_id = rsItem("SUBCAT_ID")
	  call setPermVars(rsItem,2)
	  
	  if not bSCatRead then
	    closeandgo("pic.asp")
	  else
		  
		  if bFirst then
		    pc = 0
		    foundit = false
		    do until foundit = true
			  pc = pc + 1
		      if rsItem("PIC_ID") = pic_id then
			    foundit = true
				iPageCurrent = pc
		      end if
			  rsItem.movenext
		    loop
		  end if
		  rsItem.AbsolutePage = iPageCurrent
		
	    ShareCheck = "no"	
		strPicOwner = rsItem("OWNER")
		if strPicOwner="0" or instr(strPicOwner,("|" & strUserMemberID & "|")) then
		  if instr(strPicOwner,("|" & strUserMemberID & "|")) then 
		    ShareCheck = "ok"
		  end if
		  strpicTitle = rsItem("TITLE")

  		  arg2 = cat_name & "|pic.asp?cmd=1&amp;cid=" & c_id
  		  arg3 = sub_name & "|pic.asp?cmd=2&amp;cid=" & c_id & "&amp;sid=" & s_id
  		  arg4 = strpicTitle
  		  'arg4 = strpicTitle & "|pic.asp?cmd=6&amp;cid=" & rsItem("PIC_ID")
  		  arg5 = ""
  		  arg6 = ""
  
  		  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  		  app_MainColumn_top()
		
		  hp = rsItem("FEATURED")
		  if hp = 0 then
	  	    hp = 1
	  	    sTxt = "Make this a Featured Picture'"
	  	    sImg = "Themes/" & strTheme & "/icons/featured.gif"
		  else
	  	    hp = 2
	  	    sTxt = "Remove from Featured Pictures"
	  	    sImg = "Themes/" & strTheme & "/icons/unfeature.gif"
		  end if
		  
		  strURL = rsItem("URL")
		  strPoster = rsItem("POSTER")
		  strDESC = rsItem("DESCRIPTION")
		  intpicID = rsItem("PIC_ID")
		  intHit = rsItem("Hit")
		  strPoster = rsItem("POSTER")
		  strPicCopyright = rsItem("COPYRIGHT")	
		  strPostDate = ChkDate2(rsItem("POST_DATE"))
		  dateSince = getDateDiff(strCurDateString,rsItem("POST_DATE"))
		  if not len(trim(rsItem("TURL"))) = 7 or rsItem("TURL") = "" then
            strTURL = rsItem("TURL")
 	  	  end if
		  
		  'spThemeTitle = spThemeTitle & "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow5('pic_pop.asp?mode=9&amp;cid=" & intpicID & "')""><img border=""0"" src=""images/icons/print.gif"" align=""right"" title=""Printable Version"" alt=""Printable Version"" style=""display:inline;"" hspace=""4"" /></a>"
		  if strUserMemberID > 0 and strEmail = 1 then
		    spThemeTitle = "<a href=""JavaScript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=8&amp;cid=" & intpicID & "')""><img border=""0"" src=""images/icons/icon_email.gif"" align=""right"" title=""Email this picture to a friend"" alt=""Email this picture to a friend"" style=""display:inline;"" hspace=""4"" /></a>"
		  end if
  If strUserMemberID > 0 and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,"0","0",intpicID,strUserMemberID)
	  if bookmark_id <> 0 then
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=10&amp;cid=" & bookmark_id & "');""><img src=""themes/" &  strTheme & "/icons/unbookmark.gif"" title=""Remove Bookmark for this Picture"" alt='remove bookmark' border='0' align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  else
		spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=3&amp;cmd=3&amp;cid=" & intpicID & "');""><img src=""themes/" &  strTheme & "/icons/bookmark.gif"" title=""Bookmark this Picture"" alt=""bookmark"" border=""0"" align=""right"" style=""display:inline;"" hspace=""4""></a>" 
	  end if
  end if
		  If bSCatFull Then 
		    spThemeTitle = spThemeTitle & "<a href=""admin_pic_editpic.asp?id=" & intpicID & """><img border=""0"" src=""images/icons/icon_edit_topic.gif"" align=""right"" title=""Edit Picture"" alt=""Edit Picture"" style=""display:inline;"" hspace=""4"" /></a>"
		  End If
		  If bAppFull Then 
		    spThemeTitle = spThemeTitle & "<a href=""javascript:;"" onclick=""JavaScript:openWindow('pic_pop.asp?mode=" & hp & "&amp;cid=" & intpicID & "')""><img border=""0"" src=""" & sImg & """ align=""right"" title=""" & sTxt & """ alt=""" & sTxt & """ style=""display:inline;"" hspace=""4"" /></a>"
		  End If
		  spThemeTitle = spThemeTitle & "&nbsp;" & strpicTitle
		  spThemeBlock1_open(intSkin)
		  if iPageCount > 1 then
	  	    showDaPaging iPageCurrent,iPageCount,0
		  end if
		  'displaypic()
		  showInfo()
		  if comments <> 0 then
		    spThemeTitle = "Comments:" & "&nbsp;"
		    response.Write("<div style=""margin:10px;width:90%;"">")
		    'response.Write(intpicID)
		    spThemeBlock3_open()
		      GetComments(intpicID)
		    spThemeBlock3_close()
		    response.Write("</div>")
		  end if
	    else
		  spThemeBlock1_open(intSkin) %>
		  <div class="tPlain"><p>&nbsp;</p><p>Private Picture</p><p>&nbsp;</p></div>
		<%
		end if
		spThemeBlock1_close(intSkin)
	  end if
	end if
end sub

function GetComments(picID)
	strSQL = "SELECT COMMENTS, RATE_BY, RATE_DATE, RATING FROM PIC_RATING WHERE COMMENTS NOT LIKE '' AND PIC = " & picid
	'set rs = server.CreateObject("adodb.recordset")
	'rs.Open strSQL, my_Conn
	set rs = my_Conn.execute(strSQL)
		
	dim intVotes
	dim intRating
	dim intRateBy
	do while not rs.eof
	'response.Write("HelloX<br>")
		Comments = rs("Comments")
		intRateBy = rs("RATE_BY")
		%>
<table width="90%" border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse; margin:5px;"><tr>

	<td class="tCellAlt0" width="35%" valign="top" nowrap>
	  <% If bSCatFull Then %>
	  <a href="javascript:openWindow3('pic_pop.asp?mode=7&cid=<%= picid %>&cmd=<%= intRateBy %>');"><img src="images/icons/icon_trashcan.gif" alt="Delete Comment" title="Delete Comment" width="12" height="12" border="0" /></a>
	  <% End If %>
	  By: <b><%=getMemberName(rs("RATE_BY"))%></b>
	  <br>On: <%=ChkDate2(rs("RATE_DATE"))%><%'= rs("RATE_DATE")%>
	</td>
	
    <td width="65%">
      
		<%=formatstr(Comments)%><br></font>
    </td></tr>
</table>
		<%
		rs.MoveNext
	loop
	'rs.Close
	set rs = nothing
end function

sub showInfo()
	strSQL1 = "SELECT count(*) as Comments FROM PIC_RATING WHERE COMMENTS NOT LIKE ' ' AND PIC = " & intpicID
	set rspicComments = my_Conn.execute(strSQL1)
	if not rspicComments.eof then
		Comments = rspicComments("Comments")
	else
	    Comments = 0
	end if
	set rspicComments = nothing
  if instr(strURL,"_rs.") > 0 then
     'stImg = "<img src=""" & replace(strURL,"_rs","") & """ border=""0"" alt=""Image"" title=""Local Click to view full sized picture"" />"
     stImg = "<a href=""" & replace(strURL,"_rs","") & """ target=""_blank""><img src=""" & strURL & """ border=""0"" alt=""Image"" title=""Click to view full sized picture"" /></a>"
	 'stImg = strURL
  else
     stImg = "<a href=""pic_pop.asp?mode=12&cid=" & intpicID & """ target=""_blank""><img src=""pic_pop.asp?mode=12&cid=" & intpicID & """ border=""0"" alt=""Image"" title=""Click to view full sized picture"" width=""400"" /></a>"
     'stImg = "<a href=""pic_pop.asp?mode=12&cid=" & intpicID & """ target=""_blank""><img src=""" & strURL & """ border=""0"" alt=""Image"" title=""Click to view full sized picture"" width=""400"" /></a>"
  end if
%><hr />
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<td align="center" ><%= stImg %></td></table><hr />

<table class="grid" width="100%" border="0" cellspacing="0" cellpadding="5">

<% if trim(strPicCopyright) <> "" and not isNull(strPicCopyright) then %>
<tr>
<td class="tSubTitle" width="15%"><b>Copyright:</b></td>
<td width="85%"><%= strPicCopyright %></td>
</tr>
<%end if%>
<tr>
<td class="tSubTitle" width="15%"><b>Description:</b></td> 
<td width="85%"><%= strDesc %>&nbsp;</td>
</tr>
<tr>
<td class="tSubTitle" width="15%"><b>Hits:</b></td>
<td width="85%"><%= intHit %></td>
</tr>
<tr>
<td class="tSubTitle" width="15%"><b>Rating:</b></td>
<td width="85%"><% GetRating(intpicID) %></td>
</tr>
<tr>
<td class="tSubTitle" width="15%"><b>Added on:</b></td>
<td width="85%"><% = strPostDate %></td>
</tr>
<tr>
<td class="tSubTitle" width="15%"><b>Posted by:</b></td>
<td width="85%"><% =strPoster%></td>
</tr>
<tr>
<td colspan="2" align="center"><span class="fSubTitle">Comments: <%=Comments%> </span><% if strUserMemberID > 0 then %><span class="fSubTitle"> | </span><a href="JavaScript:openWindow4('pic_pop.asp?mode=5&amp;cid=<%=intpicID%>')"><span class="fSubTitle">Add Comment/Rating</span></a><span class="fSubTitle"> | </span><a href="JavaScript:openWindow4('pic_pop.asp?mode=6&amp;cid=<%=intpicID%>')"><span class="fSubTitle">Report bad link</span></a><% end if %></td>
</tr>
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

function addPicture()
 %>
<SCRIPT type="text/javascript">

function chkGInput(strStr,params) {
var re = new RegExp("\.(" + params.replace(/,/gi,"|").replace(/\s/gi,"") + ")$","i");
    if(!re.test(strStr)) return false;
	else return true;
}
function CheckGThis(str) {
	var re;
	re = /[*'"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
function CheckName(str) {
	var re;
	re = /[\\\/:*?"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
function chkGForm() {
 if (document.forms.add_Pic.cat.value == 0) {
   alert("Select a category/subcategory");
	document.forms.add_Pic.cat.focus();
  //Dialog.alert("Select a category/subcategory", 
//				        {windowParameters: {width:300, height:100}, okLabel: "close"
//						    });
 return false;
 }
 if (document.forms.add_Pic.title.value == "") {
 alert("TITLE cannot be empty");
	document.forms.add_Pic.title.focus();
 return false;
 }
if (document.forms.add_Pic.title.value.length<3) {
 alert("TITLE must be more than 3 characters");
	document.forms.add_Pic.title.focus();
return false;
}
 if (!CheckName(document.forms.add_Pic.title.value)) {
 alert("TITLE can not contain any of the\nfollowing characters: \\ / : *  \" < > |");
	document.forms.add_Pic.title.focus();
 return false;
}
 if (!CheckName(document.forms.add_Pic.key.value)) {
 alert("Keywords can not contain any of the\nfollowing characters: \\ / : *  \" < > |");
	document.forms.add_Pic.key.focus();
 return false;
 }
 if (!CheckGThis(document.forms.add_Pic.copyright.value)) {
 alert("Copyright cannot contain any of the\nfollowing characters:  *  \" < > |");
	document.forms.add_Pic.copyright.focus();
 return false;
 }
 if (document.forms.add_Pic.url.value.length<8) {
<% If bFso Then %>
 	if (document.forms.add_Pic.file1.value == "") {
 		alert("You must either upload or link to an image");
		document.forms.add_Pic.url.focus();
 		return false;
	}
<% Else %>
 		alert("You must supply an image");
		document.forms.add_Pic.url.focus();
 		return false;
<% End If %>
 }

	if (document.forms.add_Pic.file1.value == ""){
    document.getElementById('wait').style.visibility = 'visible';
    document.getElementById('file1').style.visibility = 'hidden';
    document.getElementById('button').style.visibility = 'hidden';
	}
}

function chkFForm() {
 if (document.forms.add_Pic.url.value == "http://") {
 alert("Image URL cannot be empty");
	document.forms.add_Pic.title.focus();
 return false;
 }
 return false;
	//document.formEle.submit();
 }
</script>
<%

  	arg2 = ""
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
	
	if cat_id > 0 and isnumeric(cat_id) then
	  sSQL = "SELECT PIC_CATEGORIES.CAT_ID, PIC_CATEGORIES.CAT_NAME, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_WRITE, PIC_CATEGORIES.CG_FULL, PIC_CATEGORIES.CG_INHERIT, PIC_CATEGORIES.CG_PROPAGATE, PIC_SUBCATEGORIES.SUBCAT_ID, PIC_SUBCATEGORIES.SUBCAT_NAME, PIC_SUBCATEGORIES.SG_READ, PIC_SUBCATEGORIES.SG_WRITE, PIC_SUBCATEGORIES.SG_FULL, PIC_SUBCATEGORIES.SG_INHERIT "
	  sSQL = sSQL & "FROM PIC_CATEGORIES INNER JOIN PIC_SUBCATEGORIES ON PIC_CATEGORIES.CAT_ID = PIC_SUBCATEGORIES.CAT_ID "
	  sSQL = sSQL & "WHERE (((PIC_CATEGORIES.CAT_ID)=" & cat_id & ") AND ((PIC_SUBCATEGORIES.SUBCAT_ID)=" & sub_id & "));"
	
  	  set rsT = my_Conn.execute(sSQL)
  	  cat_name = rsT("CAT_NAME")
  	  sub_name = rsT("SUBCAT_NAME")
  	  call setPermVars(rsT,2)
  	  set rsT = nothing
	  
	  if bSCatWrite then
	    bCanSubmit = true
	  end if

  	  arg2 = cat_name & "|pic.asp?cmd=1&amp;cid=" & cat_id
  	  arg3 = sub_name & "|pic.asp?cmd=2&amp;cid=" & cat_id & "&amp;sid=" & sub_id
      arg4 = "Submit Picture|pic.asp?cmd=8"
	else
	  sSQL = "SELECT PIC_CATEGORIES.CAT_ID, PIC_CATEGORIES.CAT_NAME, PIC_CATEGORIES.CG_READ, PIC_CATEGORIES.CG_WRITE, PIC_CATEGORIES.CG_FULL, PIC_SUBCATEGORIES.SUBCAT_ID, PIC_SUBCATEGORIES.SUBCAT_NAME, PIC_SUBCATEGORIES.SG_READ, PIC_SUBCATEGORIES.SG_WRITE, PIC_SUBCATEGORIES.SG_FULL "
	  sSQL = sSQL & "FROM PIC_CATEGORIES INNER JOIN PIC_SUBCATEGORIES ON PIC_CATEGORIES.CAT_ID = PIC_SUBCATEGORIES.CAT_ID "
	  sSQL = sSQL & "ORDER BY PIC_CATEGORIES.C_ORDER, PIC_CATEGORIES.CAT_NAME, PIC_SUBCATEGORIES.C_ORDER, PIC_SUBCATEGORIES.SUBCAT_NAME;"
	
	  selectOptions = ""
  	  set rsT = my_Conn.execute(sSQL)
	    if not rsT.eof then
  	  	  cat_name = rsT("CAT_NAME")
  	  	  sub_name = rsT("SUBCAT_NAME")
  	  	  selectOptions = GetCategories(rsT)
	    else
		end if
  	  set rsT = nothing
	  
      arg2 = "Submit Picture|pic.asp?cmd=8"
	end if
	
  	'shoDebugVars()
	  
	if not bCanSubmit then
	  closeAndGo("pic.asp")
	end if
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	app_MainColumn_top()

spThemeTitle= "Submit Picture"
spThemeBlock1_open(intSkin) %>
<form method="post" action="pic_add.asp" id="add_Pic" name="add_Pic" onSubmit="return chkGForm();" enctype="multipart/form-data">
  <table class="tPlain" cellpadding="0" cellspacing="0"><tr>
            <td colspan="2"> 
              Please read <a href="javascript:openWindow3('pic_pop.asp?mode=13#Submit_New_Picture')" title="Read the FAQs page for additional support for the Pictures Area."><u><b>this page</b></u></a> before posting pictures on our site.</td>
  </tr>
  <tr>
    <td align="right">
	  Category:  
      <input type="hidden" name="memID" value="<%= strUserMemberID %>" />
              <input type="hidden" name="max" value="1" />
	</td>
	<% 
	if cat_id = "" or cat_id = 0 then %>
	<% response.Write(selectOptions)
	else
	%>
	<td>
	  <b><%=cat_name%> / <%=sub_name%></b>
      <input type="hidden" name="cat"  value="<%=sub_id%>" />
      <input type="hidden" name="parentID"  value="<%=cat_id%>" />
	</td>
	<%
	end if
	%>
  </tr>
  <tr>
    <td align="right">
	  <span class="fAlert">*</span> Title:  
	</td>
	<td><input type="text" name="title" size="40" maxlength="90" /></td>
  </tr>
  <tr>
    <td align="right" valign="top">Description:  <br /><span id="charLeft1">250 characters left.</span>  </td>
    <td><textarea rows="5" name="desc" id="desc" cols="40" onKeyUp="cntChar('desc','charLeft1','{CHAR} characters left.',250);"></textarea></td>
  </tr>
  <tr>
    <td valign="middle" align="right">Keywords for search:  </td>
    <td><input type="text" name="key" size="40" maxlength="240" /></td>
  </tr>
		  <%
		  	strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_APPID = " & intAppID
			set rsUload = my_Conn.execute(strSQL)
			uActive = rsUload("UP_ACTIVE")
			uAllowed = rsUload("UP_ALLOWEDGROUPS")
			uSize = rsUload("UP_SIZELIMIT")
			uExt = rsUload("UP_ALLOWEDEXT")
			session.Contents("uploadType") = rsUload("ID")
			session.Contents("loggedUser") = strdbntusername
			set rsUload = nothing
		  If bFso and strAllowUploads = 1 and uActive = 1 and hasAccess(uAllowed) Then
		    ast = "**"
			btxt = "<span class=""fAlert"">**</span> = link to either foreign image OR upload an image from hard disk. (<u>Don't use both</u>)<br />" %>
          <tr>
            <td align="center" valign="top" colspan="2">
			  <br />Maximum Upload file size: <b><%= uSize %> kb</b><br />
			  Allowed extentions: <b><%= uExt %></b>
            </td>
          </tr>
          <tr valign="middle">
            <td align="right" valign="top">
			  <a href="JavaScript:openWindow3('pic_pop.asp?mode=13&item=1#Uploading')" title="Click here if you need help on Uploading."><img src="images/icons/icon_smile_question.gif" border="0"></a> <span class="fAlert">**</span> Upload Image:  </td>
            <td>
              <input class="textbox" name="file1" id="file1" type="file" size="30" onchange="preview(this,'100','100')" /><br /><center>
			  <img alt="Picture will preview here" id="previewField" src="images/spacer.gif" border="0"></center>
              <br />
              <b>OR</b>, Enter the URL of an image in another site:
            </td>
          </tr>
		  <% Else
		       ast = "*"
			   btxt = "" %>
		  		<input class="textbox" name="file1" id="file1" type="hidden" value="" />
		  <% End If %>
  <tr>
    <td align="right">
      <a href="JavaScript:openWindow3('pic_pop.asp?mode=13&item=1#URL')"><img src="images/icons/icon_smile_question.gif" border="0" title="Click here if you need help on entering an URL for your picture."></a> <span class="fAlert"><%= ast %></span> URL:  
    </td>
    <td><input type="text" name="url" size="40" value="http://" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right">
      Thumbnail URL:  
    </td>
    <td><input type="text" name="turl" size="40" value="http://" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right">
      Copyright:  
    </td>
    <td><input type="text" name="copyright" size="40" maxlength="90" /></td>
  </tr>
  <tr> 
    <td align="center" colspan="2">
	 <div id="wait" name="wait" style="visibility:hidden;"><center><b>Upload in progress, please wait...</b></center><br /></div></td>
  </tr>
  <tr>
	<td><!-- Private <input type="checkbox" name="private" value="1" /> --></td>
    <td><input type="submit" value="Submit" name="B1" accesskey="s" title="Shortcut Key: Alt+S" class="button" id="button" /> <input type="reset" value="Reset" name="B2" class="button" /></td>
  </tr></table>
</form>
<hr>
<center>
<span class="fAlert">*</span> = required field<br /><%= btxt %><br />
If you did not see a category that fit your picture's content,<br />
<a href="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getMemberID(split(strwebmaster,",")(0)) %>');"><u><span class="fAlert">contact 
        us</span></u></a> and we'll be happy to consider it for you.<br />
<br />
<%	if lcase(strEmail) = "1" then%>
We will notify you by Email when the picture gets added to our database.
<%  end if%>
<br /><br /><a href="pic.asp">Back</a>
<br /><br />
</center>
<%spThemeBlock1_close(intSkin)%>
<%
end function

function GetCategories(ob)
  tStr = ""
  tStr = tStr & "<td>" & vbCRLF
  tStr = tStr & "<select name=""cat"">" & vbCRLF
  tStr = tStr & "<option value=""0""> [select one] </option>" & vbCRLF
  do while not ob.EOF
    if hasAccess(ob("CG_WRITE")) and hasAccess(ob("SG_WRITE")) then
	  bCanSubmit = true
	  sc = ob("CAT_NAME")
	  ssc = ob("SUBCAT_NAME")
	  isc = ob("SUBCAT_ID")
	  tStr = tStr & "<option value="""& isc &""">"& sc &" / "& ssc &"</option>" & vbCRLF
	end if
	ob.MoveNext
  loop
  if not bCanSubmit then
    'closeandgo("pic.asp")
  end if
  tStr = tStr & "</select>" & vbCRLF
  tStr = tStr & "</td>" & vbCRLF
  GetCategories = tStr
end function

sub showPicCats()
  sSql = "SELECT * FROM PIC_CATEGORIES ORDER BY C_ORDER"
  set rsT = my_Conn.execute(sSql)
  if not rsT.eof then
    Response.Write("<br /><div class=""tSubTitle"">Categories</div>")
    Response.Write("<div class=""menu"">")
    do until rsT.eof
	  if hasAccess(rsT("CG_READ")) then
	    Response.Write("<a href=""pic.asp?cmd=1&cid=" & rsT("CAT_ID") & """>- ")
	    Response.Write(rsT("CAT_NAME"))
	    Response.Write("<br /></a>")
	  end if
      rsT.movenext
    loop
    Response.Write("</div><br />")
  end if
  set rsT = nothing
end sub

sub menu_pictures()
	'spThemeTitle= txtMenu
	spThemeBlock1_open(intSkin)
 if bFso then
    mnu.menuName = "m_pictures"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
 else %>
	<div class="tSubTitle"><%= txtMenu %></div>
	<div class="menu">
      <a href="pic.asp?cmd=3">- <%= txtNewPics %><br /></a>
      <a href="pic.asp?cmd=4">- <%= txtPopPics %><br /></a>
      <a href="pic.asp?cmd=5">- <%= txtTopPics %><br /></a>
	<%if not strDBNTUserName = "" then%>
      <a href="pic.asp?cmd=8">- <%= txtSubPic %><br /></a>
	<%End If %>
      <a href="javascript:openWindow3('pic_pop.asp?mode=13')">- <%= txtPicsFAQ %><br /></a>
	</div>
<% End If %>
  <% If iPgType > 0 Then
   		showPicCats()
  	 end if %>
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
	<form method="get" action="pic.asp" id="formS1" name="formS1" onSubmit="return chkSrchForm1()">
	<% 
	spThemeTitle = txtSearch & ":"
	spThemeBlock3_open() %>
    <div class="tPlain" style="text-align:center;">
	<input type="text" name="search" size="15" style="margin-top:5px;margin-bottom:5px;" />
  <select name="mode" id="mode">
    <option value="0" selected>All Pictures</option>
    <option value="1">For Member</option>
  </select></div>
      <div class="fNorm" style="margin-bottom:3px;text-align:center;">
      <input type="submit" value=" <%= txtSearch %> " id="searchA" name="searchA" class="button" /><input type="hidden" name="cmd" value="7" /></div><% spThemeBlock3_close() %></form><br />
<%spThemeBlock1_close(intSkin)
end sub
%>