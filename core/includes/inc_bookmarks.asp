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

sub showMyBookmarks()
  memID = strUserMemberID
	strSQL = "SELECT DISTINCT " & strTablePrefix & "BOOKMARKS.APP_ID, " & strTablePrefix & "APPS.APP_NAME FROM " & strTablePrefix & "BOOKMARKS INNER JOIN " & strTablePrefix & "APPS ON " & strTablePrefix & "BOOKMARKS.APP_ID = " & strTablePrefix & "APPS.APP_ID WHERE (((" & strTablePrefix & "BOOKMARKS.M_ID)=" & memID & ")) ORDER BY " & strTablePrefix & "APPS.APP_NAME;"
	set rsBmAp = my_Conn.execute(strSQL)
	If rsBmAp.Eof OR rsBmAp.Bof Then
	  call showMsgBlock(1,txtNoBkmkFnd)
	Else
  	  do until rsBmAp.eof
  	  appID = rsBmAp("APP_ID")
      %>
        <form Action="cp_main.asp?cmd=7&mode=delete" method=post id="Form<%= appID %>" name="Form<%= appID %>">
	    <%
	    spThemeMM = "bkmk"
	    spThemeTitle = ucase(rsBmAp("APP_NAME")) & " BookMarks"
	    spThemeBlock1_open("1")
	  sSQL = "SELECT " & strTablePrefix & "BOOKMARKS.BOOKMARK_ID, " & strTablePrefix & "BOOKMARKS.APP_ID, " & strTablePrefix & "BOOKMARKS.ITEM_ID, " & strTablePrefix & "BOOKMARKS.SUBCAT_ID, " & strTablePrefix & "BOOKMARKS.CAT_ID, " & strTablePrefix & "BOOKMARKS.ITEM_TITLE, " & strTablePrefix & "APPS.APP_INAME, " & strTablePrefix & "APPS.APP_VIEW, " & strTablePrefix & "BOOKMARKS.M_ID FROM " & strTablePrefix & "BOOKMARKS INNER JOIN " & strTablePrefix & "APPS ON " & strTablePrefix & "BOOKMARKS.APP_ID = " & strTablePrefix & "APPS.APP_ID WHERE (((" & strTablePrefix & "BOOKMARKS.APP_ID)=" & appID & ") AND ((" & strTablePrefix & "BOOKMARKS.M_ID)=" & memID & ")) ORDER BY " & strTablePrefix & "BOOKMARKS.CAT_ID DESC, " & strTablePrefix & "BOOKMARKS.SUBCAT_ID DESC, " & strTablePrefix & "BOOKMARKS.ITEM_ID DESC;"
	  
	  set rsBmks = my_Conn.execute(sSQL)
	  
	  curType = 0
	  shoHeader = false
	  do while not rsBmks.eof
	  'do while curAppID = rs("APP_ID")
		  select case rsBmks("APP_INAME")
		    case "forums"
			  if rsBmks("ITEM_ID") = 0 then
			    if curtype <> 1 then
				  shoHeader = true
				end if
				subType = txtForums & ":"
				curType = 1
			    lnkTo = "link.asp?forum_id=" & rsBmks("SubCat_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  else
			    if curtype <> 2 and not shoHeader then
				  shoHeader = true
				end if
				subType = txtTopics & ":"
				curType = 2
			    lnkTo = "link.asp?topic_id=" & rsBmks("ITEM_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  end if
		    case "links"
			  if rsBmks("CAT_ID") <> 0 then 'cat bookmark
			    if curtype <> 1 then
				  shoHeader = true
				end if
				subType = txtCats & ":"
				curType = 1
			    lnkTo = "links.asp?cmd=1&cid=" & rsBmks("CAT_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  elseif rsBmks("SUBCAT_ID") <> 0 then 'subcat bookmark
			    if curtype <> 2 then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				sSQL = "SELECT " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_NAME, " & strTablePrefix & "M_CATEGORIES.CAT_NAME "
				sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
				sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				set rsB = my_Conn.execute(sSQL)
				if not rsB.eof then
				    itmTitle = rsB(1) & "/" & rsB(0)
				end if
				set rsB = nothing
				subType = txtSubCats & ":"
				curType = 2
			    lnkTo = "links.asp?cmd=2&sid=" & rsBmks("SUBCAT_ID")
				cls = "tCellAlt1"
			  else
			    if curtype <> 3 and not shoHeader then
				  shoHeader = true
				end if
				subType = txtItems & ":"
				curType = 3
			    lnkTo = rsBmks("APP_VIEW") & "?cmd=10&sid=" & rsBmks("ITEM_ID") & """ target=""_blank"
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  end if
		    case else
			  if rsBmks("CAT_ID") <> 0 then 'cat bookmark
			    if curtype <> 1 then
				  shoHeader = true
				end if
				subType = txtCats & ":"
				curType = 1
			    lnkTo = rsBmks("APP_VIEW") & "?cmd=1&cid=" & rsBmks("CAT_ID")
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  elseif rsBmks("SUBCAT_ID") <> 0 then 'subcat bookmark
			    if curtype <> 2 then
				  shoHeader = true
				end if
				itmTitle = rsBmks("ITEM_TITLE")
				sSQL = "SELECT " & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_NAME, " & strTablePrefix & "M_CATEGORIES.CAT_NAME "
				sSQL = sSQL & " FROM " & strTablePrefix & "M_CATEGORIES INNER JOIN " & strTablePrefix & "M_SUBCATEGORIES ON " & strTablePrefix & "M_CATEGORIES.CAT_ID = " & strTablePrefix & "M_SUBCATEGORIES.CAT_ID"
				sSQL = sSQL & " WHERE (((" & strTablePrefix & "M_SUBCATEGORIES.SUBCAT_ID)=" & rsBmks("SUBCAT_ID") & "));"
				set rsB = my_Conn.execute(sSQL)
				if not rsB.eof then
				    itmTitle = rsB(1) & "/" & rsB(0)
				end if
				set rsB = nothing
				subType = txtSubCats & ":"
				curType = 2
			    lnkTo = rsBmks("APP_VIEW") & "?cmd=2&sid=" & rsBmks("SUBCAT_ID")
				cls = "tCellAlt1"
			  else
			    if curtype <> 3 and not shoHeader then
				  shoHeader = true
				end if
				subType = txtItems & ":"
				curType = 3
				if instr(rsBmks("APP_VIEW"),"?") > 0 then
			      lnkTo = rsBmks("APP_VIEW") & rsBmks("ITEM_ID")
				else
			      lnkTo = rsBmks("APP_VIEW") & "?cmd=10&sid=" & rsBmks("ITEM_ID")
				end if
				cls = "tCellAlt1"
				itmTitle = rsBmks("ITEM_TITLE")
			  end if
		  end select
		  if shoHeader then
		    shoHeader = false
	       response.Write("<div class=""tAltSubTitle"" style=""padding-left:8px;padding-top:5px;padding-bottom:3px; text-align:left;"">&nbsp;&nbsp;<b>" & subType & "</b></div>")
		  end if
	    response.Write("<div class=""" & cls & """ style=""padding-left:8px;padding-top:5px;padding-bottom:3px; text-align:left;""><input type=""checkbox"" name=""delBookmark"" value=""" & rsBmks("BOOKMARK_ID") & """>&nbsp;&nbsp;<b><a href=""" & lnkTo & """><span class=""fNorm"">" & itmTitle & "</span></a></b></div>")
	    rsBmks.movenext
	  loop
	'loop
	 %>
	  <div align="center">
	  <input type="submit" name="del" value="<%= txtDelSelected %>&nbsp;<%= ucase(rsBmAp("APP_NAME")) %>&nbsp;<%= txtBookmks %>" ID="Submit<%= rsBmAp("APP_ID") %>" class="button"></div><%
	    spThemeBlock1_close("1") %>
        </FORM><br />
<%    rsBmAp.movenext
      loop
    end if
	set rsBmAp = nothing
	set rsBmks = nothing

end sub

sub modifyBookmarks()
  Select Case Request.QueryString("mode")
   Case "delete"
	delCnt = 0
	delBkmk = split(Request.Form("delBookmark"), ",")
	
	for ib = 0 to ubound(delBkmk)
		if isnumeric(delBkmk(ib)) then
		' Delete selected topic bookmarks
		delSQL = "DELETE FROM "& strTablePrefix & "BOOKMARKS WHERE BOOKMARK_ID = " & delBkmk(ib)
		delSQL = delSQL & " AND M_ID=" & strUserMemberID
    	executeThis(delSQL)
		delCnt = delCnt + 1
		end if
	next
	if delCnt > 0 then
	  tmpResult = tmpResult & txtSelBkDel
	else
	  tmpResult = tmpResult & txtNoSelBkmk
	end if
  	'call showMsgBlock(1,tmpResult)

   Case "deleteAll"
	'delBookmark = split(Request.Form("delBookmark"), ",")
	delSQL = "DELETE FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID = " & strUserMemberID
    executeThis(delSQL)
	tmpResult = tmpResult & txtAllBkmkDel
  	'call showMsgBlock(1,tmpResult)
	
  end select
end sub
%>