<!--#include file="config.asp" --><% 
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
dim curpagetype
curpagetype = "pictures"
%>

<!--#include file="inc_functions.asp" -->
<!-- #include file="modules/pictures/pic_functions.asp" -->
<%
dim iMode, iCmd, cid, app, sid
iMode = 0
sMode = ""
iCmd = 0
cid = 0
sid = 0
app = ""
intSkin=1

if Request("mode") <> "" and Request("mode") <> " " then
  if IsNumeric(Request("mode")) = True then
	iMode = cLng(Request("mode"))
  else
	sMode = chkString(Request("mode"),"sqlstring")
  end if
end if

if Request("cmd") <> "" and Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iCmd = cLng(Request("cmd"))
	else
		closeAndGo("stop")
	end if
end if

if Request("cid") <> "" and Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cid = cLng(Request("cid"))
	else
		closeAndGo("stop")
	end if
end if

if Request("sid") <> "" and Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sid = cLng(Request("sid"))
	else
		closeAndGo("stop")
	end if
end if
%>
<!--#include file="inc_top_short.asp" -->
<%
call setAppPerms("pictures","iName")

select case iMode
  case 1
    app = "Pic"
	adtyp = "added to"
	hp = 1
  case 2
    app = "Pic"
	adtyp = "removed from"
	hp = 0
  case 12 'display pic - redirect
	showPic2()
	'showPicAspJpeg()
  case else
end select
		  
'strUserMemberID = getmemberid(strdbntusername)
 if iMode > 0 and hasAccess(sAppRead) then
	select case sMode
	  case "editAccess" 'edit group access form
		editAccessForm()
	  case "updAccess" 'update group access
		updateAccess()
  	  case else
	end select
	  select case iMode
	    case 1, 2
		  addFrontPage(app)
		case 3 'bookmark item
		  addBookmark()
		case 4 'PIC_goto.asp
		  emailToFriend()
		case 5 'rate item
		  rateItem()
		case 6 'report bad item
		  badLink()
		case 7 'delete comment
		  deleteComment()
		case 8 'email friend
		  emailToFriend()
		case 9 'subscribe to item
		  addSubscription()
		case 10 'delete bookmark
		  delBookmark()
		case 11 'delete subscription
		  delSubscription()
  		case 13 'FAQ
		  spThemeTitle= "Pictures FAQ/Help"
		  showFAQ()
		case 14 'edit group access form
		  editAccessForm "PIC_CATEGORIES","PIC_SUBCATEGORIES","pic_pop.asp",15
		case 15 'update group access
		  updateAccess "PIC_CATEGORIES","PIC_SUBCATEGORIES","pic_pop.asp",15
	  end select
   end if
 %>
<!--#include file="inc_footer_short.asp" -->
<% 

sub showPicAspJpeg()
	'Response.Expires = 0
    pictype = iCmd
	strSQL = "SELECT * from PIC where "
	if not Session(strCookieURL & "Approval") = "256697926329" then
	strSQL = strSQL & "SHOW=1 and "
	end if
	strSQL = strSQL & "PIC_ID = " & cid & ""
	
	set rs = my_Conn.Execute (strSql)
	PIC_URL = rs("URL")
	PIC_TURL = rs("TURL")
	set rs = nothing
	if trim(PIC_TURL) <> "" then
  		thb = PIC_TURL
	else
  		thb = PIC_URL
	end if

	if pictype = 1 then
	  PIC_URL = thb
	else
	  PIC_URL = PIC_URL
	end if
	
	'closeAndGo(PIC_URL)
	
  if instr(PIC_URL,"_rs.") = 0 or instr(PIC_URL,"_sm.") = 0 then
	  'Response.write(PIC_URL)
	' create instance of AspJpeg
	  PIC_URL = Server.URLEncode(server.MapPath(PIC_URL))
	Set jpg = Server.CreateObject("Persits.Jpeg")
	
	' Open source file
	jpg.Open(PIC_URL)

	' Set resizing algorithm
	'jpg.Interpolation = Request("Interpolation")

	' Set new height and width
	'jpg.Width = Request("Width")
	'jpg.Height = Request("Height")
	
	' Sharpen resultant image
	'If Request("Sharpen") <> "0" Then 
		'jpg.Sharpen 1, Request("SharpenValue")
	'End If

	' Rotate if necessary. Only available in version 1.2
	'If Request("Rotate") = 1 Then jpg.RotateL
	'If Request("Rotate") = 2 Then jpg.RotateR

	' Perform resizing and 
	' send resultant image to client browser
	jpg.SendBinary
	closeAndGo("stop")
  else
	closeAndGo(PIC_URL)
  end if	
end sub

sub showPic2()
    pictype = iCmd
	strSQL = "SELECT * from PIC where "
	if hasAccess(1) then
	strSQL = strSQL & "ACTIVE=1 and "
	end if
	strSQL = strSQL & "PIC_ID = " & cid & ""
	
	set rs = my_Conn.Execute (strSql)
	PIC_URL = rs("URL")
	PIC_TURL = rs("TURL")
	set rs = nothing
	if trim(PIC_TURL) <> "" then
  		thb = PIC_TURL
	else
  		thb = PIC_URL
	end if

	if pictype = 1 then
	  'closeAndGo(thb)
	  Response.write(thb)
	else
	  Response.write(PIC_URL)
	  closeAndGo(PIC_URL)
	end if
end sub

sub delBookmark()
  bSQL = "DELETE FROM " & strTablePrefix & "BOOKMARKS WHERE M_ID=" & strUserMemberID & " AND BOOKMARK_ID=" & cid
  executeThis(bSQL)
  strMsg = "Bookmark Removed!"
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

sub delSubscription()
  bSQL = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & strUserMemberID & " AND SUBSCRIPTION_ID=" & cid
  executeThis(bSQL)
  strMsg = "Subscription Removed!"
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

sub addBookmark()
		 ' response.Write("Hello " & iCmd & "<br />")
	strMsg = ""
	memID = getmemberID(strDBNTUserName)
	sSql = "SELECT APP_ID,APP_BOOKMARKS FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'pictures'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	  if intBookmarks = 1 then
	    intBookmarks = rsA("APP_BOOKMARKS")
	  end if
	else
	  strMsg = "Module error in" & " PORTAL_APPS"
	end if
	set rsA = nothing
	if strMsg = "" and intBookmarks = 1 then
	  select case iCmd
	    case 1 'bookmark category
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT CAT_NAME FROM PIC_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Picture Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("CAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "BOOKMARKS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", " & cid & ", 0, 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = "Category Bookmark Added!"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Category Bookmark already exists"
	      End If
	      set rs = nothing
	    case 2 ' bookmark subcategory
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and SUBCAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT SUBCAT_NAME FROM PIC_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'item does not exist
		      strMsg = "Picture SubCategory not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("SUBCAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "BOOKMARKS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = "SubCategory Bookmark Added!"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "SubCategory Bookmark already exists"
	      End If
	      set rs = nothing
	    case 3 ' bookmark item
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and ITEM_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT TITLE FROM PIC WHERE PIC_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'item does not exist
		      strMsg = "Picture not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("TITLE")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "BOOKMARKS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, 0, " & cid & ", '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = "Bookmark Added!"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Bookmark already exists"
	      End If
	      set rs = nothing
		case else
		'do nothing
	  end select
    end if 'strMsg = ""
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

sub addSubscription()
		 ' response.Write("Hello " & iCmd & "<br />")
	strMsg = ""
	memID = getmemberID(strDBNTUserName)
	sSql = "SELECT APP_ID,APP_SUBSCRIPTIONS FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'pictures'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	  if intSubscriptions = 1 then
	    intSubscriptions = rsA("APP_SUBSCRIPTIONS")
	  end if
	else
	  strMsg = "Module error in PORTAL_APPS"
	end if
	set rsA = nothing
	
	' check for module subscription
	if strMsg = "" and intSubscriptions = 1 then
	  sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=0"
	  set rsAp = my_Conn.execute(sSql)
	  If rsAp.BOF or rsAp.EOF Then
	    ' they are not subscribed to the module
	  else
	    strMsg = "Cannot add subscription" & "<br /><br />"
	    strMsg = strMsg & "You are already subscribed" & "<br />"
		strMsg = strMsg & "to the Pictures Module"
	  end if
	end if
	
	if strMsg = "" and intSubscriptions = 1 then
	  select case iCmd
	    case 1 'subscribe to category
	       sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	       set rs = my_Conn.execute(sSql)
	       If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT CAT_NAME FROM PIC_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Picture Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("CAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", " & cid & ", 0, 0, '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "Category Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "a new picture is added to the " & itmTitle & " category!<br />"
			  
	          insSql = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS"
	          insSql = insSql & " WHERE ((APP_ID=" & intAppID & " AND ITEM_ID=0 AND M_ID=" & memID & ")"
			  insSql = insSql & " AND (CAT_ID<>0 OR SUBCAT_ID<>0));"
	          'executeThis(insSql)
			  
	          strMsg = strMsg & "All of your previous " & itmTitle & " SubCategory" & "<br />"
	          strMsg = strMsg & "subscriptions have been deleted" & ".<br />"
		     end if
		    set rsA = nothing
	      else
		    strMsg = "Category subscription already exists"
	      End If
	      set rs = nothing
	    case 2 ' subscribe to subcategory
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and SUBCAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT SUBCAT_NAME FROM PIC_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'item does not exist
		      strMsg = "Picture SubCategory not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("SUBCAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = strMsg & "SubCategory Subscription Added!" & "<br /><br />"
	          strMsg = strMsg & "You will now receive an email when" & "<br />"
	          strMsg = strMsg & "a new picture is added to the" & " " & itmTitle & " " & "SubCategory<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "SubCategory subscription already exists"
	      End If
	      set rs = nothing
	    case 3 ' Subscription module
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=0"
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT APP_NAME FROM "& strTablePrefix & "APPS WHERE APP_ID=" & intAppID
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'item does not exist
		      strMsg = "Pictures module not found"
		    else 'item does exist, lets Subscription it
		      itmTitle = "All New Pictures"
			  
	          ' Subscription doesn't already exist so add it
	          insSql = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS"
	          insSql = insSql & " WHERE APP_ID=" & intAppID & " AND M_ID=" & memID & ";"
	          executeThis(insSql)
			  
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, 0, 0, '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "'Picture' Module Subscription Added!" & "<br /><br />"
	          strMsg = strMsg & "You will now receive an email when" & "<br />"
	          strMsg = strMsg & "any new pictures are added to the database" & ".<br /><br />"
			  
	          strMsg = strMsg & "All of your previous Picture Category and" & "<br />"
	          strMsg = strMsg & "SubCategory subscriptions have been deleted" & ".<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Pictures module subscription already exists"
	      End If
	      set rs = nothing
		case else
		'do nothing
	  end select
    end if 'strMsg = ""
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

sub rateItem()
  response.Write("<br /><br />")
  if strDBNTUserName <> "" then 
    strUserMemberID = strUserMemberID 
  else 
    strUserMemberID = -1
  end if

  if iCmd = 1 then
	dim intFormRating
	dim strComments
	dim strError
	strError = ""
	
	intFormRating = ChkString(Request.Form("rating"),"sqlstring")
	if not IsNumeric(intFormRating) = True then
  	  intFormRating = 10
	end if
	strComments = ChkString(request.form("comments"), "message")

	'Check to see if they are a member.
	if strUserMemberID = -1 then
		strError = strError & "<li>Only Members can rate pictures.</li>"
		strError = strError & "<li>If you are already a member, please log in to rate this picture.</li>"
	end if
	
	' Check to see if they rated already
	strSQL = "SELECT RATE_BY FROM PIC_RATING WHERE PIC = " & cid & " AND RATE_BY = " & strUserMemberID
	set rsCheck = server.CreateObject("adodb.recordset")
	rsCheck.Open strSQL, my_Conn
	if not rsCheck.EOF then
		strError = strError & "<li>You have already rated this picture</li><li>You cannot rate it again.</li>"
	end if
	rsCheck.Close
	set rsCheck = nothing
	
	'Check to see if they entered a rating.
	if intFormRating = "" then
		strError = strError & "<li>You didn't select a rating.</li>"
	end if
	if strComments & "x" = "x" then
		strError = strError & "<li>You did not make a comment.</li>"
	end if

	if strError <> "" then
	spThemeBlock1_open(intSkin)%>
      <p align="center"><div class="fTitle">There Was A Problem.</div></p>
	  <table align="center" border="0">
	   <tr>
	    <td align="center">
		  <ul style="text-align:left;"><%=strError%></ul>
	    </td>
	   </tr>
	  </table>
	  <p align="center"><a href="JavaScript:history.go(-1)">Go Back</a></p><br />
	  <% spThemeBlock1_close(intSkin)		
	else
		
		strSQL = "SELECT VOTES, RATING FROM PIC WHERE PIC_ID = " & cid
		set rsArticleRating = server.CreateObject("adodb.recordset")
		rsArticleRating.Open strSQL, my_Conn
		
		dim intVotes
		dim intRating
		intVotes = rsArticleRating("VOTES") + 1
		intRating = rsArticleRating("RATING") + intFormRating
		rsArticleRating.Close
		set rsArticleRating = nothing

		strSQL = "UPDATE PIC SET VOTES = " & intVotes & " , RATING = " & intRating & " Where PIC_ID = " & cid
		executeThis(strSQL)
		
		strSQL = "INSERT INTO PIC_RATING ( PIC, RATING, COMMENTS, RATE_BY, RATE_DATE ) VALUES ( " & cid & " , " & intRating & " , '" & strComments & "', " & strUserMemberID & " , '" & strCurDateString & "' )"
		executeThis(strSQL)
	  spThemeBlock1_open(intSkin)%>
	  <p align="center"><div class="fTitle">&nbsp;</div></p>
	  <table align="center" border="0">
	    <tr>
	      <td align=center><br><br>
	      Thank You for rating this Picture.
		<script type="text/javascript">opener.document.location.reload();</script>
		<br><br><br><br>
		  </td>
	    </tr>
	  </table><%
	  spThemeBlock1_close(intSkin)
		
	end if
  else 'iCmd <> 1, Show rating form
	strError = ""
	'Check to see if they are a member.
	if strUserMemberID = -1 then
		strError = "<li>Only Members can rate Pictures.</li>"
		strError = strError & "<li>If you are already a member, please log in to rate this picture.</li>"
	else
		' Check to see if they rated already
		strSQL = "SELECT RATE_BY FROM PIC_RATING WHERE PIC = " & cid & " AND RATE_BY = " & strUserMemberID
		set rsCheck = server.CreateObject("adodb.recordset")
		rsCheck.Open strSQL, my_Conn
		if not rsCheck.EOF then
		  strError = "<li>You have already rated this picture</li><li>You cannot rate it again.</li>"
		end if
		rsCheck.Close
		set rsCheck = nothing
	end if
	
	if strError <> "" then
	  spThemeBlock1_open(intSkin)%>
	  <p align="center"><div class="fTitle">There Was A Problem.</div></p>
	  <table align="center" border="0">
	    <tr>
	      <td align="center"><span><ul style="text-align:left;"><% =strError %></ul></span>
		  </td>
	    </tr>
	  </table><%
	  spThemeBlock1_close(intSkin)		
	else
	  dim strLinkSQL, rsLink
	  strLinkSQL = "SELECT TITLE, PIC_ID, HIT, POST_DATE FROM PIC WHERE PIC_ID = " & cid
	  Set rsLink = Server.CreateObject("ADODB.Recordset")
	  rsLink.Open strLinkSQL, my_Conn
	  if rsLink.EOF then
		Response.Write "Picture does not exist."
	  else
		strLinkTitle = rsLink("TITLE")
		intLinkID = rsLink("PIC_ID")
		intHit = rsLink("HIT")		
		strPostDate = strtodate(rsLink("POST_DATE"))
		dateSince=DateDiff("d", Date(), strPostDate)+7 %>
		<form method="post" name="rateform" action="pic_pop.asp">
		<%
		spThemeTitle= ""
		spThemeBlock1_open(intSkin)%>
		<table><tr><td>
		<b>Add Comment/Rating</b></td></tr>
		<tr><td>
		<div class="fSmall"><ol>
		<li>Rating Scale: <span class="fAlert">1</span> = worst, <span class="fAlert">10</span> = best</li>
		<li>Please try to be objective.</li>
		<li>Only registered members can comment/rate items.</li>
		<li>You can only rate/comment an item once.</li>
		<li>You must supply a rating and comment.</li></ol>
		</div></td></tr>
		<tr><td><b><%=strLinkTitle%></b>
		<% if dateSince >= 0 then response.write icon(icnNew1,"New Item","","","align=""middle""") %><br>
		(Added : <%=formatdatetime(strPostDate, 2)%> Hits : <%=intHit%>)<br></td></tr>
		<tr><td>
		<b>Rating:</b>&nbsp;
		<select name="rating">
          <option value="">-</option>
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
        </select>
		</td></tr>
		<tr><td><b>Comments:</b>&nbsp;<br>
		<textarea rows="8" cols="35" name="comments"></textarea></td></tr>
		<tr><td align=center><input type="submit" value=" Submit " class="button" />&nbsp;&nbsp;
		<input type=reset value=" Clear " class="button" />
		<input type=hidden name="cid" value="<%= intLinkID %>" />
		<input type=hidden name="cmd" value="1" />
		<input type=hidden name="mode" value="<%=iMode%>" /></td></tr></table>
		<% spThemeBlock1_close(intSkin)%>
		</form>
<%
	end if
	rsLink.Close
	set rsLink = nothing
	end if
  end if
end sub

sub picAddHit(typ)
	lastdate = chkString(Request.Cookies("date"),"sqlstring")
	lastid = chkString(Request.Cookies("dlid"),"sqlstring")

	if lastid <> dlid then
		Response.Cookies("dlid") = dlid
		Response.Cookies("dlid").Expires = dateadd("d",7,strCurDateAdjust)
	    executeThis("UPDATE DL SET HIT = HIT + 1 Where PIC_ID =" & typ)
	end If
	dim rs
	Set rs = my_Conn.Execute("SELECT URL FROM PIC WHERE PIC_ID = " & typ)
	linkurl = rs("URL")
	set rs = nothing
  'response.Write(linkurl)
  'closeAndGo(linkurl)	
end sub

sub emailToFriend()
  response.Write("<br />")
  spThemeBlock1_open(intSkin)
  if iCmd = 6 then
	Err_Msg = ""
	if (Request.Form("YName") = "") then 
		Err_Msg = Err_Msg & "<li>You must enter your name!</li>"
	end if
	if (Request.Form("YEmail") = "") then 
		Err_Msg = Err_Msg & "<li>You Must give your email address</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid email address</li>"
		end if
	end if
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>You must enter the recipients name</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>You Must enter the recipients email address</li>"
	else
		if (EmailField(Request.Form("Email")) = 0) then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid email address for the recipient</li>"
		end if
	end if
	if (Request.Form("Msg") = "") then 
		Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
	end if
	'##  Emails Topic To a Friend.  
	if lcase(strEmail) = "1" then
	  if (Err_Msg = "") then
			strRecipientsName = chkString(Request.Form("Name"),"sqlstring")
			strRecipients = chkString(Request.Form("Email"),"sqlstring")
			strSubject = "From: " & chkString(Request.Form("YName"),"sqlstring") & " Interesting " & repstring
			strMessage = "Hello " & chkString(Request.Form("Name"),"sqlstring") & vbCrLf & vbCrLf
			strMessage = strMessage & chkString(Request.Form("Msg"),"sqlstring") & vbCrLf & vbCrLf
			strMessage = strMessage & "You received this from : " & chkString(Request.Form("YName"),"sqlstring") & " " & chkString(Request.Form("YEmail"),"sqlstring")
			sendOutEmail strRecipients,strSubject,strMessage,2,0
%>
		    <br><p align="center"><span class="fTitle">Email has been sent</span></p>
			<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">Close Window</a><br>&nbsp;</font></p>
			<% spThemeBlock1_close(intSkin)
			closeAndGo("stop")
	  else %>
		  <table>
		  <tr>
		  <td align="center">
			  <p>There Was A Problem With Your Email</p>
			  <span><ul style="text-align:left;"><%= Err_Msg %></ul></span>
			  <p><a href="JavaScript:onClick= window.close()">Close Window</A></p>
		  </td>
		  </tr>
		  </table>
			<% spThemeBlock1_close(intSkin)
			 closeAndGo("stop")
	  end if
	end if
  else 
		YName = ""
		YEmail = ""
		strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE M_NAME = '" & STRdbntUserName & "'"
		set rs = my_conn.Execute (strSql)
		if rs.EOF then
			if strLogonForMail = 1 then 
				Err_Msg = Err_Msg & "<li>You must be registered to email this item</li>"
%>
			<table>
			 <tr>
			  <td align="center">
			  <p>There Was A Problem With Your Email</p>
			  <span><ul style="text-align:left;"><%= Err_Msg %></ul></span>
			  <p><a href="JavaScript:onClick= window.close()">Close Window</A></p>
			  </td>
			 </tr>
			</table>
<%				rs.close
				set rs = nothing
				closeAndGo("stop")
			end if
		else
			'YName = Trim("" & rs("M_NAME"))
			YEmail = Trim("" & rs("M_EMAIL"))
		end if
		rs.close
		set rs = nothing %>
  
  <form action="pic_pop.asp" method="post" id="Form16" name="Form16">
  <input type="hidden" name="cmd" value="6" />
  <input type="hidden" name="mode" value="4" />
  <input type="hidden" name="cid" value="<%= cid %>" />
<%
%>
      <table width="100%"><TR>
        <TD align="center" colspan="2" class="fTitle" nowrap="nowrap"><p>Send Picture to a Friend<br>&nbsp;</p></td>
      </tr>
      <TR>
        <TD align="right" nowrap><b>Send To Name:&nbsp;</b></td>
        <TD><input type="text" name="Name" size="25" /></td>
      </tr>
      <TR>
        <TD align="right" nowrap><b>Send To Email:&nbsp;</b></td>
        <TD><input type="text" name="Email" size="25" /></td>
      </tr>                
      <tr>
        <td align="right" nowrap><b>Your Name:&nbsp;</b></td>
        <td><input name="YName" type="<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %>" value="<% = YName %>" size="25" /> <% if YName <> "" then Response.Write(YName) end if %></td>
      </tr>
      <tr>
        <td align="right" nowrap><b>Your Email:&nbsp;</b></td>
        <td><input name="YEmail" type="<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %>" value="<% = YEmail %>" size="25" /> <% if YEmail <> "" then Response.Write(YEmail) end if %></td>
      </tr> 
      <tr>
        <td colspan=2 nowrap><b>Message:</b></td>
      </tr>
      <tr>
        <td colspan=2 align=center><textarea name="Msg" cols="38" rows="5" readonly="readonly">Hi, <% =vbCrLf %>I thought you might be interested in this Picture:<%= vbCrLf & vbCrLf & strHomeUrl & "pic.asp?cmd=6&amp;cid=" & cid %></textarea></td>
      </tr>                    
      <tr>
        <td colspan=2 align=center><input class="button" type="submit" value="Send" id="Submit1" name="Submit1" /></td>
      </tr></table>
  </form>
<%
end if
spThemeBlock1_close(intSkin)
end sub

sub addFrontPage(typ)
  response.Write("<br /><br />")
  spThemeBlock1_open(intSkin)
  response.Write("<br /><br />")
  if hasAccess(1) then
	'## Forum_SQL
	strSql = "UPDATE PIC set FEATURED = " & hp
	strSql = strSql & " WHERE PIC_ID = " & cid
	executeThis(strSql)
%>
	<p align=center><b><%= uCase(typ) & " " & adtyp %> 'Featured Pictures' items</b><br></p><script type="text/javascript"> opener.document.location.reload();</script>
<%
  else %>
	<p align=center><font size="<%= strDefaultFontSize %>"><b>You do not have enough permission to perform this action.</b></font></p>
<%
  end If
  response.Write("<br /><br />")
	spThemeBlock1_close(intSkin)
end sub

sub badLink()
  if strDBNTUserName <> "" Then
	sSQL = "SELECT TITLE FROM PIC WHERE PIC_ID = " & cid
	set rsNam = my_Conn.execute(sSQL)
	  lnkNam = rsNam(0)
	set rsNam = nothing
	executeThis("UPDATE PIC SET BADLINK = " & strUserMemberID & " WHERE PIC_ID=" & cid)
    strMsg = strMsg & "<li>Thank you for reporting the bad link for: <br /><b>" & lnkNam
    strMsg = strMsg & "</b><br />We will correct the problem as soon as possible</li>"
  else 
    strMsg = strMsg & "<li>You do not have enough permission to perform this action</li>"
    strMsg = strMsg & "<li>You must be logged in to report a bad Picture link</li>"
  end if %>
	<p align="center">&nbsp;</p><%
	spThemeBlock1_open(intSkin)%>
	<table align="center" border="0">
	  <tr>
	    <td align="center"><span><ul style="text-align:left;"><% =strMsg %></ul></span></td>
	  </tr>
	</table><%
	spThemeBlock1_close(intSkin)
end sub

sub deleteComment()
  strMsg = ""
  if hasAccess(1) then
    sSql = "SELECT RATING FROM PIC_RATING WHERE PIC = " & cid & " and RATE_BY = " & iCmd
	set rsRate = my_Conn.execute(sSql)
	if not rsRate.eof then
	  strRate = rsRate("RATING")
      sSql = "SELECT VOTES, RATING FROM PIC WHERE PIC_ID = " & cid
	  set rsChk = my_Conn.execute(sSql)
	    totalVotes = rsChk("VOTES")
	    totalRate = rsChk("RATING")
	  set rsChk = nothing
	  intRating = totalRate - strRate
	  intVotes = totalVotes - 1
      sSql = "DELETE FROM PIC_RATING WHERE PIC = " & cid & " and RATE_BY = " & iCmd
	  executeThis(sSql)
	  strSQL = "UPDATE PIC SET VOTES = " & intVotes & " , RATING = " & intRating & " Where PIC_ID = " & cid
	  executeThis(strSQL)
	end if
	set rsRate = nothing
    strMsg = strMsg & "<li>Picture comment, vote and rating deleted</li>"
  else
    strMsg = strMsg & "<li>You do not have enough permission to perform this action</li>"
  end if %>
	<p align="center">&nbsp;</p><%
	spThemeBlock1_open(intSkin)%>
	<script type="text/javascript"> opener.document.location.reload();</script>
	<table align="center" border="0">
	  <tr>
	    <td align="center"><span><ul style="text-align:left;"><% =strMsg %></ul></span></td>
	  </tr>
	</table><%
	spThemeBlock1_close(intSkin)
end sub

sub showFAQ()
response.Write("<br><br>")
spThemeBlock1_open(intSkin) %>
  <table class="tPlain"><tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
<p><br />
This is the area where you can share your pictures, images and photos with the rest of the community.
 
On the &quot;Pictures&quot; menu, you will see several menu options: <br /><br />
<center>
<img src="<%= strHomeUrl %>images/faq/pics_menu.gif" width="170" height="152" alt="Pictures Nav Menu">
<br />

Figure 1: The Pictures Nav Menu.</center>

<br /><br />
Note: the images displayed here could differ from those shown on your screen, due 
to the different themes available, but the content in these boxes should remain the same.
<br /></p>
    <p>
    <br />
Click on the questions below.
<br /><br />
<b>How do I...</b>
    <ul>
	<li>...<a href="#Main_Directory">see all the pictures on this site?</a></li>
    <li>...<a href="#New_Pictures">get a list of the Newest Pictures?</a></li>
    <li>...<a href="#Popular_Pictures">get a list of the most Popular Pictures?</a></li>
    <li>...<a href="#Top_Rated_Pictures">get a list of the Top-Rated Pictures?</a></li>
    <li>...<a href="#Submit_New_Picture">add a new Picture to the site?</a></li>
    <li>...<a href="#Search_Pictures">find a picture in the Pictures Area?</a></li>
    </ul>
    </p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
    </td>
  </tr>
  
  <tr>
    <td class="tSubTitle"><b><a name="Main_Directory"></a> Main Directory:</b></td>
  </tr>
  <tr>
    <td background="images/bg_help.gif">
    <br />
    
Upon clicking on the Main Directory menu option on the Nav Bar, you will see a page that 
displays a <b>Menu</b> on the left and the Picture Categories on the right. This menu is similar 
to the menu on the top of the page (see above), except that this menu block has an added 
Search field, in case you want to search for pictures from here.  This menu block will be available 
in all the Pictures pages, and will look and work as the example menu below:<br /><br />
<center>
<a name="#Menu"></a>
<table width=140 border=0><tr><td>
<%
'spThemeTableCustomCode = "align=""center"" width=""170"" cellpadding=0" 
spThemeTitle= "Pictures Menu:"
spThemeBlock1_open(intSkin)%> 
<table class="tPlain"><tr>
    <td nowrap WIDTH=140>
	<div class="menu">
      <a href="#New_Pictures" title="Link to the Newest pictures on this site...Click on THIS item NOW for more information on this subject.">- New Pictures<br /></a>
      <a href="#Popular_Pictures" title="Link to the Most popular pictures on this site...Click on THIS item NOW for more information on this subject.">- Popular Pictures<br /></a>
      <a href="#Top_Rated_Pictures" title="Link to the Top-ten-rated pictures on this site...Click on THIS item NOW for more information on this subject.">- Top Pictures<br /></a>
<%if not strDBNTUserName = "" then%>
      <a href="#Submit_New_Picture" title="Link to the area where you can add your pictures to this site...Click on THIS item NOW for more information on this subject.">- Submit Picture<br /></a><%end if%>
      <a href="#" title="The link that took you to this window...Click on THIS item NOW for more information on this subject.">- Picture FAQ<br /></a>
	  </div></td></tr></table>
<%spThemeBlock1_close(intSkin)%>
</td></tr></table>


Figure 2: The Main Pictures Menu.</center>

<br />
<p>
On the right pane, you will see a block containing the different Picture categories available, showing the pictures 
inside those categories:
<br /><br />
<center>
<img src="<%= strHomeUrl %>images/faq/pics_MainDir1.gif" alt="The Main Pictures Directory: Pictures categories available, showing the pictures inside those categories." width="311" height="89"></center>

<br /><br />
The example above shows the Category entitled "Category Name", with one Sub-Category 
under it, entitled "Logo".  Next to the Sub-Category name, you will see a number in 
parenthesis, indicating how many actual pictures there are under that Sub-Category.
<br />
If you click on the Sub-Category link, you will be taken to a page that shows all 
the pictures in that Sub-Category.  A block will be displayed, showing the header with the 
Site's title, the  Sub-Category name, followed by the number of pictures in the sub-category 
being displayed, like this:
<br />
<b>Category: Sub_Category (# Pictures - count includes private pictures)</b>
<br /><br />
Depending on how the picture was added to the site, this block will show either a thumbnail view or 
the full image of that picture in question.
<br />
If the image file was added to the database <b>after</b> your <u>last visit</u>, then a <%= icon(icnNew1,"New Item","","","align=""middle""") %> image will be 
displayed next to the picture's <b>Title</b>, just under the header, followed by the <b>Date</b> added and the 
<b>picture description</b>.
<br />
If there is more than 1 page to display the image index for that Sub-Category, you will see the paging options at the 
bottom of the block, stating the page number you are in and the number of total pages to be displayed.  You 
can navigate through the whole picture collection for this Sub-Category this way.  The last link you see is the 
"<b><a href="#Submit_New_Picture">Add a Picture<a/></b> " link, to submit your own pictures to the database.
<br />
If you click on either the picture Title or the picture itself, the system will take you to the picture's own page. 
<br />
The page to which you arrive shows all the details of the picture you clicked on. Following is some of the information 
displayed along with the picture itself:
<UL>
<LI>Picture Title
<LI>Copyright information
<LI>Picture Description
<LI>Number of Hits (# of visitors to that picture)
<LI><a href="#Pic_Rating">Rating</a> score (see <a href="#Pic_Rating">below</a>)
<LI>Date added
<LI>Username who submitted the picture
<LI>Comments relevant to the picture displayed
<LI>Link to Rate this picture
<LI>Link to <a href="#report_bad_link">Report a bad link</a> for this picture
</UL>
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
</td>
  </tr>  
  
  <tr>
    <td class="tSubTitle"><b><a name="Pic_Rating"></a> Picture Rating:</b></td>
  </tr>
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
<br />
When you click on the "Rate this Picture" link, a small window will be displayed, showing the following:
<br /><br />
<%
spThemeTitle= "Rate this picture"
spThemeBlock1_open(intSkin)%>
  <table class="tPlain"><tr>
    <td>
      <li>Rating Scale: <span class="fAlert">1</span> = worst, <span class="fAlert">10</span> = best</li>
      <li>Please try to be objective.</li>
      <li>Only registered members can rate pictures.</li>
      <li>You can only rate a picture once.</li>
      <li>Feel free to add a comment about this picture.</li>
    </td>
  </tr>
  <tr>
    <td>
      
      <b>Rating:</b> 
      <select name="rating">
        <option selected value="">-</option>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option value="10">10</option>
      </select>
      
    </td>
  </tr>
  <tr>
    <td>
      
      <b>Comments:</b> <br />
      <textarea rows=8 cols=33 name=comments title="Enter your comments about the picture in this box...This form is deactivated. It is for demonstration purposes only.">This is just an example of where you should write comments for any given picture, along with the options above.</textarea>
      
    </td>
  </tr>
  <tr>
    <td align=center>
    <input type="button" value=" Submit " class="button" title="Click on this button to submit your Rating for this picture...This button is deactivated in this sample form. It is for demonstration purposes only.">  <input type=reset value=" Clear " class="button" title="Click on this button to Clear/Erase your selections and any comments you may have entered...This button is deactivated in this sample form. It is for demonstration purposes only."></td>
  </tr></table>
<%spThemeBlock1_close(intSkin)%>

<br /><br />
The above block is self-explanatory.  You can <b>rate</b> and <b>add a comment</b> relative to the picture you just saw.
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
</td>
  </tr>  
  
  <tr>
    <td class="tSubTitle"><b><a name="report_bad_link"></a> Reporting a bad link:</b></td>
  </tr>
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
<br />
Clicking on this link will open a small window, returning a message like the following:<br /><br />
<table border=1 align=center cellpadding=3 width="316">
<tr><td align=center width="302">

<b>Thank you for reporting the bad picture: 
<br /><br />
We will correct the problem as soon as possible
<br /><br />
Close Window</b>
</td></tr></table>
<br /><br />
Bear in mind that this message is sent to the site administrator(s), and it carries the picture's link 
faulty information, together with the username of the person who actually reported the bad link, for 
logging purposes.
<br />
</p>

<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
</td>
  </tr>  
  
  <tr>
    <td class="tSubTitle"><b><a name="New_Pictures"></a> New Pictures:</b></td>
  </tr>
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
    <br />
<p>
Selecting this menu option will display a list of pictures that have been added to 
the Pictures section <u>within the last week</u>. The page will display the last seven days 
with a number enclosed in parentheses next to the name of the day. If the number is not zero, then a 
picture has been added on that day. The New Pictures list shows the day 
the item was entered. 
<br />
To view the <b>details</b> of a picture that has been added, 
click on the name of the day and the site will display the pictures that were 
entered on that day, ordered by their title. Next, click on the <b>title</b> of the picture and 
the picture details will be displayed.
<p>The <b>Details</b> displayed are: 
<UL>
<LI># of Hits
<LI>rating
<LI>date added
<LI>author/source
<LI>author's email/website
<LI>username of the picture's poster and 
<LI>comments
</UL>
</p>
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br /></td>
  </tr>
  
  <tr>

    <td class="tSubTitle"><b><a name="Popular_Pictures"></a> Popular Pictures:</b></td>
  </tr>
  
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
<br /><p>
Selecting this menu option will display the top-10 pictures, ordered by hit count. You will 
also see the details under the pictures. When you click on the items you 
can switch to the details page and vote/rate the pictures too.</p>
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
</td>
  </tr>
  
  <tr>
    <td class="tSubTitle">
    <b><a name="Top_Rated_Pictures"></a> Top-Rated Pictures:</b></td>
  </tr>
  
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
	  <br /><p>
	Selecting this menu option will display the top-10-rated pictures. You will see a 
	hit count number along with date added, rating and votes given to the picture in question. 
	When you click on the items you can switch to the details page.</p><br />
	<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p><br />
	<a name="Submit_New_Picture"></a><br /></td>
  </tr>
  
  <tr>
    <td class="tSubTitle">
    <b> Submitting New Pictures:</b></td>
  </tr>
  
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
	<br />
	<p align="left"><b>To submit your pictures:</b><br />
	The first thing you have to do to submit one or more pictures to the site is to 
	select the Category where you wish to place your picture.  <u>Make sure you select the right Category</u>. 
	<br />The Category combo-box selector will look like this:
	<br />
	<b>Category:</b> <select><option>Category 1</option><option>Category 2</option><option>Category 3</option><option>etc, etc.</option></select>
	<br />
	Then, enter your picture's <b>Title</b> (required field) and its <b>Description</b> in the corresponding fields.
	<br />
	You can also include some <b>Keywords for search</b> related to the picture you are now 
	adding, which will eventually help your peer users to search within the Picture Gallery and find that 
	particular picture or the rest of the pictures in the collection.
	<br /><br />
	Next, specify the <i>current location</i> of the picture you want to submit.  You have <b>two</b> options for adding pictures to 
	the <%=strSiteTitle%>'s Picture Gallery: Using a picture located or hosted <u>at another website</u>, or <b>upload</b> your 
	picture to <%=strSiteTitle%>'s server from your own local hard disk.  Following is a brief explanation of both options:</p>
  <a name="URL"></a>
  <b>1). <u>External Reference</u> to <%=strSiteTitle%>:</b>
  <br />One way to add your pictures to this site is writing the web address (or URL) 
  of a picture that is currently being hosted on another website, different or foreign to <%=strSiteTitle%>.  For instance, if 
  you know that a picture resides (is located or hosted) in a website called <b>http://www.website.com</b>, and that picture's 
  address is <b>http://www.website.com/photos/<span class="fAlert">my_photo.jpg</span></b>, then you would enter that 
  URL (web address) into the <b>URL</b> field.
  <br />
  Make sure the address of your picture starts with the <b><i>http://</i></b> prefix.<br />
  This option has one <u>advantage</u>: since the picture itself is hosted or "lives" in <u>another</u> website, different to <%=strSiteTitle%>, 
  that picture will not "waste" any available (and valuable) disk space on <%=strSiteTitle%>'s server, while it will be displayed on 
  <%=strSiteTitle%> and keep being hosted on the <i>other</i> server (the other website's server). 
  <br />
  The <u>disadvantage</u> is that the picture depends on the good performance of the other site for it 
  to be properly displayed on <%=strSiteTitle%>. <br /><br />

  <a name="Uploading"></a>
  <b>2). <u>Uploading</u> your pictures from your hard disk:</b> 
  <br />The other way to add pictures to the site is to literally copy a picture that you have saved in your hard disk and "paste it" on 
  <%=strSiteTitle%>'s server hard disk (we will call this process "<b>Uploading a Picture"</b>).
  <br />
  To upload a picture from your local disk, press the <input type="button" class="button" value="Browse" title="Use this button to start searching for your picture on your local hard disk...This button is deactivated on this page. It is for demonstration purposes only."> button next to the 
  field entitled  "<b>Upload Image:</b>".
  <br />
  This will launch a dialog window that helps you select your picture in any folder in your computer. Select the image 
  you want to upload, and press the <b>Open</b> button (or simply double-click the file desired).  Once you have selected your file, 
  its location or path (in your disk) will be inserted into the Upload Image field.
  <br /><br />
  The allowed extensions for your picture are <b>.gif</b> and <b>.jpg</b>
  <br /><br />
  <span class="fAlert"><b>NOTE: Make sure that your picture is <u>less than 1000 KB</u></span>.</b>  
  This is for obvious reasons.  If we were to allow pictures larger than that (weight in KB), we would rapidly fill our servers with picture 
  files too big and heavy, which would result in making the site slower when displaying these images.
  <br />
  <span class="fAlert"><b>IMPORTANT:</b></span> Keep in mind that the 2 methods for submitting an image to the site are <b><u>exclusive</u></b>, meaning that you have to use 
  <u>either one or the other</u>.  If using <a href="#URL">Method #1</a>, then <b>don't</b> enter anything in the "Upload Image" field.  Conversely, if you 
  use <a href="#Uploading">Method #2</a>, then <b>don't</b> enter anything in the "URL" field.
  <br /><br />
  
  If you are using method 1 (External Reference to <%=strSiteTitle%>), and you know that the picture you're adding has a <b>thumbnail</b> image, 
  and you know the location (URL) of that thumbnail image, you can enter its URL into the <b>Thumbnail URL</b> field.  If you leave this 
  field empty, the picture's actual URL will be used in place of the thumbnail (somehow having both functions: Picture & Thumbnail in one).  
  Make sure you leave the http:// prefix intact if not using this field. 
  <br />
  You can enter any <b>Copyright</b> information related to the picture you're submitting.  Remember that many pictures or photos that are 
  published on the Internet have certain Copyrights granted by default, while you should give credit to the author of the picture, if you know it, or the 
  site's name from which you have "harvested" the picture.  This is not a required field.
  <br /><br />
  Finally, check the "<b>Private</b>" checkbox <b><u>if</u></b> you want your picture to only be available to you.  What this means is that <u><b>only you</b></u> 
  (<i>and the administrator</i>) will be able to see this picture, and that it will not be displayed to other people in the general collection or 
  Gallery.
<br /><br />
  Once you have finished with all the picture's information, press the <input type="button" class="button" value="Submit" title="Click on this button to submit your picture to the site...This button is deactivated in this window. It is for demonstration purposes only."> button in order to send it for approval, before it's actually authorized by an administrator to be published. You will be notified by 
  Email when your picture is authorized and gets added to our database.  All images submitted will undergo the authorization process, again, 
  for obvious reasons. </p>

<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />

	</td>
  </tr>
  
  <tr>
    <td class="tSubTitle"><b><a name="Search_Pictures"></a> Search Pictures:</b></td>
  </tr>
  <tr>
    <td background="<%= strHomeUrl %>images/bg_help.gif">
<br /><p>
  You can use the search area on the <b><a href="#Menu">left menu</a></b> by entering your search term or terms into the Search box and then click 
  the <input type="button" class="button" value="Search" title="Click on this button to proceed with your search...This button is deactivated in this sample menu. It is for demonstration purposes only."> button to search through the Pictures section. Pictures with 
  matching text will be shown in the display box on the right. 
  <br />
  You can also set the number of results by selecting any of the radio buttons search options (see "<a href="#Menu">left menu</a>"). Each of these radio buttons 
  corresponds to whether you want 10, 20 or 30 results per page (refer to the paging section, above).  The Search routine looks 
  through the Pictures Titles and Descriptions when trying to find a match. The titles of the results that are displayed 
  are hyperlinks tied to the thumbnails, which can be clicked to take you to the page with the details of the picture or pictures found.
</p>
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page" /></a></p>
<br /><br />
	</td>
  </tr></table>
<% spThemeBlock1_close(intSkin)
end sub
 %>