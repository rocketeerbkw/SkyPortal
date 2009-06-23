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
curpagetype = "article"
%>

<!-- #INCLUDE FILE="lang/en/article_lang.asp" -->
<!-- #include file="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #INCLUDE FILE="Modules/articles/article_functions.asp" -->
<!--include file="modules/rss/rss_functions.asp" -->
<%
dim iMode, iCmd, cid, app
iMode = 0
iCmd = 0
cid = 0
app = ""
intSkin = 1

if IsNumeric(Request("mode")) = True then
	iMode = cLng(Request("mode"))
else
  if Request("mode") <> "" then
	sMode = Request("mode")
  else
	closeAndGo("stop")
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
else
	'closeAndGo("stop")
end if

if Request("sid") <> "" and Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sid = cLng(Request("sid"))
	else
		closeAndGo("stop")
	end if
end if

select case iMode
  case 1
    app = "ARTICLE"
	adtyp = "added to"
	hp = 1
  case 2
    app = "ARTICLE"
	adtyp = "removed from"
	hp = 0
  case 10 'FAQ
	spThemeTitle = "Articles FAQ"
	showFAQ()
  case else
end select

if sMode = "print" then
    spThemeShortBodyTag = " leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" bgcolor=""#FFFFFF"""
end if
%>
<!--#include file="inc_top_short.asp" -->
<%
setAppPerms curpagetype,"iName"

  if sMode <> "" and cid >= 0 then
	select case sMode
	  case "print" 'print item
		printItem()
	end select
  end if
  if sMode <> "" and cid >= 0 and strUserMemberID > 0 then
	select case sMode
	  case "rate"
		Call mod_rateItem(cid,intAppID,item_tbl,item_fld,"TITLE",art_Comments,art_Rate)
	  case "emailitem"
		emailToFriend()
	  case "addsub"
		call addSubscription()
	  case "delsub"
		call mod_delSubscription(cid)
	  case "addbook"
		call addBookmark()
	  case "delbook"
		call mod_delBookmark(cid)
	  case "editAccess" 'edit group access form
		editAccessForm()
	  case "updAccess" 'update group access
		updateAccess()
	end select
  end if

 if iMode > 0 and hasAccess("1,2") then
	  select case iMode
	    case 1, 2
		  Call mod_addFeatured(app,adtyp)
		case 4 'email item
		  'emailToFriend()
		case 5 'print item
		  'printItem()
		case 6 'rate item
		  'rateItem()
		  'sMode = iMode
		  'Call mod_rateItem(cid,intAppID,item_tbl,item_fld,"TITLE",art_Comments,art_Rate)
	  end select 
   end if
 %>
<!--#include file="inc_footer_short.asp" -->
<% 
sub addBookmark()
		 ' response.Write("Hello " & iCmd & "<br />")
	strMsg = ""
	memID = strUserMemberID
	set rsA = nothing
	if strMsg = "" and intBookmarks = 1 then
	  select case iCmd
	    case 1 'bookmark category
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that article exists
		    sSql = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article Category not found"
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
	        'Verify that subcat exists
		    sSql = "SELECT SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article SubCategory not found: " & cid
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
	        'Verify that article exists
		    sSql = "SELECT TITLE FROM ARTICLE WHERE ARTICLE_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article not found"
		    else 'article does exist, lets bookmark it
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
	sSql = "SELECT APP_ID,APP_SUBSCRIPTIONS FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'article'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	  if intSubscriptions = 1 then
	    intSubscriptions = rsA("APP_SUBSCRIPTIONS")
	  end if
	else
	  strMsg = "Module error in PORTAL_APPS"
	end if
	
	' check for module subscription
	if strMsg = "" and intSubscriptions = 1 then
	  sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=0"
	  set rsAp = my_Conn.execute(sSql)
	  If rsAp.BOF or rsAp.EOF Then
	    ' they are not subscribed to the module
	  else
	    strMsg = "Cannot add subscription" & "<br>"
	    strMsg = strMsg & "You are already subscribed" & "<br>"
		strMsg = strMsg & "to the Article Module"
	  end if
	end if
	
	set rsA = nothing
	if strMsg = "" and intSubscriptions = 1 then
	  select case iCmd
	    case 1 'bookmark category
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that article exists
		    sSql = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("CAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", " & cid & ", 0, 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = strMsg & "Category Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "an Article is added to the " & itmTitle & " category!<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Category subscription already exists"
	      End If
	      set rs = nothing
	    case 2 ' bookmark subcategory
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and SUBCAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that article exists
		    sSql = "SELECT SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article SubCategory not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("SUBCAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = strMsg & "SubCategory Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "an Article is added to the " & itmTitle & " SubCategory!<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "SubCategory subscription already exists"
	      End If
	      set rs = nothing
	    case 3 ' bookmark module
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=0"
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that article exists
		    sSql = "SELECT APP_NAME FROM "& strTablePrefix & "APPS WHERE APP_ID=" & intAppID
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Article module not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = "All New Articles"
			  
	          ' Subscription doesn't already exist so add it
			  'Delete existing Article subscriptions
	          insSql = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS"
	          insSql = insSql & " WHERE APP_ID=" & intAppID & " and M_ID=" & memID & ";"
	          executeThis(insSql)
			  
			  'Add module subscription
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, 0, 0, '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "Article Module Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "an Article is added to the database!<br />"
			  
	          strMsg = strMsg & "All of your previous Article Category and" & "<br />"
	          strMsg = strMsg & "SubCategory subscriptions have been deleted" & ".<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Article module subscription already exists"
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

sub printItem()
	strSQL1 = "SELECT count(*) as Comments FROM " & strTablePrefix & "M_RATING WHERE COMMENTS NOT LIKE ' ' AND ITEM_ID = " & cid & " AND APP_ID=" & intAppID & ""

	set rsArticleComments = server.CreateObject("adodb.recordset")
	rsArticleComments.Open strSQL1, my_Conn
		
	dim intVotes
	dim intRating
	if not rsArticleComments.eof then
		Comments = rsArticleComments("Comments")
	else
	    Comments = 0
	end if
	rsArticleComments.Close
	set rsArticleComments = nothing

	strSQL = "SELECT * from ARTICLE where ACTIVE = 1 and ARTICLE_ID = " & cid
	'response.Write(strSql)
	set rs = my_Conn.Execute (strSQL)
	if rs.eof then
	  Response.Write "Item not found"
	  set rs = nothing
	  closeAndGo("stop")
	end if
	
	dim strPoster
	strPoster = rs("POSTER")
	if len( Trim(strPoster)) > 0 then
	strPoster = strPoster 
	else
	strPoster = "Anonymous"
	end if

	dim strAuthor
	'strAuthor = rs("AUTHOR")
	if len( Trim(strAuthor)) > 0 then
	strAuthor = strAuthor
	else
	strAuthor = "n/a"
	end if
	
	dim strAuthorEmail
	'strAuthorEmail = rs("AUTHOR_EMAIL")
	if len( Trim(strAuthorEmail)) > 0 then
	strAuthorEmail = strAuthorEmail
	else
	strAuthorEmail = "n/a"
	end if

	title =rs("TITLE")
	parentid = rs("PARENT_ID")
	catid = rs("CATEGORY")
	postdate = ChkDate(rs("POST_DATE"))
	mainContent = formatStr(rs("CONTENT"))
	mainContent = HTMLDecode(mainContent)
	mainContent = replace(mainContent,"tiny_mce/",strHomeurl & "tiny_mce/")
	
	rs.close
	set rs = nothing
	 %>
<style type="text/css">
table td{color:#000000;}
</style>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
<td colspan="2" align="center" valign="top"><p align="center">
<a href="javascript:window.print()"><font color="#000000">Send To Printer</font></a></p></td>
</tr>
<tr>
<td  width="20%" nowrap><font color="#000000"><b>Title:</b></font></td>
<td  width="80%"><font color="#000000"><b><% =title%></b></font><br></td>
</tr>
<tr>
<td  width="20%" nowrap><font color="#000000"><b>Author:</b></font></td>
<td  width="80%"><font color="#000000"><b><% =strAuthor%></b></font><br></td>
</tr>
<tr>
<td  width="20%" nowrap><font color="#000000"><b>Author Email:</b></font></td>
<td  width="80%"><font color="#000000"><b><% =strAuthorEmail%></b></font><br></td>
</tr>
<tr>
<td  width="20%" nowrap><font color="#000000"><b>Date Posted:</b></font></td>
<td  width="80%"><font color="#000000"><b><%= postdate %></b></font><br></td>
</tr>
<tr>
<td width="100%" colspan="2">
<hr size="1"><table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<td><font color="#000000"><% =mainContent%></font></td></table><hr size="1"></td>
</tr>
<tr>
<td width="100%" colspan="2">
<%
  Response.Write("<p><hr /></p>")
  Response.Write("<p><font color=""#000000"">")
  Response.Write("<b>" & strSiteTitle & " :</b></font> <a href=""" & strHomeURL & """>")
  Response.Write("<font color=""#000000"">" & strHomeURL & "</font></a>")
  Response.Write("</p>")
  Response.Write("<p><font color=""#000000"">")
  Response.Write("<b>"  & strCopyright & "</b> ")
  Response.Write("</p>")
%>
    <p align="center"><a href="JavaScript:onClick= window.close()"><font color="#000000">Close Window</font></a></p><p>&nbsp;</p>
    </td>
  </tr>
</table>
<%
  closeAndGo("stop")
end sub

sub emailToFriend()
  response.Write("<br><br>")
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
		if (rs.EOF and rs.BOF) then
			if strLogonForMail = 1 then 
				Err_Msg = Err_Msg & "<li>You must be registered to email this article</li>"
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
  
  <form action="article_pop.asp" method=post id=Form1 name=Form1>
  <input type=hidden name="cmd" value="6">
  <input type=hidden name="mode" value="emailitem">
  <input type=hidden name="cid" value="<%= cid %>">
<%
%>
      <table><TR>
        <TD align="center" colspan="2" class="fTitle" nowrap="nowrap"><p>Send Article to a Friend<br>&nbsp;</p></td>
      </tr>
      <TR>
        <TD align="right" nowrap><b>Send To Name:&nbsp;</b></td>
        <TD><input type=text name="Name" size=25></td>
      </tr>
      <TR>
        <TD align="right" nowrap><b>Send To Email:&nbsp;</b></td>
        <TD><input type=text name="Email" size=25></td>
      </tr>                
      <tr>
        <td align="right" nowrap><b>Your Name:&nbsp;</b></td>
        <td><input name=YName type=<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YName %>" size=25> <% if YName <> "" then Response.Write(YName) end if %></td>
      </tr>
      <tr>
        <td align="right" nowrap><b>Your Email:&nbsp;</b></td>
        <td><input name=YEmail type=<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YEmail %>" size=25> <% if YEmail <> "" then Response.Write(YEmail) end if %></td>
      </tr> 
      <tr>
        <td colspan=2 nowrap><b>Message:</b></td>
      </tr>
      <tr>
        <td colspan=2 align=center><textarea name="Msg" cols="38" rows=5 readonly>Hi, <% =vbCrLf %>I thought you might be interested in this Article:<%= vbCrLf & vbCrLf & strHomeUrl & "article_read.asp?item=" & cid %></textarea></td>
      </tr>                    
      <tr>
        <td colspan=2 align=center><input class="button" type=submit value="Send" id=Submit1 name=Submit1></td>
      </tr></table>
  </form>
<%
end if
spThemeBlock1_close(intSkin)
end sub

sub addFrontPage(typ)
  response.Write("<br><br>")
  spThemeBlock1_open(intSkin)
  response.Write("<br><br>")
  if hasAccess(1) then
	strSql = "UPDATE " & typ & " set FEATURED = " & hp
	strSql = strSql & " WHERE " & typ & "_ID = " & cid
	executeThis(strSql)
%>
	<P align=center><b><%= uCase(typ) & " " & adtyp %> home page items</b><br></P><script type="text/javascript"> opener.document.location.reload();</script>
<%
  else %>
	<p align=center><b>Only administrators can perform this action.</b></p>
<%
  end If
  response.Write("<br><br>")
	spThemeBlock1_close(intSkin)
end sub

sub showFAQ()
spThemeBlock1_open(intSkin) %>
  <table><tr>
    <td>
<p><br>
This is the place where you can share your writings/articles. When you click on &quot;Articles&quot; link you will see a menu on the left and the categories on the right. You can also use the menu at the top of the page. Click on the below questions in order to get the answer.<br>
<br></p></td>
  </tr>
  <tr>
    <td>
    <p><b>How do I...</b>
    <ul>
    <li><a href="#New_Articles">...get a list of the Newest Articles?</a></li>
    <li><a href="#Popular_Articles">...get a list of the most Popular Articles?</a></li>
    <li><a href="#Top_Rated_Articles">...get a list of the Top Rated Articles?</a></li>
    <li><a href="#Adding_New_Articles_to_the_Articles">...add a new Article to the site?</a></li>
    <li><a href="#Search_Articles">...find an article in the Articles?</a></li>
    <li><a href="#Pop-Up_Article_Detail_View">...view the Details of an Article?</a></li>
    </ul>
    <br></p>
    </td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="New_Articles"></a>&nbsp;New Articles</b></td>
  </tr>
  <tr>
    <td><br>
<p>
Selecting this menu option will display a list of articles that have been added to 
the Articles section within the last week. The page will display the last seven days 
with a number next to the name of the day. If the number is not zero, then an 
article has been added on that day. The New Articles list shows the day 
the item was entered, to view the details of an item that has been entered, 
click on the name of the day and the site will display the articles that were 
entered on that day by their title. Next click on the title of the article and 
the article details will be displayed.<p>Details are: Hits, rating, 
date added, author/source, author's email/website, posted by and comments.</p>
<p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" width="15" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br /></td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Popular_Articles"></a>&nbsp;Popular Articles</b></td>
  </tr>
  <tr>
    <td>
<br><p>
Selecting this menu option will display top 10 articles by hit count. You will 
also see the details under the articles. When you click on the items you 
can switch to the details page and vote too.</p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of Page"></a></p>
<br /><br />
  <tr>
    <td class="tSubTitle"><b><a name="Top_Rated_Articles"></a>&nbsp;Top Rated Articles</b></td>
  </tr>
  <tr>
    <td>
	  <br><p>Selecting this menu option will display top 10 rated articles. You will see a 
hit count number along with date added, rating and votes given. When you click on the items you can switch to the details page.</p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Adding_New_Articles_to_the_Articles"></a>&nbsp;Adding New Articles</b></td>
  </tr>
  <tr>
    <td>
<br>
<p>You can submit an article in two ways:
Choosing to "Submit Article" from the Articles top menu is the first choice that is made when entering a new article on the Articles section. Other way is first selecting a category then a subcategory where the article wanted to be located and click on the "Add an Article" link towards the bottom.<p>
Selecting in either way it will open a submit form for you to fill out. Fill in the provided form. The first selection that can be made on the page is to select a proper category for the article that will be submitted. Although you will see the category name here if you firstly selected the subcategory then submit. If you did not see a category that fit your article's description, contact with the admins of the site and they'll be happy to add it for you :)<p>
Then you have to complete the "Title", "Content", "Summary" and "Your Email" areas those required. You can also use Forum Codes and Smilies when writing the content. You can go on with the other fields as needed but they are not the required ones.<p>
Once you have filled out all of the items for your article, click on the "Submit" button to submit your article to the site administrators. You can use "Reset" button to clear the form. If the site is configured to require approval for new articles then you will receive a message that your article submission will be reviewed by the administrator and if approved will then appear on the Articles section. Until the article is approved it will not show up on the Articles. If the article is not approved by the Administrator the article will be deleted and not show up on the Articles, although you will receive an email to let you know whether or not your article has been approved. If the site is configured to accept the requests online then a summary page about your article will be shown and added to the Articles database immediately.</p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Search_Articles"></a>&nbsp;Search Articles</b></td>
  </tr>
  <tr>
    <td>
<br><p>
You can use the search area on the left directly by writing search term or terms into the Search box then click the Search button to search through the articles on the Articles section. Articles with matching text will be shown in the display box on the right. Also you can set the number of results by selecting thru the radio buttons. The Search routine looks through Article Descriptions when trying to find a match. The titles of the results that are displayed are hyperlinks which can be clicked on to pop-up another page with the details of the article.
</p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Pop-Up_Article_Detail_View"></a>&nbsp;Pop-Up Article Detail View</b></td>
  </tr>
  <tr>
    <td>
<br><p>
Clicking on the title of an article will pop-up another page with the Article Details. This page will show the Article Title followed by the contents of the article. Other details such as Hits, Rating, Added Date, Author, etc will be shown at the bottom. Only the administrators can edit the article even if you have submitted the article. You can also rate the article and make a comment about.
</p>
    <p align=right><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0" alt="Top of this Page"></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td align="center">
	<p></p></td>
  </tr></table>
<% spThemeBlock1_close(intSkin)
end sub

sub createXmlForm()
  dim sfTitle, sfSummary, sfURL, sfPublic, sfActive, sfFeatured
  dim sfWL, sfYH, sfBL, sfNG, sfGG, gRead, sfSCATid, sfID
  sfSCATid = 0
  sfMode = 98
  sfPublic = 1
  sfFeatured = 1
  sUrl = ""
	
	sfPHTZone = "EDT"
	sfPHDisplay = "10"
	sfPHRead = strHomeURL & "ArticleRead.asp?item="
	sfPHImage = strHomeURL & "files/rss/images/rss_news.gif"
	
	sfPFTable = "ARTICLE"
	sfPFid = "ARTICLE_ID"
	sfPFTitle = "TITLE"
	sfPFAuthor = "POSTER"
	sfPFAuthorInfo = "POSTER_EMAIL"
	sfPFSummary = "SUMMARY"
	sfPFPostDate = "POST_DATE"
	sfPFWhere = "ACTIVE=1 AND CATEGORY=" & sid
	sfPFOrderBy = "ARTICLE_ID DESC"
	
	sfMCat = 0
	sfMSCat = sid
	
	sfPModuleID = intAppID
	sfActive = 1
	sfWL = 1
	sfYH = 1
	sfBL = 1
	sfNG = 1
	sfGG = 1
	gRead = "1,2,3"
  
    sSql = "SELECT * FROM PORTAL_M_SUBCATEGORIES WHERE SUBCAT_ID=" & sid
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  sfTitle = rsA("SUBCAT_NAME")
	  sfSummary = rsA("SUBCAT_SDESC")
	  sUrl = replace(rsA("SUBCAT_NAME")," ","_")
	  if len(sUrl) > 10 then
	    sUrl = left(sUrl,10)
	  end if
	  sUrl = sUrl & "_" & sid & ".xml"
	end if
	set rsA = nothing
	sfURL = sUrl
	
  spThemeTitle = "Add Site RSS Feed"
  spThemeBlock1_open(intSkin) %>
<script type="text/javascript">
function val_xmlfeed()
{
//var at=document.getElementById("email").value.indexOf("@")
var c_cat=document.getElementById("cat").value
var c_title=document.getElementById("sfTitle").value
var c_file=document.getElementById("sfFile").value
var alMsg = ""
submitOK="true"

if (c_cat == 0)
 {
 alMsg += "\nPlease select a Category for your feed\n";
 submitOK="false";
 }
if (c_title.length>49)
 {
 alMsg += "\nYour TITLE must be less than 50 characters\n";
 submitOK="false";
 }
if (c_title.length<1)
 {
 alMsg += "\nPlease fill in your TITLE\n";
 submitOK="false";
 }
 if (!CheckSql(c_title)) {
 alMsg += "\nYour TITLE cannot contain any of the\nfollowing characters:  \\ / * \" < > | [ ] ? ;\n";
 submitOK="false";
 }
if (c_file.length<1)
 {
 alMsg += "\nPlease fill in your Feed URL\n";
 submitOK="false";
 }
if (submitOK=="false")
 {
 alert(alMsg);
 return false
 }
}
function CheckSql(str) {
	var re;
	re = /[\\\/\]\[:\;*?"<>%|]/gi;
	if (re.test(str)) return false;	
	else return true;
}

</script>
  <form name="sfFeed" method="post" id="sfFeed" action="SkyPublisher.asp" onSubmit="return val_xmlfeed();">
        <table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="40%">&nbsp;</td>
            <td width="60%">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" align="center" class="tTitle">Create a new Site Feed 
              XML link for your site</td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class="tSubTitle">General Feed Information 
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td align="right"><b>Add to XML Category:</b> </td>
				<%
				rssCat = sfSCATid
				call xml_SubCats("xml","cat")
				%>
          </tr>
          <tr> 
            <td align="right"><b>Add to Feed Library:</b> </td>
				<%
				rssCat = sfSCATid
				call xml_SubCats("skyfeedreader","feedLib")
				%>
          </tr>
          <tr> 
            <td align="right"><b>Friendly Name:</b> </td>
            <td> 
              <input name="sfTitle" type="text" id="sfTitle" value="<%= sfTitle %>">
            </td>
          </tr>
          <tr> 
            <td align="right" valign="top"><b>Friendly Summary:</b> </td>
            <td>
              <textarea name="sfSummary" rows="5" wrap="VIRTUAL" id="sfSummary"><%= sfSummary %></textarea>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class="tSubTitle">Construction of this Site Feed 
            </td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" align="center">
              <fieldset style="margin:5px;width:450px;">
			  <legend><b>XML Configuration</b></legend>
              <table border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td width="35%">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" align="center"></td>
          </tr>
          <tr> 
            <td align="right"><b>RSS Title:</b> </td>
            <td> 
              <input name="sfPHTitle" type="text" id="sfPHTitle" value="<%= sfTitle %>">
            </td>
          </tr>
          <tr> 
            <td align="right" valign="top"><b>RSS Description:</b> </td>
            <td>
              <textarea name="sfPHSummary" cols="25" rows="5" wrap="VIRTUAL" id="sfPHSummary"><%= sfSummary %></textarea>
            </td>
          </tr>
          <tr> 
            <td align="right"><b>RSS Image:</b> </td>
            <td> 
              <input name="sfPHImage" type="text" id="sfPHImage" value="<%= sfPHImage %>">
            </td>
          </tr>
          <tr> 
            <td align="right"><b>RSS Filename:</b> </td>
            <td> 
              <input name="sfFile" type="text" id="sfFile" value="<%= sfURL %>">
			</td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
              </table>
              </fieldset>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" align="center">
              <input name="cmd" type="hidden" id="cmd" value="<%= sfMode %>">
              <input name="mode" type="hidden" id="mode" value="11">
  			  <input name="app" type="hidden" id="app" value="<%= curpagetype %>">
              <input type="submit" name="Submit" value="Submit">
			</td>
          </tr>
          <tr> 
            <td>&nbsp;
  <input name="sfMCat" type="hidden" id="sfMCat" value="<%= sfMCat %>">
  
  <input name="sfPHDisplay" type="hidden" id="sfPHDisplay" value="<%= sfPHDisplay %>">
  <input name="sfPHTZone" type="hidden" id="sfPHTZone" value="<%= sfPHTZone %>">
  <input name="sfPHRead" type="hidden" id="sfPHRead" value="<%= sfPHRead %>">
  
  <input name="sfPModuleID" type="hidden" id="sfPModuleID" value="<%= intAppID %>">
  <input name="sfMCat" type="hidden" id="sfMCat" value="<%= sfMCat %>">
  <input name="sfMSCat" type="hidden" id="sfMSCat" value="<%= sfMSCat %>">
  <input name="sfPFTable" type="hidden" id="sfPFTable" value="<%= sfPFTable %>">
  <input name="sfPFid" type="hidden" id="sfPFid" value="<%= sfPFid %>">
  <input name="sfPFTitle" type="hidden" id="sfPFTitle" value="<%= sfPFTitle %>">
  <input name="sfPFAuthor" type="hidden" id="sfPFAuthor" value="<%= sfPFAuthor %>">
  <input name="sfPFAuthorInfo" type="hidden" id="sfPFAuthorInfo" value="<%= sfPFAuthorInfo %>">
  <input name="sfPFSummary" type="hidden" id="sfPFSummary" value="<%= sfPFSummary %>">
  <input name="sfPFPostDate" type="hidden" id="sfPFPostDate" value="<%= sfPFPostDate %>">
  <input name="sfPFWhere" type="hidden" id="sfPFWhere" value="<%= sfPFWhere %>">
  <input name="sfPFOrderBy" type="hidden" id="sfPFOrderBy" value="<%= sfPFOrderBy %>">
  			</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
      </form>
  <%
  spThemeBlock1_close(intSkin)
end sub
 %>