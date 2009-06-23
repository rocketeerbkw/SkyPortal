<!-- #include file="config.asp" -->
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
dim curpagetype
curpagetype = "downloads"
%><!-- #include file="modules/downloads/dl_config.asp" -->
<!-- #include file="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="modules/downloads/dl_functions.asp" -->
<%
dim iMode, sMode, iCmd, cid, sid, app, intSkin
iMode = 0
sMode = ""
iCmd = 0
cid = 0
sid = 0
app = ""
intSkin=1


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
    app = "DL"
	adtyp = "added to"
	hp = 1
  case 2
    app = "DL"
	adtyp = "removed from"
	hp = 0
  case 12 'FAQ
	spThemeTitle = "Downloads FAQ"
	showFAQ()
  case else
end select
%>
<!--#include file="inc_top_short.asp" -->
<%
setAppPerms CurPageType,"iName"

'strUserMemberID = getmemberid(strdbntusername)
  if sMode <> "" and cid >= 0 and strUserMemberID > 0 then
	select case sMode
	  case "rate"
		call mod_rateItem(cid,intAppID,"DL","DL_ID","NAME",dl_Comments,dl_Rate)
	  case "emailitem"
		'emailToFriend()
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
	  'Response.Write("Hello!")
	  'response.End()
		updateAccess()
	end select
  end if
 if iMode > 0 and cid >= 0 then
	  select case iMode
	    case 1, 2
		  if strUserMemberID > 0 then
		    Call mod_addFeatured(app,adtyp)
		  end if
		case 4 'dl_goto.asp
		  if strUserMemberID > 0 or dl_GuestsCanDL then
		    dlGoTo(cid)
		  end if
		case 6 'report bad item
		  if strUserMemberID > 0 then
		    badLink()
		  end if
		case 7 'delete comment
		  'deleteComment(cid,intAppID,"DL","DL_ID")
		case 8 'email friend
		  if strUserMemberID > 0 then
		    emailToFriend()
		  end if
	  end select
   end if
 %>
<!--#include file="inc_footer_short.asp" -->
<% 
sub addBookmark()
		 ' response.Write("Hello " & iCmd & "<br />")
	strMsg = ""
	memID = getmemberID(strDBNTUserName)
	sSql = "SELECT APP_ID,APP_BOOKMARKS FROM "& strTablePrefix & "APPS WHERE APP_iNAME = '" & CurPageType & "'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	  if intBookmarks = 1 then
	    intBookmarks = rsA("APP_BOOKMARKS")
	  end if
	else
	  strMsg = "Module error in PORTAL_APPS"
	end if
	set rsA = nothing
	if strMsg = "" and intBookmarks = 1 then
	  select case iCmd
	    case 1 'bookmark category
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'download does not exist
		      strMsg = "Download Category not found"
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
	        'Verify that download exists
		    sSql = "SELECT SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'download does not exist
		      strMsg = "Download SubCategory not found"
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
	        'Verify that download exists
		    sSql = "SELECT NAME FROM DL WHERE DL_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'download does not exist
		      strMsg = "Download not found"
		    else 'download does exist, lets bookmark it
		      itmTitle = rsA("NAME")
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
	sSql = "SELECT APP_ID,APP_SUBSCRIPTIONS FROM "& strTablePrefix & "APPS WHERE APP_iNAME = '" & CurPageType & "'"
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
	    strMsg = "Cannot add subscription" & "<br />"
	    strMsg = strMsg & "You are already subscribed" & "<br />"
		strMsg = strMsg & "to the Download Module"
	  end if
	end if
	
	if strMsg = "" and intSubscriptions = 1 then
	  select case iCmd
	    case 1 'subscribe to category
	       sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	       set rs = my_Conn.execute(sSql)
	       If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'download does not exist
		      strMsg = "Download Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("CAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", " & cid & ", 0, 0, '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "Category Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "a new Download is added to the " & itmTitle & " category!<br />"
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
	        'Verify that download exists
		    sSql = "SELECT SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'download does not exist
		      strMsg = "Download SubCategory not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = rsA("SUBCAT_NAME")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = strMsg & "SubCategory Subscription Added!" & "<br /><br />"
	          strMsg = strMsg & "You will now receive an email when" & "<br />"
	          strMsg = strMsg & "a new Download is added to the" & " " & itmTitle & " " & "SubCategory<br />"
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
	        'Verify that download exists
		    sSql = "SELECT APP_NAME FROM "& strTablePrefix & "APPS WHERE APP_ID=" & intAppID
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'item does not exist
		      strMsg = "Downloads module not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = "All New Downloads"
	          ' item doesn't already exist so add it
			  
	          insSql = "DELETE FROM "& strTablePrefix & "SUBSCRIPTIONS"
	          insSql = insSql & " WHERE APP_ID=" & intAppID & " AND M_ID=" & memID & ";"
	          executeThis(insSql)
			  
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, 0, 0, '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "Downloads Module Subscription Added!" & "<br /><br />"
	          strMsg = strMsg & "You will now receive an email when" & "<br />"
	          strMsg = strMsg & "any new Downloads are added to the database" & ".<br /><br />"
			  
	          strMsg = strMsg & "All of your previous Download Category and" & "<br />"
	          strMsg = strMsg & "SubCategory subscriptions have been deleted" & ".<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Downloads module subscription already exists"
	      End If
	      set rs = nothing
		case else
		'do nothing
	  end select
    end if 'strMsg = ""
  response.Write("<br /><br /><br />")
    strMsg = jsReloadOpener & strMsg
    call showMsgBlock(1,strMsg)
  response.Write("<br /><br /><br />")
end sub

sub dlGoTo(typ)
	lastdate = chkString(Request.Cookies("date"),"sqlstring")
	lastid = chkString(Request.Cookies("dlid"),"sqlstring")
  	dlid = typ

	if lastid <> dlid then
		Response.Cookies("dlid") = dlid
		Response.Cookies("dlid").Expires = dateadd("d",7,strCurDateAdjust)
	    executeThis("UPDATE DL SET HIT = HIT + 1 Where DL_ID =" & dlid)
		resetCoreConfig()
	end If
	dim rs
	Set rs = my_Conn.Execute("SELECT URL FROM DL WHERE DL_ID = " & dlid)
	linkurl = rs("URL")
	set rs = nothing
  'response.Write(linkurl)
  closeAndGo(linkurl)	
end sub

sub emailToFriend()
  response.Write("<br /><br />")
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
		    <br /><p align="center"><span class="fTitle">Email has been sent</span></p>
			<p><a href="JavaScript:onClick= window.close()">Close Window</a><br />&nbsp;</p>
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
  
  <form action="dl_pop.asp" method="post" id="Form16" name="Form16">
  <input type="hidden" name="cmd" value="6" />
  <input type="hidden" name="mode" value="8" />
  <input type="hidden" name="cid" value="<%= cid %>" />
<%
%><table cellpadding="0" cellspacing="0" width="100%">
      <TR>
        <TD align="center" colspan="2" class="fTitle" nowrap="nowrap"><p>Send info to a Friend<br />&nbsp;</p></td>
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
        <td colspan=2 align=center><textarea name="Msg" cols="38" rows="5" readonly="readonly">Hi, <% =vbCrLf %>I thought you might be interested in this file available for download at <%=strSiteTitle%>:<%= vbCrLf & vbCrLf & strHomeUrl & "dl.asp?cmd=6&amp;cid=" & cid %></textarea></td>
      </tr>                    
      <tr>
        <td colspan=2 align=center><input class="button" type="submit" value="Send" id="Submit1" name="Submit1" /></td>
      </tr></table>
  </form>
<%
end if
spThemeBlock1_close(intSkin)
end sub

sub badLink()
  if strDBNTUserName <> "" Then
	sSQL = "SELECT NAME FROM DL WHERE DL_ID = " & cid
	set rsNam = my_Conn.execute(sSQL)
	  lnkNam = rsNam(0)
	set rsNam = nothing
	executeThis("UPDATE DL SET BADLINK = " & strUserMemberID & " WHERE DL_ID=" & cid)
    strMsg = strMsg & "<li>Thank you for reporting the bad link for: <br /><b>" & lnkNam
    strMsg = strMsg & "</b><br />We will correct the problem as soon as possible</li>"
  else 
    strMsg = strMsg & "<li>You do not have enough permission to perform this action</li>"
    strMsg = strMsg & "<li>You must be logged in to report a bad download link</li>"
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

sub showFAQ()
spThemeBlock1_open(intSkin) %>
  <table><tr>
    <td>
<p><br />
This section is a download manager to be used as a collection of popular 
downloads. When you click on &quot;Downloads&quot; link you will see a menu on the left 
and the categories on the right. You can use menu items as described below -the 
same menu links also appear on the upper bar under the logo: <br />
<br /></p></td>
  </tr>
  <tr>
    <td>
    <p><b>How do I...</b>
    <ul>
    <li><a href="#New_Downloads">get a list of the Newest Downloads?</a></li>
    <li><a href="#Popular_Downloads">get a list of the most Popular Downloads?</a></li>
    <li><a href="#Top_downloads">get a list of the Top Downloads?</a></li>
    <li><a href="#Adding_New_Files_to_the_Downloads">Submit new files to the Downloads?</a></li>
    <li><a href="#Search_Downloads">Find a program in the Downloads?</a></li>
    <li><a href="#Pop-Up_Download_Detail_View">view the Details of a Download?</a></li>
    </ul>
    <br /></p>
    </td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="New_Downloads"></a>&nbsp;New Downloads</b></td>
  </tr>
  <tr>
    <td><br />
<p>
Selecting this menu option will display a list of files that have been added to 
the Downloads within the last week. The page will display the last seven days 
with a number next to the name of the day. If the number is not zero, then a 
download item has been added on that day. The New Downloads list shows the day 
the item was entered, to view the details of an item that has been entered, 
click on the name of the day and the site will display the downloads that were 
entered on that day by their title. Next click on the title of the download and 
the download details will be displayed. Details are: Size, hits, license, 
platform, publisher, language, date added, rating and description.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br /></td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Popular_Downloads"></a>&nbsp;Popular Downloads</b></td>
  </tr>
  <tr>
    <td>
<br /><p>
Selecting this menu option will display top 10 downloads by hit count. You will 
also see a hit count number next to the files. When you click on the items you 
can switch to the details page.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br />
  <tr>
    <td class="tSubTitle"><b><a name="Top_downloads"></a>&nbsp;Top Downloads</b></td>
  </tr>
  <tr>
    <td>
	  <br /><p>Selecting this menu option will display top 10 rated downloads. You will see a 
hit count number along with rating and votes given. When you click on the items you can switch to the details page.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Adding_New_Files_to_the_Downloads"></a>&nbsp;Adding New Files to the Downloads</b></td>
  </tr>
  <tr>
    <td>
<br />
<p>Choosing to &quot;Submit File&quot; is the first choice that is made when entering a new 
file on the Downloads.</p>
<p>Selecting this menu item will open a submit form for you to fill out. 
Alternatively you can add a file by going to a subcategory that matches with the 
file you're supplying. You will see &quot;Add a file&quot; link towards the bottom, click 
it. Fill in the provided form. The first selection that can be made on the page 
is to select a proper category for the file that will be uploaded. If you did 
not see a category that fit your program's description, contact with the admins 
of the site and they'll be happy to add it for you.</p>
<p>Then you have to complete the &quot;Program Title&quot;, &quot;URL of the file(s)&quot;, &quot;Your 
Email&quot; and &quot;Description&quot; areas those required.</p>
<p>You can go on with the other fields as needed but they are not the required 
ones.</p>
<p>Once you have filled out all of the items for your file, click on the &quot;Submit&quot; 
button to submit your file to the site administrators. If the site is configured 
to require approval for new downloads then you will receive a message that your 
file submission will be reviewed by the administrator and if approved will then 
appear on the Downloads. Until the file is approved it will not show up on the 
Download. If the file is not approved by the Administrator the download will be 
deleted and not show up on the Downloads, although you will receive an email to 
let you know whether or not your file has been approved. If the site is 
configured to accept the requests online then a summary page about your file 
will be shown and added to the downloads database immediately.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Search_Downloads"></a>&nbsp;Search Downloads</b></td>
  </tr>
  <tr>
    <td>
<br /><p>
You can use the search area on the left directly by writing search term or terms 
into the Search box then click the Search button to Search through the files on 
the downloads section. Files with matching text will be shown in the display box 
on the right. The Search routine looks through File Descriptions when trying to 
find a match. The titles of the results that are displayed are hyperlinks which 
can be clicked on to pop-up another page with the details of the download.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td class="tSubTitle"><b><a name="Pop-Up_Download_Detail_View"></a>&nbsp;Pop-Up Download Detail View</b></td>
  </tr>
  <tr>
    <td>
<br /><p>
Clicking on the title of a download will pop-up another page with the Download 
Details. This page will show the File Title followed by the Username of the 
person who entered the file. Only the administrators can edit the download even 
if you have submitted the file. If you click the Download text, the site will be 
opened for that file and will be ready for downloading.<br />
You can also report a bad URL and rate the file by clicking on the respective 
texts.</p>
    <p align=right><a href="#top"><%= icon(icnGoUp,"Top of this Page","","","align=""right""") %></a></p>
<br /><br />
	</td>
  </tr>
  <tr>
    <td align="center">
	<p></p></td>
  </tr></table>
<% spThemeBlock1_close(intSkin)
end sub
 %>