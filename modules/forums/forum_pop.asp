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
curpagetype = "forums"
%>
<!--#include file="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="modules/forums/forum_functions.asp" -->
<%
dim iMode, iCmd, cid, app
iMode = 0
iCmd = 0
cid = 0
app = ""

if IsNumeric(Request("mode")) = True then
	iMode = cLng(Request("mode"))
else
	closeAndGo("stop")
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

%>
<!--#include file="inc_top_short.asp" -->
<%
select case iMode
  case 11 'make/remove moderator
	if chkForumModerator(cid, getMemberName(iCmd)) then
    	'Remove from db
			strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
			strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID=" & cid
			strSql = strSql & " AND " & strTablePrefix & "MODERATOR.MEMBER_ID=" & iCmd
            executeThis(strSql)
	  		response.Write("<script type=""text/javascript"">opener.document.location.reload();</script>")
            closeAndGo("?mode=10&cid=" & cid & "&cmd=12")
    else
    	'add to db
			strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
			strSql = strSql & "(FORUM_ID"
			strSql = strSql & ", MEMBER_ID"
			strSql = strSql & ") VALUES (" 
			strSql = strSql & cid
			strSql = strSql & ", " & iCmd
			strSql = strSql & ")"
            executeThis(strSql)
            closeAndGo("?mode=10&cid=" & cid & "&cmd=12")
    end if    
  case 2
  case 5 'print topic
    spThemeShortBodyTag = " leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" bgcolor=""#FFFFFF"""
  case else
end select
if hasSubTheme then
 intSkin=51
else
 intSkin = 1
end if

	strMsg = ""
	memID = getmemberID(strDBNTUserName)
	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	else
	  strMsg = "Module error in PORTAL_APPS"
	end if
	set rsA = nothing

 if iMode > 0 and hasAccess(2) and strMsg = "" then
	  select case iMode
	    case 1 'Lock item
		  'popLock()
		case 2 'unlock item
		  'popUnlock()
		case 3 'delete item
		  'popDelete()
		case 4 'email item
		  emailToFriend()
		case 5 'print item
		  printItem()
		case 6 'bookmark item
		  addBookmark()
		case 7 'subscribe to item
		  addSubscription()
		case 8 'bookmark item
		  delBookmark()
		case 9 'subscribe to item
		  delSubscription()
		case 10 'make moderator
		  makeModerator()
		case 12 'display IP
		  DisplayIP()
	  end select 
   end if
 %>
<!--#include file="inc_footer_short.asp" -->
<% 
sub DisplayIP()
	usr = (chkForumModerator(iCmd, STRdbntUserName))
	if hasAccess(1) then 
		usr = 1
	end if
	if usr then
		if cid > 0 then
			strSql = "SELECT " & strTablePrefix & "TOPICS.T_IP, " & strTablePrefix & "TOPICS.T_SUBJECT "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & cid
			set rsIP = my_Conn.Execute(strSql)
			if rsIP.eof or rsIP.bof then
			  IP = txtNotRec
			else
			  IP = rsIP("T_IP")
			end if
		else
			if sid > 0 then
				strSql = "SELECT " & strTablePrefix & "REPLY.R_IP "
				strSql = strSql & " FROM " & strTablePrefix & "REPLY "
				strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & sid
				set rsIP = my_Conn.Execute(strSql)
				if rsIP.eof then
				  IP = txtNotRec
				else
				  IP = rsIP("R_IP")
				end if
			end if
		end if
		set rsIP = nothing
  	  response.Write("<br /><br />")
	  spThemeBlock1_open(intSkin)
%>
		<p>&nbsp;</p>
		<P align=center><b><%= txtUsrIP %>:</b><br />
		<%= IP %></P>
		<p>&nbsp;</p><br />
<%
	  spThemeBlock1_close(intSkin)
	else %>
<%
	end If
end sub

sub makeModerator()
  if hasAccess(1) then
  strMsg = "&nbsp;"
  response.Write("<br /><br />")
	  spThemeBlock1_open(intSkin)
      if iCmd = 12 then
	    response.Write("<script type=""text/javascript"">opener.document.location.reload();</script>")
		strMsg = "Forum moderators updated"
	  end if %>
	  <p align="center"><div class="fSubTitle"><%= strMsg %></div></p>
	  <table align="center" border="0">
	    <tr>
	      <td align="center">
		  <%
    'get forum name
	strSql = "select F_SUBJECT from PORTAL_FORUM where FORUM_ID = " & cid
	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open strSql, my_conn
	strForumName = rs("F_SUBJECT")
	rs.close
	set rs = nothing
	
    response.write("<center>Please select a moderator</center>")
    response.write("<center>for forum <b>" & strForumName & "</b></center>")
    response.write("<center>Members with an <img src=""images/icons/icon_mod.gif"" /> icon<br />are already moderators for this forum</center>" & vbcrlf & vbcrlf)
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"
    set rs = Server.CreateObject("ADODB.RecordSet")
    rs.Open strSql, my_conn
    	response.write("<table border=""0""><tr><td><ul>")
        	if not(rs.eof) then
            	do while not(rs.eof)
				if chkForumAccess(rs("MEMBER_ID"),cid) then
                  response.write("<li>")
				  isMod = false
            	  if chkForumModerator(cid, rs("M_NAME")) = "1" then
				    isMod = true
                	response.write("<img src=""images/icons/icon_mod.gif"" />")
                  end if
                  response.write("<a href=""forum_pop.asp?mode=11&cid=" & cid & "&cmd=" & rs("MEMBER_ID") & """>")
                  response.write(rs("M_NAME"))
                  response.write("</a></li>")
				end if
                rs.movenext
                Loop
           Else
           		response.write("<li><b>No members found...")
           End if
       response.write("</ul></td></tr></table>")
    rs.close
    set rs = nothing %>
		  </td>
	    </tr>
	  </table><%
	  spThemeBlock1_close(intSkin)
  response.Write("<br /><br /><br />")
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
	if strMsg = "" then
	  select case iCmd
	    case 1 'bookmark category
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that category exists
		    sSql = "SELECT CAT_NAME FROM "& strTablePrefix & "CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Forum Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = chkString(rsA("CAT_NAME"),"sqlstring")
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
	    case 2 ' bookmark forum
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and SUBCAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that article exists
		    sSql = "SELECT F_SUBJECT FROM "& strTablePrefix & "FORUM WHERE FORUM_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Forum not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = chkString(rsA("F_SUBJECT"),"sqlstring")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "BOOKMARKS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = "Bookmark Added!"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Bookmark already exists"
	      End If
	      set rs = nothing
	    case 3 ' bookmark item
	      sSql ="SELECT * FROM "& strTablePrefix & "BOOKMARKS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and ITEM_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that topic exists
		    sSql = "SELECT T_SUBJECT FROM "& strTablePrefix & "TOPICS WHERE TOPIC_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'topic does not exist
		      strMsg = "Topic not found"
		    else 'topic does exist, lets bookmark it
		      itmTitle = chkString(rsA("T_SUBJECT"),"sqlstring")
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
	
	' check for module subscription
	if strMsg = "" then
	  sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=0"
	  set rsAp = my_Conn.execute(sSql)
	  If rsAp.BOF or rsAp.EOF Then
	    ' they are not subscribed to the module
	  else
	    strMsg = "Cannot add subscription" & "<br />"
	    strMsg = strMsg & "You are already subscribed" & "<br />"
		strMsg = strMsg & "to the Forums Module"
	  end if
	end if
	
	set rsA = nothing
	if strMsg = "" then
	  select case iCmd
	    case 1 'subscribe to category
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that item exists
		    sSql = "SELECT CAT_NAME FROM "& strTablePrefix & "CATEGORIES WHERE CAT_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Forum Category not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = chkString(rsA("CAT_NAME"),"sqlstring")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", " & cid & ", 0, 0, '" & itmTitle & "')"

	          executeThis(insSql)
	          strMsg = strMsg & "Category Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "a Topic is added to the " & itmTitle & " forum category!<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Category subscription already exists"
	      End If
	      set rs = nothing
	    case 2 ' subscribe to forum
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and SUBCAT_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that forum exists
		    sSql = "SELECT F_SUBJECT FROM "& strTablePrefix & "FORUM WHERE FORUM_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'article does not exist
		      strMsg = "Forum not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = chkString(rsA("F_SUBJECT"),"sqlstring")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, " & cid & ", 0, '" & itmTitle & "')"
		
	          executeThis(insSql)
	          strMsg = strMsg & "Forum Subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "a Topic is added to the " & itmTitle & " Forum!<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Forum subscription already exists"
	      End If
	      set rs = nothing
	    case 3 ' subscribe to topic
	      sSql ="SELECT * FROM "& strTablePrefix & "SUBSCRIPTIONS WHERE M_ID=" & memID & " and APP_ID=" & intAppID & " and CAT_ID=0 and SUBCAT_ID=0 and ITEM_ID=" & cid
	      set rs = my_Conn.execute(sSql)
	      If rs.BOF or rs.EOF Then
	        'Verify that topic exists
		    sSql = "SELECT T_SUBJECT FROM "& strTablePrefix & "TOPICS WHERE TOPIC_ID=" & cid
		    set rsA = my_Conn.execute(sSql)
		    if rsA.eof then 'topic does not exist
		      strMsg = "Topic not found"
		    else 'item does exist, lets bookmark it
		      itmTitle = chkString(rsA("T_SUBJECT"),"sqlstring")
	          ' Bookmark doesn't already exist so add it
	          insSql = "INSERT INTO "& strTablePrefix & "SUBSCRIPTIONS ("
	          insSql = insSql & "M_ID, APP_ID, CAT_ID, SUBCAT_ID, ITEM_ID, ITEM_TITLE) VALUES ("
	          insSql = insSql & memID & ", " & intAppID & ", 0, 0, " & cid & ", '" & itmTitle & "')"
	          executeThis(insSql)
			  
	          strMsg = strMsg & "Forum topic subscription Added!<br /><br />"
	          strMsg = strMsg & "You will now receive an email when<br />"
	          strMsg = strMsg & "a new reply is added to this topic!<br />"
		    end if
		    set rsA = nothing
	      else
		    strMsg = "Forum topic subscription already exists"
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

'## Get Original Posting
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.T_MESSAGE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " &  cid 

set rs4 = my_Conn.Execute (strSql)
if rs4.EOF then
	rs4.close
	set rs4 = nothing
	response.write "Either the Topic was not found or you are not authorized to view it"
	closeAndGo("stop")
end if

Forum_ID = rs4("FORUM_ID")
strDBNTUserName = chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")
if strPrivateForums = "1" then
	if (not hasAccess(1)) then
		result = chkForumAccess(strUserMemberID,Forum_ID)
			if result = "False" or result = "FALSE" then
			rs4.close
			set rs4 = nothing
			response.write "You do not have access to the forum where this Topic resides"
			closeAndGo("stop")
		end if
	end if
end if
	
' - Get all replies from DB
strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.REPLY_ID, " 
strSql = strSql & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " 
strSql = strSql & strTablePrefix & "REPLY.R_DATE, " & strTablePrefix & "REPLY.R_MESSAGE " 
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
strSql = strSql & " AND   TOPIC_ID = " & cid & " "
strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"

set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.open  strSql, my_Conn
response.Clear()
	 %>
<html>
<head>
<title><% =strSiteTitle %></title>
</head>
<body bgColor="#FFFFFF" text="#000000" onLoad="window.print()">
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
<td width="100%" colspan="2">
<font color="#000000">
<p><b><% =rs4("T_Subject")%></b></p>
<b>Link:</b> <a href="<%=strHomeURL%>link.asp?topic_id=<%=cid %>"><%=strHomeURL%>link.asp?topic_id=<%=cid %></a><br />
<b>Date:</b> <% = strCurDate %></p>
<p>
<p>
<b>Subject:</b> <% =rs4("T_Subject")%><br />
<b>Topic author:</b> <% =rs4("M_NAME")%></br>
<b>Posted on:</b> <% =ChkDate(rs4("T_DATE")) %> <% =ChkTime(rs4("T_DATE")) %><br />
<b>Message:</b><br /><p><% = replace(formatStr(rs4("T_MESSAGE")),"tiny_mce/", strHomeURL & "tiny_mce/")  %></p>
<%
  if rs3.EOF or rs3.BOF then  
    Response.Write ""
  else
	rs3.movefirst
	do until rs3.EOF
		Response.Write("<hr /></p>")
		Response.Write("<p>")
		Response.Write("<b>Reply author:</b> " & rs3("M_NAME") & "<br />")
		Response.Write("<b>Date:</b> " & ChkDate(rs3("R_DATE")) & " " & ChkTime(rs3("R_DATE")) & "<br />")
		Response.Write("<b>Message:</b></p><p>" & formatStr(replace(rs3("R_MESSAGE"),"tiny_mce/", strHomeURL & "tiny_mce/")) & "</p>")
 
		rs3.MoveNext
	loop
  end if

  rs3.close
  set rs3 = Nothing
  rs4.close
  set rs4 = Nothing

  Response.Write("<p><hr /></p>")
  Response.Write("<p>")
  Response.Write("<b>" & strSiteTitle & " </b>: <a href=""" & strHomeURL & """>" & strHomeURL & "</a>")
  Response.Write("</p>")
  Response.Write("<p>")
  Response.Write("<b>"  & strCopyright & "</b> ")
  Response.Write("</p>")
%>
    <p align="center"><a href="javascript:window.print()"><font color="#000000">Send To Printer</font></a> | <a href="JavaScript:onClick= window.close()"><font color="#000000">Close Window</font></a></p><p>&nbsp;</p>
    </td>
  </tr>
</table>
</body>
</html>
<%
  closeAndGo("stop")
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
  
  <form action="forum_pop.asp" method=post id=Form1 name=Form1>
  <input type=hidden name="cmd" value="6">
  <input type=hidden name="mode" value="4">
  <input type=hidden name="cid" value="<%= cid %>">
<%
%>
      <table><TR>
        <TD align="center" colspan="2" class="fTitle" nowrap="nowrap"><p>Send Topic to a Friend<br />&nbsp;</p></td>
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
        <td colspan=2 align=center><textarea name="Msg" cols="38" rows=5 readonly>Hi, <% =vbCrLf %>I thought you might be interested in this Topic:<%= vbCrLf & vbCrLf & strHomeUrl & "link.asp?topic_id=" & cid %></textarea></td>
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
  response.Write("<br /><br />")
  spThemeBlock1_open(intSkin)
  response.Write("<br /><br />")
  if hasAccess(1) then
	'
	strSql = "UPDATE " & typ & " set FEATURED = " & hp
	strSql = strSql & " WHERE " & typ & "_ID = " & cid
	executeThis(strSql)
%>
	<P align=center><b><%= uCase(typ) & " " & adtyp %> home page items</b><br /></P><script type="text/javascript"> opener.document.location.reload();</script>
<%
  else %>
	<p align=center><b>Only administrators can perform this action.</B></p>
<%
  end If
  response.Write("<br /><br />")
	spThemeBlock1_close(intSkin)
end sub %>