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



CurPageType="forums"
%>
<!--#include file="config.asp" -->  
<!-- #include file="lang/en/forum_core.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<!--#include file="modules/forums/forum_functions.asp" -->
<%
select case Request.QueryString("mode") 
	case "UStick"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("user"),"sqlstring")) = "1") _
			or (hasAccess(1)) _
			or (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1") then

				'
				strSql = "Update " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_INPLACE = 0"
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & chkString(Request.Form("TOPIC_ID"),"sqlstring")

				my_Conn.Execute (strSql)
%>
<P align=center><span class="fTitle"><b>Topic Un-Sticky</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			Else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Stick Topics</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		Else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Stick Topics</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
		
	case "Sticky"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("user"),"sqlstring")) = "1") _
			or (hasAccess(1)) _
			or (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1") then

				'
				strSql = "Update " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_INPLACE = 1"
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cLng(Request.Form("TOPIC_ID"))

				my_Conn.Execute (strSql)
%>
<P align=center><span class="fTitle"><b>Topic made Sticky!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			Else %>
<P align=center><span class="fTitle"><b>No Permissions to "Sticky" this Topic</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		Else %>
<P align=center><span class="fTitle"><b>No Permissions to "Sticky" this Topic</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
		
	case "CloseTopic"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("user"),"sqlstring")) = "1") _
			or (hasAccess(1)) _
			or (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1") then

				'
				strSql = "Update " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_STATUS = 0 "
				
				if Request.Form("InPlace") = "1" then
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 1"
				else
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 0"
				end if
				
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cLng(Request.Form("TOPIC_ID"))

				my_Conn.Execute (strSql)
%>
<P align=center><span class="fTitle"><b>Topic Locked!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			Else %>
<P align=center><b><span class="fTitle">No Permissions to Lock Topic</span></b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		Else %>
<P align=center><span class="fTitle"><b>No Permissions to Lock Topic</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 

	case "CloseForum"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("user"),"sqlstring")) = "1") or (hasAccess(1)) or (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1") then

				'
				strSql = "Update " & strTablePrefix & "FORUM "
				strSql = strSql & " SET " & strTablePrefix & "FORUM.F_STATUS = 0 "
				strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & chkstring(request.form("FORUM_ID"), "sqlstring") 

				my_Conn.Execute (strSql)
%>
<P align=center><span class="fTitle"><b>Forum Locked!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			else %>
<P align=center><span class="fTitle"><b>No Permissions to Lock Forum</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		else %>
<P align=center><span class="fTitle"><b>No Permissions to Lock Forum</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 

	case "CloseCategory"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if hasAccess(1) then

				'
				strSql = "Update " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET " & strTablePrefix & "CATEGORY.CAT_STATUS = 0 "
				strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & chkstring(request.form("CAT_ID"), "sqlstring") 

				my_Conn.Execute (strSql)
%>
<P align=center><span class="fTitle"><b>Category Locked!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			else %>
<P align=center><b><span class="fTitle">No Permissions to Lock Category</span></b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		else %>
<P align=center><b><span class="fTitle">No Permissions to Lock Category</span></b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
	case else 
%>
<P><span class="fTitle"><% if Request.Querystring("mode") = "Topic" then Response.Write("Lock Topic:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("TOPIC_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Forum" then Response.Write("Lock Forum:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("FORUM_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Category" then Response.Write("Lock Category:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("CAT_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Member" then Response.Write("Lock Member") %><% if Request.Querystring("mode") = "STopic" then Response.Write("Make Topic Sticky:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("TOPIC_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "UTopic" then Response.Write("Un-Stick Topic:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("TOPIC_TITLE"),"sqlstring") & "</span>") %></span></p>

<p><span class="fAlert"><b>NOTE:</b></span>  
<%		select case Request.QueryString("mode") %>
<%			case "Member" %>
Only Administrators can lock a Member.
<%			case "Category" %>
Only Administrators can lock a Category.
<%			case "Forum" %>
Only Moderators and Administrators can lock a Forum.
<%			case "Topic" %>
Only Moderators and Administrators can lock a Topic.
<%			case "STopic" %>
Only Moderators and Administrators<br />can make a Topic sticky.
<%			case "UTopic" %>
Only Moderators and Administrators<br />can un-stick a Topic.
<%		end select %>
</p>
<script language="JavaScript" type="text/JavaScript">
function focuspass() { document.forms.Form10.Pass.focus(); }
window.onload=focuspass;
</script>
<form name="Form10" action="forum_pop_lock.asp?mode=<% if Request.Querystring("mode") = "Topic" then Response.Write("CloseTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("CloseForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("CloseCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("LockMember") %><% if Request.Querystring("mode") = "STopic" then Response.Write("Sticky") %><% if Request.Querystring("mode") = "UTopic" then Response.Write("UStick") %>" method=post>
<input type=hidden name="Method_Type" value="<% if Request.Querystring("mode") = "Topic" then Response.Write("CloseTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("CloseForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("CloseCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("LockMember") %>">
<input type=hidden name="TOPIC_ID" value="<% =cLng(Request.QueryString("TOPIC_ID")) %>">
<input type=hidden name="FORUM_ID" value="<% =cLng(Request.QueryString("Forum_ID")) %>">
<input type=hidden name="CAT_ID" value="<% =cLng(Request.QueryString("CAT_ID")) %>">
<input type=hidden name="MEMBER_ID" value="<% =cLng(Request.QueryString("MEMBER_ID")) %>">
<%
spThemeTableCustomCode = "align=""center"" width=""75%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")

		if strAuthType="db" then %>
      <tr>
                <td class="tCellAlt0" align=right nowrap><b>User 
                  Name:&nbsp; </b></td>
        <td class="tCellAlt0"><input class="textbox" type=text name="User" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size=20></td>
      </tr>
      <tr>
                <td class="tCellAlt0" align=right nowrap><b>Password:&nbsp; 
                  </b></td>
        <td class="tCellAlt0"><input class="textbox" type=Password name="Pass" size=20></td>
      </tr>
<%		else %>
<%			if strAuthType="nt" then %>
      <tr>
                <td class="tCellAlt0" align=right nowrap><b>NT 
                  Account:&nbsp; </b></td>
        <td class="tCellAlt0"><%=Session(strCookieURL & "userid")%></td>
      </tr>
<%			end if %>
<%		end if %>   			
      <tr>
        <td class="tCellAlt0" colspan=2 align=center><Input class="button" type="Submit" value="  Send  "></td>
      </tr>
      
<%
	if Request.QueryString("mode") = "Topic" Then  
		response.write "<tr>"
		strSQL = "SELECT " & strTablePrefix & "TOPICS.T_INPLACE FROM " & strTablePrefix & "TOPICS "
		strSql = strSQL & "WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cLng(Request.QueryString("TOPIC_ID"))
		set rs = my_conn.Execute(strSql)
		response.write "<td class=""tCellAlt0"" align=right><b>Make sticky: </td>"
		response.write "<td class=""tCellAlt0"" align=left><Input type=""Checkbox"" value=""1"" name=""InPlace"""
		if rs("T_INPLACE") = 1 then response.write "checked" 
		response.write "></td>"
		response.write "</tr>" 
 		rs.close
	End If 
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
</form>
<% end select %><!--#include file="inc_footer_short.asp" -->
