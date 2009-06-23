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
<!--#include file="modules/forums/forum_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<%
Select case Request.QueryString("mode") 
	case "OpenTopic"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("User"),"sqlstring")) = "1") or _
			(hasAccess(1)) or _
			(chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1") then

				'
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_STATUS = 1 "
				if Request.Form("InPlace") = "1" then
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 1"
				else
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 0"
				end if
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cLng(Request.Form("TOPIC_ID"))

				my_Conn.Execute (strSql)

%>
<p align=center><span class="fTitle"><b>Topic Un-Locked!</b></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Lock Topic</b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Back</a></p>
<%			end if %>
<%		else %>
          <P align=center><span class="fTitle"><b>No 
            Permissions to Un-Lock Topic</b><br />
            <br />
            <a href="JavaScript: onClick= history.go(-1)">Back</a></p>
<%
		end if 
	case "OpenForum"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), chkString(Request.Form("user"),"sqlstring")) = "1") or (hasAccess(1)) or (chkForumModerator(chkString(Request.Form("FORUM_ID"),"sqlstring"), Session(strCookieURL & "userid")) = "1")then

				'
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET " & strTablePrefix & "FORUM.F_STATUS = 1 "
				strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & cLng(Request.Form("FORUM_ID"))

				my_Conn.Execute (strSql)

%>
<p align=center><span class="fTitle"><b>Forum Un-Locked!</b></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Lock Forum</b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Lock Forum</b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
	case "OpenCategory"
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
		
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if hasAccess(1) then
				strSql = "UPDATE " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET " & strTablePrefix & "CATEGORY.CAT_STATUS = 1 "
				strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & cLng(Request.Form("CAT_ID"))

				executeThis(strSql)
%>
<p align=center><span class="fTitle"><b>Category Un-Locked!</b></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Lock Category</b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%			end if %>
<%		else %>
<P align=center><span class="fTitle"><b>No Permissions to Un-Lock Category</b><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
	case else 
%>
<P><span class="fTitle">Un-Lock <% if Request.Querystring("mode") = "Topic" then Response.Write("Topic:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("TOPIC_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Forum" then Response.Write("Forum:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("FORUM_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Category" then Response.Write("Category:<br /><span class=""fAlert"">"&chkstring(Request.Querystring("CAT_TITLE"),"sqlstring") & "</span>") %><% if Request.Querystring("mode") = "Member" then Response.Write("Member") %></span></p>

<p><span class="fAlert"><b>NOTE:</b></span>  
<%				if Request.QueryString("mode") = "Member" then %>
Only Administrators can un-lock a Member.
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
Only Administrators can un-lock a Category.
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
Only Administrators can un-lock a Forum.
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
Only Moderators and Administrators can un-lock a Topic.
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</p>
<script language="JavaScript" type="text/JavaScript">
function focuspass() { document.forms.Form10.Pass.focus(); }
window.onload=focuspass;
</script>
<form id="Form10" name="Form10" action="forum_pop_open.asp?mode=<% if Request.Querystring("mode") = "Topic" then Response.Write("OpenTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("OpenForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("OpenCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("UnLockMember") %>" method=Post>
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
        <td class="tCellAlt0" align=right nowrap><b>User Name:</b></td>
        <td class="tCellAlt0"><input class="textbox" type=Text name="User" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size=20></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align=right nowrap><b>Password:</b></td>
        <td class="tCellAlt0"><input class="textbox" type=Password name="Pass" size=20></td>
      </tr>
<%				else %>
<%					if strAuthType="nt" then %>
      <tr>
        <td class="tCellAlt0" align=right nowrap><b>NT Account:</b></td>
        <td class="tCellAlt0"><%=Session(strCookieURL & "userid")%></td>
      </tr>
<%					end if %>
<%				end if %>   	
      <tr>
        <td class="tCellAlt0" colspan=2 align=center><Input class="button" type=Submit value="Send"></td>
      </tr>

<%
	if Request.QueryString("mode") = "Topic" Then  
		response.write "<tr>"
		strSQL = "SELECT " & strTablePrefix & "TOPICS.T_INPLACE FROM " & strTablePrefix & "TOPICS "
		strSql = strSQL & "WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cLng(Request.QueryString("TOPIC_ID")) 
		set rs = my_conn.Execute(strSql)
		response.write "<td class=""tCellAlt0"" align=right><b>Lock In Place</td>"
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
</b></b></b></b></b></b></b></b>
<% end select %><!--#include file="inc_footer_short.asp" -->
