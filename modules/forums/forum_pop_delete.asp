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
if Request("CAT_ID") <> "" then
	if IsNumeric(Request("CAT_ID")) = True then Cat_ID = cLng(Request("CAT_ID")) else Cat_ID = 0
end if
if Request("FORUM_ID") <> "" then
	if IsNumeric(Request("FORUM_ID")) = True then Forum_ID = cLng(Request("FORUM_ID")) else Forum_ID = 0
end if
if Request("TOPIC_ID") <> "" then
	if IsNumeric(Request("TOPIC_ID")) = True then Topic_ID = cLng(Request("TOPIC_ID")) else Topic_ID = 0
end if
if Request("REPLY_ID") <> "" then
	if IsNumeric(Request("REPLY_ID")) = True then Reply_ID = cLng(Request("REPLY_ID")) else Reply_ID = 0
end if

if (Cat_ID + Forum_ID + Topic_ID + Reply_ID) < 1 then
	Response.Write	"      <p align=""center""><span class=""fTitle""><b>The URL has been modified!</b></span></p>" & vbNewLine & _
			"      <p align=""center""><span class=""fTitle""><b>Possible Hacking Attempt!</b></span></p>" & vbNewLine
	clostAndGo("stop")
end if
 
if strAuthType = "db" then
	strDBNTUserName = chkString(Request.Form("User"),"sqlstring")
end if

if Request.QueryString("mode") = "DeleteReply" then 
	  	  if strAuthType = "db" then
			tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
	  	  else
			tmpPass = ""
	  	  end if
	mLev = cint(ChkUser3(strDBNTUserName, tmpPass, Reply_ID)) 
	if hasAccess(2) then  '## is Member
	  if (chkForumModerator(Forum_ID, strDBNTUserName) = "1") or (mLev = 1) or (hasAccess(1)) then '## is Allowed

			strSql = "SELECT R_AUTHOR"
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE REPLY_ID = " & Reply_ID & " "

			set rs = my_Conn.Execute (strSql)
			
			if not(rs.eof or rs.bof) then
				deleteCount rs("R_AUTHOR")
			end if
			
			' - Delete reply
			strSql = "DELETE FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & Reply_ID

			my_Conn.Execute strSql
			
			set rs = Server.CreateObject("ADODB.Recordset")

			' - Get last_post and last_post_author for Topic
			strSql = "SELECT R_DATE, R_AUTHOR"
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID & " "
			strSql = strSql & " ORDER BY R_DATE DESC"

			set rs = my_Conn.Execute (strSql)
			
			if not(rs.eof or rs.bof) then
				strLast_Post = rs("R_DATE")
				strLast_Post_Author = rs("R_AUTHOR")
			end if			
			if (rs.eof or rs.bof) or IsNull(strLast_Post) or IsNull(strLast_Post_Author) then  'topic has no replies
				set rs2 = Server.CreateObject("ADODB.Recordset")

				' - Get post_date and author from Topic
				strSql = "SELECT T_AUTHOR, T_DATE "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID & " "
				
				set rs2 = my_Conn.Execute (strSql)
			
				strLast_Post = rs2("T_DATE")
				strLast_Post_Author = rs2("T_AUTHOR")
				
				rs2.Close
				set rs2 = nothing
				
			end if
			
			rs.Close
			set rs = nothing
			
			' - Decrease count of replies to individual topic by 1
			strSql = "UPDATE " & strTablePrefix & "TOPICS "
			strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_REPLIES = " & strTablePrefix & "TOPICS.T_REPLIES - 1"
			if strLast_Post <> "" then 
				strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author & ""
				end if
			end if
			strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Topic_ID

			my_Conn.Execute strSql

			' - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			rs.Close
			set rs = nothing

			' - Decrease count of total replies in Forum by 1
			strSql =  "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET " & strTablePrefix & "FORUM.F_COUNT = " & strTablePrefix & "FORUM.F_COUNT - " & 1 & " "
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Forum_ID

			my_Conn.Execute strSql

			' - Decrease count of total replies in Totals table by 1
			strSql = "UPDATE " & strTablePrefix & "TOTALS "
			strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT - 1"


			my_Conn.Execute strSql
%>
<P><span class="fTitle">Reply Deleted!</span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%		Else %>
<P><span class="fTitle">No Permissions to Delete Reply</span></p>
<p><a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%		end if %>
<%	Else %>
<P><span class="fTitle">No Permissions to Delete Reply</span></p>
<p><a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%	end if

else
	if Request.QueryString("mode") = "DeleteTopic" then
	  if strAuthType = "db" then
		tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
		tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
	  else
		tmpPass = "nt"
		tmpPass2 = "nt"
	  end if
		if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
			if (chkForumModerator(Forum_ID, STRdbntUserName) = "1") or (hasAccess(1)) then
				delAr = split(Topic_ID, ",")
				for i = 0 to ubound(delAr) 

					' - count total number of replies of TOPIC_ID  in Reply table

					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & cint(delAr(i))
					rs.Open strSql, my_Conn
					risposte = rs("cnt")
					rs.close
					set rs = nothing
									
					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT T_POLL "
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cint(delAr(i))

					rs.Open strSql, my_Conn
					if not(rs.eof or rs.bof) then
						dPoll_id = rs("T_POLL")
					else
						dPoll_id = 0
					end if
					rs.close
					set rs = nothing

					if dPoll_id <> 0 then
						strSql = "DELETE FROM " & strTablePrefix & "POLLS "
						strSql = strSql & " WHERE " & strTablePrefix & "POLLS.POLL_ID = " & dPoll_id
						my_Conn.Execute strSql
					
						strSql = "DELETE FROM " & strTablePrefix & "POLL_ANS "
						strSql = strSql & " WHERE " & strTablePrefix & "POLL_ANS.POLL_ID = " & dPoll_id
						my_Conn.Execute strSql
					end if

					' - Delete the actual topics
					strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cint(delAr(i))
					my_Conn.Execute strSql

					' - Delete all replys related to the topics
					strSql = "DELETE FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & cint(delAr(i))

					my_Conn.Execute strSql	' - Get last_post and last_post_author for Forum
					strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR"
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "			
					strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
					strSql = strSql & " ORDER BY T_LAST_POST DESC"
					
					set rs = my_Conn.Execute (strSql)
			
					if not rs.eof then
						rs.movefirst
						strLast_Post = rs("T_LAST_POST")
						strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
					else
						strLast_Post = ""
						strLast_Post_Author = ""
					end if
			
					rs.Close
					set rs = nothing

					' - Update count of replies to a topic in Forum table
					strSql = "UPDATE " & strTablePrefix & "FORUM "
					strSql = strSql & " SET " & strTablePrefix & "FORUM.F_COUNT = " & strTablePrefix & "FORUM.F_COUNT - " & cint(risposte) + 1
					strSql = strSql & " ,   " & strTablePrefix & "FORUM.F_TOPICS = " & strTablePrefix & "FORUM.F_TOPICS - "	& 1				
					if strLast_Post <> "" then 						
						strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
						if strLast_Post_Author <> "" then
							strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
						end if
					end if

					strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Forum_ID
					my_Conn.Execute strSql  
					' - Update total TOPICS in Totals table

					strSql = "UPDATE " & strTablePrefix & "TOTALS "
					strSql = strSql & " SET " & strTablePrefix & "TOTALS.T_COUNT = " & strTablePrefix & "TOTALS.T_COUNT - " & 1
					strSql = strSql & ",    " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT - " & cint(risposte) + 1
					my_Conn.Execute strSql					

				next
%>
<p>&nbsp;</p><P align=center><span class="fTitle"><b>Topic Deleted!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%			Else %>
<p>&nbsp;</p><P align=center><span class="fTitle"><b>No Permissions to Delete Topic</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1) ">Go Back to Re-Authenticate</a></p>
<%			end if %>	  
<%		Else %>
<p>&nbsp;</p><P align=center><span class="fTitle"><b>No Permissions to Delete Topic</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
		end if 
	else 
		if Request.QueryString("mode") = "DeleteForum" then
	  	  if strAuthType = "db" then
			tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
			tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
	  	  else
			tmpPass = "nt"
			tmpPass2 = "nt"
	  	  end if
			if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
				if hasAccess(1) then
					delAr = split(Forum_ID, ",")
					for i = 0 to ubound(delAr) 
						' - Delete all replys related to the topics
						strSql = "DELETE FROM " & strTablePrefix & "REPLY "
						strSql = strSql & " WHERE FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						' - Delete the actual topics
						strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
						strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						' - Delete the moderators
						strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
						strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						' - Delete the actual forums
						strSql = "DELETE FROM " & strTablePrefix & "FORUM "
						strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						
						' - count total number of replies in Reply table

						set rs = Server.CreateObject("ADODB.Recordset")

						strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
						strSql = strSql & " FROM " & strTablePrefix & "REPLY "
						
						rs.Open strSql, my_Conn
						risreply = rs("cnt")
						rs.close

						set rs = nothing

						set rs = Server.CreateObject("ADODB.Recordset")

						' - count total number of Topics in Topics table
						strSql = "SELECT count(" & strTablePrefix & "TOPICS.TOPIC_ID) AS cnt "
						strSql = strSql & " FROM " & strTablePrefix & "TOPICS "

						rs.Open strSql, my_Conn
						rispost = rs("cnt")
						rs.close

						set rs = nothing

						' - Update total topics and posts in Totals table
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & risreply + rispost
						strSql = strSql & ",    " & strTablePrefix & "TOTALS.T_COUNT = " & cint(rispost)

						my_Conn.Execute strSql
					next
%>
<P align=center><span class="fTitle"><b>Forum Deleted!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%				Else %>
<P align=center><span class="fTitle"><b>No Permissions to Delete Forum</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1) ">Go Back to Re-Authenticate</a></p>
<%				end if %>	  
<%			Else %>
<P align=center><span class="fTitle"><b>No Permissions to Delete Forum</span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
			end if 
		else
			if Request.QueryString("mode") = "DeleteCategory" then
	  	  	  if strAuthType = "db" then
				tmpPass = pEncrypt(pEnPrefix & chkString(Request.Form("pass"),"sqlstring"))
				tmpPass2 = Request.Cookies(strUniqueID & "User")("PWord")
	  	  	  else
				tmpPass = "nt"
				tmpPass2 = "nt"
	  	  	  end if
				if hasAccess(2) and (tmpPass = tmpPass2) then  '## is Member
					if hasAccess(1) then
						delAr = split(Cat_ID, ",")
						for i = 0 to ubound(delAr) 
							' - Delete all replys related to the topics
							strSql = "DELETE FROM " & strTablePrefix & "REPLY "
							strSql = strSql & " WHERE " & strTablePrefix & "REPLY.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql
							
							' - Delete the actual topics
							strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
							strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql

							' - Delete the actual forums
							strSql = "DELETE FROM " & strTablePrefix & "FORUM "
							strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql

							' - Delete the actual category
							strSql = "DELETE FROM " & strTablePrefix & "CATEGORY "
							strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql
							
							
							' - count total number of replies in Reply table
							set rs = Server.CreateObject("ADODB.Recordset")

							strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
							strSql = strSql & " FROM " & strTablePrefix & "REPLY "

							rs.Open strSql, my_Conn
							risreply = rs("cnt")
							rs.close
	
							set rs = nothing

						
							' - count total number of Topics in Topics table
							set rs = Server.CreateObject("ADODB.Recordset")

							strSql = "SELECT count(" & strTablePrefix & "TOPICS.TOPIC_ID) AS cnt "
							strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
							
							rs.Open strSql, my_Conn
							rispost = rs("cnt")
							rs.close
							
							P_Count = cint(risreply + rispost)
							T_Count = cint(rispost)
							' - Update total topics and posts in Totals table
							strSql = "UPDATE " & strTablePrefix & "TOTALS "
							strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & P_Count & " "
							strSql = strSql & ",    " & strTablePrefix & "TOTALS.T_COUNT = " & T_Count & " "
							my_Conn.Execute strSql
						next
%>
<P align=center><span class="fTitle"><b>Category Deleted!</b></span></p>
<script type="text/javascript"> 
opener.document.location.reload();
window.close();
</script>
<%					else %>
<P align=center><span class="fTitle"><b>No Permissions to Delete Category</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1) ">Go Back to Re-Authenticate</a></p>
<%					end if %>	  
<%				else %>
<P align=center><span class="fTitle"><b>No Permissions to Delete Category</b></span><br />
<br />
<a href="JavaScript: onClick= history.go(-1)">Go Back to Re-Authenticate</a></p>
<%
				end if 
			else
%>

<P><span class="fTitle">Delete 
<%				if Request.QueryString("mode") = "Member" then %>
Member
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
Category
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
Forum
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
Topic
<%							else %>
<%								if Request.QueryString("mode") = "Reply" then %>
Reply
<%								end if %>
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</span></p>

<p><span class="fAlert"><b>NOTE:</b></span>  
<%				if Request.QueryString("mode") = "Member" then %>
Only Administrators can delete a Member.
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
Only Administrators can delete a Category.
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
Only Administrators can delete Forums.
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
Only Moderators and Administrators can delete Topics.
<%							else %>
<%								if Request.QueryString("mode") = "Reply" then %>
Only the Author, Moderators and Administrators can delete Replies.
<%								end if %>
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</p>

<script language="JavaScript" type="text/JavaScript">
function focuspass() { document.forms.Form10.Pass.focus(); }
window.onload=focuspass;
</script>
<form action="forum_pop_delete.asp?mode=<% if Request.QueryString("mode") = "Member" then Response.Write("DeleteMember")%><% if Request.QueryString("mode") = "Category" then Response.Write("DeleteCategory")%><% if Request.QueryString("mode") = "Forum" then Response.Write("DeleteForum")%><% if Request.QueryString("mode") = "Topic" then Response.Write("DeleteTopic")%><%if Request.QueryString("mode") = "Reply" then Response.Write("DeleteReply")%>" method="post" id="Form10" name="Form10">
<input type=hidden name="REPLY_ID" value="<%= Reply_ID %>">
<input type=hidden name="TOPIC_ID" value="<% =Topic_ID %>">
<input type=hidden name="FORUM_ID" value="<% =Forum_ID %>">
<input type=hidden name="CAT_ID" value="<% =Cat_ID %>">
<input type=hidden name="MEMBER_ID" value="<% =Member_ID %>">

<%
spThemeTableCustomCode = "align=""center"" width=""65%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
	if strAuthType="db" then %>
      <tr>
        <td class="tCellAlt0" align=right nowrap><b>User Name:</b></td>
        <td class="tCellAlt0"><input type=text name="User" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size=20></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align=right nowrap><b>Password:</b></td>
        <td class="tCellAlt0"><input type=Password name="Pass" size=20></td>
      </tr>
<%				else %>
      <tr>
        <td class="tCellAlt0" align=right nowrap><b>NT Account:</b></td>
        <td class="tCellAlt0"><%=Session(strCookieURL & "userid")%></td>
      </tr>
<%				end if %>
      <tr>
        <td class="tCellAlt0" colspan=2 align=center><Input class="button" type=Submit value="Send" id=Submit1 name=Submit1></td>
      </tr>
<%Response.Write("</table>")
spThemeBlock1_close(intSkin)%>
</form>
<%
			end if
		end if 
	end if 
end if 
%><!--#include file="inc_footer_short.asp" -->