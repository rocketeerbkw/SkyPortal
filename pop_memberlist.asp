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

%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<%
pageMode = trim(chkString(request.querystring ("pageMode"),"sqlstring"))
if Request.Querystring("mode") = "" then
	sMode = ""
else
	sMode = chkString(Request.QueryString("mode"),"sqlstring")
end if
if Request.Querystring("M_NAME") = "" then
	sName = ""
else
	sName = chkString(Request.QueryString("M_NAME"),"sqlstring")
end if
'response.Write(pageMode & "<br />") %>
<script type="text/javascript">
function ChangePage(pageNum){
document.location.href="pop_memberlist.asp?mode=<%=sMode%>&m_name=<%=sName%>&pageMode=<%=pageMode%>&pagesize=30&method=postsdesc&whichpage="+pageNum;
}
</script>
<%
if pageMode <> "" then
select case pageMode
	case "search"
%>
<script type="text/javascript">
function SearchMember(m_id, m_name) {
		pos = opener.document.SearchForm.SearchMember.length;
		if (pos<=1) {
		  opener.document.SearchForm.SearchMember.length +=1;
		  pos +=1;
		}  
		opener.document.SearchForm.SearchMember.options[pos-1].value = m_id;	
		opener.document.SearchForm.SearchMember.options[pos-1].text = m_name;
		opener.document.SearchForm.SearchMember.options[pos-1].selected = true;
}
</script>

<%	case "shoall" 
	  frm = chkString(Request.QueryString("frm"),"sqlstring")
	  sel = chkString(Request.QueryString("sel"),"sqlstring")
%>
<script type="text/javascript">
function AddMember(m_id, m_name) {

for (i=0; i<opener.document.<%= frm %>.<%= sel %>.length; i++) {
if (opener.document.<%= frm %>.<%= sel %>.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.<%= frm %>.<%= sel %>.length;
		opener.document.<%= frm %>.<%= sel %>.length +=1;
		opener.document.<%= frm %>.<%= sel %>.options[pos].value = opener.document.<%= frm %>.<%= sel %>.options[pos-1].value;	
		opener.document.<%= frm %>.<%= sel %>.options[pos].text = opener.document.<%= frm %>.<%= sel %>.options[pos-1].text;
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].value = m_id;	
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].text = m_name;
		opener.document.<%= frm %>.<%= sel %>.options[pos-1].selected = true;
}
</script>
<%	case "allowmember"%>
<script type="text/javascript">
function AddAllowedMember(m_id, m_name) {

for (i=0; i<opener.document.PostTopic.AuthUsers.length; i++) {
if (opener.document.PostTopic.AuthUsers.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.PostTopic.AuthUsers.length;
		opener.document.PostTopic.AuthUsers.length +=1;
		opener.document.PostTopic.AuthUsers.options[pos].value = opener.document.PostTopic.AuthUsers.options[pos-1].value;	
		opener.document.PostTopic.AuthUsers.options[pos].text = opener.document.PostTopic.AuthUsers.options[pos-1].text;
		opener.document.PostTopic.AuthUsers.options[pos-1].value = m_id;	
		opener.document.PostTopic.AuthUsers.options[pos-1].text = m_name;
		opener.document.PostTopic.AuthUsers.options[pos-1].selected = true;
}
</script>

<%	case "pmBan"%>
<script type="text/javascript">
function AddBannedMember(m_id, m_name) {

for (i=0; i<opener.document.PostTopic.BlockedUsers.length; i++) {
if (opener.document.PostTopic.BlockedUsers.options[i].value==m_id) {
	//user already added
	alert("<%= txtMemAlrAdd %>")
	return;
	}
}

		pos = opener.document.PostTopic.BlockedUsers.length;
		opener.document.PostTopic.BlockedUsers.length +=1;
		opener.document.PostTopic.BlockedUsers.options[pos].value = opener.document.PostTopic.BlockedUsers.options[pos-1].value;	
		opener.document.PostTopic.BlockedUsers.options[pos].text = opener.document.PostTopic.BlockedUsers.options[pos-1].text;
		opener.document.PostTopic.BlockedUsers.options[pos-1].value = m_id;	
		opener.document.PostTopic.BlockedUsers.options[pos-1].text = m_name;
		opener.document.PostTopic.BlockedUsers.options[pos-1].selected = true;
}
</script>

<%
end select

mypage = cLng(Request("whichpage"))
if mypage = 0 then
	mypage = 1
end if
mypagesize = cLng(Request.Querystring("pagesize"))
if mypagesize = 0 then
	mypagesize = 30
end if

If Request.QueryString("mode") = "search" then
mypagesize = 20
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '" & trim(chkstring(Request("M_NAME"), "sqlstring")) & "%' "

Set rsMembers = Server.CreateObject("ADODB.RecordSet")
rsMembers.open  strSql, my_conn, 3

if not (rsMembers.EOF or rsMembers.BOF) then  '## No categories found in DB
	rsMembers.movefirst
	rsMembers.pagesize = mypagesize

	maxpages = cint(rsMembers.pagecount)
end if
%>
<%
spThemeBlock1_open(intSkin)
%>
<% Call Paging() %><br/>
<table width="95%" cellspacing="0" cellpadding="0" align="center" class="tCellAlt1">
      <tr>
        <td align="center" class="tSubTitle"><%= txtMemName %>:</td></tr>

<% If rsMembers.EOF or rsMembers.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center"><span class="fSubTitle"><b><%= txtNoMemFnd %></b></span>
        <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
        </td>
      </tr>
<% Else 
	currMember = 0 %>
<%
	i = 0
	rsMembers.cacheSize = 30
	rsMembers.moveFirst
	rsMembers.pageSize = myPageSize
	maxPages = cint(rsMembers.pageCount)
	maxRecs = cint(rsMembers.pageSize)
	rsMembers.absolutePage = myPage
	howManyRecs = 0
	rec = 1
	do until rsMembers.Eof or rec = 31 
		if i = 1 then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if

memId = rsMembers("MEMBER_ID")
memName = ChkString(rsMembers("M_NAME"),"display")

select case pageMode
	case "search"
             Call selectMemSearch()
	case "shoall"
             Call selMemAllow()
	case "allowmember"
             Call selectMemAllow()
    case "pmBan"
             Call selectMemPmBan()
    case "pm"
             Call selectMemPm()
	case "all"
             Call selectAllMem()
	case "games"
             Call selectMemGames()
	case else
		response.write "ERROR!!!"
		response.end
end select

		currMember = rsMembers("MEMBER_ID")
		rsMembers.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop %>
<tr><td>
<% shoPopMembSrchBlk() %><br/>
</td></tr>
</table>
<table>
<tr><td align="right"><br />
<% Call Paging() %>
</td></tr>
      <tr>
        <td align="center"><br /><p><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p></td>
      </tr>
<% end if%> 
      </table>
  <%
  spThemeBlock1_close(intSkin)

else

strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE, " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE, " & strMemberTablePrefix & "MEMBERS.M_DATE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
if hasAccess(1) then
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
else
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
end if
if pageMode = "pmBan" then
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS=1"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL < 3"
end if
if pageMode = "pm" then
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS=1"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE=1"
end if
strSql = strSql & " ORDER BY M_POSTS DESC"

Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.pageSize = 30
	rs.cacheSize = 30
	'rs.open  strSql, my_conn
  	rs.Open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

if not (rs.EOF or rs.BOF) then  '## No categories found in DB
	i = 0
	'rs.moveFirst
	
	maxPages = cint(rs.pageCount)
	maxRecs = cint(rs.recordcount)
	
  	If myPage > maxPages Then myPage = maxPages
  	If myPage < 1 Then myPage = 1
	
	'rs.absolutePage = iPageCurrent
	if strDBType <> "mysql" then
	'if myPage > 1 then
	  rs.absolutePage = myPage
	end if
	howManyRecs = 0
	rec = 1
end if
%>

<%
spThemeBlock1_open(intSkin)
%>
<% Call Paging() %><br/>
<table border="0" cellspacing="0" cellpadding="1" width="100%" class="tCellAlt1">
      <tr>
        <td align="center" width="100%" class="tSubTitle"><%= txtMemName %>:</td>
      </tr>
<% If rs.EOF or rs.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center"><span class="fSubTitle"><b><%= txtNoMemFnd %></b></span>
        <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
        </td>
      </tr>
<% Else %>
<%	currMember = 0 %>
<%
	'rs.movefirst
	'rs.pagesize = mypagesize

	'maxpages = cint(rs.pagecount)
	
	do until rs.Eof or rec = 31 
		if i = 1 then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if

memId = rs("MEMBER_ID")
memName = ChkString(rs("M_NAME"),"display")
select case pageMode
	case "search"
             Call selectMemSearch()
	case "shoall"
             Call selMemAllow()
	case "allowmember"
             Call selectMemAllow()
    case "pmBan"
             Call selectMemPmBan()
	case "pm"
             Call selectMemPm()
	case "all"
             Call selectAllMem()
	case "games"
             Call selectMemGames()
	case else
		response.write txtError & "!!"
		response.end
end select

		currMember = rs("MEMBER_ID")
		rs.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop 
end if 
%>

<tr><td>
<% shoPopMembSrchBlk() %><br/>
    </td>
  </tr>
</table><br/>
	  <% Call Paging() %>
<%
spThemeBlock1_close(intSkin)%>
<% end if
end if

sub selectMemSearch()%>
      <tr class="<% =CColor %>">
        <td class="fNorm">
			  <!-- <a href="<%'= strHomeURL %>cp_main.asp?cmd=8&member=<% '=memID %>" target="_new"><img src="<%'= strHomeUrl %>images/icons/icon_profile.gif" alt="<%'= txtViewProf %>" border="0"  style="cursor:hand;margin-top:2px;" align="middle"></a> -->
			  &nbsp;&nbsp;&nbsp;<img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0" align="middle">
        	<a href="javascript:void(0)" onclick="SearchMember('<%=memID%>', '<% =memName %>'); window.close();" title="<%= txtAddMem %>"><b><% =memName %></b></a>
        	  </td>
      </tr>
<%
end sub

sub selMemAllow()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0">&nbsp;
        	<a href="javascript:void(0)" onclick="AddMember('<%=memID%>', '<% =memName %>');" title="<%= txtAddMem %>"><b><% =memName %></b></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemAllow()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><img src="<%= strHomeUrl %>Themes/<%= strTheme %>/icons/arrow1.gif" border="0">&nbsp;
        	<a href="javascript:void(0)" onclick="AddAllowedMember('<%=memID%>', '<% =memName %>');" title="<%= txtAddMem %>"><b><% =memName %></b></a>&nbsp;&nbsp;
			  <a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemPmBan()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><a href="javascript:;" onclick="AddBannedMember('<%=memID%>', '<% =memName %>');" style="cursor:hand">&nbsp;
        <b><% =memName %></b></a>
        </td>
      </tr>
<%
end sub

sub selectMemPm()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><a onclick="opener.document.PostTopic.sendto.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" title="<%= txtSndMsg %>" alt="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>&nbsp;&nbsp;<a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" title="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub selectAllMem()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><a onclick="opener.document.PostForm.member.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" title="<%= txtSndMsg %>" alt="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>
        	  </td>
      </tr>
<%
end sub

sub selectMemGames()%>
      <tr class="<% =CColor %>">
        <td class="fNorm"><a onclick="opener.document.Bank.member.value+='<% =memName %>'; window.close()" style="cursor:hand">&nbsp;<img src="<%= strHomeUrl %>images/icons/pm.gif" width="11" height="17" alt="<%= txtSndMsg %>" title="<%= txtSndMsg %>" border="0">&nbsp;
        <b><% =memName %></b></a>&nbsp;&nbsp;<a href="<%= strHomeURL %>cp_main.asp?cmd=8&member=<% =memID %>" target="_new"><img src="<%= strHomeUrl %>images/icons/icon_profile.gif" alt="<%= txtViewProf %>" title="<%= txtViewProf %>" border="0"  style="cursor:hand"></a>
        	  </td>
      </tr>
<%
end sub

sub Paging()
	Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
	Response.Write "<tr><td valign=""top"" align=""right"" class=""fNorm"">"
	Response.Write "<b>" & txtPage & ":&nbsp;</b></td>"
	Response.Write "<td valign=""top"" align=""right"">"
	
	if maxpages > 1 then
		if Request("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = chkString(Request("whichpage"),"sqlstring")
		end if
		Response.Write("<form name=""PageNum"" method=""post"" action=""pop_memberlist.asp?pageMode=" & pageMode & """>") & vbNewLine
		Response.Write("<select name=""whichpage"" size=""1"" onchange=""ChangePage(this.value)"">") & vbNewLine
		'Response.Write("<select name=""whichpage"" size=""1"" onchange=""submit()"">") & vbNewLine
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				Response.Write "<option value=""" & counter &  """>" & counter & vbNewLine
			else
				Response.Write "<option value=""" & counter &  """ selected=""selected"">" & counter & vbNewLine
			end if
		next
		Response.Write("</select></form>")
	end if
	Response.Write "</td></tr></table>"
end sub 

sub shoPopMembSrchBlk() %>
<hr />
 <form action="pop_memberlist.asp?mode=search&pageMode=<% = pageMode %>" method="post" name="SearchMembers">
<table cellpadding="0" cellspacing="0" width="100%">
  <tr><td align="center" valign="top">
	<% 
	goingto = "pop_memberlist.asp?mode=search&pageMode=" & pageMode & "&M_NAME="
	goingtostart = "pop_memberlist.asp?pageMode=" & pageMode & ""
	arrAlpha = split(txtAlphabet,",")
	midCnt = round((ubound(arrAlpha)+1)/2)
	if midCnt*2 = ubound(arrAlpha)+1 then
	  midCnt = midCnt + 1
	else
	  midCnt = midCnt + 2
	end if
	dCnt = 0
	response.Write("<a rel=""nofollow"" href=""" & goingtostart & """ title=""" & txtSearch & " " & txtAll & """><small>" & txtAll & "</small></a>&nbsp;")
	for xa = 0 to ubound(arrAlpha)
	response.Write("&nbsp;<a rel=""nofollow"" href=""" & goingto & "" & arrAlpha(xa) & """ title=""" & txtSearch & " " & arrAlpha(xa) & """><small>" & arrAlpha(xa) & "</small></a>")
	dCnt = dCnt + 1
	if dCnt = midCnt then
	response.Write "<br />"
	end if
	next
	%>
	</td>
  </tr>
 <tr><td align="center"><hr/>
 <input type="text" name="M_NAME" size="15">
  </td></tr><tr><td align="center">
   <INPUT type="submit" value="<%= txtSearch %>" id="submit1" name="submit1" border="0" width="40" height="25" class="button">
  </td></tr>
</table>
</form> 
<%
end sub
%><!--#include file="inc_footer_short.asp" -->