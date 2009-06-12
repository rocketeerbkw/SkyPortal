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
CurPageType = "core"
%>
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<%
sMsg = ""
iPgType = 0
iMode = 0
c_id = 0
s_id = 0
hasPerm = false

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("stop")
	end if
else
  closeAndGo("stop")
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("stop")
	end if
end if
if Request("cid") <> "" or  Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		c_id = cLng(Request("cid"))
	else
		closeAndGo("stop")
	end if
end if
if Request("sid") <> "" or  Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		s_id = cLng(Request("sid"))
	else
		closeAndGo("stop")
	end if
end if

%>
<!--#include file="inc_top_short.asp" -->
<br />
<%
'if not hasPerm then
  select case iPgType
    case 1 'delete member
	  if intIsSuperAdmin and iMode = 1 and c_id > 0 then
	    delMember(c_id)
	  elseif intIsSuperAdmin and iMode = 0 and c_id > 0 then
        strTxtTitle = txtChkDelMem
		strTxtTitle2 = getMemberName(c_id)
		sMode = 1
	    showFrm()
	  else
	    showNoPerms()
	  end if
    case 2 'lock member
	  if hasAccess(1) and iMode = 1 and c_id > 0 then
	    lockMember(c_id)
	  elseif hasAccess(1) and iMode = 0 and c_id > 0 then
        strTxtTitle = txtChkLckMem
		strTxtTitle2 = getMemberName(c_id)
		sMode = 1
	    showFrm()
	  else
	    showNoPerms()
	  end if
    case 3 'unlock member
	  if hasAccess(1) and iMode = 1 and c_id > 0 then
	    unLockMember(c_id)
	  elseif hasAccess(1) and iMode = 0 and c_id > 0 then
        strTxtTitle = txtChkUnLckMem
		strTxtTitle2 = getMemberName(c_id)
		sMode = 1
	    showFrm()
	  else
	    showNoPerms()
	  end if
    case 4 'display IP address
	  DisplayIP()
    case 5 'display groups
	  portalGroups()
    case 6 'display preview
	  showPreview()
	case 7 ' show messenger contact
	  showMessengers()
	case 8 ' show portal avatars
	  avatarLegend()
	case 9 ' smilie legend
	  smilieLegend()
	case 10
	  forumCode()
	case 11
	  editGroupForm(c_id)
	case 12
	  updateGroup(c_id)
    case else
  end select
'else 'show no permissions
  'showNoPerms()
'end if
%>
<!--#include file="inc_footer_short.asp" -->
<% 
sub showFrm()
response.Write("<p>&nbsp;</p>")
spThemeBlock1_open(intSkin) %>
<table class="tPlain" width="100%">
      <tr>
        <td class="tCellAlt0" colspan=2 align=center><p><%= strTxtTitle %><br />
		<b><%= strTxtTitle2 %></b></p><p>&nbsp;</p></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="center" width="50%">
<form action="pop_portal.asp?cmd=<%= iPgType %>&mode=<%= sMode %>" method="post" id="Form10" name="Form10">
<input type="hidden" name="cid" value="<%= c_id %>">
		<Input class="button" type="Submit" value=" <%= txtYes %> " id="Submit1" name="Submit1">
</form>
		</td>
        <td class="tCellAlt0" align="center" width="50%">
		<Input class="button" type="button" onclick="javascript:window.close();" value=" <%= txtNo %> " id="Submit12" name="Submit12">
		</td>
      </tr></table>
<%
spThemeBlock1_close(intSkin)%>
<%
end sub

sub showNoPerms()
  spThemeBlock1_open(intSkin) %>
<p>&nbsp;</p>
<p align=center><span class="fTitle"><b><%= txtNoPermDelMem %></b></span></p>
<p>&nbsp;</p>
<%
  spThemeBlock1_close(intSkin)
end sub

sub delMember(m_id) 
	intPostcount = 0
	intReplycount = 0
	memberName = getMemberName(m_id)
	'######### Delete bookmarks ###########
	strSql = "DELETE FROM " &strTablePrefix & "BOOKMARKS WHERE M_ID = " & m_id
	executeThis(strSql)
	
	'######### Delete Subscriptions ###########
	strSql = "DELETE FROM " &strTablePrefix & "SUBSCRIPTIONS WHERE M_ID = " & m_id
	executeThis(strSql)
	
	'######### Delete front page user ###########
	strSql = "DELETE FROM " &strTablePrefix & "FP_USERS WHERE fp_uid = " & m_id
	executeThis(strSql)
	
	'######### Delete from groups ###########
	strSql = "DELETE FROM " &strTablePrefix & "GROUP_MEMBERS WHERE G_MEMBER_ID = " & m_id
	executeThis(strSql)
	
	'######### Delete cp config ###########
	strSql = "DELETE FROM " &strTablePrefix & "CP_CONFIG WHERE MEMBER_ID = " & m_id
	executeThis(strSql)
	
	'######## Delete personal avatar folder, if there is one. ##################
	if bFso = true then
		remotePath = "\members\" & m_id
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		strFolder = server.MapPath(remotePath)
		If objFSO.FolderExists(strFolder) = true Then
			objFSO.DeleteFolder strFolder
		End If
		set objFSO = nothing
	end if
	
	'######## Delete all PMs from or to this member ###########
	strSQL = "DELETE FROM " & strTablePrefix & "PM "
	strSql = strSql & "WHERE M_TO = " & m_id & " OR M_FROM = " & m_id & ";"
	executeThis(strSql)
	
  '######## delete from classifieds module
  if chkApp("classifieds","USERS") then
	strSql = "DELETE FROM CLASSIFIED WHERE POSTER = '" & memberName & "';"
	executeThis(strSql)
  end if
	
  '######## delete or disable from weblinks module
  if chkApp("links","USERS") then
	'######### Delete article rating ###########
	strSql = "SELECT ITEM_ID, RATING FROM "& strTablePrefix &"M_RATING WHERE RATE_BY = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
	  do until rs.eof
	    strSql = "UPDATE LINKS SET VOTES=VOTES-1,RATING=RATING-"& rs("RATING") &" WHERE LINK_ID="& rs("LINK") & ";"
	    executeThis(strSql)
	    rs.movenext
	  loop
	end if
	set rs = nothing
	'######### Update articles ###########
	strSql = "UPDATE LINKS SET POSTER='n/a' WHERE POSTER = '" & memberName & "';"
	executeThis(strSql)
  end if
	
  '######## delete or disable from downloads module
  if chkApp("downloads","USERS") = 12 then
	'######### Delete downloads rating ###########
	strSql = "SELECT ITEM_ID, RATING FROM "& strTablePrefix &"M_RATING WHERE RATE_BY = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
	  do until rs.eof
	    strSql = "UPDATE DL SET VOTES=VOTES-1,RATING=RATING-"& rs("RATING") &" WHERE DL_ID="& rs("DL") & ";"
	    executeThis(strSql)
	    rs.movenext
	  loop
	end if
	set rs = nothing
	'######### Update articles ###########
	strSql = "UPDATE DL SET UPLOADER='n/a', EMAIL='' WHERE UPLOADER = '" & memberName & "';"
	executeThis(strSql)
  end if
	
  '######## delete or disable from article module
  if chkApp("article","USERS") = 12 then
	'######### Delete article rating ###########
	strSql = "SELECT ITEM_ID, RATING FROM "& strTablePrefix &"M_RATING WHERE RATE_BY = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
	  do until rs.eof
	    strSql = "UPDATE ARTICLE SET VOTES=VOTES-1,RATING=RATING-"& rs("RATING") &" WHERE ARTICLE_ID="& rs("ARTICLE") & ";"
	    executeThis(strSql)
	    rs.movenext
	  loop
	end if
	set rs = nothing
	'######### Update articles ###########
	strSql = "UPDATE ARTICLE SET POSTER='n/a' WHERE POSTER = '" & memberName & "';"
	executeThis(strSql)
	strSql = "UPDATE ARTICLE SET AUTHOR='n/a' WHERE AUTHOR = '" & memberName & "';"
	executeThis(strSql)
  end if
	
  '######## delete or disable from pictures module
  if chkApp("pictures","USERS") then
	'######### Delete picture rating ###########
	strSql = "SELECT ITEM_ID, RATING FROM "& strTablePrefix &"M_RATING WHERE RATE_BY = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
	  do until rs.eof
	    strSql = "UPDATE PIC SET VOTES=VOTES-1,RATING=RATING-"& rs("RATING") &" WHERE PIC_ID="& rs("PIC") & ";"
	    executeThis(strSql)
	    rs.movenext
	  loop
	end if
	set rs = nothing
	'######### Update pictures ###########
	strSql = "UPDATE PIC SET POSTER='n/a' WHERE POSTER = '" & memberName & "';"
	executeThis(strSql)
  end if
  
  if chkApp("forums","USERS") then
	' - Remove the member from the moderator table
	strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
	strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & m_id & ";"
	executeThis(strSql)
	' - Remove the member from forum 'allowed members' table
	strSql = "DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS "
	strSql = strSql & " WHERE " & strTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID = " & m_id & ";"
	executeThis(strSql)
	
	'######### get topic count #########
	strSql = "SELECT COUNT(T_AUTHOR) AS POSTCOUNT "
	strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & "WHERE T_AUTHOR = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
		intPostcount = rs("POSTCOUNT")
	end if
	set rs = nothing
	
	'######### get reply count #########
	strSql = "SELECT COUNT(R_AUTHOR) AS REPLYCOUNT "
	strSql = strSql & "FROM " & strTablePrefix & "REPLY "
	strSql = strSql & "WHERE R_AUTHOR = " & m_id & ";"
	set rs = my_Conn.Execute(strSql)
	if not rs.eof then
		intReplycount = rs("REPLYCOUNT")
	end if
	set rs = nothing
  end if
							
	if ((intReplycount + intPostCount) = 0) then
		'######### Delete the Member - Member has no posts or replies
		strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id & ";"
		executeThis(strSql)
	else					
		'######### disable account 
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 0"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_EMAIL = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL = 0"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME = 'n/a'"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL = '" & strHomeUrl & "files/avatars/noavatar.gif'"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE = 'deleted'"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_MSN = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG = ' '"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS = 1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_GOLD = 1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_REP = 1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_GLOW = ''"
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id & ";"
		executeThis(strSql)
	end if
	  ' - Update total of Members in Totals table
	  strSql = "UPDATE " & strTablePrefix & "TOTALS "
	  strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT=" & strTablePrefix & "TOTALS.U_COUNT-1;"
	  executeThis(strSql)
	
	'######### Delete Ratings/Comments ###########
	strSql = "DELETE FROM "& strTablePrefix &"M_RATING WHERE RATE_BY="& m_id
	executeThis(strSql)

  spThemeBlock1_open(intSkin)
%>
<p>&nbsp;</p>
<p align=center><span class="fTitle"><b><%= txtMemberDel %></b></span></p>
<p>&nbsp;</p>
<%
  spThemeBlock1_close(intSkin)
end sub

sub lockMember(m_id)
  if instr(strWebMaster,"" & lcase(getMemberName(m_id)) & ",") <> 0 then
	if intIsSuperAdmin then
		strSql = "Update " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 0 "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id
		executeThis(strSql)
		strMSG = txtSAMemLckd
	else
		strMSG = txtNoLckSAdmin
	end if
  else
	strSql = "Update " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 0 "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id
	executeThis(strSql)
	strMSG = txtMemLckd
  end if
  spThemeBlock1_open(intSkin)
%>
<p>&nbsp;</p>
<p align=center><span class="fTitle"><b><%= txtMemLckd %></b></span></p>
<p>&nbsp;</p>
<script type="text/javascript"> 
opener.document.location.reload();
//window.close();
</script>
<%
  spThemeBlock1_close(intSkin)
end sub

sub unLockMember(m_id)
  if hasAccess(1) then
    strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1 "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id
	executeThis(strSql)
  end if
  spThemeBlock1_open(intSkin)
%>
<p>&nbsp;</p>
<p align=center><span class="fTitle"><b><%= txtMemUnLckd %></b></span></p>
<p>&nbsp;</p>
<script type="text/javascript"> 
opener.document.location.reload();
//window.close();
</script>
<%
  spThemeBlock1_close(intSkin)
end sub

sub DisplayIP()
	usr = (chkForumModerator(strRqForumID, STRdbntUserName))
	if hasAccess(1) then 
		usr = 1
	end if
	if usr = 1 then
		if strRqTopicID <> "" then
			strSql = "SELECT " & strTablePrefix & "TOPICS.T_IP, " & strTablePrefix & "TOPICS.T_SUBJECT "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & strRqTopicID
			set rsIP = my_Conn.Execute(strSql)
			if rsIP.eof or rsIP.bof then
			  IP = txtNotRec
			else
			  IP = rsIP("T_IP")
			end if
		else
			if strRqReplyID <> "" then
				strSql = "SELECT " & strTablePrefix & "REPLY.R_IP "
				strSql = strSql & " FROM " & strTablePrefix & "REPLY "
				strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & strRqReplyID
				set rsIP = my_Conn.Execute(strSql)
				if rsIP.eof then
				  IP = txtNotRec
				else
				  IP = rsIP("R_IP")
				end if
			end if
		end if
		set rsIP = nothing
%>
		<P align=center><b><%= txtUsrIP %>:</b><br />
		<%= IP %></P><br />
<%
	else %>
<%
	end If
end sub

sub portalGroups()
If iMode = 2 and c_id > 0 Then
sSQL = "SELECT * FROM " & strTablePrefix & "GROUPS WHERE G_ID = " & c_id
set rsGrp = my_Conn.execute(sSQL)
if not rsGrp.eof then
 gid = rsGrp("G_ID")
 gname = rsGrp("G_NAME")
 giname = rsGrp("G_INAME")
 gdesc = rsGrp("G_DESC")
 gcreate = strtodate(rsGrp("G_CREATE"))
 gactive = rsGrp("G_ACTIVE")
 gaddmem = rsGrp("G_ADDMEM")
 if trim(rsGrp("G_MODIFIED")) <> "" then
  gmodify = strtodate(rsGrp("G_MODIFIED"))
 end if
end if
Set rsGrp = nothing
%>
	<table border=0 cellpadding=5 cellspacing=0 class=3dpanel width=100%>
		<tr><td><b><%= txtGrpNam %>:</b></td>
		<td><%= gname %></td></tr>
		<tr><td><b><%= txtDesc %>:</b></td>
		<td><textarea name="info" style="width:240;" rows="3" readonly><%= gdesc %></textarea></td></tr>
		<% If gaddmem Then %>
		<tr><td><b><%= txtCreatOn %>:</b></td>
		<td><%= gcreate %></td></tr>
		<tr><td><b><%= txtLstModfy %>:</b></td>
		<td><%= gmodify %></td></tr>
		<tr><td valign=top>&nbsp;</td><td></td></tr>
		<tr><td valign=top><b><%= txtMembers %>:</b></td><td>
			<select name="members" style="width: 240px;" size="10">
			<% getGrpMembers(gid) %>
			</select>
			<br />*&nbsp;<%= txtGrpLdr %>
		</td></tr>
		<% End If %>
	</table>
<% ElseIf iMode = 1 Then  ' list groups to add to modules 
	if request.QueryString("grps") <> "" then
	  tGrps = chkstring(request.QueryString("grps"),"sqlstring")
	else
	  tGrps = ""
	end if%>
	<script type="text/javascript">
	
	function aGrp(m_id,m_name) {
	var tFrm="<%= chkstring(request.QueryString("frm"),"sqlstring") %>";
	var tObj="<%= chkstring(request.QueryString("sel"),"sqlstring") %>";
	var oObj=window.opener.document[tFrm];
	for (i=0; i<oObj[tObj].length; i++) {
	  if (oObj[tObj].options[i].value==m_id) {
		//user already added
		alert("<%= txtGrpAlAdd %>!");
		return;
	  }
	}

		pos = oObj[tObj].length;
		oObj[tObj].length +=1;
		oObj[tObj].options[pos].value = oObj[tObj].options[pos-1].value;	
		oObj[tObj].options[pos].text = oObj[tObj].options[pos-1].text;
		oObj[tObj].options[pos-1].value = m_id;	
		oObj[tObj].options[pos-1].text = m_name;
		oObj[tObj].options[pos-1].selected = true;
}

	function getCurGroups() {
	var tFm="<%= chkstring(request.QueryString("frm"),"sqlstring") %>";
	var tObj="g_write";
	var oObj=window.opener.document[tFm];
	for (i=0; i<oObj[tObj].length; i++) {
	  if (oObj[tObj].options[i].value!=0) {
	  document.write('<tr><td>');
	  document.write('<a href="javascript:void();" onclick="openWindow3(&#39;pop_portal.asp?cmd=5&mode=2&cid='+oObj[tObj].options[i].value+'&#39;)">');
	  document.write('<img src="<%= icnInfo %>" alt="" border="0"></a>&nbsp;');
	  //document.write('');
	  document.write('<a href="javascript:void();" onclick="aGrp(&#39;'+oObj[tObj].options[i].value+'&#39;, &#39;'+oObj[tObj].options[i].text+'&#39;);" title="<%= txtAddMem %>">');
	  document.write('<b>'+oObj[tObj].options[i].text+'</b></a>');
	  document.write('</td></tr>');
	  }
	}
}
</script>
	<table border=0 cellpadding=5 cellspacing=0 class=3dpanel width=100%>
		<tr><td>
		<% spThemeTitle = txtClkGrpAdd
		   spThemeBlock1_open(intSkin)
		   Response.Write("<table class=""tPlain"">")
		   ListGroups(tGrps)
		   Response.Write("</table>")
		   spThemeBlock1_close(intSkin) %>
		</td></tr>
	</table>
<% End If
end sub

sub getGrpMembers(gid)
	sSQL = "select * from " & strTablePrefix & "GROUP_MEMBERS where G_GROUP_ID = " & gid & " ORDER BY G_GROUP_LEADER DESC"
	response.Write(sSQL & "<br />")
	set rsGrp = my_Conn.execute(sSQL)
	if not rsGrp.eof then
	  do until rsGrp.eof
	    ast = ""
	  	if rsGrp("G_GROUP_LEADER") = 1 then
		  ast = "*"
		end if
	    response.Write("<option value=""" & rsGrp("G_MEMBER_ID") & """>" & getmembername(rsGrp("G_MEMBER_ID")) & " " & ast & "</option>" & vbnewline)
	    rsGrp.movenext
	  loop
	end if
	set rsGrp = nothing
end sub

Sub ListGroups(grps) 
  if grps = "0" then %>
	<script type="text/javascript">
    getCurGroups();
	</script>
	<%
  else
    if grps = "" then
 	  set rsGrp = my_Conn.execute("select G_ID, G_NAME, G_INAME from " & strTablePrefix & "GROUPS ORDER BY G_ACTIVE, G_NAME")
	else
 	  set rsGrp = my_Conn.execute("select G_ID, G_NAME, G_INAME from " & strTablePrefix & "GROUPS WHERE G_ID IN (" & grps & ") ORDER BY G_NAME")
	end if
	if not rsGrp.eof then
	  do until rsGrp.eof
	    meID = rsGrp("G_ID")
		meName = rsGrp("G_NAME")
		selectMemAllow meID,meName
		rsGrp.movenext
	  loop
	end if
    set rsGrp = nothing
  end if
end Sub


sub selectMemAllow(memID,memName)%>
      <tr>
        <td><a href="javascript:;" onclick="openWindow3('pop_portal.asp?cmd=5&mode=2&cid=<%= memID %>')"><%= icon(icnInfo,txtInfo,"","","") %></a>&nbsp;
        	<a href="javascript:;" onclick="aGrp('<%= memID %>','<%= memName %>')" title="<%= txtAddMem %>"><b><% =memName %></b></a>
        </td>
      </tr>
<%
end sub

sub showPreview()
  if Request.Cookies("strMessagePreview") <> "" then
	strMessagePreview = trim(Request.Cookies("strMessagePreview"))
	Response.Cookies("strMessagePreview") = ""
	if strMessagePreview = "" or IsNull(strMessagePreview) then
		strMessagePreview = "[center][b]< " & txtNoTxtPrev & " >[/b][/center]"
	end If
	strPreview = replace(replace(Formatstr(ChkString(strMessagePreview,"message")),"''","'"),"images/Smilies/", "" & strHomeURL & "images/Smilies/")
  elseif Request.Cookies("strSignaturePreview") <> "" then
	strSignaturePreview = Request.Cookies("strSignaturePreview")
	Response.Cookies("strSignaturePreview") = ""
	if strSignaturePreview = "" or IsNull(strSignaturePreview) then
		strSignaturePreview = "[center][b]< " & txtNoTxtPrev & " >[/b][/center]"
	end if
	strPreview = replace(replace(Formatstr(ChkString(strSignaturePreview,"message")),"''","'"),"images/Smilies/", "" & strHomeURL & "images/Smilies/")
  end if
  spThemeTitle = txtPreview
  spThemeBlock1_open(intSkin)
%>
<p>&nbsp;</p>
<p align="left"><%= strPreview %></p>
<p>&nbsp;</p>
<%
  spThemeBlock1_close(intSkin)
end sub

sub showMessengers()
  select case iMode
	case 1 ' ICQ %>
	  <p><span class="fTitle"><%= txtSndICQMsg %></span></p>
	  <form action="http://wwp.mirabilis.com/scripts/WWPMsg.dll" method="get">
	  <input type="hidden" name="subject" value="SkyDogg">
	  <input type="hidden" name="to" value="<%= cLng(Request.QueryString("ICQ")) %>">
	  <%
	  spThemeBlock1_open(intSkin)
	  %>
	  <table class="tPlain" width="100%">
      <TR>
        <TD class="tCellAlt0" align=right nowrap><%= txtSndToNam %>:</td>
        <TD class="tCellAlt0"><%= chkString(Request.QueryString("M_NAME"),"sqlstring") %></td>
      </tr>
      <TR>
        <TD class="tCellAlt0" align=right nowrap><%= txtSndICQNum %>:</td>
        <TD class="tCellAlt0">
		<img src="http://online.mirabilis.com/scripts/online.dll?icq=<%= cLng(Request.QueryString("ICQ")) %>&img=5" border="0" align="absmiddle"><% =cLng(Request.QueryString("ICQ")) %></td>
      </tr>
      <TR>
        <TD class="tCellAlt0" align=right nowrap><%= txtUNam %>:</td>
        <TD class="tCellAlt0">
		<input type="text" name="from" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD class="tCellAlt0" align=right nowrap><%= txtUEml %>:</td>
        <TD class="tCellAlt0">
		<input type="text" name="fromemail" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD class="tCellAlt0" Colspan=2 nowrap><%= txtMsg %></td>
      </TR>
      <TR>
        <TD class="tCellAlt0" Colspan=2>
		<textarea name="body" rows="3" cols="40" wrap="Virtual"></textarea></td>
      </TR>
      <TR>
        <TD class="tCellAlt0" Colspan=2 align=center><input type="submit" value="<%= txtSend %>"></td>
      </tr></table>
	  <%
	  spThemeBlock1_close(intSkin) %>
	  </form>

<%	case 2 'AIM %>
	  <p align="center"><b><%= txtAIMinst %>.</b></p>
	  <%
	  spThemeBlock1_open(intSkin)
	  %>
	  <table class="tPlain" width="100%">
      <TR><TD class="tCellAlt0" align=center nowrap>
	  <a href="aim:goIM?screenname=<% =chkString(Request.QueryString("AIM"),"sqlstring")  %>">
	  <%= txtSndMsg %></a></td>
      </tr></table>
	  <%
	  spThemeBlock1_close(intSkin)

	case 3 'msn
	  spThemeBlock1_open(intSkin) %>
	  <table class="tPlain"><TR>
	  <TD class="tCellAlt0" align=center nowrap><br />
	  <a href="http://messenger.msn.com/" target="_blank">
	  <img src="<%= strHomeUrl %>images/icons/icon_msn1.gif" alt="" border="0" align="absmiddle"></a>
	  </td></tr>
	  <TR><TD class="tCellAlt0" align=center nowrap><br />
	  <%= chkString(Request.QueryString("M_NAME"),"sqlstring") %>'s <%= txtMSNinfo %>:<br /><br />
	  <b><%= chkString(replace(Request.QueryString("msn"),"[no-spam]@","@"),"sqlstring") %></b><br /><br />
	  </td></tr>
	  <TR><TD class="tCellAlt0" align=center nowrap>&nbsp;</td>
	  </tr></table>
      <%
      spThemeBlock1_close(intSkin)
  end select
end sub

sub avatarLegend()
  response.Write("<br /><br />")
  spThemeBlock1_open(intSkin)
  %>
<script type="text/javascript">
function changeAvatar(a_url) {
   for(var x=0; x < opener.document.formEle.url2.length; x++) {
    if(opener.document.formEle.url2.options[x].value == a_url) {
     opener.document.formEle.url2.selectedIndex = x
     opener.document.formEle.url2.options[x].selected = true
     opener.document.formEle.url2.value = opener.document.formEle.url2.options[x].value
     opener.document.formEle.url2.fireEvent('onchange');
    }
  }
}
</script>
  <table class="tPlain">
	<tr>
      <td class="tSubTitle"><a name="avatars"></a><span class="fSubTitle">
	  <b><%= txtAvatars %></b></span></td></tr>
    <tr><td class="tCellAlt1" align="center">
    <%= txtSiteAvtrs %><br /><br />
    <table border="0" align="center" cellpadding="5">
		<% 
		strSql = "SELECT " & strTablePrefix & "AVATAR2.A_HSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_WSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"
		set rsav = my_Conn.Execute(strSql)

		strSql = "SELECT " & strTablePrefix & "AVATAR.A_ID" 
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_URL"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_NAME"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_MEMBER_ID"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_MEMBER_ID = 0 "
		strSql = strSql & " ORDER BY " & strTablePrefix & "AVATAR.A_NAME ASC;"
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.cachesize = 20
		rs.open strSql, my_Conn, 3
		if rs.EOF or rs.BOF then  '## No replies found in DB %>
          <tr><td class="tCellAlt0" colspan="5"><b><%= txtNoAvFnd %></b></td></tr>
		  <%
		else
			rs.movefirst
			rs.pagesize = strPageSize
			maxpages = cint(rs.pagecount)
			intRowCounter = 0
			%>
              <tr valign="top"><td>
			     <table border="0" align="center">
			       <tr>
					 <%
			do until rs.EOF %>
			  <td class="tCellAlt1" align="center">
			  <a href="javascript:void(0)" onclick="changeAvatar('<%=rs("A_URL")%>'); window.close()" title="<%= txtAddMem %>">
			  <img src="<%= strHomeURL & rs("A_URL") %>" border=<% =rsav("A_BORDER") %> hspace=0 alt="<% =rs("A_NAME") %>"></a><br />
			  <% =rs("A_NAME") %></font></td>
			  <%
			  rs.MoveNext
			  intRowCounter = intRowCounter + 1
			  if (intRowCounter mod 7) = 0 and not rs.EOF then %>
			  </tr><tr>
			  <%
			  end if
			loop %>
			</tr></table>
			<%
		end if
		rs.close
		set rs = nothing
		set rsav = nothing %>
                </td>
              </tr>
            </table>
          </td>
        </tr></table>
  <%
  spThemeBlock1_close(intSkin)
end sub

sub smilieLegend() %>
<script type="text/javascript">
<!-- hide
function insertsmilie(smilieface){
	AddText(smilieface);
	window.close();
}
function AddText(NewCode) {
	if (window.opener.document.PostTopic.Message.createTextRange && window.opener.document.PostTopic.Message.caretPos) {
		var caretPos = window.opener.document.PostTopic.Message.caretPos;
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? NewCode + ' ' : NewCode;
	}
	else {
		window.opener.document.PostTopic.Message.value+=NewCode
	}
	setfocus();
}
function setfocus() {
  window.opener.document.PostTopic.Message.focus();
}
// -->
</script>

  <%
  spThemeBlock1_open(intSkin)
  %>
  <table class="tPlain" width="100%">
  <tr>
    <td class="tSubTitle"><a name="smilies"></a><b><%= txtInsSmile %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt1">
    <p><%= txtSiteSmiles %>&nbsp;<% =strSiteTitle %>:<br />
    <table border="0" align="center" cellpadding="5">
      <tr valign="top">
        <td>
        <table border="0" align="center">
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:)]');"><img border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/smile.gif"></a></td>
            <td class="tCellAlt1">[:)]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:D]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/big.gif"></a></td>
            <td class="tCellAlt1">[:D]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[8D]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/cool.gif"></a></td>
            <td class="tCellAlt1">[8D]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:I]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/blush.gif"></a></td>
            <td class="tCellAlt1">[:I]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:p]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/tongue.gif"></a></td>
            <td class="tCellAlt1">[:P]</td>
         </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[}:)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/evil.gif"></a></td>
            <td class="tCellAlt1">[}:)]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[;)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/wink.gif"></a></td>
            <td class="tCellAlt1">[;)]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:o)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/clown.gif"></a></td>
            <td class="tCellAlt1">[:o)]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[B)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/blackeye.gif"></a></td>
            <td class="tCellAlt1">[B)]</td>
          </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[8]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/8ball.gif"></a></td>
            <td class="tCellAlt1">[8]</td>
          </tr>
        </table>
        </td>
        <td>
        <table border="0" align="center">
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:(]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/sad.gif"></a></td>
              <td class="tCellAlt1">[:(]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[8)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/shy.gif"></a></td>
              <td class="tCellAlt1">[8)]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:0]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/shock.gif"></a></td>
              <td class="tCellAlt1">[:O]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:(!]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/angry.gif"></a></td>
              <td class="tCellAlt1">[:(!]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[xx(]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/dead.gif"></a></td>
              <td class="tCellAlt1">[xx(]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[|)]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/sleepy.gif"></a></td>
              <td class="tCellAlt1">[|)]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[:X]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/kisses.gif"></a></td>
              <td class="tCellAlt1">[:X]</td>
            </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[^]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/approve.gif"></a></td>
              <td class="tCellAlt1">[^]</td>
           </tr>
            <tr>
              <td class="tCellAlt1"><a href="Javascript:insertsmilie('[V]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/dissapprove.gif"></a></td>
              <td class="tCellAlt1">[V]</td>
           </tr>
          <tr>
            <td class="tCellAlt1"><a href="Javascript:insertsmilie('[?]');"><img alt border="0" hspace="10" src="<%= strHomeURL %>images/Smilies/question.gif"></a></td>
            <td class="tCellAlt1">[?]</td>
          </tr>
        </table>
        </td>
      </tr>
    </table></p>
    </td>
  </tr></table>
  <%
  spThemeBlock1_close(intSkin)
end sub

sub forumCode()
response.Write("<br /><br />")
spThemeBlock1_open(intSkin)
%>
<table class="tPlain">
  <tr>
    <td class="tSubTitle"><a name="format"></a><b><%= txtFCHowTo %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt1">
    <p><%= txtFCTitle %></p>
    <blockquote>
      <p><%= txtFCbold %></p>

      <p><%= txtFCitalic %></p>

      <p><%= txtFCunderline %></p>

      <p><%= txtFCalignL %>
      </p>

      <p><%= txtFCalign %>
      </p>

      <p><%= txtFCalignR %>
      </p>

      <p><%= txtFCpre %>
      </p>

      <p><%= txtFCstrike %>
      </p>

      <p><%= txtFCmarq %>
      </p>
      
      <p><%= txtFCsupS %>
      </p>
      
      <p><%= txtFCsubS %>
      </p>
      
      <p><%= txtFCtt %>
      </p>
      
      <p><%= txtFChl %>
      </p>
      
      <p><%= txtFChr %></p>

      <p>&nbsp; </p>

      <p><b><%= txtFCFontC %></b><br />
	  	<%= txtFCFontC2 %>
	  	<%= txtFCred %>
	  	<%= txtFCblue %>
	  	<%= txtFCpink %>
	  	<%= txtFCbrown %>
	  	<%= txtFCblack %>
	  	<%= txtFCorange %>
	  	<%= txtFCviolet %>
	  	<%= txtFCyellow %>
	  	<%= txtFCgreen %>
	  	<%= txtFCgold %>
	  	<%= txtFCwhite %>
	  	<%= txtFCpurple %>
      </p>
      <p>&nbsp; </p>
      <p><b><%= txtFChead %></b><br />
        <%= txtFChTxt %><br />
        <table border=0>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h1]</b><%= txtFCtxt %><b>[/h1]</b> =
            </td>
            <td>
            <h1><%= txtFCtxt %></h1>
            </td>
          </tr>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h2]</b><%= txtFCtxt %><b>[/h2]</b> =
            </td>
            <td>
            <h2><%= txtFCtxt %></h2>
            </td>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h3]</b><%= txtFCtxt %><b>[/h3]</b> =
            </td>
            <td>
            <h3><%= txtFCtxt %></h3>
            </td>
          </tr>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h4]</b><%= txtFCtxt %><b>[/h4]</b> =
            </td>
            <td>
            <h4><%= txtFCtxt %></h4>
            </td>
          </tr>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h5]</b><%= txtFCtxt %><b>[/h5]</b> =
            </td>
            <td>
            <h5><%= txtFCtxt %></h5>
            </td>
          </tr>
          <tr>
            <td>
            <%= txtFCex %>&nbsp;<b>[h6]</b><%= txtFCtxt %><b>[/h6]</b> =
            </td>
            <td>
            <h6><%= txtFCtxt %></h6>
            </td>
          </tr>
        </table>
      </p>
      <p>&nbsp; </p>
      <p><b>Font Sizes:</b><br />
        <%= txtFCex %>&nbsp;<b>[size=1]</b><%= txtFCtxt %><b>[/size=1]</b> = <font size=1><%= txtFCtxt %></font id=size1><br />
        <%= txtFCex %>&nbsp;<b>[size=2]</b><%= txtFCtxt %><b>[/size=2]</b> = <font size=2><%= txtFCtxt %></font id=size2><br />
        <%= txtFCex %>&nbsp;<b>[size=3]</b><%= txtFCtxt %><b>[/size=3]</b> = <font size=3><%= txtFCtxt %></font id=size3><br />
        <%= txtFCex %>&nbsp;<b>[size=4]</b><%= txtFCtxt %><b>[/size=4]</b> = <font size=4><%= txtFCtxt %></font id=size4><br />
        <%= txtFCex %>&nbsp;<b>[size=5]</b><%= txtFCtxt %><b>[/size=5]</b> = <font size=5><%= txtFCtxt %></font id=size5><br />
        <%= txtFCex %>&nbsp;<b>[size=6]</b><%= txtFCtxt %><b>[/size=6]</b> = <font size=6><%= txtFCtxt %></font id=size6>
      </p>

      <p>&nbsp; </p>

      <p><%= txtFCbul %></p>

      <p><%= txtFCal %></p>

      <p><%= txtFCnu %></p>

      <p><%= txtFCcode %></p>

      <p><%= txtFCquo %></p>

<%	if (strIMGInPosts = "1") then %>
      <p><%= txtFCimg %></p>
<%	end if %>
    </blockquote>
    </td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)
end sub
%>
