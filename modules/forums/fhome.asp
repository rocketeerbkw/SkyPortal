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
CurPageType = "forums"
cnter = 1
%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" --><%
CurPageInfoChk = "1"
function CurPageInfo () 
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString) 
	PageName = txtForums 
	PageAction = txtViewing & "<br />" 
	PageLocation = "fhome.asp" 
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function
%>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<table width="100%" cellpadding="0" cellspacing="0">
<tr>
  <td class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1) %>
	<% menu_fp() %></td>
  <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
	' get module id
	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	end if
	
  arg1 = txtForums & "|fhome.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	
if IsEmpty(Session(strUniqueID & "last_here_date")) then
	Session(strUniqueID & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if

if strShowStatistics <> "1" then
	set rs1 = Server.CreateObject("ADODB.Recordset")

	'
	strSql = "SELECT " & strTablePrefix & "TOTALS.P_COUNT, " & strTablePrefix & "TOTALS.T_COUNT, " & strTablePrefix & "TOTALS.U_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "TOTALS"

	rs1.open strSql, my_Conn

	Users = rs1("U_COUNT")
	Topics = rs1("T_COUNT")
	Posts = rs1("P_COUNT")

	rs1.Close
	set rs1 = nothing
end if

' - Get all Forums From DB
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_STATUS, " 
strSql = strSql & strTablePrefix & "CATEGORY.CAT_NAME, " & strTablePrefix & "CATEGORY.CAT_ORDER "
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_ORDER ASC"
strSql = strSql &  ", " & strTablePrefix & "CATEGORY.CAT_NAME ASC;"

set rs = my_Conn.Execute (strSql)

spThemeBlock1_open(intSkin)

ShowLastHere = (hasAccess(2))
%><table width="100%" cellpadding="0" cellspacing="0">
<tr><td width="100%" align="center">
<table id="fCat" width="100%" height="10" border="0" cellspacing="1" cellpadding="5" class="tBorder" align="center">
      <tr>
        <td width="5%" align=center class="tSubTitle" valign="top"><%
if (hasAccess(1) or mlev = 3) then  
    PostingOptions()
end if %></td>
        <td align=center class="tSubTitle" valign="top"><b><%= txtForum %></b></td>
        <td width="5%" align=center class="tSubTitle" valign="top"><b><%= txtTopics %></b></td>
        <td width="5%" align=center class="tSubTitle" valign="top"><b><%= txtPosts %></b></td>
        <td width="10%" align=center class="tSubTitle" valign="top"><b><%= txtLstPost %></b></td>
<% 
if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then 
%>
        <td width="15%" align=center class="tSubTitle" valign="top"><b><%= txtModerators %></b></td>
<%
end if %>
        <td width="15%" align=center class="tSubTitle">
		<b><%= txtOptions %></b>
		</td>
      </tr>
<% 
if rs.EOF or rs.BOF then
%>
      <tr>
        <td colspan="<% if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then Response.Write("7") else Response.Write("6")%>"><span class="fAltSubTitle"><b><%= txtNoCatFrmFnd %></b></span></td>
      </tr>
<%
else
	intPostCount  = 0
	intTopicCount = 0
	intForumCount = 0
	strLastPostDate = ""
	do until rs.EOF 

		' - Build SQL to get forums via category
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_STATUS, " 
		strSql = strSql & strTablePrefix & "FORUM.CAT_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " 
		strSql = strSql & strTablePrefix & "FORUM.F_URL, " 
		strSql = strSql & strTablePrefix & "FORUM.F_DESCRIPTION, " 
		strSql = strSql & strTablePrefix & "FORUM.CAT_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_TOPICS, " 
		strSql = strSql & strTablePrefix & "FORUM.F_COUNT, " 
		strSql = strSql & strTablePrefix & "FORUM.F_LAST_POST, " 
		strSql = strSql & strTablePrefix & "FORUM.F_STATUS, " 
		strSql = strSql & strTablePrefix & "FORUM.F_TYPE, " 
		strSql = strSql & strTablePrefix & "FORUM.F_PRIVATEFORUMS,  " 
		strSql = strSql & strTablePrefix & "FORUM.FORUM_ORDER, " 
		strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " 
		strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME " 
		strSql = strSql & "FROM " & strTablePrefix & "FORUM " 
		strSql = strSql & "LEFT JOIN " & strMemberTablePrefix & "MEMBERS ON "
		strSql = strSql & strTablePrefix & "FORUM.F_LAST_POST_AUTHOR = "
		strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rs("CAT_ID") & " "
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.FORUM_ORDER ASC"
		strSql = strSql &  ", " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rsForum =  my_Conn.Execute(strSql)
		chkDisplayHeader = true
		
		if rsForum.eof or rsForum.bof then
%>
      <tr>
        <td class="tAltSubTitle" colspan="<% if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then Response.Write("6") else Response.Write("5") end if %>">&nbsp;<img src="themes/<%= strTheme %>/icon_min.gif">&nbsp;<b><% =ChkString(rs("CAT_NAME"),"display") %></b></td>
		<td class="tAltSubTitle" align="center" valign="top" nowrap>
<%
			if (hasAccess(1) or mlev = 3) then 
			  call CategoryAdminOptions()
			end if 
%>		</td>
      </tr>
<%
			Response.Write	"  <tr>" & vbCrLf & _
							"    <td class=""fSubTitle"" colspan="""
			if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then 
				Response.Write "7"
			else 
				Response.Write "6"
			end if
			Response.Write """><b>" & txtNoFrmFnd & "</b></td>" & vbCrLf
			Response.Write "  </tr>" & vbCrLf
		else
			blnHiddenForums = false
			do until rsForum.Eof
				if ChkDisplayForum(rsForum("FORUM_ID")) then
					if rsForum("F_TYPE") <> "1" then 
						intPostCount  = intPostCount + rsForum("F_COUNT")
						intTopicCount = intTopicCount + rsForum("F_TOPICS")
						intForumCount = intForumCount + 1
						if rsForum("F_LAST_POST") > strLastPostDate then 
							strLastPostDate = rsForum("F_LAST_POST")
							intLastPostForum_ID = rsForum("FORUM_ID")
						end if
					end if
					if chkDisplayHeader then
							catHide = ""
							catImg = "min"
							catAlt = txtCollapse
						if request.Cookies(strUniqueID & "hide")("fCat" & rsForum("CAT_ID") & "") <> "" then
						  if request.Cookies(strUniqueID & "hide")("fCat" & rsForum("CAT_ID") & "") = "1" then
							catHide = "none"
							catImg = "max"
							catAlt = txtExpand
						  end if
						end if  ' " & spThemeBlock_forumCategoryCell & "
						Response.Write	"      <tr>" & vbcrlf & _
										"        <td class=""tAltSubTitle"" colspan="""
						if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then
							Response.Write "6"" width=""100%" 
						else
							Response.Write "5"" width=""100%" 
						end if
						Response.Write """ valign=top>&nbsp;<img name=""fCat" & rsForum("CAT_ID") & "Img"" id=""fCat" & rsForum("CAT_ID") & "Img"" src=""Themes/" & strTheme & "/icon_" & catImg & ".gif"" onClick=""javascript:mwpHS('fCat','" & rsForum("CAT_ID") & "','tbody');"" style=""cursor:pointer;"" alt="""  & catAlt &  """><span class=""fAltSubTitle"">&nbsp;&nbsp;<b>" & ChkString(rs("CAT_NAME"),"display") & "</b></span></td>" & vbcrlf 
							Response.Write "        <td class=""tAltSubTitle"" align=center valign=top nowrap><b>"
						if (hasAccess(1) or mlev = 3) then 
							call CategoryAdminOptions()
						end if 
							Response.Write "</b></td>" & vbcrlf 
						Response.Write "      </tr><tbody id=""fCat" & rsForum("CAT_ID") & """ style=""display:" & catHide & ";"">" & vbcrlf
						chkDisplayHeader = false
					end if
					
					if sCColor = "tCellAlt0" then
					  sCColor = "tCellAlt2"
					else
					  sCColor = "tCellAlt0"
					end if
					
					Response.Write	"<tr class=""" & sCColor & """ width=""100%"" onMouseOver=""this.className='tCellHover';"" onMouseOut=""this.className='" & sCColor & "';"">" & vbcrlf & _
									"<td width=""5%"" align=center valign=top>" & vbcrlf
					if rsForum("F_TYPE") = 0 then
						if rs("CAT_STATUS") = 0 then 
							Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>"
							if rsForum("F_LAST_POST") > Session(strUniqueID & "last_here_date") then
								Response.Write "<img src=""images/icons/icon_folder_new_locked.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" alt=""" & txtCatLok & """ title=""" & txtCatLok & """ /></a>"
							else
								Response.Write "<img src=""images/icons/icon_folder_locked.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" alt=""" & txtCatLok & """ title=""" & txtCatLok & """ /></a>"
							end if     
						else 
							if rsForum("F_STATUS") <> 0 then
								Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>" & ChkIsNew(rsForum("F_LAST_POST")) & "</a>"
							else
								Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>"
								if rsForum("F_LAST_POST") > Session(strUniqueID & "last_here_date") then
									Response.Write "<img src=""images/icons/icon_folder_new_locked.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" alt=""" & txtFrmLok & """ title=""" & txtFrmLok & """ /></a>"
								else
									Response.Write "<img src=""images/icons/icon_folder_locked.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" alt=""" & txtFrmLok & """ title=""" & txtFrmLok & """ /></a>"
								end if
							end if
						end if
					else 
						if  rsForum("F_TYPE") = 1 then 
							Response.Write "<a href=""" & rsForum("F_URL") & """ target=""_blank""><img src=""images/icons/icon_url.gif"" height=""16"" width=""16"" border=""0"" hspace=""0"" /></a>"
						end if
					end if 
					Response.Write	"        </td>" & vbcrlf
%>
        <td width="45%" class="fNorm"<% if rsForum("F_TYPE") = 1 then Response.Write(" colspan=4") %> valign="top">
		<a href="<% if rsForum("F_TYPE") = 0 then Response.Write("forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath")) else if rsForum("F_TYPE") = 1 then Response.Write(rsForum("F_URL") & """ target=""_blank") end if %>"> 
			<% =ChkString(rsForum("F_SUBJECT"),"display") %></a><br />
        
        <span class="fSmall"></span><%= htmldecode(rsForum("F_DESCRIPTION")) %></td>
<%
					if rsForum("F_TYPE") = 0 then
						if IsNull(rsForum("F_TOPICS")) then 
%>
        <td align="center" valign="top" class="fNorm" width="5%">0</td>
<%
						else 
%>
        <td align="center" valign="top" class="fNorm" width="5%"><% =rsForum("F_TOPICS") %></td>
<%
						end if 
						if IsNull(rsForum("F_COUNT")) then 
%>
        <td align="center" valign="top" class="fNorm" width="5%">0</td>
<%
						else 
%>
        <td align="center" valign="top" class="fNorm" width="5%"><% =rsForum("F_COUNT") %></td>
<%
						end if 
						if IsNull(rsForum("MEMBER_ID")) then
							strLastUser = ""
						else
							strLastUser = "<br />" & txtby & ": " 
							strLastUser = strLastUser & "<a href=""cp_main.asp?cmd=8&member="& rsForum("MEMBER_ID") & """>"
							strLastUser = strLastUser & ChkString(rsForum("M_NAME"),"display") & "</a>"
						end if
%>
        <td align="center" valign="top" width="10%" nowrap="nowrap"><span class="fSmall"><b><% =ChkDate(rsForum("F_LAST_POST")) %></b><br />
        <% =ChkTime(rsForum("F_LAST_POST")) %><%= strLastUser%></span></td>
<%
					else 
						if rsForum("F_TYPE") = 1 then 
							'## Do Nothing 
						end if 
					end if 
					if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then 
%>
        <td align="left" class="fSmall" valign="top">
<%
						if (listForumModerators(rsForum("FORUM_ID")) <> "") then
							Response.Write(listForumModerators(rsForum("FORUM_ID")))
						else
							Response.Write("&nbsp;")
						end if 
%>
        </td>
<%
					end if 
%>
        <td align="center" valign="top" nowrap="nowrap">
		<%
			if intSubscriptions = 1 and hasAccess(2) and (strForumSubscription = 1 or strForumSubscription = 3) then 
			  subscription_id = chkIsSubscribed(intAppID,"0",rsForum("FORUM_ID"),"0",strUserMemberID)
			  if subscription_id <> 0 then
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=9&amp;cid=" & subscription_id &"');""><img src=""themes/" &  strTheme & "/icon_pmread.gif"" title=""" & txtUnSubScrFm & """ alt=""" & txtUnSubScr & """ border=""0"" /></a>&nbsp;" 
			  else
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=7&amp;cmd=2&amp;cid="&rsForum("FORUM_ID")&"');""><img src=""themes/" &  strTheme & "/icon_pmold.gif"" title=""" & txtSubScrFm & """ alt=""" & txtSubScr & """ border=""0"" /></a>&nbsp;" 
			  end if
			end if
			strSQL = "SELECT TOPIC_ID FROM " & strTablePrefix & "ARCHIVE_TOPICS WHERE FORUM_ID=" & rsForum("FORUM_ID")
			set rsArchive = my_conn.execute (strSQL)
			if not rsArchive.eof then
			  Response.Write "<a href=""forum_archive.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>"
			  Response.Write "<img src=""images/icons/article.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtVwArchPst & """ alt=""" & txtVwArchPst & """ /></a>&nbsp;"
			end if
			rsArchive.close
			if (hasAccess(1) or mlev = 3) then  
			  call ForumAdminOptions
			end if  %>
		</td>
      </tr>
<% 	  else
		lnHiddenForums = true
	  end if ' ChkDisplayForum() 
	  rsForum.MoveNext
	 loop %>
	 </tbody>
<%	end if
		rs.MoveNext
	loop
end if 

forumMess()
WriteStatistics()
%>
  <tr>
    <td colspan="<% if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then Response.Write("7") else Response.Write("6")%>">
    <table width="100%">
      <tr>
        <td class="fNorm">
        <img title="<%= txtNewPosts %>" alt="<%= txtNewPosts %>" src="images/icons/icon_folder_new.gif" />&nbsp;<%= txtNewLstVst %>.<br />
        <img title="<%= txtOldPosts %>" alt="<%= txtOldPosts %>" src="images/icons/icon_folder.gif" />&nbsp;<%= txtNoNewLstVst %>.<br />
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table></td></tr></table>
<% '
spThemeBlock1_close(intSkin)%>
</td></tr></table>
<%
set rs = nothing 
set rsForum = nothing 
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
sub PostingOptions() 
	if (hasAccess(1)) or (lcase(strNoCookies) = "1") then 
		Response.Write "<a href=""forum_post.asp?method=Category""><img border=0 src=""images/icons/icon_folder_new_topic.gif"" title=""" & txtCrtNewCat & """ alt=""" & txtCrtNewCat & """ height=""15"" width=""15"" border=""0"" /></a>" & vbcrlf 
	else
		Response.Write "&nbsp;" & vbcrlf 
	end if 
end sub 

function ChkIsNew(dt)
	if rsForum("F_STATUS") <> 0 then
		if dt > Session(strUniqueID & "last_here_date") then
			ChkIsNew =  "<img src=""images/icons/icon_folder_new.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtNewPosts & """ alt=""" & txtNewPosts & """ />" 
		Else
			ChkIsNew = "<img src=""images/icons/icon_folder.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtOldPosts & """ alt=""" & txtOldPosts & """ />" 
		end if
	else
		ChkIsNew = "<img src=""images/icons/icon_folder_locked.gif"" height=""15"" width=""15"" border=""0"" hspace=""0"" title=""" & txtFrmLok & """ alt=""" & txtFrmLok & """ />"
	end if
end function

sub CategoryAdminOptions() 
  cnter = cnter + 1  %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onMouseOver="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onMouseOut="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtAdminOpts %>" alt="<%= txtAdminOpts %>" border="0" hspace="0" /></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:115px; z-index:100; display:none; position:absolute; right:50px;"> <%
'spThemeTitle= "Category Options "
'spThemeBlock3_open()
Response.Write("<table width=""105""><tr><td align=""center"" nowrap=""nowrap"">")
Response.Write("<b>" & txtCatOpts & "</b><br />")
	if (hasAccess(1)) or (lcase(strNoCookies) = "1") then 
		if (rs("CAT_STATUS") <> 0) then 
			Response.Write "          <a href=""JavaScript:openWindow('forum_pop_lock.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"jsURLPath") & "')""><img src=""images/icons/icon_lock.gif"" title=""" & txtLokCat & """ alt=""" & txtLokCat & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		else
			Response.Write "          <a href=""JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "')""><img src=""images/icons/icon_unlock.gif"" title=""" & txtUnlokCat & """ alt=""" & txtUnlokCat & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		end if 
		if (rs("CAT_STATUS") <> 0) or (hasAccess(1)) then
			Response.Write "          <a href=""forum_post.asp?method=EditCategory&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"urlpath") & """><img src=""images/icons/icon_pencil.gif"" title=""" & txtEdit & "&nbsp;" & txtCat & """ alt=""" & txtEdit & "&nbsp;" & txtCat & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		end if
			Response.Write "          <a href=""JavaScript:openWindow('forum_pop_delete.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"JSurlpath") & "')""><img src=""images/icons/icon_trashcan.gif"" title=""" & txtDelete & "&nbsp;" & txtCat & """ alt=""" & txtDelete & "&nbsp;" & txtCat & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		if (rs("CAT_STATUS") <> 0) or (hasAccess(1)) then
			Response.Write "<a href=""forum_post.asp?method=Forum&CAT_ID=" & rs("CAT_ID") & "&type=0""><img src=""images/icons/icon_folder_new_topic.gif"" title=""" & txtCrtNewFrm & """ alt=""" & txtCrtNewFrm & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		end if 
		if (rs("CAT_STATUS") <> 0) or (hasAccess(1)) then
			Response.Write "<a href=""forum_post.asp?method=URL&CAT_ID=" & rs("CAT_ID") & "&type=1""><img src=""images/icons/icon_url.gif"" title=""" & txtCrtNewWbLnk & """ alt=""" & txtCrtNewWbLnk & """ border=""0"" hspace=""0"" /></a>" & vbcrlf
		end if 
	else
		Response.Write "&nbsp;" & vbcrlf
	end if %><br />
<center><a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><span class="fSmall"><%= txtClose %></span></a></center>
<%
Response.Write("</td></tr></table>")
'spThemeBlock3_close() %>
</div>
<%
end sub 

sub ForumAdminOptions()
  cnter = cnter + 1 %>
          <a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1');"><img src="themes/<%= strTheme %>/icons/toolbox.gif" onMouseOver="javascript:this.src='themes/<%= strTheme %>/icons/toolbox_active.gif';" onMouseOut="javascript:this.src='themes/<%= strTheme %>/icons/toolbox.gif';" title="<%= txtFrmOpts %>" alt="<%= txtFrmOpts %>" border="0" hspace="0" align="absmiddle" /></a>
<div id="fadminOpts<%= cnter %>" class="spThemeNavLog" style="width:100px; z-index:100; display:none; position:absolute; right:50px;">
<%  'cnter = 1
'spThemeTitle= "Forum Options "
'spThemeBlock3_open()
Response.Write("<table width=""100"" align=""center""><tr><td align=""center"" nowrap=""nowrap"">")
Response.Write("<b>" & txtFrmOpts & ":</b><br />")
ForumAdminOptions1()
 %><br />
<a href="javascript:;" onclick="javascript:mwpHSs('fadminOpts<%= cnter %>','1'); shwFm('formEle');"><span class="fSmall"><%= txtClose %></span></a>
<% Response.Write("</td></tr></table>")
'spThemeBlock3_close() %>
</div>
<%
End sub

sub ForumAdminOptions1() 
	if (hasAccess(1)) or (chkForumModerator(rsForum("FORUM_ID"), strDBNTUserName) = "1") or (lcase(strNoCookies) = "1") then
		if rsForum("F_TYPE") = 0 then
			if rs("CAT_STATUS") = 0 then
				if (hasAccess(1)) then 
%>
          <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnlokCat %>" alt="<%= txtUnlokCat %>" border="0" hspace="0" /></a>
<%
				end if
			else 
				if rsForum("F_STATUS") = 1 then 
%>
          <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_lock.gif" title="<%= txtLkFrm %>" alt="<%= txtLkFrm %>" border="0" hspace="0" /></a>
<%
				else 
%>
          <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="images/icons/icon_unlock.gif" title="<%= txtUnLkFrm %>" alt="<%= txtUnLkFrm %>" border="0" hspace="0" /></a>
<%
				end if 
			end if
		end if
		if rsForum("F_TYPE") = 0 then
			if (rs("CAT_STATUS") <> 0 and rsForum("F_STATUS") <> 0) or (hasAccess(1) or mlev = 3) then 
%>
          <a href="forum_post.asp?method=EditForum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>&type=0"><img src="images/icons/icon_pencil.gif" title="<%= txtEdFrmProp %>" alt="<%= txtEdFrmProp %>" border="0" hspace="0" /></a>
<%
			end if
		else 
			if rsForum("F_TYPE") = 1 then 
%>
          <a href="forum_post.asp?method=EditURL&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>&type=1"><img src="images/icons/icon_pencil.gif" title="<%= txtEdWbLnk %>" alt="<%= txtEdWbLnk %>" border="0" hspace="0" /></a>
<%
			end if 
		end if 
		if (hasAccess(1)) or (lcase(strNoCookies) = "1") then 
%>
          <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath") %>')"><img src="images/icons/icon_trashcan.gif" title="<%= txtDelFrm %>" alt="<%= txtDelFrm %>" border="0" hspace="0" /></a>
<%
		end if
		Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=10&amp;cid="& rsForum("FORUM_ID") &"');""><img src=""images/icons/icon_mod.gif"" title=""" & txtAsnModFrm & """ alt=""" & txtModerator & """ border=""0"" /></a>" 
		if rsForum("F_TYPE") = 0 then
			if (hasAccess(1)) or (lcase(strNoCookies) = "1") then 
%>
          <a href="forum_post.asp?method=Topic&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="<%= txtCreNewTop %>" alt="<%= txtCreNewTop %>" height="15" width="15" border="0" /></a>
<%
			end if
		end if 
	else
		Response.Write "&nbsp;"
	end if
end sub 

sub WriteStatistics() 
	Dim Forum_Count
	Dim NewMember_Name, NewMember_Id, Member_Count
	Dim LastPostDate, LastPostLink

	set rs = Server.CreateObject("ADODB.Recordset")
	
	Forum_Count = intForumCount

	' - Get newest membername and id from DB
	strSql = "SELECT M_NAME, MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_STATUS=1 AND MEMBER_ID > 0"
	strSql = strSQL & " ORDER BY M_DATE desc;"
	set rs = my_Conn.Execute(strSql)
	if not rs.EOF then
		NewMember_Name = ChkString(rs("M_NAME"), "display") 
		NewMember_Id = rs("MEMBER_ID")
	else 
		NewMember_Name = ""
	end if
    
	' - Get Active membercount from DB 
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS WHERE M_POSTS > 0 AND M_STATUS=1"
	
	set rs = my_Conn.Execute(strSql)
	
	if not rs.EOF then
		Member_Count = rs("U_COUNT")
	else
		Member_Count = 0
	end if
	set rs = nothing
	
	LastPostDate = ""
 	LastPostLink = ""
	LastPostAuthorLink = ""
	
	if not (intLastPostForum_ID = "") then	
		' - Get lastPostDate and link to that post from DB
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, "
		strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & "FROM " & strTablePrefix & "FORUM, " & strTablePrefix & "TOPICS, "
		strSql = strSql & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "TOPICS.CAT_ID "
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & intLastPostForum_ID & " "
		strSql = strSql & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

	 	set rs = my_Conn.Execute(strSql)
 	
 		if not rs.EOF then
			LastPostDate = ChkDate(rs("T_LAST_POST")) & ChkTime(rs("T_LAST_POST"))
			LastPostLink = "forum_topic.asp?TOPIC_ID=" & rs("TOPIC_ID") & "&FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID")
			LastPostLink = LastPostLink  & "&Topic_Title=" & ChkString(rs("T_SUBJECT"),"urlpath")
			LastPostLink = LastPostLink  & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath") 
			LastPostAuthorLink = " by: "
			strMember_ID = rs("MEMBER_ID")
			strM_NAME = ChkString(rs("M_NAME"),"display") 
			if strUseExtendedProfile then
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""cp_main.asp?cmd=8&member="& strMember_ID & """>"
			else
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""JavaScript:openWindow2('cp_main.asp?cmd=8&member=" & strMember_ID & "')"">"
			end if
            		LastPostAuthorLink = LastPostAuthorLink  & strM_NAME & "</a>"
		end if
	end if
	'rs.close
	set rs = nothing

	'ActTopicCount = 0
	if not isNull(Session(strUniqueID & "last_here_date")) then 
		'Response.Write "Hello<br />" & Session(strUniqueID & "last_here_date")
		if not blnHiddenForums then
			' - Get ActiveTopicCount from DB
			strSql = "SELECT COUNT(" & strTablePrefix & "TOPICS.T_LAST_POST) AS NUM_ACTIVE "
			strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & "WHERE ((" & strTablePrefix & "TOPICS.T_LAST_POST)>'"& Session(strUniqueID & "last_here_date") & "')"

			set rsZ = my_Conn.Execute(strSql)
			if not rsZ.EOF then
				ActTopicCount = rsZ("NUM_ACTIVE")
			else
				ActTopicCount = 0
			end if
			set rsZ = nothing
		end if
	end if

	set rs1 = Server.CreateObject("ADODB.Recordset")

	' SkyPortal
	strSql = "SELECT " & strTablePrefix & "TOTALS.U_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "TOTALS"

	rs1.open strSql, my_Conn

	Users = rs1("U_COUNT")

	rs1.Close
	set rs1 = nothing

	ShowLastHere = (hasAccess(2))
%>
      <tr>
        <td class="tSubTitle" colspan="<% if (strShowModerators = "1") or (hasAccess(1) or mlev = 3) then Response.Write("7") else Response.Write("6") end if %>"><%= txtStats %></td></tr><tr>
<%
	if ShowLastHere then 
%>
        <td class="fNorm" colspan="<% if ((strShowModerators = "1") or (hasAccess(1) or mlev = 3)) then Response.Write("7") else Response.Write("6") end if %>">
        <%= txtLstVisitOn %>&nbsp;<% =ChkDate(Session(strUniqueID & "last_here_date")) %> <% =ChkTime(Session(strUniqueID & "last_here_date")) %>
        </td>
	  </tr>
	  <tr>
<%
	end if 
	if intPostCount > 0 then 
%>
        <td class="fNorm" colspan="<% if ((strShowModerators = "1") or (hasAccess(1) or mlev = 3)) then Response.Write("7") else Response.Write("6") end if%>">
		<% 
		txtFrmStCnts = replace(txtFrmStCnts,"[$Member_Count$]",Member_Count)
		txtFrmStCnts = replace(txtFrmStCnts,"[$Users$]",Users)
		txtFrmStCnts = replace(txtFrmStCnts,"[$intPostCount$]",intPostCount)
		txtFrmStCnts = replace(txtFrmStCnts,"[$intForumCount$]",intForumCount)
		 %>
		<%= txtFrmStCnts %>&nbsp;<a href="<%= lastPostLink %>"><u><%= lastPostDate %></u></a>&nbsp;<%= LastPostAuthorLink %>.
          </td>
        </tr>
        <tr>
<%
	end if
%>      
        <td class="fNorm" colspan="<% if ((strShowModerators = "1") or (hasAccess(1) or mLev = 3)) then Response.Write("7") else Response.Write("6") end if%>">
		<% 
		txtActvTopLstVst = replace(txtActvTopLstVst,"[$intTopicCount$]",intTopicCount)
		txtActvTopLstVst = replace(txtActvTopLstVst,"[$ActiveTopicCount$]",ActiveTopicCount)
		txtActvTopLstVst = replace(txtActvTopLstVst,"[$a$]","<a href=""forum_active_topics.asp""><u>")
		txtActvTopLstVst = replace(txtActvTopLstVst,"[$/a$]","</u></a>")
		 %>
		<%= txtActvTopLstVst %>
        </td>
      </tr>
<%
	if NewMember_Name <> "" then 
%>
      <tr>          
        <td class="fNorm" colspan="<% if ((strShowModerators = "1") or (hasAccess(1) or mLev = 3)) then Response.Write("7") else Response.Write("6") end if%>">
        <%= txtWcmNewMem %>
	<%	Response.Write "<a href=""cp_main.asp?cmd=8&amp;member="& NewMember_Id & """>"
		Response.Write "&nbsp;<b>" & NewMember_Name & "</b></a>." & vbcrlf %>
          </td>
        </tr>
    <tr>
        <td class="fNorm" colspan="<% if ((strShowModerators = "1") or (hasAccess(1) or mLev = 3)) then Response.Write("7") else Response.Write("6") end if%>">
        
    <a href="active_users.asp"><u><%= txtActvUsrs %></u></a>:&nbsp;<%=strOnlineMembersCount%>&nbsp;<%= txtMembers %>&nbsp;<%= txtAnd %>&nbsp;<%=strOnlineGuestsCount%>&nbsp;<%= txtGuests %></td>
    </tr>
<%
	end if 
end sub 
%>