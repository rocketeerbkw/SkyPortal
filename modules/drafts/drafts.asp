<%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2006 Dogg Software All Rights Reserved
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

'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|              Coded by Brandon Williams.             |'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'

%>
<!--#include file="config.asp" --> 
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<%
PageTitle = "Drafts"

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("stop")
	end if
end if
if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	end if
end if

hasEditor = true
strEditorElements = "newDraft"
editorfull = false

CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = "Drafts"
	PageAction = txtViewing & "<br>" 
	CurPageInfo = PageAction & PageName
end function
%>
<!--#include file="inc_top.asp" -->
<%
'Here's where we call our delete function used for the AJAX
if iPgType = 2 then
	delDraft()
elseif iPgType = 3 then
	saveDraftAJAX()
end if

setAppPerms "drafts","iName"

if not bAppRead then
    closeAndGo("default.asp")
end if

tmpVMsg = ""
	
'breadcrumb here
  arg1 = "Drafts|drafts.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""



If not hasAccess("2") Then ' Not Logged in %>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
<td width="200" class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)
Menu_fp() 
affiliateBanners()
%>
</td>
<td class="mainPgCol" valign="top">
<%
intSkin = getSkin(intSubSkin,2)
  if tmpVMsg <> "" then
    showMsgBlock 0,tmpVMsg
  else
	spThemeBlock1_open(intSkin) %>
	<table border="0" cellpadding="0" cellspacing="0" width="60%" align=center>
	<tr align=center><td><p>&nbsp;</p><p align="center"><span class="fSubTitle"><%= txtLgnToVwPg %></span>
	<br /><br /><%= txtNoRegis %>&nbsp;<a href="policy.asp"><u><%= txtRegNow %></u></a>.</p>
	<p>&nbsp;</p>
	</td></tr></table>
<%  spThemeBlock1_close(intSkin)
  end if
 %>
	</td></tr>
</table>
<% Else %>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
<td width="200" class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)

showUserMenu() 
%>
</td>
<td class="mainPgCol" valign="top">
<%
intSkin = getSkin(intSubSkin,2)
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
 select case iPgType
  case 1
  	tmpVMsg = saveDraft()
  case 2
 	tmpVMsg = delDraft()
  case else
end select

if tmpVMsg <> "" then
	showMsgBlock 0,tmpVMsg
end if
%>
<form action="drafts.asp?cmd=1&memID=<% =strUserMemberID %>" method="post">
<input type="hidden" name="test" value="test" />
<%
showDrafts(strUserMemberID)
%> </form> <%
End if 
%>

</td></tr></table>
<!--#include file="inc_footer.asp" -->

<%
sub showUserMenu()
spThemeTitle = txtUsrOpts
spThemeBlock1_open(intSkin) %>
<table width="100%">
<tr><td valign="top">
<%
sSQL = "select " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.M_DONATE from " & strMemberTablePrefix & "MEMBERS where " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strdbntusername & "'"
set rsP = my_Conn.execute(sSQL)
	if varBrowser = "ie" then
		if trim(rsP("M_GLOW")) <> "" or rsP("M_DONATE") > 0 then %> <a href="javascript:;"><img src="images/icons/icon_color.gif" onClick="openWindow('pop_glow.asp?cmd=2&id=<% =rsP("MEMBER_ID") %>')" alt="<%= txtEditGlo %>" title="<%= txtEditGlo %>" border="0"></a>&nbsp;&nbsp;&nbsp;<b><%= displayName(ChkString(strdbntusername,"display"),rsP("M_GLOW")) %></b>
  <% 	else %>
				&nbsp;<b><%= strdbntusername %></b>
	  <% 
		end if
	Else %>
		&nbsp;<b><%= strdbntusername %></b><% 
	End If  
set rsP = nothing

response.Write("<hr />")
cp_userMenu() 
%>
</td></tr></table>
<%
spThemeBlock1_close(intSkin)
end sub
%>
