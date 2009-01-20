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
'#################################################################################
'## NET IPGATE v2.3.0 Orig Idea by alex042@aol.com(c)Aug 2002, 
'## rewritten by www.gpctexas.net patrick@gpctexas.net  released March 31, 2004
'##
'## rewritten by Hawk92 November 2004
'#################################################################################
'pgType = "manager"

%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" --><%

temp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If temp<>"" Then 
	userip = temp
Else
	userip = Request.ServerVariables("REMOTE_ADDR")
End If
strIPGateCss=1
pagereq=Request.ServerVariables("Path_Info")

if Request.ServerVariables("QUERY_STRING") <> "" then
	pagereq=pagereq &"?"& Request.ServerVariables("QUERY_STRING")
end if

headcss="class=""tTitle"""
headnocss=""
headfontnocss=""

catcss="class=""tSubTitle"""
catnocss=""
catfontnocss=""

forumcss=""
forumnocss=""
fontnocss=""
font2nocss=""
ClassText = ""
ClassBouton = "class=""button"""

bannedcount=0
watchedcount=0
blockedcount=0
'referer=Request.ServerVariables("HTTP_REFERER")
referer=pagereq
qry=request.querystring("qry")
userhost=request.servervariables("REMOTE_HOST")
memberid=trim(chkString(request.form("memberid"),"SQLString"))
startip=Trim(request.form("startip"))
startdate=trim(request.form("startdate"))
enddate=trim(request.form("enddate"))
usercomment=trim(chkString(request.form("usercomment"),"SQLString"))
userstatus=request.form("userstatus")
dbpagekey=trim(chkString(request.form("dbpagekey"),"SQLString"))
userdate=strCurDateString
tempdate=StrIPGateExp
fromadmin=1 %>

<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
  		ipGateConfigMenu("1")
  		response.Write("<hr />")
  		menu_admin()%>
<%
spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtIPGmgr & "|admin_ipgate.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>

<%
Set rs5 = Server.CreateObject("ADODB.Recordset")

StrSql = "SELECT * "
StrSql = StrSql & "FROM " & strTablePrefix & "IPLIST"
rs5.Open StrSql, my_Conn
do until (rs5.eof)
   if rs5("IPLIST_STATUS") = 0 then bannedcount=bannedcount+1
   if rs5("IPLIST_STATUS") = 1 then watchedcount=watchedcount+1
   if rs5("IPLIST_STATUS") = 2 then blockedcount=blockedcount+1
   
   dbrecord = rs5("IPLIST_STARTIP") & "."
   dbrecordarr = split(dbrecord,".")
   useriparr = userip & "."
   useriparr = split(userip,".")
   if dbrecordarr(0) =  useriparr(0) then
		if dbrecordarr(1) =  useriparr(1)then
			if dbrecordarr(2)  = useriparr(2) then
				if dbrecordarr(3) = "" or dbrecordarr(2) = "" then
				   if rs5("IPLIST_STATUS") = 0 then
					%><center>
<span class="fAlert"><h4><b><%=replace(txtIPGIPRangeBan,"[%marker_userip%]", userip)%></b></h4></span>
</center><%
					warning = "Yes"
				   end if
				end if
			end if
		end if
	end if

rs5.MoveNext
Loop
if rs5.state = 1 then rs5.close
set rs5 = nothing

spThemeBlock1_open(intSkin)

Select Case request.querystring("ViewPage")

case "MainMenu"
%><br />

<div align="center">
  <center>
  <table border="1" cellspacing="0" width="550" style="border-collapse: collapse;width:550px;" class="grid" cellpadding="0">
    <tr>
      <td width="100%" colspan="2" align="center" height="26" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
      <% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & "  " & strIPGateVer & " " & txtAdministration%><% if strIPGatecCss = 0 then response.write ("</font>") %></b>
	  </td>
    </tr>
    <tr>
      <td width="50%" align="center" height="36" <% if strIPGateCss=1 then response.write(catcss) else response.write(headnocss) end if %>>
      <b><% if strIPGateCss = 0 then response.write(headfontnocss)%>
	  <%=txtIPTitleMainCurSet%>
      <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      <td width="50%" align="center" height="36" <% if strIPGateCss=1 then response.write(catcss) else response.write(headnocss) end if %>>
      <% if strIPGateCss = 0 then response.write(headfontnocss)%>
	  <%=txtIPTitleURWtchdLstsCount%>
	  <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
    </tr>
    <tr>
      <td width="50%" align="center" valign="top" height="161">
      <table border="1" height="100%" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPBanningIs%>&nbsp; 
		  <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
          <% if strIPGateBan=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPLockDownIs%>&nbsp; 
		  <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
          <% if strIPGateLck=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPCookiesAre%>&nbsp; <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <% if strIPGateCok=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPAllLoggingIs%>&nbsp; <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <% if strIPGateTyp=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPUserLoggingIs%>&nbsp; <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
          <% if strIPGateLog=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %>
		  <% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="70%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPRedirIs%>&nbsp; <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
          <td width="30%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <% if strIPGateMet=1 then response.write(txtIPOn) else response.write(txtIPOff) end if %>
		  <% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
      </table>
      </td>
      <td width="50%" align="center" valign="top" height="161">
      <table border="1" height="100%" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
        <tr height=""23"">
          <td width="76%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPNumBanUsers%>:&nbsp;
		  <% if strIPGatecCss = 0 then response.write ("</font>") %>
		  <img src="images/spacer.gif" border="0" height="17" width="1"></td>
          <td width="24%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
          <%=bannedcount%><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="76%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPNumWtchdUsers%>:<% if strIPGatecCss = 0 then response.write ("</font>") %><img src="images/spacer.gif" border="0" height="17" width="1"></td>
          <td width="24%" align="center" 
		  <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
		  <% if strIPGateCss = 0 then response.write(fontnocss)%><%=watchedcount%><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr height="23">
          <td width="76%" align="right" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
		  <%=txtIPNumRestrctUsers%>:<% if strIPGatecCss = 0 then response.write ("</font>") %><img src="images/spacer.gif" border="0" height="17" width="1"></td>
          <td width="24%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
          <%=blockedcount%><% if strIPGatecCss = 0 then response.write ("</font>") %>&nbsp;
          </td>
        </tr>
        <tr><td colspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><img src="images/spacer.gif" border="0" height="43" width="1"></td></tr>
        <tr height="23">
          <td width="100%" align="center" height="23" colspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><% if strIPGateCss = 0 then response.write(fontnocss)%>
          <a href="JavaScript:openWindow5('pop_help.asp?mode=3')"><%=txtIPGateHelp%>&nbsp;&nbsp;<img src="images/icons/icon_mi_10.gif" align="absmiddle" border="0"></a><% if strIPGatecCss = 0 then response.write ("</font>") %><img src="images/spacer.gif" border="0" height="16" width="1"></td>
        </tr>
        <tr rowspan="2" ><td colspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>><img src="images/spacer.gif" border="0" height="19" width="1"></td></tr>
      </table>
      </td>
    </tr>
  </table>
  <br />
  <A href="ADMIN_HOME.ASP"><%= fontnocss %><%=txtIPRtrnToAdminHome%></font></a>
  </center>
</div>
<br />
<%

case "Settings"

if request.querystring("Save") = 1 then
%>
<p align="center"><span class="fSubTitle"><%=txtIPSttngsSvd%></span></p>
<% end if%>
<form method="POST" action="admin_ipgate.asp?Viewpage=Write_Configuration" name="form1">
  <div align="center">
    <center>
    <br />
    <table border="1" style="border-collapse: collapse;width:600px;" class="grid" width="90%" id="AutoNumber4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" cellpadding="4">
      <tr>
        <td width="100%" colspan="2" align="center" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
        <b><% if strIPGateCss = 0 then response.write(headfontnocss)%><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " "  & txtIPSysSttngs%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></b></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPBanning%><br /></b>
        <input type="radio" value="1" name="strIPGateBan" <%= chkRadioB(strIPGateBan,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateBan" value="0" <%= chkRadioB(strIPGateBan,0,true)%>><%=txtIPOff%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQBanning%></b><br /><%=txtIPFAQBanningDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPLoggingUsers%><br />
        </b>
        <input type="radio" value="1" name="strIPGateLog" <%= chkRadioB(strIPGateLog,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateLog" value="0" <%= chkRadioB(strIPGateLog,0,true)%>><%=txtIPOff%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQLoggingUsers%></b><br /><%=txtIPFAQLoggingUsersDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPLoggingAll%><br /></b>
        <input type="radio" value="1" name="strIPGateTyp" <%= chkRadioB(strIPGateTyp,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateTyp" value="0" <%= chkRadioB(strIPGateTyp,0,true)%>><%=txtIPOff%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQLoggingAll%></b><br /><%=txtIPFAQLoggingAllDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPCookies%><br />
        </b>
        <input type="radio" value="1" name="strIPGateCok" <%= chkRadioB(strIPGateCok,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateCok" value="0" <%= chkRadioB(strIPGateCok,0,true)%>><%=txtIPOff%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQCookies%></b><br /><%=txtIPFAQCookiesDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPLockdown%><br />
        </b>
        <input type="radio" value="1" name="strIPGateLck" <%= chkRadioB(strIPGateLck,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateLck" value="0" <%= chkRadioB(strIPGateLck,0,true)%>><%=txtIPOf%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQLockdown%></b><br /><%=txtIPFAQLockdownDesc%>
        <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPRedirection%></b><br />
        <input type="radio" name="strIPGateMet" value="1" <%= chkRadioB(strIPGateMet,0,false)%>><%=txtIPOn%>
        <input type="radio" name="strIPGateMet" value="0" <%= chkRadioB(strIPGateMet,0,true)%>><%=txtIPOff%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPRedirectionDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPBanMsg%></b><br />
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="strIPGateMsg" size="<% if strIPGateCss=1 then response.write(52) else response.write(45) end if %>" maxLength="75" value=" <%= chkExistElse(strIPGateMsg,txtIPNoAccReqPage) %>"><br />
        <b><%=txtIPForumLockdownMsg%></b><br />
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="strIPGateLkMsg" size="<% if strIPGateCss=1 then response.write(52) else response.write(45) end if %>" maxLength="75" value=" <%= chkExistElse(strIPGateLkMsg,txtIPNoAccReqPage) %>"><br />
        <b><%=txtIPNoPageAccMsg%></b><br />
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="strIPGateNoAcMsg" size="<% if strIPGateCss=1 then response.write(52) else response.write(45) end if %>" maxLength="75" value=" <%= chkExistElse(strIPGateNoAcMsg,txtIPNoAccReqPage) %>"><br />
        <b>Warning Message</b><br />
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="strIPGateWarnMsg" size="<% if strIPGateCss=1 then response.write(52) else response.write(45) end if %>" maxLength="75" value=" <%= chkExistElse(strIPGateWarnMsg,txtIPNoAccReqPage) %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPMsgDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="41%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPLogExp%></b></b><br />
        <select name="strIPGateexp">
        <option value="0" <%= chkSelect(strIPGateexp,0) %>><%=txtIPAllLogs%></option>
        <option value="1" <%= chkSelect(strIPGateexp,1) %>>2 <%=txtDays%></option>
        <option value="7" <%= chkSelect(strIPGateexp,7) %>>1 <%=txtWeek%></option>
        <option value="14" <%= chkSelect(strIPGateexp,14) %>>2 <%=txtWeeks%></option>
        <option value="21" <%= chkSelect(strIPGateexp,21) %>>3 <%=txtWeeks%></option>
        <option value="28" <%= chkSelect(strIPGateexp,28) %>>4 <%=txtWeeks%></option>
        <option value="35" <%= chkSelect(strIPGateexp,35) %>>5 <%=txtWeeks%></option>
        <option value="42" <%= chkSelect(strIPGateexp,42) %>>6 <%=txtWeeks%></option>
        <option value="49" <%= chkSelect(strIPGateexp,49) %>>7 <%=txtWeeks%></option>
        <option value="56" <%= chkSelect(strIPGateexp,56) %>>8 <%=txtWeeks%></option>
        <option value="63" <%= chkSelect(strIPGateexp,63) %>>9 <%=txtWeeks%></option>
        <option value="70" <%= chkSelect(strIPGateexp,70) %>>10 <%=txtWeeks%></option>
        </select><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="71%" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQLogExp%></b><br /><%=txtIPFAQLogExpDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="100%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %> colspan="2">
        <input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtIPChangeSettings%>" name="B1">&nbsp;<input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="reset" value="Reset" name="B2"></td>
      </tr>
    </table>
    </center>
  </div>
  <input type="hidden" name="strIPGateVer" value="2.3.0">
</form>
<form method="POST" action="admin_ipgate.asp?ViewPage=deletelog&qry=<%=StrIpgateExp%>">
  <p align="center">
  <input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtIPEraseAllExpLogs%>
  <% Select Case strIPGateexp
  	Case "0"
		response.write(".")
	Case "1"
		response.write(txtIPFrm2DaysAndOlder)
	Case "7", "14", "21", "28", "35", "42", "48", "56", "63", "70"
		tmpIPGateexp = cint(strIPGateexp)
		response.write(replace(txtIPFrmWeeksAndOlder, "[%marker_weeks%]", tmpIPGateexp/7) & ".")
	Case else
		response.write(".")
	end select
	%>
  <%
' ML Changes:
' Got rid of this gnarly 12 level if then else and replaced it with a select case
' so that you only ned 2 ml vars for all 12 choices.
'   if strIPGateexp = "0" then response.write(".") else if strIPGateexp = "1" then response.write(" from 2 days ago and older.") else if strIPGateexp = "7" then response.write(" from 1 week ago and older.") else if strIPGateexp = "14" then response.write(" from 2 weeks ago and older.") else if strIPGateexp = "21" then response.write(" from 3 weeks ago and older.") else if strIPGateexp = "28" then response.write(" from 4 weeks ago  and older.") else if strIPGateexp = "35" then response.write(" from 5 weeks ago and older.") else if strIPGateexp = "42" then response.write(" from 6 weeks ago and older.") else if strIPGateexp = "49" then response.write(" from 7 weeks ago and older.") else if strIPGateexp = "56" then response.write(" from 8 weeks ago  and older.") else if strIPGateexp = "63" then response.write(" from 9 weeks ago and older.") else if strIPGateexp = "70" then response.write(" from 10 weeks ago and older.")
%>" name="B1"></p>
</form>
<br /><%

case "UserSettings"
if qry <> "" then
		Set rs = Server.Createobject("ADODB.Recordset")
		StrSql="SELECT * FROM " & strTablePrefix & "IPLOG "
		StrSql=StrSql & "WHERE IPLOG_ID = " & qry & ";"
		
		rs.Open StrSql, strConnString
	end if
%>
<form method="POST" action="admin_ipgate.asp?ViewPage=addipupdate&qry=startip" name="ipgate_user_settings">
  <div align="center">
    <center>
    <br />
    <table border="1" style="border-collapse: collapse;width:600px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" width="90%" id="AutoNumber5">
      <tr>
        <td width="100%" colspan="2" align="center" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPBanUser%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPUsrName%><br />
        </b><% if qry <> "" then %>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="memberid" value="<% if rs("IPLOG_MEMBERID") <> "" then Response.Write (rs("IPLOG_MEMBERID")) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %><p>
        </p>
        </td>
        <% else %> <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="memberid"> <% end if %>
        </td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPUsrNameDesc%><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPStartDate%><br />
        </b><input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="startdate" size="<% if strIPGateCss=1 then response.write(31) else response.write(29) end if %>" value="<%=date()%>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" rowspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQStartDate%></b><br /><%=txtIPFAQStartDateDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b>end date<br />
        </b>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="enddate" size="<% if strIPGateCss=1 then response.write(31) else response.write(29) end if %>" value="<%=dateadd("m", 1, date())%>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=lcase(txtStatus)%><br />
        </b><select size="1" name="userstatus">
        <option value="0"><%=txtIPBannedUser%></option>
        <option value="1"><%=txtIPWatchedUser%></option>
        <option value="2"><%=txtIPBlockedAccess%></option>
        <option value="3"><%=txtIPExpireCookie%></option>
        </select><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPStatusNoCSS%>
        <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPCommentOpt%>
		<input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="usercomment" size="<% if strIPGateCss=1 then response.write(32) else response.write(30) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>&nbsp;</td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPBlockedPagesbyIPDesc%></b><br />
		<% Set rs8 = Server.CreateObject("ADODB.Recordset")
		   pgkySql = "SELECT * "
		   pgkySql = pgkySql & "FROM " & strTablePrefix & "PAGEKEYS order by PAGEKEYS_PAGEKEY asc"
	
		   rs8.Open pgkySql, strConnString %>
        <select size="10" name="dbpagekey" multiple><% do until (rs8.eof)%>
        <option value="<%=rs8("PAGEKEYS_PAGEKEY")%>"><%=rs8("PAGEKEYS_PAGEKEY")%>
        </option>
        <%				
					rs8.MoveNext
    				Loop
    	%></select> <% 
        if rs8.State = 1 then rs8.Close
    	set rs8=nothing
        if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQPageKey%></b><br /><%=txtIPFAQPageKeyDesc%><%=txtIPEditPageKey%>
		<a href="admin_ipgate.asp?ViewPage=pagekeys"><%=txtIPClickHere%></a>
&nbsp;<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtSubmit%>" name="B1">&nbsp;<input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="reset" value="Reset" name="B2"></p>
  <input type="hidden" name="startip" value="0.0.0.0">
</form>
<br />
<%
if qry <> "" then
		rs.close
	      set rs = nothing
	end if

case "IPBanning"

if qry <> "" then
		Set rs = Server.Createobject("ADODB.Recordset")
		StrSql="SELECT * FROM " & strTablePrefix & "IPLOG "
		StrSql=StrSql & "WHERE IPLOG_ID = " & qry & ";"
		
		rs.Open StrSql, strConnString
	end if
%>
<form method="POST" action="admin_ipgate.asp?ViewPage=addipupdate&qry=startip" name="ipaddress_settings">
  <div align="center">
    <center>
    <br />
    <table border="1" style="border-collapse: collapse;width:600px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" width="95%" id="AutoNumber6">
      <tr>
        <td width="100%" colspan="2" align="center" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPBanAddSubnet%></b>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPAddSubnet%><br />
        </b><% if qry <> "" then %>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="startip" value="<% if rs("IPLOG_IP") <> "" then Response.Write (rs("IPLOG_IP")) end if %>"><p>
        <% else %> <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="startip"> <% end if %> <br />
        <%=txtIPYourIP%><%=userip%><br />
        <%=txtIPYourHost%><%=userhost%> <% if strIPGatecCss = 0 then response.write ("</font>") %></p>
        </td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQIPandHost%></b><br /><%=txtIPFAQIPandHostDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPStartDate%><br />
        </b><input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="startdate" size="<% if strIPGateCss=1 then response.write(31) else response.write(29) end if %>" value="<%=date()%>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" rowspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQStartDate%></b><br /><%=txtIPFAQStartDateDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b>txtIPEndDate<br />
        </b>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="enddate" size="<% if strIPGateCss=1 then response.write(31) else response.write(29) end if %>" value="<%=dateadd("m", 1, date())%>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=lcase(txtStatus)%><br />
        </b><select size="1" name="userstatus">
        <option value="0"><%=txtIPBannedUser%></option>
        <option value="1"><%=txtIPWatchedUser%></option>
        <option value="2"><%=txtIPBlockedAccess%></option>
        <option value="3"><%=txtIPExpireCookie%></option>
        </select><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPStatusNoCSS%>
        <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPCommentOpt%>
        </b><input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="usercomment" size="<% if strIPGateCss=1 then response.write(34) else response.write(30) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>&nbsp;</td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<%=txtIPBlockedPagesbyIPDesc%>
		<% Set rs8 = Server.CreateObject("ADODB.Recordset")
		   pgkySql = "SELECT * "
		   pgkySql = pgkySql & "FROM " & strTablePrefix & "PAGEKEYS order by PAGEKEYS_PAGEKEY asc"
	
		   rs8.Open pgkySql, strConnString %>
        <select size="10" name="dbpagekey" multiple><% do until (rs8.eof)%>
        <option value="<%=rs8("PAGEKEYS_PAGEKEY")%>"><%=rs8("PAGEKEYS_PAGEKEY")%>
        </option>
        <%				
					rs8.MoveNext
    				Loop
    	%></select> <% 
        if rs8.State = 1 then rs8.Close
    	set rs8=nothing
        if strIPGatecCss = 0 then response.write ("</font>") %></p>
        </td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQPageKey%></b><br /><%=txtIPFAQPageKeyDesc%><%=txtIPEditPageKey%>
		<a href="admin_ipgate.asp?ViewPage=pagekeys"><%=txtIPClickHere%></a>
&nbsp;<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="Submit" name="B1">&nbsp;<input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="reset" value="Reset" name="B2"></p>
</form>
<br />
<%
if qry <> "" then
		rs.close
	      set rs = nothing
	end if

Case "Write_Configuration"
	
	Err_Msg = ""
	
	if Request.Form("strIPGateBan") <> "0" then
		if trim(Request.Form("strIPGateMsg")) = "" then
			Err_Msg = Err_Msg & "<li>" & txtIPAlertMsgErr & "</li>"
		end if
	end if

	if Request.Form("strIPGateLck") <> "0" then
		if (Request.Form("strIPGateBan")) = "0" then
			Err_Msg = Err_Msg & "<li>" & txtIPLockdownErr & "</li>"
		end if
	end if
	
	if Err_Msg = "" then
	
		for each key in Request.Form 
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLstring"))
			end if
		next
		
		Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
		
		closeAndGo("admin_ipgate.asp?ViewPage=Settings&save=1")
		
	else
		Response.Write	"      <h4 align=""center""><p>" & txtIPDetErr & "</p></h4>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><h5><ul>" & Err_Msg & "</ul></h5></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      <h6 align=""center""><p><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></p></h6>" & vbNewLine
	end if

Case "Logs"
Set rs = Server.CreateObject("ADODB.Recordset")

	Select Case strDBType
		Case "mysql" 
      		strSql = "SELECT * from " & strTablePrefix & "IPLOG order by IPLOG_ID desc LIMIT 0,500"
      	Case "access" 
      		strSql = "SELECT TOP 500 * from " & strTablePrefix & "IPLOG IPLOG order by IPLOG_ID desc"
      	Case else
     			strSql = "SELECT * from " & strTablePrefix & "IPLOG order by IPLOG_ID desc"
	end select

	rs.Open StrSql, strConnString
	if request.querystring("success")= 1 then
	%>
<p align="left"><%=replace(txtIPOldEntriesRemoved,"[%marker_date%]", dateadd("d", (-1* strIPGateexp), date()))%></p>
<% end if 
if request.querystring("singlesuccess") = 1 then
%>
<p align="left"><%=txtIPLogEntryRemoved%></p>
<% end if %>
<div align="center">
  <center>
  <br />
  <table border="1" style="border-collapse: collapse;width:600px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" width="90%">
    <tr>
      <td colspan="5" width="1013" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
      <p align="center"><% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPLogAdmin%></b>
      <% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
    </tr>
    <tr>
      <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtMember%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPIP%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPPageAccessed%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtDate%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtAction%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
    </tr>
    <% do while not rs.EOF and not rs.BOF %>
    <tr>
      <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><% if rs("IPLOG_MEMBERID")= "" then response.write ("Guest") else response.write (rs("IPLOG_MEMBERID")) end if%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><a href="http://ws.arin.net/cgi-bin/whois.pl?queryinput=<%=rs("IPLOG_IP")%>" target="_blank"><%=rs("IPLOG_IP")%></a><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><a href="<%=rs("IPLOG_PATHINFO")%>"><%=rs("IPLOG_PATHINFO")%></a><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=ChkDate(rs("IPLOG_DATE"))%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
      <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
      <p><% if strIPGateCss = 0 then response.write(fontnocss)%>[<a href="admin_ipgate.asp?ViewPage=UserSettings&qry=<%= rs("IPLOG_ID") %>"><%=txtIPAddUser%></a>] [<a href="admin_ipgate.asp?ViewPage=IPBanning&qry=<%= rs("IPLOG_ID") %>"><%=txtIPAddIP%></a>] [<a href="admin_ipgate.asp?ViewPage=deletelogip&qry=<%=rs("IPLOG_ID")%>"><%=txtIPDel%></a>]<% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
    </tr>
    <%
	rs.MoveNext
	loop
	if rs.State = 1 then rs.Close
	%>
    <tr>
      <td colspan="5" width="100%" align="center"><br />
     <form method="POST" action="admin_ipgate.asp?ViewPage=deletelog&qry=<%=StrIpgateExp%>">
        <input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtIPEraseAllExpLogs%>
  <% Select Case strIPGateexp
  	Case "0"
		response.write(".")
	Case "1"
		response.write(txtIPFrm2DaysAndOlder)
	Case "7", "14", "21", "28", "35", "42", "48", "56", "63", "70"
		tmpIPGateexp = cint(strIPGateexp)
		response.write(replace(txtIPFrmWeeksAndOlder, "[%marker_weeks%]", tmpIPGateexp/7) & ".")
	Case else
		response.write(".")
	end select
	%>
  <%
' ML Changes:
' Got rid of this gnarly 12 level if then else and replaced it with a select case
' so that you only ned 2 ml vars for all 12 choices.
'   if strIPGateexp = "0" then response.write(".") else if strIPGateexp = "1" then response.write(" from 2 days ago and older.") else if strIPGateexp = "7" then response.write(" from 1 week ago and older.") else if strIPGateexp = "14" then response.write(" from 2 weeks ago and older.") else if strIPGateexp = "21" then response.write(" from 3 weeks ago and older.") else if strIPGateexp = "28" then response.write(" from 4 weeks ago  and older.") else if strIPGateexp = "35" then response.write(" from 5 weeks ago and older.") else if strIPGateexp = "42" then response.write(" from 6 weeks ago and older.") else if strIPGateexp = "49" then response.write(" from 7 weeks ago and older.") else if strIPGateexp = "56" then response.write(" from 8 weeks ago  and older.") else if strIPGateexp = "63" then response.write(" from 9 weeks ago and older.") else if strIPGateexp = "70" then response.write(" from 10 weeks ago and older.")
%>" name="B1">
      </form>
      <% if strIPGatecCss = 0 then response.write ("</font>") %>
      </td>
    </tr>
  </table>
  </center>
</div>

<%
Case "deletelog"
	Set rs = Server.CreateObject("ADODB.Recordset")
	StrIPGateExp = DateToStr(dateAdd("d",-StrIPGateExp,now()))
	StrSql = "DELETE FROM " & strTablePrefix & "IPLOG WHERE IPLOG_DATE < '" & StrIPGateExp & "'"
	rs.Open StrSql, strConnString 
	if rs.State = 1 then rs.Close
	
closeAndGo("admin_ipgate.asp?ViewPage=Logs&success=1")

Case "deletelogip"
	Set rs = Server.CreateObject("ADODB.Recordset")

	StrSql = "DELETE FROM " & strTablePrefix & "IPLOG WHERE IPLOG_ID=" & qry & ";"
	rs.Open StrSql, strConnString 
	if rs.State = 1 then rs.Close
	
	closeAndGo("admin_ipgate.asp?ViewPage=Logs&singlesuccess=1")

Case "adminip"

Set rs = Server.CreateObject("ADODB.Recordset")

	strSql = "SELECT  * from " & strTablePrefix & "IPLIST order by IPLIST_MEMBERID asc"
	rs.Open StrSql, strConnString 
	%>
	<div align="center">
<br />
<table border="1" style="border-collapse: collapse;width:550px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" width="90%">
  <tr>
    <td colspan="9" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
    <p align="center"><% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGATE & " " & strIPGateVer & " " & txtIPAdministration%></b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
  </tr>
  <tr>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Member</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>IP/Range/Host</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Start Date</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>End Date</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Comment</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Status</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Blocked Pages</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(catcss) else response.write(catnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><b>Action</b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
  </tr>
  <% do while not rs.EOF and not rs.BOF %>
  <tr>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=rs("IPLIST_MEMBERID")%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=rs("IPLIST_STARTIP")%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=strtodate(rs("IPLIST_STARTDATE"))%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=strtodate(rs("IPLIST_ENDDATE"))%><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><%=rs("IPLIST_COMMENT")%>&nbsp;<% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><% 	userstatus = rs("IPLIST_STATUS")
				Select Case userstatus
				Case 0 		Response.Write "Banned"
				Case 1		Response.Write "Watched"
				Case 2		Response.Write "Blocked Access"
				Case 3      response.write "Expire Cookie"
				end select
			%> <% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%><a href="admin_ipgate.asp?ViewPage=pagekeys">See 
    Blocked Pages Setting</a><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
    <td <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <p><% if strIPGateCss = 0 then response.write(fontnocss)%>[<a href="admin_ipgate.asp?ViewPage=edituser&qry=<%=rs("IPLIST_ID")%>">edit</a>] 
    [<a href="admin_ipgate.asp?ViewPage=deleteip&qry=<%=rs("IPLIST_ID")%>">del</a>]<% if strIPGatecCss = 0 then response.write ("</font>") %></p>
    </td>
  </tr>
  <%
	rs.MoveNext
	loop
	if rs.State = 1 then rs.Close
	%>
  <tr>
    <td colspan="9" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>&nbsp;</td>
  </tr>
</table></div>
<br />&nbsp;<%

Case "edituser"

Set rs = Server.CreateObject("ADODB.Recordset")

strSql = "SELECT * from " & strTablePrefix & "IPLIST WHERE IPLIST_ID =" & qry & ";"
	rs.Open StrSql, strConnString 

    %> </p>
<br />
<div align="center">
  <form method="POST" action="admin_ipgate.asp?ViewPage=editipupdate&qry=<%=rs("IPLIST_ID")%>">
    <table border="1" style="border-collapse: collapse;width:600px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" width="90%" id="AutoNumber7">
      <tr>
        <td width="100%" colspan="2" align="center" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPEditUserIPRec%><% if strIPGatecCss = 0 then response.write ("</font>") %></b>
        <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPUsrName%><br />
        </b><% if qry <> "" then %>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="memberid" value="<% if rs("IPLIST_MEMBERID") <> "" then Response.Write (rs("IPLIST_MEMBERID")) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <% else %> <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="memberid"> <% end if %>
        </td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPEnterUsername%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" size="<% if strIPGateCss=1 then response.write(34) else response.write(32) end if %>" name="startip" value="<% if rs("IPLIST_STARTIP") <> "" then Response.Write (rs("IPLIST_STARTIP")) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQIPandHost%></b><br /><%=txtIPFAQIPandHostDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPStartDate%><br />
        </b>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="startdate" size="<% if strIPGateCss=1 then response.write(33) else response.write(29) end if %>" value="<% if rs("IPLIST_STARTDATE") <> "" then %><%=strtodate(rs("IPLIST_STARTDATE"))%><% end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" rowspan="2" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
		<b><%=txtIPFAQStartDate%></b><br /><%=txtIPFAQStartDateDesc%>
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPEndDate%><br />
        </b>
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="enddate" size="<% if strIPGateCss=1 then response.write(33) else response.write(29) end if %>" value="<% if rs("IPLIST_ENDDATE") <> "" then %><%=strtodate(rs("IPLIST_ENDDATE"))%><% end if %>"></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><b><%=lcase(txtStatus)%><br />
        </b><select name="userstatus" size="1">
        <option VALUE="0" <% if rs("IPLIST_STATUS")= "0" then Response.Write(" selected") %>>
        Banned User</option>
        <option VALUE="1" <% if rs("IPLIST_STATUS")= "1" then Response.Write(" selected") %>>
        Watched</option>
        <option VALUE="2" <% if rs("IPLIST_STATUS")= "2" then Response.Write(" selected") %>>
        Blocked Access</option>
        <option value="3" <% if rs("IPLIST_STATUS")= "3" then Response.Write(" selected") %>>
        Expire Cookie</option>
        </select><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPStatusNoCSS%>
        <% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPCommentOpt%>
		
        <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="usercomment" size="<% if strIPGateCss=1 then response.write(34) else response.write(30) end if %>" value="<% if rs("IPLIST_COMMENT")<> "" then Response.Write (rs("IPLIST_COMMENT")) end if %>"><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>&nbsp;</td>
      </tr>
      <tr>
        <td width="38%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPPagesBlocked%>
        <% Set rs8 = Server.CreateObject("ADODB.Recordset")
		   pgkySql = "SELECT * "
		   pgkySql = pgkySql & "FROM " & strTablePrefix & "PAGEKEYS order by PAGEKEYS_PAGEKEY asc"
	
		   rs8.Open pgkySql, strConnString %>
        <select size="10" name="dbpagekey" multiple><% do until (rs8.eof)%>
        <option value="<%=rs8("PAGEKEYS_PAGEKEY")%>"><%=rs8("PAGEKEYS_PAGEKEY")%>
        </option>
        <%				
					rs8.MoveNext
    				Loop
    	%></select> <% 
        if rs8.State = 1 then rs8.Close
    	set rs8=nothing
        if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td width="62%" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%><%=txtIPPagesBlockedDesc%><%=txtIPEditPageKey%>
		<a href="admin_ipgate.asp?ViewPage=pagekeys"><%=txtIPClickHere%></a><br />&nbsp;
		<% if strIPGatecCss = 0 then response.write ("</font>") %></td>
      </tr>
    </table>
    <p><input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtSubmit%>" name="B1">&nbsp;<input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="reset" value="Reset" name="B2"></p>
  </form>
  <% if rs.State = 1 then rs.Close %>
</div>
<br />
<%    
Case "deleteip"
Set rs = Server.CreateObject("ADODB.Recordset")

	StrSql = "DELETE FROM " & strTablePrefix & "IPLIST WHERE IPLIST_ID=" & qry & ";"
	rs.Open StrSql, strConnString 
	if rs.State = 1 then rs.Close

	Response.Write	"      <h4 align=""center""><p>IP List Updated!</p></h4>" & vbNewLine & _
			"      <meta http-equiv=""Refresh"" content=""2; URL=admin_ipgate.asp?ViewPage=adminip"">" & vbNewLine & _
			"      <h5 align=""center""><p>Congratulations!</p></h5>" & vbNewLine & _
			"      <h6 align=""center""><p><a href=""admin_ipgate.asp?ViewPage=adminip"">" & txtIPBackToAdmin & "</a></p></h6>" & vbNewLine

Case "editipupdate"

	if qry <> "" & (userip <> startip & userstatus <> 0) then 
		
		Err_Msg = ""
		
	
		if (Request.Form("startip") = userip) AND (Request.Form("userstatus") = 0) then Err_Msg = Err_Msg & txtIPCantBanSelf end if

		if ubound(split(Request.Form("startip"),".")) < 2 then Err_Msg = Err_Msg & txtIPDetErr end if
		'if ubound(split(Request.Form("startip"),".")) < 3 then Err_Msg = Err_Msg & "<li>The Starting IP Address is not complete.</li>" end if

		
		if Err_Msg = "" then
	Set rs = Server.CreateObject("ADODB.Recordset")

			if memberid = "" then memberid = 0 end if
		
			strSql = "UPDATE " & strTablePrefix & "IPLIST "
			strSql = strSql & " SET IPLIST_MEMBERID = '" & memberid & "', IPLIST_STARTIP = '" & startip & "', IPLIST_STARTDATE = '" & DateToStr(startdate) & "', IPLIST_ENDDATE = '" & DateToStr(enddate) & "', IPLIST_COMMENT = '" & usercomment & "', IPLIST_STATUS = '" & userstatus & "' "
			strSql = strSql & " WHERE IPLIST_ID = " & qry & ";"
			rs.Open StrSql, strConnString 
			if rs.State = 1 then rs.Close

			Response.Write	"      <h4 align=""center"">" & txtIPListUpdated & "</h4>" & vbNewLine & _
					"      <meta http-equiv=""Refresh"" content=""2; URL=admin_ipgate.asp?ViewPage=adminip"">" & vbNewLine & _
					"      <h5 align=""center""><p>Congratulations!</p></h5>" & vbNewLine & _
					"      <h6 align=""center""><p><a href=""admin_ipgate.asp?ViewPage=adminip"">" & txtIPBackToAdmin & "</a></p></h6>" & vbNewLine
		else
			Response.Write	"      <h4 align=""center""><p>" & txtIPDetErr & "</p></h4>" & vbNewLine & _
					"      <table align=""center"" border=""0"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td><h5><ul>" & Err_Msg & "</ul></h5></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      <h6 align=""center""><p><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></p></h6>" & vbNewLine
		end if
	
	end if

Case "addipupdate"

	if qry <> "" & (userip <> startip & userstatus <> 0) then 
	
		Err_Msg = ""

		if (Request.Form("startip") = userip) AND (Request.Form("userstatus") = 0) then Err_Msg = Err_Msg & txtIPCantBanSelf end if

		if ubound(split(Request.Form("startip"),".")) < 2 then Err_Msg = Err_Msg & txtIPDetErr end if
		
		if Err_Msg = "" then
	Set rs = Server.CreateObject("ADODB.Recordset")

			if memberid = "" then memberid = 0 end if

			strSql = "INSERT into " & strTablePrefix & "IPLIST (IPLIST_MEMBERID, IPLIST_STARTIP, IPLIST_STARTDATE, IPLIST_ENDDATE, IPLIST_COMMENT, IPLIST_STATUS)"
			strSql = strSql & "values ('" & memberid & "','" & startip & "','" & DateToStr(startdate) & "','" & DateToStr(enddate) & "','" & usercomment & "','" & userstatus & "')"
			rs.Open StrSql, strConnString 
			if rs.State = 1 then rs.Close
			qry =""

			Response.Write	"      <p align=""center"">IP List Updated!</p>" & vbNewLine & _
					"      <meta http-equiv=""Refresh"" content=""2; URL=admin_ipgate.asp?ViewPage=adminip"">" & vbNewLine & _
					"      <h4 align=""center""><p>" & txtIPCongrats & "</p></h4>" & vbNewLine & _
					"      <h5 align=""center""><p><a href=""admin_ipgate.asp?ViewPage=adminip"">" & txtIPBackToAdmin & "</a></p></h5>" & vbNewLine
		else
			Response.Write	"      <h4 align=""center""><p>" & txtIPDetErr & "</h4>" & vbNewLine & _
					"      <table align=""center"" border=""0"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td><h5 align=""center""><ul>" & Err_Msg & "</ul></h5></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      <h6 align=""center""><p><a href=""JavaScript:history.go(-1)"">" & txtGoBackData & "</a></p></h6>" & vbNewLine
		end if

	end if
	
case "pagekeys" 
%>
<div align="center">
<br />
<table border="1" style="border-collapse: collapse;width:600px;" class="grid" cellpadding="4" cellspacing="<% if strIPGateCss=1 then response.write(0) else response.write(1) end if %>" align="center" width="65%" id="AutoNumber8">
  <tr>
    <td width="100%" colspan="2" align="center" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
    <% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPEditBlockedPages%>
    </b><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
  </tr>
  <tr>
    <td width="50%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <%=txtIPPagesAlreadyBlocked%>
	<% Set rs8 = Server.CreateObject("ADODB.Recordset")
		   pgkySql = "SELECT * "
		   pgkySql = pgkySql & "FROM " & strTablePrefix & "PAGEKEYS order by PAGEKEYS_PAGEKEY asc"
	
		   rs8.Open pgkySql, strConnString %>
    <select size="10" name="dbpagekey" multiple><% do until (rs8.eof)%>
    <option value="<%=rs8("PAGEKEYS_ID")%>"><%=rs8("PAGEKEYS_PAGEKEY")%>
    </option>
    <%	rs8.MoveNext
    	Loop %></select> <% 
        if rs8.State = 1 then rs8.Close
    	set rs8=nothing
 if strIPGatecCss = 0 then response.write ("</font>")
%> </p>
    </td>
    <td width="50%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form7_Validator(theForm)
{

  if (theForm.addpage.value == "")
  {
    alert("<%=txtIPAddPageErr%>");
    theForm.addpage.focus();
    return (false);
  }

  if (theForm.addpage.value.length < 1)
  {
    alert("<%=txtIPAddPage1CharErr%>");
    theForm.addpage.focus();
    return (false);
  }
  return (true);
}
//--></script><form method="POST" action="admin_ipgate.asp?ViewPage=addpgkey" name="FrontPage_Form7" onsubmit="return FrontPage_Form7_Validator(this)" language="JavaScript">
      <p><b><%=txtIPAddAPage%><br />
      </b>
	  <input <% if strIPGateCss = 1 then response.write(ClassText)%> type="text" name="addpage" size="<% if strIPGateCss=1 then response.write(27) else response.write(23) end if %>"><br />
      <%=txtIPAddAPageDesc%><br />
      <br />
      <input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtIPAddThisPage%>" name="B1"></p>
    </form>
&nbsp; <%if strIPGatecCss = 0 then response.write ("</font>")%> </td>
  </tr>
  <tr>
    <td width="50%" align="left" valign="top" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
    <%=txtIPAddPageDesc%>
	</td>
    <td width="50%" align="center" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
   <script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form8_Validator(theForm)
{

  if (theForm.removepagekey.selectedIndex < 0)
  {
    alert("<%=txtIPRemSelErr%>");
    theForm.removepagekey.focus();
    return (false);
  }

  if (theForm.removepagekey.selectedIndex == 0)
  {
    alert("<%=txtIPRevSel1stErr%>");
    theForm.removepagekey.focus();
    return (false);
  }
  return (true);
}
//--></script>
<form method="POST" action="admin_ipgate.asp?ViewPage=removepgkey" onsubmit="return FrontPage_Form8_Validator(this)" language="JavaScript" name="FrontPage_Form8">
      <p><b>Remove a Page<br />
      </b><% Set rs6 = Server.CreateObject("ADODB.Recordset")
		   pgkySql = "SELECT * "
		   pgkySql = pgkySql & "FROM " & strTablePrefix & "PAGEKEYS order by PAGEKEYS_PAGEKEY asc"
	
		   rs6.Open pgkySql, strConnString %>
      <select size="1" name="removepagekey">
      <option value="0"><%=txtIPSelectAPage%></option>
      <% do until (rs6.eof)%>
      <option value="<%=rs6("PAGEKEYS_ID")%>"><%=rs6("PAGEKEYS_PAGEKEY")%>
      </option>
      <% rs6.MoveNext
        Loop
      %></select> <% 
    if rs6.State = 1 then rs6.Close
    set rs6=nothing
 if strIPGatecCss = 0 then response.write ("</font>")
%> <br />
      <br />
      <input <% if strIPGateCss = 1 then response.write(ClassBouton)%> type="submit" value="<%=txtIPRemSelected%>" name="B1"></p>
    </form>
    <%if strIPGatecCss = 0 then response.write ("</font>")%></td>
  </tr>
</table>
</div>
<br />

<%
case "addpgkey"

Set rs7 = Server.CreateObject("ADODB.Recordset")

StrSql = "INSERT into " & strTablePrefix & "PAGEKEYS (PAGEKEYS_PAGEKEY) "
StrSql = StrSql & "values ('" & trim(chkString(request.form("addpage"),"SQLString")) & "')"

rs7.Open StrSql, strConnString 

if rs7.State = 1 then rs7.Close
set rs7 = nothing	

closeAndGo("admin_ipgate.asp?ViewPage=pagekeys")

case "removepgkey"

deleteme=request.form("removepagekey")

Set rs = Server.CreateObject("ADODB.Recordset")

StrSql = "DELETE FROM " & strTablePrefix & "PAGEKEYS WHERE PAGEKEYS_ID=" & deleteme & ";"
rs.Open StrSql, strConnString 
if rs.State = 1 then rs.Close
closeAndGo("admin_ipgate.asp?ViewPage=pagekeys")

case "logarchive"

If Request.Form("mode")="archive" Then %>
<br />
<p></p>
<div align="center">
  <center>
  <table border="1" cellpadding="4" cellspacing="0" width="600" style="border-collapse: collapse;width:600px;" class="grid">
    <tr>
      <td colspan="5" width="100%" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
      <p align="center"><% if strIPGateCss = 0 then response.write(headfontnocss)%><b><%=strSiteTitle & " " & txtIPGate & " " & strIPGateVer & " " & txtIPLogArchiving%>
      </b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
    </tr>
    <tr><td colspan="5">
<% Set archiveObj = Server.CreateObject("Scripting.FileSystemObject")
 if (0 <> Err) Then 
  Response.Write txtIPLogArchivingErr & Err.description
 Else
 txtbase=Server.MapPath("ipgate_logarchive.asp")
 scriptname = Split(txtbase,"\")
 txtpath=""
 For i=0 to ubound(scriptname)-1
  txtPath=txtpath &"\" & scriptname(i)
 Next
txtpath=right(txtpath,Len(txtpath)-1)
txtpath=txtpath & "\files\ipgatelogs\" & Replace(FormatDateTime(Date(),2),"/","") & Replace(FormatDateTime(Time(),4),":","") & ".Log"
strIPGateArchive=chkString(Request.Form("strIPGateArchive"),"sqlstring")
strIPGateArchive=DateToStr(DateAdd("d",-strIPGateArchive,now()))
Set rs = Server.CreateObject("ADODB.Recordset")
strSql = "SELECT * from " & strTablePrefix & "IPLOG WHERE IPLOG_DATE < '" & strIPGateArchive & "' order by IPLOG_ID desc"
rs.Open StrSql, strConnString 
If Not rs.EOF And Not rs.BOF Then
  Set myDefaultFile = archiveobj.CreateTextFile(txtpath,True,False)
  strline="---" &replace(txtIPLogFileMsg, "[%marker_site%]", strSiteTitle) & " " & FormatDateTime(Date(),2)& " ---"
  myDefaultFile.WriteLine(strLine)
  strline= txtIPArcLogHeader
  myDefaultFile.WriteLine(strLine)
  Response.Write replace(txtIPArchOlderThan,"[%marker_date%]", ChkDate(strIPGateArchive)) & "<br />"
Else
 Response.Write txtIPArchNoRecsErr & "<br />"
End if 
do While Not rs.EOF And Not rs.BOF
  strLine=rs("IPLOG_DATE")&"|"& rs("IPLOG_ID")&"|"& rs("IPLOG_IP")&"|"& rs("IPLOG_MEMBERID")&"|"& rs("IPLOG_PATHINFO")
  myDefaultFile.WriteLine(strLine)
rs.MoveNext
Loop
if rs.State = 1 Then rs.Close
myDefaultFile.Close
Err.Clear
Set myDefaultFile = Nothing
Set archiveObj = Nothing

If Request.Form("logclear")="on" Then
	Set rs = Server.CreateObject("ADODB.Recordset")
	StrSql = "DELETE FROM " & strTablePrefix & "IPLOG WHERE IPLOG_DATE < '" & strIPGateArchive & "'"
	rs.Open StrSql, strConnString 
	if rs.State = 1 Then rs.Close
	Response.Write replace(txtIPLogsOlderThan,"[%marker_date%]", ChkDate(strIPGateArchive))& "<br />"
End If

Response.Write "</td></tr></table><p>&nbsp;</p><p>&nbsp;</p>"
End If

Else
bannedcount=0
watchedcount=0
blockedcount=0
'referer=Request.ServerVariables("HTTP_REFERER")
referer=pagereq
qry=chkString(request.querystring("qry"),"sqlstring")
userhost=request.servervariables("REMOTE_HOST")
memberid=Trim(chkString(Request.form("memberid"),"SQLString"))
startip=Trim(chkString(request.form("startip"),"sqlstring"))
startdate=trim(chkString(request.form("startdate"),"sqlstring"))
enddate=trim(chkString(request.form("enddate"),"sqlstring"))
usercomment=trim(chkString(request.form("usercomment"),"SQLString"))
userstatus=trim(chkString(request.form("userstatus"),"sqlstring"))
dbpagekey=trim(chkString(request.form("dbpagekey"),"SQLString"))
userdate=strCurDateString


Set rs5 = Server.CreateObject("ADODB.Recordset")

StrSql = "SELECT * "
StrSql = StrSql & "FROM " & strTablePrefix & "IPLIST"
rs5.Open StrSql, strConnString
do until (rs5.eof)
   if rs5("IPLIST_STATUS") = 0 then bannedcount=bannedcount+1
   if rs5("IPLIST_STATUS") = 1 then watchedcount=watchedcount+1
   if rs5("IPLIST_STATUS") = 2 then blockedcount=blockedcount+1
   
   dbrecord = rs5("IPLIST_STARTIP") & "."
   dbrecordarr = split(dbrecord,".")
   useriparr = userip & "."
   useriparr = split(userip,".")
   if dbrecordarr(0) =  useriparr(0) then
		if dbrecordarr(1) =  useriparr(1)then
			if dbrecordarr(2)  = useriparr(2) then
				if dbrecordarr(3) = "" or dbrecordarr(2) = "" then
				   if rs5("IPLIST_STATUS") = 0 then
					%><center><b>
<h4><font color="#FF0000"><%=replace(txtIPGIPRangeBan,"[%marker_userip%]", userip)%></font></h4>
</b></center><%
					warning = "Yes"
				   end if
				end if
			end if
		end if
	end if

rs5.MoveNext
Loop
if rs5.state = 1 then rs5.close
set rs5 = nothing

	Set rs = Server.CreateObject("ADODB.Recordset")

	Select Case strDBType
		Case "mysql" 
      		strSql = "SELECT * from " & strTablePrefix & "IPLOG order by IPLOG_ID desc LIMIT 0,500"
      	Case "access" 
      		strSql = "SELECT TOP 150 * from " & strTablePrefix & "IPLOG IPLOG order by IPLOG_ID desc"
      	Case else
     			strSql = "SELECT * from " & strTablePrefix & "IPLOG order by IPLOG_ID desc"
	end select

	rs.Open StrSql, strConnString 

	%><br />
<p></p>
<br />
<p></p>
<div align="center">
  <center>
  <table border="1" cellpadding="4" cellspacing="0" width="600" style="border-collapse: collapse;width:600px;" class="grid">
    <tr>
      <td colspan="5" width="100%" <% if strIPGateCss=1 then response.write(headcss) else response.write(headnocss) end if %>>
      <p align="center"><% if strIPGateCss = 0 then response.write(headfontnocss)%><%=strSiteTitle & " " & txtIPGate & "  " & strIPGateVer & " " & txtIPLogArchiving%>
      </b><% if strIPGatecCss = 0 then response.write ("</font>") %></p>
      </td>
    </tr>
   
    <% 
FSOstatus = False  
Err = 0
writeError=False
on error Resume Next
Set testObj = Server.CreateObject("Scripting.FileSystemObject")
if (0 = Err) Then FSOstatus = true Else FSOstatus = False
If FSOstatus = True Then
txtbase=Server.MapPath("ipgate_logarchive.asp")
scriptname = Split(txtbase,"\")
txtpath=""
For i=0 to ubound(scriptname)-1
 txtPath=txtpath &"\" & scriptname(i)
Next
txtpath=right(txtpath,Len(txtpath)-1)
txtpath=txtpath & "\files\ipgatelogs\testfso.txt"
Set myDefaultFile = testobj.CreateTextFile(txtpath,True,False)
Response.Write(Err.Description)
myDefaultFile.WriteLine(txtIPTestFileMsg)
If Err.number <> 0 Then
writeStatus=False
Else
writeStatus=True
End If
myDefaultFile.Close
Err.Clear
Set myDefaultFile = Nothing

Set testObj = Nothing
If FSOstatus=True And writestatus=True Then
  enabled=True
Else
  enabled=False
End If
End If

%>
 <tr>
 <td align="right" <%response.write(headnocss)%>><%response.write(headfontnocss)%><b><%=txtIPFSO%>:</b></font></td>
 <td <%=catnocss%>><b><% if fsostatus=False then Response.Write(txtIPNotEnabled) Else Response.Write(txtIPEnabled) End If %></b></td>
 <td align="right" <%response.write(headnocss)%>><%response.write(headfontnocss)%><b><%=txtIPFolders%>:</b></font></td>
 <td colspan="2" <%=catnocss%>><b><% if writestatus=False then Response.Write(txtIPWritable) Else Response.Write(txtIPWritable) End If %></b></td>

 </tr>
  <% If enabled=False Then %>  
    <tr>
    <td colspan="5" <% If strIPGateCss=1 Then Response.write(catcss) Else Response.write(catnocss) End If %>><% =fontnocss%>
	<%=txtIPLogErrDesc%>
    </font></td>
    </tr>
    <tr><td colspan="5" align="center" <%=headnocss%>><% =headfontnocss%><b><%=txtIPLogArchIs%><% if enabled=True then response.Write(txtIPEnabled) else response.Write(txtIPNotEnabled)%></b></font></td></tr>
   <% Else %> 
   <form name="frmarchive" method="post" action="admin_ipgate.asp?ViewPage=logarchive">
   <tr><td colspan="5" align="center" <%=headnocss%>><% =headfontnocss%><b><%=txtIPLogArchIs%><% if enabled=True then response.Write(txtIPEnabled) else response.Write(txtIPNotEnabled)%></b></font></td></tr>    
      <tr><td><input type="hidden" name="mode" value="archive">
    <% If strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPArchLogPeriod%></b><br />
        <select name="strIPGateArchive">
        <option value="0" <%= chkSelect(strIPGateexp,0) %><%=txtIPAllLogs%></option>
        <option value="1" <%= chkSelect(strIPGateexp,1) %>>2 <%=txtDays%></option>
        <option value="7" <%= chkSelect(strIPGateexp,7) %>>1 <%=txtWeek%></option>
        <option value="14" <%= chkSelect(strIPGateexp,14) %>>2 <%=txtWeeks%></option>
        <option value="21" <%= chkSelect(strIPGateexp,21) %>>3 <%=txtWeeks%></option>
        <option value="28" <%= chkSelect(strIPGateexp,28) %>>4 <%=txtWeeks%></option>
        <option value="35" <%= chkSelect(strIPGateexp,35) %>>5 <%=txtWeeks%></option>
        <option value="42" <%= chkSelect(strIPGateexp,42) %>>6 <%=txtWeeks%></option>
        <option value="49" <%= chkSelect(strIPGateexp,49) %>>7 <%=txtWeeks%></option>
        <option value="56" <%= chkSelect(strIPGateexp,56) %>>8 <%=txtWeeks%></option>
        <option value="63" <%= chkSelect(strIPGateexp,63) %>>9 <%=txtWeeks%></option>
        <option value="70" <%= chkSelect(strIPGateexp,70) %>>10 <%=txtWeeks%></option>
        </select><% if strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td colspan="4" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
        <br />
        <%=txtIPArchLogPeriodDesc%><% If strIPGatecCss = 0 Then Response.write ("</font>") %>
    </td></tr>
    <tr><td>
    <% If strIPGateCss = 0 then response.write(fontnocss)%><b><%=txtIPClearLogsAfter%></b><br />
     <input name="logclear" type="checkbox" CHECKED>
        <% If strIPGatecCss = 0 then response.write ("</font>") %></td>
        <td colspan="4" align="left" <% if strIPGateCss=1 then response.write(forumcss) else response.write(forumnocss) end if %>>
        <% if strIPGateCss = 0 then response.write(fontnocss)%>
        <br />
        <%=txtIPConfirmLogCleared%>
          </td></tr>
      <tr></tr>   <td colspan="5" > 
      &nbsp;&nbsp;&nbsp;<input type="submit"  value="<%=txtIPArchiveLogs%>">     
      </td></tr>
      </form>

     <% End If %>
                 
  </table><br />
  </center>
</div>

<%
  end if
case else 
	closeAndGo("admin_ipgate.asp?ViewPage=MainMenu")
end select 
%>
		<% spThemeBlock1_close(intSkin) %>
	    	</td>
	  	  </tr>
		</table>
<!--#include file="inc_footer.asp"-->
<%
Else
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
End IF

sub ipgateConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_ipgate"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
  else
	if typ = 1 then
	  cls = "block"
	  icn = "min"
	  alt = "Collapse"
	else
	  cls = "none"
	  icn = "max"
	  alt = "Expand"
	end if
	 'onclick="javascript:mwpHSs('block12<%= typ ','0');"    %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:location.reload();"><span style="margin: 2px;"><img name="blockFP<%= typ %>Img" id="blockFP<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b><%=txtIPGateMenu%></b></div>
    <div class="menu" id="blockFP<%= typ %>" style="display: <%= cls %>;">
    <a href="admin_home.asp"><%= icn_bar %><%=txtAdminHome%><br /></a>
	<a href="admin_ipgate.asp?ViewPage=MainMenu"><%= icn_bar %><%=txtIPMainMenu%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=adminip"><%= icn_bar %><%=txtIPAdmin%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=UserSettings"><%= icn_bar %><%=txtIPUserBan%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=IPBanning"><%= icn_bar %><%=txtIPIPBan%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=Logs"><%= icn_bar %><%=txtIPViewLogs%><br /></a>
	<a href="admin_ipgate.asp?ViewPage=logarchive"><%= icn_bar %><%=txtIPArchiveLogs%><br /></a>
	<a href="admin_ipgate.asp?ViewPage=deletelog&qry=<%=StrIpgateExp%>"><%= icn_bar %><%=txtIPEraseOldLogs%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=pagekeys"><%= icn_bar %><%=txtIPEditBlockedPages%><br /></a>
    <a href="admin_ipgate.asp?ViewPage=Settings"><%= icn_bar %><%=txtIPSettings%><br /></a>
    <a href="JavaScript:openWindow5('pop_help.asp?mode=3')"><%= icn_bar %><%=txtIPGateHelp%><br /></a>
	</div>
  <%
  end if
end sub
%>