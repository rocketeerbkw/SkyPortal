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
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::: %>
	</head>
	<body<%=spThemeBodyTag%>>
<a name="top"></a>
<% spThemeStart() %>
<% headerTop() %>
<% spThemeHeader_open() %>
<table cellspacing="0" cellpadding="0" width="100%" border="0">
  <tr>
	<td align="left" valign="middle"><a href="default.asp"><img title="<% =strSiteTitle %>" alt="<% =strSiteTitle %>" border="0" class="logo" src="<%= strHomeUrl %>Themes/<%= strTheme %>/<%= subTheme %><%= strTitleImage %>" /></a></td>
    <td align="right" valign="top">
	  	<% 
		if strLoginType = 0 and strdbntusername = "" and strAuthType = "db" and strNewReg = 1 then
			showloginbox()
		else			
		  Select Case strHeaderType
			case 0
				shoNotta()
			case 1, 2
				showheaderBanner()
			case 3
				showIcons()
			case 4
				showOther()
			case else
				shoNotta()
		  End Select
		end if
		 %>
	</td>
  </tr>
</table>
<% spThemeHeader_close() %>

<% sub shoNotta() %>
<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
<tr><td width="100%">&nbsp;</td></tr></table>
<% end sub

sub showHeaderBanner()
  response.Write("<div class=""sp_Banner"">")
  displayABanner 1,strHeaderType,10,,false
  response.Write("</div>")
end sub

Sub writeFlash(swfImg) %>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="468" height="60" id="Flash_Banner" align=""><param name=movie value="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bannerID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bannerName) %>"><param name=quality value=high><embed src="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bannerID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bannerName) %>" quality="high" name="Flash_Banner" height="60" width="468" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>
<% end sub %>

<% Sub showloginbox() %>
    <table class="sp_Header_Login" cellpadding="0" cellspacing="0" style="border-collapse: collapse;" align="right">
		<tr>
		  <td align="right" valign="middle">
            <form action="<% =Request.ServerVariables("URL") %>" method="post" id="formb1" name="formb1">
          <table width="100%" border="0" cellpadding="3" cellspacing="0">
              <input type="hidden" name="Method_Type" value="login" />
              <tr> 
                <td width="90" align="center" valign="middle"><b>&nbsp;<%= txtUsrName %>:</b><br />
                  &nbsp;<input class="textbox" type="text" name="Name" size="10" />
                </td>
                <td width="90" align="center" valign="middle"><b><%= txtPass %>:</b><br />
                  <input class="textbox" type="password" name="Password" size="10" />
                </td>
                <td width="75" align="center" valign="middle">&nbsp;<input class="btnLogin" type="submit" value="<%= txtLogin %>" id="submitx1" name="submitx1" />
				</td>
              </tr>
              <tr> 
                <td colspan="3" align="center"> 
                  <input type="checkbox" name="SavePassWord" value="true" checked />
                  <span class="fSmall"><%= txtSvPass %>&nbsp;&nbsp;</span>
                  <%if (lcase(strEmail) = "1") then %>
                  <a href="password.asp"><span class="fSmall"><%= txtForgotPass %>?</span></a>&nbsp;&nbsp; 
                  <% end if 
				  if strNewReg = 1 then %>
                  <br /><span class="fSmall"><%= txtNotMember %>?</span> 
                  <a href="policy.asp"><span class="fSmall"><%= txtRegNow %>!</span></a>
				  <% End If %>
				  </td>
              </tr>
          </table>
            </form>
                </td>
			</tr>
    </table>
<% End Sub %>

<% sub showIcons() %>
<table width="84%" border="0" cellspacing="0" cellpadding="0" class="sp_headerIcons" align="center">
  <tr align="center" valign="middle"> 
    <td width="12%"><a href="fhome.asp" title="<%= txtView %> <%= txtForum %>"><img title="<%= txtView %> <%= txtForum %>" alt="<%= txtView %> <%= txtForum %>" src="Themes/<%= strTheme %>/forums.gif" border="0" /></a></td>
    <td width="12%"><a href="events.asp" title="<%= txtView %> <%= txtCalendar %>"><img title="<%= txtView %> <%= txtCalendar %>" alt="<%= txtView %> <%= txtCalendar %>" src="Themes/<%= strTheme %>/events.gif" border="0" /></a></td>
    <td width="12%"><a href="article.asp" title="<%= txtView %> Articles"><img title="<%= txtView %> Articles" alt="<%= txtView %> <%= txtArticles %>" src="Themes/<%= strTheme %>/articles.gif" border="0" /></a></td>
    <td width="12%"><a href="dl.asp" title="<%= txtView %> <%= txtDownloads %>"><img title="<%= txtView %> <%= txtDownloads %>" alt="<%= txtView %> <%= txtDownloads %>" src="Themes/<%= strTheme %>/dl.gif" border="0" /></a></td>
    <td width="12%"><a href="links.asp" title="<%= txtView %> <%= txtLinks %>"><img title="<%= txtView %> <%= txtLinks %>" alt="<%= txtView %> <%= txtLinks %>" src="Themes/<%= strTheme %>/links.gif" border="0" /></a></td>
    <td width="12%"><a href="pic.asp" title="<%= txtView %> <%= txtPics %>"><img title="<%= txtView %> <%= txtPics %>" alt="<%= txtView %> <%= txtPics %>" src="Themes/<%= strTheme %>/pic.gif" border="0" /></a></td>
    <td width="12%"><a href="classified.asp" title="<%= txtView %> <%= txtClassifieds %>"><img title="<%= txtView %> <%= txtClassifieds %>" alt="<%= txtView %> <%= txtClassifieds %>" src="Themes/<%= strTheme %>/features.gif" border="0" /></a></td>
  </tr>
  <tr align="center" valign="middle"> 
    <td><a href="fhome.asp" title="<%= txtView %> <%= txtForum %>"><b><%= txtForum %></b></a></td>
    <td><a href="events.asp" title="<%= txtView %> <%= txtCalendar %>"><b><%= txtEvents %></b></a></td>
    <td><a href="article.asp" title="<%= txtView %> <%= txtArticles %>"><b><%= txtArticles %></b></a></td>
    <td><a href="dl.asp" title="<%= txtView %> <%= txtDownloads %>"><b><%= txtDownloads %></b></a></td>
    <td><a href="links.asp" title="<%= txtView %> <%= txtLinks %>"><b><%= txtLinks %></b></a></td>
    <td><a href="pic.asp" title="<%= txtView %> <%= txtPics %>"><b><%= txtPics %></b></a></td>
    <td><a href="classified.asp" title="<%= txtView %> <%= txtClassifieds %>"><b><%= txtClassifieds %></b></a></td>
  </tr>
</table>
<% end sub %>

<% sub showOther() %>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td width="100%">&nbsp;
<!-- Add code here to display in the header area
	 when "other" is selected as the Header Type -->
</td></tr></table>
<% end sub %>