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

spThemeTitle = ""
spThemeShortBodyTag = "onLoad=""window.focus()"""
subTheme=""
hasSubTheme = false

sub spThemeStart_short()
  Response.Write "<table class=""spTheme"" align=""center"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  Response.Write "<tr><td>" & vbcrlf
end sub
'
sub spThemeStart()
  Response.Write "<table class=""spPageContainer"" align=""center"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  Response.Write "<tr><td class=""spPageLeft"" valign=""top""></td><td valign=""top"">" & vbcrlf
  Response.Write "<table class=""spThemePage"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
  Response.Write "<tr><td width=""100%"" valign=""top"">" & vbcrlf
end sub

sub spThemeHeader_style()
  addCSSfile("Themes/" & strTheme & "/style_core.css")
  addCSSfile("Modules/custom_styles.css")
end sub

sub spThemeHeader_open()
      response.Write "<div class=""spHeader"">" & vbcrlf
        response.Write "<div class=""spHeader_tl"">" & vbcrlf
          response.Write "<div class=""spHeader_tr"">" & vbcrlf
            response.Write "<div class=""spHeader_tc""></div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
        response.Write "<div class=""spHeader_ml"">" & vbcrlf
          response.Write "<div class=""spHeader_mr"">" & vbcrlf
            response.Write "<div class=""spHeader_content"">" & vbcrlf
end sub

sub spThemeHeader_close()
			response.Write "</div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
        response.Write "<div class=""spHeader_bl"">" & vbcrlf
          response.Write "<div class=""spHeader_br"">" & vbcrlf
            response.Write "<div class=""spHeader_bc""></div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
      response.Write "</div>" & vbcrlf
end sub

sub spThemeNavBar_open()
  Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
  Response.Write "<tr>" & vbcrlf
  Response.Write "<td class=""sp_NavLeft"" align=""left""></td>" & vbcrlf
  Response.Write "<td class=""sp_NavTile"">" & vbcrlf
end sub

sub spThemeNavBar_close()
  Response.Write "</td>" & vbcrlf
  Response.Write "<td class=""sp_NavRite"" align=""right""></td>" & vbcrlf
  Response.Write "</tr>" & vbcrlf
  Response.Write "</table>" & vbcrlf
end sub

spThemeBlock_subTitleCell = "class=""tSubTitle"""

mm = 0
catHide = "block"
catImg = "min"
catAlt = txtCollapse

sub spThemeBlock1_open(tbNum)
	  catHide = "block"
	  catImg = "min"
	  catAlt = txtCollapse
	  mm = mm + 1
	  mwpTb = spThemeMM
	  if mwpTb = "" then
	    mwpTb = "t" & mm
	  end if
	  
      response.Write "<div class=""spThemeBlock"&tbNum&""">" & vbcrlf
	  if trim(spThemeTitle) = "" then
        response.Write "<div class=""spThemeBlock"&tbNum&"_sh_tr"">" & vbcrlf
          response.Write "<div class=""spThemeBlock"&tbNum&"_sh_tl"">" & vbcrlf
            response.Write "<div class=""spThemeBlock"&tbNum&"_sh_tc"">" & vbcrlf
	  else
        response.Write "<div class=""spThemeBlock"&tbNum&"_tr"">" & vbcrlf
          response.Write "<div class=""spThemeBlock"&tbNum&"_tl"">" & vbcrlf
            response.Write "<div class=""spThemeBlock"&tbNum&"_tc"">" & vbcrlf
		 if trim(spThemeMM) <> "" and spThemeMM <> "&nbsp;" then
		 	showMM spThemeMM,mm,tbNum
		 end if 
              response.Write "<h4>" & spThemeTitle & "</h4>" & vbcrlf
	  end if
            response.Write "</div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
        response.Write "<div class=""spThemeBlock"&tbNum&"_ml"" id=""" & mwpTb & """ style=""display:" & catHide & """>" & vbcrlf
          response.Write "<div class=""spThemeBlock"&tbNum&"_mr"">" & vbcrlf
            response.Write "<div class=""spThemeBlock"&tbNum&"_content"">" & vbcrlf
 
	spThemeMM = ""
	spThemeTitle = ""
end sub

sub spThemeBlock1_close(tbNum)
			response.Write "</div>"
          response.Write "</div>"
        response.Write "</div>"
        response.Write "<div class=""spThemeBlock"&tbNum&"_bl"">"
          response.Write "<div class=""spThemeBlock"&tbNum&"_br"">"
            response.Write "<div class=""spThemeBlock"&tbNum&"_bc""></div>"
          response.Write "</div>"
        response.Write "</div>"
      response.Write "</div>"
end sub

sub showMM(nam,num,tid)
	if request.Cookies(strUniqueID & "hide")("" & nam & "") <> "" then
		if request.Cookies(strUniqueID & "hide")("" & nam & "") = "1" then
			catHide = "none"
			catImg = "max"
			catAlt = txtExpand
		else
			catHide = "block"
			catImg = "min"
			catAlt = txtCollapse
		end if
	else
	  if left(nam,2) = "x_" then
		catHide = "none"
		catImg = "max"
		catAlt = txtExpand
	  end if
	end if %>
	<span class="spThemeblock<%=tid%>MinMax" style="display:inline; float:right; position:relative;"><img name="<%= nam %>Img" id="<%= nam %>Img" src="Themes/<%= strTheme %>/icon_<%= catImg %>.gif" onclick="javascript:mwpHSx('<%= nam %>');" style="cursor:pointer;" alt="<%= catAlt %>" title="<%= catAlt %>" /></span>
<%
end sub

sub spThemeBlock2_open()
  '::<!-- rounded corners - top **** -->
  Response.Write "<div class=""spThemeBlockRC"">"
  Response.Write "<div class=""rtop""><div class=""r1""></div><div class=""r2""></div><div class=""r3""></div><div class=""r4""></div></div>"
  response.Write "<div class=""spThemeBlockRC_content"">" & vbcrlf
  if spThemeTitle <> "" then
	response.Write "<h1>" & spThemeTitle & "</h1>" & vbcrlf
    spThemeTitle = ""
  end if
end sub

sub spThemeBlock2_close()
  Response.Write "</div>"
  '<!-- rounded corners - bottom **** -->
  Response.Write "<div class=""rbottom""><div class=""r4""></div><div class=""r3""></div><div class=""r2""></div><div class=""r1""></div></div>"
  Response.Write "</div>"
end sub

sub spThemeBlock3_open()
	'Response.Write "<div class=""tPlain"" style=""padding:8px;""></div>" & vbcrlf
	response.Write "<fieldset>" & vbcrlf
	if spThemeTitle <> "" then
	  response.Write "<legend><b>" & spThemeTitle & "&nbsp;</b></legend>" & vbcrlf
	  spThemeTitle = ""
	end if
end sub

sub spThemeBlock3_close()
  response.Write "</fieldset>" & vbcrlf
end sub

sub spThemeBlock4_open()
	Response.Write "<div style=""padding:8px;"">" & vbcrlf
	response.Write "<fieldset>" & vbcrlf
	if spThemeTitle <> "" then
	  response.Write "<legend><b>" & spThemeTitle & "&nbsp;</b></legend>" & vbcrlf
	  spThemeTitle = ""
	end if
end sub

sub spThemeBlock4_close()
  response.Write "</fieldset></div>" & vbcrlf
end sub

sub spThemeSmallBlock_open()
  Response.Write "<table class=""spThemeSmallBlock"" " & spThemeTableCustomCode & ">" & vbcrlf 
  spThemeTableCustomCode = ""
end sub

sub spThemeSmallBlock_close()
  Response.Write "</table>" & vbcrlf
end sub

sub spThemeFooterBlock_open()
      response.Write "<div class=""spFooter"">" & vbcrlf
        response.Write "<div class=""spFooter_tl"">" & vbcrlf
          response.Write "<div class=""spFooter_tr"">" & vbcrlf
            response.Write "<div class=""spFooter_tc""></div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
        response.Write "<div class=""spFooter_ml"">" & vbcrlf
          response.Write "<div class=""spFooter_mr"">" & vbcrlf
            response.Write "<div class=""spFooter_content"">" & vbcrlf
end sub

sub spThemeFooterBlock_close()
			response.Write "</div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
        response.Write "<div class=""spFooter_bl"">" & vbcrlf
          response.Write "<div class=""spFooter_br"">" & vbcrlf
            response.Write "<div class=""spFooter_bc""></div>" & vbcrlf
          response.Write "</div>" & vbcrlf
        response.Write "</div>" & vbcrlf
      response.Write "</div>" & vbcrlf
end sub

sub spThemeEnd()
  Response.Write "</td></tr></table>" & vbcrlf
  Response.Write "</td><td class=""spPageRight""></td></tr></table>" & vbcrlf
  'Response.Write "</div>" & vbcrlf
end sub
%>