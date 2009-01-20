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
function headerTop()
  'Contents of this will appear at the top of the header block on every page
  'response.Write("<div class=""spHeadTop""><center>test this: Header top</center></div>")
  response.Write("<div class=""spHeadTop""></div>")
end function

function navBarTop()
  'Contents of this will appear at the top of the navbar block on every page
  'response.Write("<div class=""spNavBarTop""><center>test this: navbar top</center></div>")
  response.Write("<div class=""spNavBarTop""></div>")
end function

function navBarBottom()
  'Contents of this will appear at the bottom of the navbar block on every page
  'response.Write("<div class=""spNavBarBottom""><center>test this: navbar bottom</center></div>")
  response.Write("<div class=""spNavBarBottom""></div>")
end function

function footerTop()
  'Contents of this will appear at the top of the footer block on every page
  'response.Write("<div class=""spFooterTop""><center>test this: Footer top</center></div>")
  response.Write("<div class=""spFooterTop""></div>")
end function

function eachPageLoad()
'this function executes on each page load

end function

function checkOncePerDay()
'this function will only execute once per day

end function

function getDonor_Level(fM_ID)

end function

function getStar_Level(fM_LEVEL, fM_POSTS)

  dim Star_Level
	Star_Level = ""
	select case fM_LEVEL
		case "1"
			if (fM_POSTS < intRankLevel1) then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColor1 & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColor2 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor2 & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColor3 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor3 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor3 & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColor4 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor4 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor4 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor4 & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColor5 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor5 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor5 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor5 & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColor5 & ".gif"" border=""0"" alt="""" />"
		case "2" 
			if fM_POSTS < intRankLevel1 then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorMod & ".gif"" border=""0"" alt="""" />"
		case "3" 
			if (fM_POSTS < intRankLevel1) then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" />"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & "<img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" /><img src=""Themes/" & strTheme & "/Stars/" & strRankColorAdmin & ".gif"" border=""0"" alt="""" />"
		case else  
			Star_Level = Star_Level & "Error"
	end select

  getStar_Level = Star_Level
end function

	  
sub customMetaTags() %>
<%
  if cust_meta <> "" then
    cust_meta = "<meta http-equiv=""Content-Type"" content=""text/html; charset=" & strCharset & """ />" & vbcrlf & cust_meta
  else
	addToMeta "http-equiv","Content-Type","text/html; charset=" & strCharset & ""
	'addToMeta "http-equiv","EXPIRES","never"
	addToMeta "http-equiv","imagetoolbar","no"
	addToMeta "name","DISTRIBUTION","GLOBAL"
	'addToMeta "name","ROBOTS","INDEX, FOLLOW"
	'addToMeta "name","REVISIT-AFTER","7 DAYS"
	addToMeta "name","RESOURCE-TYPE","DOCUMENT"
  end if
end sub

'::::: ADD YOUR CUSTOM FUNCTIONS BELOW THIS LINE :::::::::::::::::::

'::::: ADD YOUR CUSTOM FUNCTIONS ABOVE THIS LINE :::::::::::::::::::
%>