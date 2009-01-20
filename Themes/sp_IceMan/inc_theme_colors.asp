<%
':::::::::::::::::::::::::::::::::::::::::
'::: | SKYPORTAL.NET(c)Dogg Software | :::
':::::::::::::::::::::::::::::::::::::::::
'::: This file only loads once, when skin is first installed.
'::: More Skin adjustments in Admin Home >> Managers >> Skin Manager
'::: If you've come this far, you are now part of the scene :)
'::: Frosty...
%>
<%
':: Theme Folder Name
thmFolder = "sp_IceMan"				               

':: Theme Author
thmAuthor = "<a href=http://www.frozenwinds.com/>R.Frost</a>"

':: Theme Description
thmDescription = "IceMan Skin Series - IceMan-Pro Sp Version"

':: Your Site Logo Image
thmLogoImage = "Site_Logo.jpg"

':: SubSkin Value - Upto 3 Different Theme Box Styles
':: Divided into 3 columns - Left, Main, Right
':: Look in style_blocks.css for theme block styles
':: Look in style_layout.css for column styles

thmSubSkin = 3  'value set when skin first installed 
	
':: thmSubSkin possible values = 0, 1, 2, 3
':: 0 = all 3 columns use ThemeBlock 1 - .spThemeBlock1
':: 1 = Left .spThemeBlock2 - Main and Right .spThemeBlock1
':: 2 = Left and Main .spThemeBlock1 - Right .spThemeBlock3
':: 3 = Left .spThemeBlock2 - Main .spThemeBlock1 - Right .spThemeBlock3
':: Can also adjust thmSubSkin in Admin Home >> Skin Manager >> SubSkin Value:

':: No need to modify anything below here

Session("thmFolder") = thmFolder
Session("thmAuthor") = thmAuthor
Session("thmDescription") = thmDescription
Session("thmLogoImage") = thmLogoImage
Session("thmSubSkin") = thmSubSkin

newThm = replace(replace(request("tName"),"<",""),">","")
thmFolder = replace(replace(request("tFolder"),"<",""),">","")
whereto =  request.ServerVariables("HTTP_REFERER") & "?tName=" & newThm & "&tFolder=" & thmFolder & "&cmd=1"
'whereto = "admin_config_themes.asp?tName=" & newThm & "&tFolder=" & thmFolder
response.Redirect whereto
%>