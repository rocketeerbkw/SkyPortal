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
function app_LeftColumn()

'mnu.menuName = "m_pictures"
'mnu.title = "clickMenu2 v1"
'mnu.template = 4
'mnu.thmBlk = 0
'mnu.shoExpanded = 1
'mnu.canMinMax = 0
'mnu.GetMenu()

    getMenu(intAppID)
  	intShow = 4
  	photos_sm("new")
end function

function app_MainColumn_top()
	modFeatures()
  	'intShow = 3
  	'intDir = 1
  '	photos_sm("featured")
end function

function app_MainColumn_bottom()
	intShow = 6
	photos_lg("new")
end function

function app_RightColumn()
	intShow = 3
	photos_sm("featured")
	intShow = 3
	photos_sm("top")
	intShow = 3
	photos_sm("rated")
end function

function app_Footer()
end function
%>