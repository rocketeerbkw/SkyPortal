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
<!--#include file="inc_top_short.asp" -->
<!--#include file="modules/drafts/drafts_functions.asp" -->
<script language="javascript" type="text/javascript" src="tiny_mce/tiny_mce_popup.js"></script>
<%
if Request("mode") <> " " and NOT isNull(Request("mode")) then
	if isNumeric(Request("mode")) then
		iMode = Request("mode")
	else
		closeAndGo("stop")
	end if
end if

spThemeBlock1_open(intSkin)

select case iMode
	case 1
		call popDrafts(strUserMemberID)
	case 2
		tmpVMsg = saveDraft()
		showMsgBlock 0,tmpVMsg
end select

spThemeblock1_close(intSkin)
%>
<!--#include file="inc_footer_short.asp" -->
