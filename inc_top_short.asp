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


'set my_Conn= Server.CreateObject("ADODB.Connection")
'my_Conn.Open strConnString

strDBNTFUserName = chkString(Request.Form("Name"),"display")
tempArr = split(strWebMaster, ",")
strSiteOwner = tempArr(0)

jsReloadOpener = "<script type=""text/javascript"">opener.document.location.reload();</script>"
jsCloseWindow = "<script type=""text/javascript"">window.close();</script>"

':: check login status - member or guest
chkLoginStatus()

':: populate member groups array
bldArrUserGroup()

':: build array for app access
'bldArrAppAccess()

if trim(curpagetype) <> "" and trim(curpagetype) <> "home" and trim(curpagetype) <> "core" and trim(curpagetype) <> "PM" then
    if not chkApp(curpagetype,"USERS") then
	  my_Conn.close
	  set my_Conn = nothing %>
	  <script type="text/javascript">
	  window.close();
	  </script>
	  <%
	  'closeAndGo("stop")
	end if
end if
%>
<!--#include file="includes/inc_ipgate.asp" -->
<!--#include file="includes/inc_theme.asp" -->
<!--#include file="includes/inc_emails.asp" -->
<!--#include file="includes/inc_editor.asp" -->
<!--include file="fp_custom.asp" -->
<% 
getPageSkin(arrGroups(0,0))

If (not hasAccess(2) and strLockDown <> 0 and lockDownOverRide <> "1") or strLockDown = "" Then
  if strAuthType <> "db" then 
'do nothing 
  else %>
	<html>
	<head>
	<%spThemeHeader_style()%>
	</head>
    <body>
	<br /><br /><br /><br /><center><b><%=txtMustBMember1%><br /><%=txtToPartic%>.</b></center><br /><br /><br /><br />
	<p align="center"><a href="JavaScript:onclick= window.close();"><%= txtCloseWin %></a></p>
	</body>
	</html>
    <%
	closeAndGo("stop")
  end if
end if
 %>
<!-- This page is generated by Sky Portal.net-->
<html>

<head>
<%
getMetaTags()
%>
<title><% =strSiteTitle %></title>
<script type="text/JavaScript">
<!--
var js_welcome = ""
 var js_none = "<%= txtNone %>";
 var js_member = "<%= txtMember %>";
 var js_admin = "<%= txtAdmin %>";
 
//  month/day arrays
 var js_months_lng=new Array("<%= txtJanuary %>", "<%= txtFebruary %>", "<%= txtMarch %>", "<%= txtApril %>", "<%= txtMay %>", "<%= txtJune %>", "<%= txtJuly %>", "<%= txtAugust %>", "<%= txtSeptember %>", "<%= txtOctober %>", "<%= txtNovember %>", "<%= txtDecember %>");
 var js_days_lng=new Array("<%= txtSunday %>", "<%= txtMonday %>", "<%= txtTuesday %>", "<%= txtWednesday %>", "<%= txtThursday %>", "<%= txtFriday %>", "<%= txtSaturday %>", "<%= txtSunday %>");
 
// pop-up calendar items
 var js_calendar = "<%= txtCalendar %>"
 var js_frm = "<%= txtForm %>"
 var js_frmfld = "<%= txtFrmFld %>"
 var js_notfnd = "<%= txtNotFound %>"
 var yxLinks=new Array("[<%= txtClose %>]", "[<%= txtClear %>]");
 
 var jsUniqueID = "<%= strUniqueID %>"

 // preload min-max images
  var mmImages = new Array(4);
  mmImages[0] = "Themes/<%= strTheme %>/icon_max.gif";
  mmImages[1] = "Themes/<%= strTheme %>/icon_min.gif";
  mmImages[2] = "Themes/<%= strTheme %>/icon_max1.gif";
  mmImages[3] = "Themes/<%= strTheme %>/icon_min1.gif";
// -->
</script>
<% 
addJSfile("includes/scripts/core.js")
addJSfile("includes/scripts/prototype.js")
addJSfile("includes/scripts/sp_ajax.js")
addJSfile("includes/scripts/effects.js")
addJSfile("includes/scripts/window.js")
addJSfile("includes/scripts/cal2.js")
addJSfile("modules/custom_scripts.js")
getJSFiles()

if (curpagetype = "pm" and iMode = 2) and not thispage = "monitor" then %>
<script type="text/javascript" src="includes/scripts/menu_com.js"></script>
<% end if %>
<% if hasEditor = true and editorType = "tinymce" and strAllowHtml = 1 then %>
<script type="text/javascript" src="tiny_mce/tiny_mce.js"></script>
<% End If %>
<script type="text/javascript">
<!--
calFormat="<%= strDateFormat %>";
var popwin = null;
function popUpWind(mypage,myname,w,h,scr,resiz){
LeftPosition = (screen.width) ? (screen.width-w)/2 : 0;
TopPosition = (screen.height) ? (screen.height-h)/2 : 0;
settings =
'height='+h+',width='+w+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scr+',toolbar=no,resizable='+resiz+',menubar=no'
popwin = window.open(mypage,myname,settings)
}
function autoReload() {
	document.ReloadFrm.submit()
}
function SetLastDate() {
	document.LastDateFrm.submit()
}
function openWindow3(url) {
  popupWin = window.open(url,'pop_col','width=400,height=450,scrollbars=yes')
}

function shoGlow() {
var newGlow = document.forms.Form1.strGlowColor.value;
document.all['glowname'].style.filter='glow(color:'+newGlow+',strength:4); width:100%';
}
function shoText() {
var newGlow = document.forms.Form1.strTxColor.value;
document.all['glowname'].color=newGlow;
}
<% If strAllowHtml <> 1 Then %>
function openWindowPM(url) {
  popupWin = window.open(url,'pm_pop_send','resizable,width=590,height=510,top=35,left=120,scrollbars=yes');
}
<% Else %>
function openWindowPM(url) {
  popupWin = window.open(url,'pm_pop_send','resizable,width=635,height=550,top=30,left=120,scrollbars=yes');
}
<% End If %>
// -->
</script>
<%
spThemeHeader_style()
getCSSfile()
%>
<!--#include file="includes/inc_editor.asp"-->
</head>
<% if thispage = "monitor" then
  spThemeShortBodyTag = ""
end if %>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%=spThemeShortBodyTag%>>
<table class="spTheme" width="100%" height="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td align="center" valign="top">
    <div><center>