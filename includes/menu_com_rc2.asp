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
%> 
<script language="JavaScript" type="text/javascript">
<!--
	//var myMaxWin;
function mymax() {
	myMaxWin = window.open('pop_myMax.asp','myMax','left=0,top=0,width=790,height=540,resizable=1,scrollbars=yes');
}
-->
</script>
<!-- Menu bar #1. -->
<%
Response.Write("<div class=""menuBar"" style=""width:100%;"">")
Response.Write("<a href=""default.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu1');"" >" & txtHome & "</a>")
if hasAccess(2) then
Response.Write("<a href=""members.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu2');"">" & txtMembers & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("forums","USERS") then
Response.Write("<a href=""fhome.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu3');"">" & txtForum & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("events","USERS") then
Response.Write("<a href=""events.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu4');"">" & txtEvents & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("article","USERS") then
Response.Write("<a href=""article.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu5');"">" & txtArticles & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("downloads","USERS") then
Response.Write("<a href=""dl.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu6');"">" & txtDownloads & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("links","USERS") then
Response.Write("<a href=""links.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu7');"">" & txtLinks & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("pictures","USERS") then
Response.Write("<a href=""pic.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu8');"">" & txtPics & "</a>")
end if
if (hasAccess("2,3") or (not hasAccess(2) and strLockDown = 0)) and chkApp("classifieds","USERS") then
Response.Write("<a href=""classified.asp"" class=""menuButton"" onmouseover=""buttonmouseover(event, 'menu9');"">" & txtClassifieds & "</a>")
end if
Response.Write("</div>" & vbcrlf)
%>
<!-- End Main menus. -->
		
<!-- Sub menu Home. --><%
Response.Write("<div id=""menu1"" class=""nav_menu"" onmouseover=""menuMouseover(event)"">")
if (not hasAccess(2) and strNewReg = 1) or hasAccess(1) then
  if hasAccess(1) and strLockDown = 1 then
    Response.Write("<a href=""register.asp?mode=Register"" class=""nav_menuItem"">" & txtRegister & "</a>")
  else
    Response.Write("<a href=""policy.asp"" class=""nav_menuItem"">" & txtRegister & "</a>")
  end if
Response.Write("<div class=""menuItemSep""></div>")
end if
'Response.Write("<a class=""nav_menuItem"" href=""features.asp"">" & txtFeatures & "</a>")
'Response.Write("<a class=""nav_menuItem"" href=""includes/faq/faq.asp"">" & txtFAQ & "</a>")
Response.Write("<a href=""javascript:openWindowPM('pm_pop.asp');"" class=""nav_menuItem"">" & txtContactUs & "</a>")
Response.Write("<div class=""menuItemSep""></div>")
Response.Write("<a class=""nav_menuItem"" href=""#"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu1_4');""><span class=""menuItemText"">Example</span><span class=""menuItemArrow"">&#9654;</span></a>")
Response.Write("</div>" & vbcrlf)
%>
		<!-- Sub menu Members. --><%
if hasAccess(2) then
  Response.Write("<div id=""menu2"" class=""nav_menu"" onmouseover=""menuMouseover(event)"">")
  if intMyMax or intIsSuperAdmin then
  Response.Write("<a href=""javascript:;"" onclick=""mymax();"" class=""nav_menuItem"">" & txtMyMax & "</a>")
  end if
   Response.Write("<a href=""cp_main.asp"" class=""nav_menuItem"">" & txtCtrlPnl & "</a>")
  Response.Write("<a class=""nav_menuItem"" href=""javascript:;"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu2_4');""><span class=""menuItemText"">" & txtMyProfile & "</span><span class=""menuItemArrow"">&#9654;</span></a>")
  if chkApp("PM","USERS") then
    Response.Write("<a class=""nav_menuItem"" href=""javascript:;"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu2_3');""><span class=""menuItemText"">" & txtPvtMsgs & "</span><span class=""menuItemArrow"">&#9654;</span></a>")
  end if
  if intBookmarks then
   Response.Write("<a href=""cp_main.asp?cmd=7"" class=""nav_menuItem"">" & txtBookmks & "</a>")
  end if
  if intSubscriptions then
   Response.Write("<a href=""cp_main.asp?cmd=6"" class=""nav_menuItem"">" & txtSubsc & "</a>")
  end if
  Response.Write("<a href=""members.asp"" class=""nav_menuItem"">" & txtMbrLst & "</a>")
  Response.Write("<a href=""active_users.asp"" class=""nav_menuItem"">" & txtActvUsrs & "</a>")
  if hasAccess(1) then
  	Response.Write("<a href=""site_monitor.asp"" class=""nav_menuItem"" target=""_search"">" & txtMxMon & "</a>")
  end if
  if hasAccess(1) and chkApp("forums","USERS") then
    Response.Write("<div class=""menuItemSep""></div>")
    Response.Write("<a href=""forum_report_post_moderate.asp"" class=""nav_menuItem"">" & txtRptdPst & "</a>")
  end if
  if hasAccess(1) then
    Response.Write("<a href=""admin_home.asp"" class=""nav_menuItem"">" & txtAdminOpts & "</a>")
  end if
  Response.Write("</div>" & vbcrlf)
end if

%><!-- Sub menu Forums. --><%
if chkApp("forums","USERS") then
Response.Write("<div id=""menu3"" class=""nav_menu"">")
Response.Write("<a href=""fhome.asp"" class=""nav_menuItem"">" & txtFrmHome & "</a>")
Response.Write("<a href=""forum_active_topics.asp"" class=""nav_menuItem"">" & txtActvTopics & "</a>")
Response.Write("<a href=""forum_search.asp"" class=""nav_menuItem"">" & txtSrchFrms & "</a>")
Response.Write("<div class=""menuItemSep""></div>")
Response.Write("<a href=""forum_faq.asp?page=forums"" class=""nav_menuItem"">" & txtFrmFAQ & "</a>")
Response.Write("</div>" & vbcrlf)
end if

%><!-- Sub menu Events. --><%
if chkApp("events","USERS") then
  Response.Write("<div id=""menu4"" class=""nav_menu"" onmouseover=""menuMouseover(event)"">")
  Response.Write("<a href=""events.asp"" class=""nav_menuItem"">" & txtCalendar & "</a>")
  Response.Write("<a href=""events.asp?mode=newEvents"" class=""nav_menuItem"">" & txtNewEvnts & "</a>")
  Response.Write("<a href=""events.asp?mode=eventList"" class=""nav_menuItem"">" & txtUpcomEvnts & "</a>")
  if hasAccess(2) then
  	Response.Write("<a class=""nav_menuItem"" href=""javascript:;"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu3_4');""><span class=""menuItemText"">" & txtSubEvnt & "</span><span class=""menuItemArrow"">&#9654;</span></a>")
  	Response.Write("<a href=""events.asp?mode=eventSubscription"" class=""nav_menuItem"">" & txtSubscribe & "</a>")
  	Response.Write("<a href=""events.asp?mode=eventReminder"" class=""nav_menuItem"">" & txtReminders & "</a>")
  end if
  Response.Write("<a href=""events.asp?mode=search"" class=""nav_menuItem"">" & txtSrchEvnts & "</a>")
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('modules/events/faq_events.asp');"" class=""nav_menuItem"">" & txtEvntsFAQ & "</a>")
  Response.Write("</div>" & vbcrlf)
end if

%><!-- Sub menus Articles. --><%
if chkApp("article","USERS") then
  Response.Write("<div id=""menu5"" class=""nav_menu"">")
  Response.Write("<a href=""article.asp"" class=""nav_menuItem"">" & txtMainDir & "</a>")
  Response.Write("<a href=""article.asp?cmd=3"" class=""nav_menuItem"">" & txtNewArts & "</a>")
  Response.Write("<a href=""article.asp?cmd=4"" class=""nav_menuItem"">" & txtPopArts & "</a>")
  Response.Write("<a href=""article.asp?cmd=5"" class=""nav_menuItem"">" & txtTopArts & "</a>")
  if hasAccess(2) then
	Response.Write("<a href=""article.asp?cmd=7"" class=""nav_menuItem"">" & txtSubArt & "</a>")
  end if
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('article_pop.asp?mode=10');"" class=""nav_menuItem"">" & txtArtFAQ & "</a>")
  Response.Write("</div>" & vbcrlf)
end if

%><!-- Sub menus Downloads. --><%
if chkApp("downloads","USERS") then
  Response.Write("<div id=""menu6"" class=""nav_menu"">")
  Response.Write("<a href=""dl.asp"" class=""nav_menuItem"">" & txtMainDir & "</a>")
  Response.Write("<a href=""dl.asp?cmd=3"" class=""nav_menuItem"">" & txtNewDL & "</a>")
  Response.Write("<a href=""dl.asp?cmd=4"" class=""nav_menuItem"">" & txtPopDL & "</a>")
  Response.Write("<a href=""dl.asp?cmd=5"" class=""nav_menuItem"">" & txtTopDL & "</a>")
  if hasAccess(2) then
    Response.Write("<a class=""nav_menuItem"" href=""dl_add_form.asp"">" & txtSubDL & "</a>")
  end if
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('dl_pop.asp?mode=12');"" class=""nav_menuItem"">" & txtDLFAQ & "</a>")
  Response.Write("</div>" & vbcrlf)
end if

%><!-- Sub menus Links. --><%
if chkApp("links","USERS") then
  Response.Write("<div id=""menu7"" class=""nav_menu"">")
  Response.Write("<a href=""links.asp"" class=""nav_menuItem"">" & txtMainDir & "</a>")
  Response.Write("<a href=""links.asp?cmd=3"" class=""nav_menuItem"">" & txtNewLinks & "</a>")
  Response.Write("<a href=""links.asp?cmd=4"" class=""nav_menuItem"">" & txtPopLinks & "</a>")
  Response.Write("<a href=""links.asp?cmd=5"" class=""nav_menuItem"">" & txtTopLinks & "</a>")
  Response.Write("<a href=""javascript:;"" onclick=""window.open('links_pop.asp?mode=4&amp;cid=0');"" class=""nav_menuItem"">" & txtRndmLink & "</a>")
  if hasAccess(2) then
    Response.Write("<a href=""links.asp?cmd=8"" class=""nav_menuItem"">" & txtSubLink & "</a>")
  end if
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('links_pop.asp?mode=12');"" class=""nav_menuItem"">" & txtLinkFAQ & "</a>")
  Response.Write("</div>")
end if

%><!-- Sub menus Photos. --><%
if chkApp("pictures","USERS") then
  Response.Write("<div id=""menu8"" class=""nav_menu"">")
  Response.Write("<a href=""pic.asp"" class=""nav_menuItem"">" & txtMainDir & "</a>")
  Response.Write("<a href=""pic.asp?cmd=3"" class=""nav_menuItem"">" & txtNewPics & "</a>")
  Response.Write("<a href=""pic.asp?cmd=4"" class=""nav_menuItem"">" & txtPopPics & "</a>")
  Response.Write("<a href=""pic.asp?cmd=5"" class=""nav_menuItem"">" & txtTopPics & "</a>")
  if hasAccess(2) then
    Response.Write("<a href=""pic.asp?cmd=8"" class=""nav_menuItem"">" & txtSubPic & "</a>")
  end if
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('pic_pop.asp?mode=13');"" class=""nav_menuItem"">" & txtPicsFAQ & "</a>")
  Response.Write("</div>")
end if

%><!-- Sub menus Classifieds. --><%
if chkApp("classifieds","USERS") then
  Response.Write("<div id=""menu9"" class=""nav_menu"">")
  Response.Write("<a href=""Classified.asp"" class=""nav_menuItem"">" & txtMainDir & "</a>")
  Response.Write("<a href=""Classified.asp?cmd=3"" class=""nav_menuItem"">" & txtNewClass & "</a>")
  if hasAccess(2) then
    Response.Write("<a href=""Classified.asp?cmd=4"" class=""nav_menuItem"">" & txtSubClass & "</a>")
  end if
  Response.Write("<div class=""menuItemSep""></div>")
  Response.Write("<a href=""javascript:openWindow3('classified_pop.asp?mode=10');"" class=""nav_menuItem"">" & txtClassFAQ & "</a>")
  Response.Write("</div>")
end if

%><!-- Sub menus for Members. --><%
if hasAccess(2) then
Response.Write("<div id=""menu2_4"" class=""nav_menu"">")
Response.Write("<a href=""cp_main.asp?cmd=8&member=" & strUserMemberID & """ class=""nav_menuItem"">" & txtViewProf & "</a>")
Response.Write("<a href=""cp_main.asp?cmd=9"" class=""nav_menuItem"">" & txtEditProf & "</a>")
Response.Write("<a href=""cp_main.asp?cmd=1&mode=AvatarEdit"" class=""nav_menuItem"">" & txtEditAvatar & "</a>")
Response.Write("</div>")

Response.Write("<div id=""menu2_3"" class=""nav_menu"">")
Response.Write("<a href=""pm.asp"" class=""nav_menuItem"">" & txtPMinbox & "</a>")
'Response.Write("<a class=""nav_menuItem"" href=""pm.asp?cmd=1"">" & txtPMoutBx & "</a>")
Response.Write("<a href=""pm.asp?cmd=2"" class=""nav_menuItem"">" & txtPMcompose & "</a>")
Response.Write("</div>")
end if

%><!-- Sub menus for Events. --><%
if hasAccess(2) and chkApp("events","USERS") then
Response.Write("<div id=""menu3_4"" class=""nav_menu"">")
Response.Write("<a href=""events.asp?mode=add&date=" & strCurDate & """ class=""nav_menuItem"">" & txtSglEvnt & "</a>")
Response.Write("<a href=""events.asp?mode=addseries&date=" & strCurDate & """ class=""nav_menuItem"">" & txtRecurEvnt & "</a>")
Response.Write("</div>")
end if

%><!-- Sub menus for Home menu. --><%
Response.Write("<div id=""menu1_4"" class=""nav_menu"" onmouseover=""menuMouseover(event)"">")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4 Item 1</a>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4 Item 2</a>")
Response.Write("<a class=""nav_menuItem"" href=""#"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu1_4_3');""><span class=""menuItemText"">Menu 1-4 Item 3</span><span class=""menuItemArrow"">&#9654;</span></a>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4 Item 4</a>")
Response.Write("</div>")

Response.Write("<div id=""menu1_4_3"" class=""nav_menu"" onmouseover=""menuMouseover(event)"">")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3 Item 1</a>")
Response.Write("<div class=""menuItemSep""></div>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3 Item 2</a>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3 Item 3</a>")
Response.Write("<div class=""menuItemSep""></div>")
Response.Write("<a class=""nav_menuItem"" href=""#"" onclick=""return false;"" onmouseover=""menuItemMouseover(event, 'menu1_4_3_4');""><span class=""menuItemText"">Menu 1-4-3 Item 4</span><span class=""menuItemArrow"">&#9654;</span></a>")
Response.Write("</div>")

Response.Write("<div id=""menu1_4_3_4"" class=""nav_menu"">")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3-4 Item 1</a>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3-4 Item 2</a>")
Response.Write("<a class=""nav_menuItem"" href=""#"">Menu 1-4-3-4 Item 3</a>")
Response.Write("</div>")
%>