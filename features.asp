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

CurPageType = "core"
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = "Features"
	PageAction = "Viewing<br />" 
	PageLocation = "features.asp"
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<table border="0" width="100%" align="center" cellpadding="0" cellspacing="0">
<tr>
<td width="180" valign="top" class="leftPgCol" nowrap>
<input id="xtest" name="xtest" class="button" type="button" value="test" onclick="openDialog2('n','dogin','New Window','250','300');">
<% 
'openDialog2(n,ob,t,w,h)
intSkin = getSkin(intSubSkin,1)
others_fp2()
menu_fp()
'showLoginBlock("")
'Call showPasswordBlock2(1,"Login","",0,0,1)
 %>
</td>
<td valign="top" class="mainPgCol" align="left">
<%
intSkin = getSkin(intSubSkin,2)
	'strCurDateString = DateToStr(DateAdd("h", strTimeAdjust , Now()))
	'strCurDateAdjust = strToDate(strCurDateString)
	'strCurDate = ChkDate(strCurDateString)
	
	'response.Write("Session.LCID: " & Session.LCID & "<br />")
	'response.Write("intMemberLCID: " & intMemberLCID & "<br />")
	'response.Write("intPortalLCID: " & intPortalLCID & "<br />")
	'response.Write("strTimeAdjust: " & strTimeAdjust & "<br /><br />")
	'response.Write("strCurDateString: " & strCurDateString & "<br />")
	'response.Write("strCurDateAdjust: " & strCurDateAdjust & "<br />")
	'response.Write("strCurDate: " & strCurDate & "<br />")
	
	'current_date = "20060218000000"  'post date
	'past_date = "20060213000000"  'end date
	'future_date = "20060227000000"  'end date
	'response.Write("<br />Date difference: " & DateDiff("d", chkDate2(current_date), chkDate2(past_date)))
	'response.Write("<br />Date difference: " & DateDiff("d", chkDate2(current_date), strCurDateAdjust))
	'getDateDiff(datetostr(strCurDate),datetostr(strPostDate))
	'response.Write("getDateDiff: " & getDateDiff(current_date, past_date) & " returns: positive")
	'response.Write("getDateDiff: " & getDateDiff(current_date, future_date) & " returns: negative")
	'response.Write("<br />getDateDiff: " & getDateDiff(future_date, current_date) & " returns: positive")
	'response.Write("<br />getDateDiff: " & getDateDiff(past_date, current_date) & " returns: negative")
	'response.Write("<br />getDateDiff: " & getDateDiff(future_date, past_date) & " returns: positive")
	'response.Write("<br />getDateDiff: " & getDateDiff(past_date, future_date) & " returns: negative")


  'response.Write("chkApp: " & chkApp("pictures","USERS") & "<br />")
'call setAppPerms("pictures","INAME")
'call setAppPerms("2","id")
  'response.Write("sAppRead: " & sAppRead & "<br />")
  'response.Write("sAppWrite: " & sAppWrite & "<br />")
  'response.Write("sAppFull: " & sAppFull & "<br /><br />")
  
  'response.Write("read:" & hasAccess(sAppRead) & "<br />")
  'response.Write("read:" & hasAccess(getAppPerms("pictures","read","id")) & "<br />")
  'response.Write("write:" & getAppPerms("PM","write","") & "<br />")
  'response.Write("full:" & getAppPerms("PM","full","") & "<br />")
%>
 <!-- Begin BidVertiser code -->
<SCRIPT LANGUAGE="JavaScript1.1" SRC="http://bdv.bidvertiser.com/BidVertiser.dbm?pid=7682&bid=76642"></SCRIPT>
<noscript><a href="http://www.bidvertiser.com">pay per click</a></noscript>
<!-- End BidVertiser code --> 
<%
spThemeTitle= "Features"
spThemeBlock1_open(intSkin)
'spThemeBlock2_open()
%><table class="tPlain">
<tr><td class="tCellAlt1" align="center">
                  
      <TABLE cellSpacing=0 cellPadding=10>
        <TBODY>
          <TR> 
            <TD vAlign=top height=15><A id=forum 
                        href="fhome.asp"><IMG 
                        src="Themes/<%= strTheme %>/forum.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=forum><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Community Forums</B><br />
              <FONT size=2>A place where you can ask questions, gain knowledge and communicate with people</FONT></FONT></LABEL>
            </TD>
            <TD vAlign=top height=15><A id=article 
                        href="article.asp"><IMG 
                        src="Themes/<%= strTheme %>/article.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=article><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Articles Manager</B><br />
              <FONT size=2>Post and share your articles here</FONT></FONT></LABEL>
            </TD>
          </TR>
          <TR> 
            <TD vAlign=top height=15><A id=download 
                        href="dl.asp"><IMG 
                        src="Themes/<%= strTheme %>/download.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=download><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Download Manager</B><br />
              <FONT size=2>A collection of popular downloads</FONT></FONT></LABEL>
            </TD>
            <TD vAlign=top height=15><A id=event 
                        href="events.asp"><IMG 
                        src="Themes/<%= strTheme %>/event.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=event><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Event Calendar</B><br />
              <FONT size=2>Add events such as birthdays, meetings. Displays upcoming events conveniently 
              on the navbar</FONT></FONT></LABEL>
            </TD>
          </TR>
          <TR> 
            <TD vAlign=top height=15><A id=link 
                        href="links.asp"><IMG 
                        src="Themes/<%= strTheme %>/link.gif" 
                      border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=link><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Links Manager</B><br />
              <FONT size=2>Make URL's available for your friends, employees or partners with our links 
              manager</FONT></FONT></LABEL>
            </TD>
            <TD vAlign=top height=15><A id=pic 
                        href="pic.asp"><IMG 
                        src="Themes/<%= strTheme %>/pics.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=pic><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Pictures Manager</B><br />
              <FONT size=2>View photo albums by most popular, top rated or newest</FONT></FONT></LABEL>
            </TD>
          </TR>
          <TR> 
            <TD vAlign=top height=15><A id=classified 
                        href="classified.asp"><IMG 
                        src="Themes/<%= strTheme %>/classified.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=classified><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Classifieds Manager</B><br />
              <FONT size=2>Post advertisements in our online classifieds manager</FONT></FONT></LABEL>
            </TD>
            <TD vAlign=top height=15><A id=news 
                        href="fnews.asp"><IMG 
                        src="Themes/<%= strTheme %>/news.gif" 
                      border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=news><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>News Archive</B><br />
              <FONT size=2>Read previous news topics from SkyPortal front page</FONT></FONT></LABEL>
            </TD>
          </TR>
          <TR> 
            <TD vAlign=top height=15><A id=support 
                        href="fhome.asp"><IMG 
                        src="Themes/<%= strTheme %>/support.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL 
                        for=support><FONT face="Verdana, Arial, Helvetica" 
                        size=4><B>Support Forum</B><br />
              <FONT size=2>Discuss technical questions related to SkyPortal</FONT></FONT></LABEL>
            </TD>
            <TD vAlign=top height=15><A id=cp 
                        href="cp_main.asp"><IMG 
                        src="Themes/<%= strTheme %>/ctp.gif" 
                        border=0></A></TD>
            <TD vAlign=top width="50%" height=15> 
              <LABEL for=cp><FONT 
                        face="Verdana, Arial, Helvetica" size=4><B>Control Panel</B><br />
              <FONT size=2>Modify your personal preferences</FONT></FONT></LABEL>
            </TD>
          </TR>
        </TBODY>
      </TABLE>
</td></tr></table>
<%
'spThemeBlock2_close()
spThemeBlock1_close(intSkin)%>

</td>
</tr>
</table>
<!--#include file="inc_footer.asp" -->

<%
function others_fp2()
'spThemeMM = "othrs"
spThemeTitle= "Support SkyPortal"
'spThemeBlock1_open(intSkin)
spThemeBlock2_open()%>
  <p>Please help support the continued development of SkyPortal by making your donation today.</p>
  <p><a href="http://www.skyportal.net/site_donation.asp"><img src="http://www.skyportal.net/images/donation_sp.gif" border="0" alt="" title="Help support the SkyPortal Development" /></a></p>
<%'spThemeBlock1_close(intSkin)
spThemeBlock2_close()
end function

%>
