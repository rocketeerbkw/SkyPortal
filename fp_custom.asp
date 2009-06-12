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
<!--include file="modules/forums/fp_forums.asp" -->
<!--include file="modules/downloads/fp_dl.asp" -->
<!--include file="modules/articles/fp_articles.asp" -->
<!--include file="modules/pictures/fp_pic.asp"-->
<!--include file="modules/links/fp_links.asp"-->
<!--include file="modules/classifieds/fp_classified.asp"-->
<% 
  PTcnt = 0
function cntPendTsks()
	sSql = "SELECT * FROM " & strTablePrefix & "MODS WHERE M_CODE='pndTskCnt'"
	set rsP = my_Conn.execute(sSql)
	if not rsP.eof then
	  do until rsP.eof
	    execute("Call " & rsP("M_VALUE"))
	    rsP.movenext
	  loop
	end if
	set rsP = nothing
  
  if strEmailVal = 7 then
  ' Pending MEMBERS count
    PTcnt = PTcnt + getCount("M_NAME",strTablePrefix & "MEMBERS_PENDING","M_LEVEL = -1")
  end if
  if strEmailVal = 8 then
  ' Pending MEMBERS count
    PTcnt = PTcnt + getCount("M_NAME",strTablePrefix & "MEMBERS_PENDING","M_LEVEL = -2")
  end if
  cntPendTsks = "&nbsp;(" & PTcnt & ")"
end function

sub closeObjects()
 set mnu = nothing
 set oRss = nothing
 set oSFS = nothing
 set oSpData = nothing
end sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site projects box
' :::::::::::::::::::::::::::::::::::::::::::::::
function projects_fp()
spThemeMM = "prjct"
spThemeTitle= txtProjStat
spThemeBlock1_open(intSkin)%>
<table border="0" width="100%"><tr><td width="100%" class="tCellAlt1">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td><b><%= strSiteTitle %></b></td></tr>
<tr><td align="left"><table width="100%" border="1" cellspacing="0" cellpadding="0" class="tBorder"><tr><td width="95%" bgcolor="#CCCCCC"><img src="images/icons/bar.gif" width="87%" height="15" alt="" /></td><td bgcolor="whitesmoke"><span class="fSmall">87%</span></td></tr></table></td></tr>
<tr><td><b><%= txtHapBDay %></b></td></tr>
<tr><td align="left"><table width="100%" border="1" cellspacing="0" cellpadding="0" class="tBorder"><tr><td width="95%" bgcolor="#CCCCCC"><img src="images/icons/bar.gif" width="35%" height="15" alt="" /></td><td bgcolor="whitesmoke"><span class="fSmall">35%</span></td></tr></table></td></tr>
</table>
</td></tr></table>
<%spThemeBlock1_close(intSkin)
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		OTHER LINKS box
' :::::::::::::::::::::::::::::::::::::::::::::::
function others_fp()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_donate','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end function

function m_aspin()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_rateus','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end function

Sub writeFlash2(swfImg,bID,bName) %>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="88" height="31" id="abImage" align=""><param name=movie value="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bName) %>"><param name=quality value=high><embed src="<%= swfImg %>?clickTAG=<%= strHomeUrl %>banner_link.asp?id=<%= bID %>&ctTarget=_blank&txtStr=<%= server.urlencode(bName) %>" quality="high" name="abImage" height="31" width="88" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>
<% 
end sub

sub modFeatures()
    spThemeBlock1_open(intSkin) %>
	<p>All the themeblocks that you see to the left, right, on top (this block) and on bottom of the MAIN page block are controlled by a single <b>'Modules/<%= CurPageTitle %>/<%= CurPageTitle %>_custom.asp'</b> file. You can show any themeblock that is available on the homepage here as well. Or you can create your own function. This way, people can change the layout to what they like, and also keeping their layout in a file that will not be included with future upgrades.<br /><br />if you want to delete this text, delete  <b>modFeatures()</b> from <b>'Modules/<%= CurPageTitle %>/<%= CurPageTitle %>_custom.asp'</b>. The <b><i>text</i></b> for this block is found in a function named <b>modFeatures()</b> and is located in <b>fp_custom.asp</b> at the very bottom of the file.</p>
<%  spThemeBlock1_close(intSkin)
end sub

' insert new functions and subs above this line

%>
