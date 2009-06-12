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

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site announcements
' :::::::::::::::::::::::::::::::::::::::::::::::
function announce_fp()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_announce_fp','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site welcome message
' :::::::::::::::::::::::::::::::::::::::::::::::
function welcome_fp()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_welcome_fp','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site skin changer box
' :::::::::::::::::::::::::::::::::::::::::::::::
function theme_changer()
  spThemeMM = "sknchgr"
  spThemeTitle = txtSknChgr
  spThemeBlock1_open(intSkin)
  %>
    <div class="spThemeChanger">
    <form name="themechanger" method="post" action="">
    <span class="fNorm">Select Skin:</span><br />
      <select name="thm" onChange="submit();">
        <% 
	ssSQL = "select C_STRAUTHOR, C_TEMPLATE, C_STRFOLDER, C_SKINLEVEL from portal_colors ORDER BY C_TEMPLATE"
	set rsThm = my_Conn.execute(ssSQL)
	if rsThm.eof then
		'strAuth = "anonymous"
	else
	  do until rsThm.eof
	    if hasAccess(rsThm("C_SKINLEVEL")) then
		  if rsThm("C_STRFOLDER") = strTheme then
	        Response.Write("<option value="""& rsThm("C_STRFOLDER") &""" selected=""selected"">"& rsThm("C_TEMPLATE") &"</option>")
		    strAuth = rsThm("C_STRAUTHOR")
		  else
	        Response.Write("<option value="""& rsThm("C_STRFOLDER") &""">"& rsThm("C_TEMPLATE") &"</option>")
		  end if
		end if
	    rsThm.movenext
	  loop
	end if
	set rsThm = nothing
	%>
      </select>
  </form>
    <span class="fSmall"><br /><%= txtAuthor %>:<b> <%= strAuth %> </b></span>
    </div>
  <% spThemeFooter = ""
  spThemeBlock1_close(intSkin)
end function 

sub affiliateBanners()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_affiliateBanners','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end sub

sub login_box()
  tDiv = "ajxBlkDiv" & randomNum(99999)
  tsA = "<div id=""" & tDiv & """>" & vbcrlf
  tsA = tsA & "<script language=""JavaScript"" type=""text/JavaScript"">" & vbcrlf
  tsA = tsA & "ajax_UpdateBlock('skyportal_ajax.asp','"& tDiv &"','sajx_login_box','','','','');"& vbcrlf
  tsA = tsA & "</script></div>" & vbcrlf
  Response.Write tsA
end sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		site searchbox
' :::::::::::::::::::::::::::::::::::::::::::::::
function search_fp()
%>
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
function RefreshS() {
if (document.SearchForm.news.checked) {
	window.location ="forum_search.asp?mode=news";
} else {
	window.location ="search.asp";
}
}
function checklength() {
if (document.srcform1.search.value.length < 3) {
alert('<%= txtSrchLen %>');
return false;
}
}
// done hiding -->
</script>

<%
spThemeMM = "ssrch"
spThemeTitle= txtSearch
'spThemeTitle = spThemeTitle & " [" & intSkin & "]"
spThemeBlock1_open(intSkin)
  'spThemeTitle= txtSrchFor & ":"
  'spThemeBlock3_open() %>
<form name="srcform1" action="site_search.asp" method="post" id="srcform1" onsubmit="return checklength()">
	<input type="text" name="search" size="15" style="margin-top:5px;" value="<%=search%>" /><br />
      <input type="submit" value=" <%= txtSearch %> " id="searchA" name="searchA" class="button" /><br />
</form>
<%
  'spThemeBlock3_close()
spThemeBlock1_close(intSkin)
end function

%>
