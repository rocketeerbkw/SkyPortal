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
CurPageType = "home"
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<script type="text/javascript" src="includes/scripts/menu_com.js"></script>
<table border="0" width="100%" align="center">
<tr>
<td valign="top" width="100%">
<%

'If myMax is turned OFF and they are not a SuperAdmin
' then stop the page from loading
if intMyMax = 0 and intIsSuperAdmin = 0 then
    closeandgo("stop")
end if

' If they are not a member and logged in
' then stop the page from loading
membID = 0
up_sticky = false
if hasAccess(2) then
  if intMyMax = 1 then
    membID = getmemberid(strdbntusername)
	disabl = " disabled=""disabled"""
  else
	up_sticky = true
	disabl = ""
  end if
else
  closeandgo("stop")
end if

if request("mode") = "reset" and membID > 0 then
  mmSQL = "DELETE FROM PORTAL_FP_USERS where fp_uid = " & membID
  executeThis(mmSQL)
  response.Write("<script type=""text/javascript""> opener.document.location.reload();</script>")
  response.Write("<center><b>" & txtConfigUpd & "!</b></center>")
end if

if request.Form("mode") = "update" and trim(strDBNTuserName) <> "" then
  left_col = chkString(request.Form("left_select"),"sqlstring")
  main_col = chkString(request.Form("main_select"),"sqlstring")
  right_col = chkString(request.Form("right_select"),"sqlstring")
  sSQL = "UPDATE PORTAL_FP_USERS SET "
  sSQL = sSQL & "fp_leftcol = '" & left_col & "'"
  sSQL = sSQL & ",fp_maincol = '" & main_col & "'"
  sSQL = sSQL & ",fp_rightcol = '" & right_col & "'"
  if up_sticky and membID = 0 then
    left_sticky = chkString(request.Form("left_sticky"),"sqlstring")
    main_sticky = chkString(request.Form("main_sticky"),"sqlstring")
    right_sticky = chkString(request.Form("right_sticky"),"sqlstring")
    sSQL = sSQL & ",fp_leftsticky = '" & left_sticky & "'"
    sSQL = sSQL & ",fp_mainsticky = '" & main_sticky & "'"
    sSQL = sSQL & ",fp_rightsticky = '" & right_sticky & "'"
  end if
  sSQL = sSQL & " WHERE fp_uid = " & membID
  executeThis(sSQL)
  response.Write("<script type=""text/javascript""> opener.document.location.reload();</script>")
  response.Write("<center><b>" & txtConfigUpd & "!</b></center>")
  
elseif request.Form("mode") = "insert" and trim(strDBNTuserName) <> "" then
  left_col = chkString(request.Form("left_select"),"sqlstring")
  main_col = chkString(request.Form("main_select"),"sqlstring")
  right_col = chkString(request.Form("right_select"),"sqlstring")
  sSQL = "INSERT INTO PORTAL_FP_USERS ("
  sSQL = sSQL & "fp_leftcol,fp_maincol,fp_rightcol,fp_uid"
  sSQL = sSQL & ")VALUES("
  sSQL = sSQL & "'" & left_col & "','" & main_col & "','" & right_col & "'," & membID & ");"
  executeThis(sSQL)
  response.Write("<script type=""text/javascript""> opener.document.location.reload();</script>")
  response.Write("<center><b>" & txtConfigUpd & "!</b></center>")
  'response.Write("main_col: " & main_col & "<br />")
  'response.Write("right_col: " & right_col & "<br />")
end if

b_desc = ""
l_options = ""
m_options = ""
r_options = ""
l_select = ""
m_select = ""
r_select = ""

' populate the options for the "add" select box
mmSQL = "select * from PORTAL_FP where FP_ACTIVE = 1 order by FP_NAME, FP_FUNCTION"
set rsMM = my_Conn.execute(mmSQL)

if not rsMM.eof then
  do until rsMM.eof
    b_desc = b_desc & "block_descr['" & rsMM("FP_NAME") & ":" & rsMM("FP_FUNCTION") & "'] = '" & rsMM("FP_DESC") & "';" & vbcrlf
	tmpStr = "<option value=""" & rsMM("FP_NAME") & ":" & rsMM("FP_FUNCTION") & """>" & rsMM("FP_NAME") & "</option>" & vbcrlf
	select case rsMM("FP_COLUMN")
	  case 1
	    l_options = l_options & tmpStr
	  case 2
	    m_options = m_options & tmpStr
	  case 3
	    r_options = r_options & tmpStr
	  case 4
	    l_options = l_options & tmpStr
	    r_options = r_options & tmpStr
	end select
    rsMM.movenext
  loop
end if
set rsMM = nothing

' get the data from the db
' populate the select boxes with the users config
' If mymax is turned OFF for all members and you are a super admin,
' you will be setting the default layout.

':: default layout data ::
  l_sticky = ""
  m_sticky = ""
  r_sticky = ""
  mmSQL = "select * from PORTAL_FP_USERS where fp_uid = 0"
  set rsMM = my_Conn.execute(mmSQL)
  if rsMM("fp_leftsticky") <> "" then
    l_stick = split(rsMM("fp_leftsticky"),",")
    for fp = 0 to ubound(l_stick)
	  l_sticky = l_sticky & "<option value=""" & l_stick(fp) & """>" & split(l_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  if rsMM("fp_mainsticky") <> "" then
    m_stick = split(rsMM("fp_mainsticky"),",")
    for fp = 0 to ubound(m_stick)
	  m_sticky = m_sticky & "<option value=""" & m_stick(fp) & """>" & split(m_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  if rsMM("fp_rightsticky") <> "" then
    r_stick = split(rsMM("fp_rightsticky"),",")
    for fp = 0 to ubound(r_stick)
	  r_sticky = r_sticky & "<option value=""" & r_stick(fp) & """>" & split(r_stick(fp),":")(0) & "</option>" & vbcrlf
    next
  end if
  left_col = rsMM("fp_leftcol")
  main_col = rsMM("fp_maincol")
  right_col = rsMM("fp_rightcol")
  set rsMM = nothing

  rmod = "update"
  if intMyMax = 1 and membID > 0  then
    mmSQL = "select * from PORTAL_FP_USERS where fp_uid = " & membID
    set rsMM = my_Conn.execute(mmSQL)
    if rsMM.eof then
      rmod = "insert"
    else
      left_col = rsMM("fp_leftcol")
      main_col = rsMM("fp_maincol")
      right_col = rsMM("fp_rightcol")
    end if
    set rsMM = nothing
  end if
  
  l_col = split(left_col,",")
  for fp = 0 to ubound(l_col)
	l_select = l_select & "<option value=""" & l_col(fp) & """>" & split(l_col(fp),":")(0) & "</option>" & vbcrlf
  next
  m_col = split(main_col,",")
  for fp = 0 to ubound(m_col)
	m_select = m_select & "<option value=""" & m_col(fp) & """>" & split(m_col(fp),":")(0) & "</option>" & vbcrlf
  next
  r_col = split(right_col,",")
  for fp = 0 to ubound(r_col)
	r_select = r_select & "<option value=""" & r_col(fp) & """>" & split(r_col(fp),":")(0) & "</option>" & vbcrlf
  next
%>
<script type="text/javascript">
var block_descr = new Array();
<%= b_desc %>
</script>
<%

spThemeTitle= txtMMhp
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">") %>
<tr><td width="100%">
<p><b><%= txtMMCHP %></b><br /><%= txtPMM1a %>.<br /><br /><%= txtPMM1b %>.<br /><br /><%= txtPMM1c %>.</p>
<form method="post" action="pop_mymax.asp" onsubmit="return select_options();">
<input type="hidden" name="cmd" value="" />
<input type="hidden" name="mode" value="<%= rmod %>" />
<input type="hidden" name="name" value="" />
<table border="1" width="750" align="center">
<tr class="tTitle"><td valign="top" align="center" width="30%">
<b><%= txtLftColBlk %></b></td>
<td valign="top" align="center" width="30%">
<b><%= txtMainColBlk %></b></td>
<td valign="top" align="center"><b><%= txtRgtColBlk %></b></td></tr>
<!--  START sticky items  -->
<tr><td valign="top">
<table><tr><td valign="top"><span class="fNorm"><%= txtLfStky %>:</span><br />
<select multiple="multiple" style="text-align:left;" id="left_sticky" name="left_sticky" size="4"<%= disabl %>>
<%= l_sticky %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_up_block('left_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('left_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_sticky', 'right_sticky');" value=" <%= txtRight %> " /><br />
<input type="button" class="details1" onclick="remove_block('left_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_sticky', 'left_select');" value="<%= txtUnstick %>" />
<% else %>
	&nbsp;
<% end if %>
</td></tr></table>
</td><td valign="top">
<table><tr><td valign="top"><span class="fNorm"><%= txtMnStky %>:</span><br />
<select multiple="multiple" style="text-align:left;" id="main_sticky" name="main_sticky" size="4"<%= disabl %>>
<%= m_sticky %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_up_block('main_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('main_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="remove_block('main_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('main_sticky', 'main_select');" value="<%= txtUnstick %>" />
<% else %>
	&nbsp;
<% end if %>
</td></tr></table>
</td><td valign="top">
<table><tr><td valign="top"><span class="fNorm"><%= txtRtStky %>:</span><br />
<select multiple="multiple" style="text-align:left;" id="right_sticky" name="right_sticky" size="4"<%= disabl %>>
<%= r_sticky %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_up_block('right_sticky');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('right_sticky');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_sticky', 'left_sticky');" value=" <%= txtLeft %> " /><br />
<input type="button" class="details1" onclick="remove_block('right_sticky');" value="<%= txtRemove %>" /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_sticky', 'right_select');" value="<%= txtUnstick %>" />
<% else %>
	&nbsp;
<% end if %>
</td></tr></table>
</td></tr>
<!--  end sticky items  -->
<tr><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="left_select" name="left_select" size="10">
<%= l_select %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_left_right_block('left_select', 'left_sticky');" value="Sticky" /><br />
<% end if %>
<input type="button" class="details11" onclick="move_up_block('left_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details11" onclick="move_down_block('left_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details11" onclick="move_left_right_block('left_select', 'right_select');" value=" <%= txtRight %> " /><br />
<input type="button" class="details11" onclick="remove_block('left_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="main_select" name="main_select" size="10">
<%= m_select %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_left_right_block('main_select', 'main_sticky');" value="<%= txtSticky %>" /><br />
<% end if %>
<input type="button" class="details1" onclick="move_up_block('main_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('main_select');" value=" <%= txtDown %> " /><br />
<!-- <input type="button" class="details1" onclick="move_left_right_block('main_select', 'right_select');" value="Move Right" /> -->
<input type="button" class="details1" onclick="remove_block('main_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="right_select" name="right_select" size="10">
<%= r_select %>
</select>
</td><td align="center">
<% if intMyMax = 0 and membID = 0  then %>
<input type="button" class="details1" onclick="move_left_right_block('right_select', 'right_sticky');" value="<%= txtSticky %>" /><br />
<% end if %>
<input type="button" class="details1" onclick="move_up_block('right_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('right_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_select', 'left_select');" value=" <%= txtLeft %> " /><br />
<input type="button" class="details1" onclick="remove_block('right_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td></tr>
<tr><td align="center" nowrap>
<select style="text-align:left;" id="left_add" name="left_add" onchange="show_description('left_add');">
<option value=""><%= txtAddLftCol %>..</option>
<%= l_options %>
</select><br />
<input type="button" class="button" style="margin-top:5px;" onclick="add_block('left_select', 'left_add');" value="<%= txtAdd %>" />
</td><td align="center" nowrap>
<select style="text-align:left;" id="main_add" name="main_add" onchange="show_description('main_add');">
<option value=""><%= txtAddMnCol %>..</option>
<%= m_options %>
</select><br />
<input type="button" class="button" style="margin-top:5px;" onclick="add_block('main_select', 'main_add');" value="<%= txtAdd %>" />
</td><td align="center" nowrap>
<select style="text-align:left;" id="right_add" name="right_add" onchange="show_description('right_add');">
<option value=""><%= txtAddRtCol %>..</option>
<%= r_options %>
</select><br />
<input type="button" class="button" style="margin-top:5px;" onclick="add_block('right_select', 'right_add');" value="<%= txtAdd %>" />
</td></tr>
<tr><td colspan="3"><div id="instructions"></div>

<center><input type="submit" value="<%= txtSubmit %>" />&nbsp;
<input type="button" value="<%= txtResetDefBlk %>" onclick="window.location='pop_myMax.asp?mode=reset';" /></center>
</td></tr>
</table>
</form>
</td></tr>
<% Response.Write("</table>")
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer_short.asp" -->
