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
'::
'::  SkyPages originally developed by Machete for SkyPortal
'::  Code modified to the current version by SkyDogg
'::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
bUseAspInput = true
pgType = "manager"

b_desc = ""
l_options = ""
m_options = ""
r_options = ""
l_select = ""
mt_select = ""
mb_select = ""
r_select = ""
pg_content = ""
pg_acontent = ""
		m_title = ""
		m_description = ""
		m_keywords = ""
		m_expires = ""
		m_rating = ""
		m_distribution = ""
		m_robots = ""

p_id = 0
cmd = 0
iMode = 0

if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		cmd = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
end if

if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
end if

if Request("p_id") <> "" or  Request("p_id") <> " " then
	if IsNumeric(Request("p_id")) = True then
		p_id = cLng(Request("p_id"))
	else
		closeAndGo("default.asp")
	end if
end if

hasEditor = true  
strEditorType = "advanced"
strEditorElements = "Message"
editorFull = true
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<!--#include file="lang/en/core_admin.asp" -->
<!--#include file="inc_top.asp" -->
<%
If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then

if iMode = 1 or iMode = 2 then 
		Err_Msg = ""
		if (Request.Form("pg_name")) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter A Page Name!</li>"
		end if
		if (Request.Form("pg_title")) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter A Page Title!</li>"
		end if
		if (Request.Form("pg_display")) = 0 and len(Request.Form("pg_otherurl")) < 4   then 
				Err_Msg = Err_Msg & "<li>Your Alternate Page Name is invalid.  Please enter a proper page name!</li>"
		end if
end if

if Err_Msg = "" then
Select case iMode
	case 1 'create new page
		if request.Form("pg_delete")="" then
			pg_delete=0
		else
			pg_delete=1
		end if
		left_col = request.Form("left_select")
		maintop_col = request.Form("maintop_select")
		mainbottom_col = request.Form("mainbottom_select")
		right_col = request.Form("right_select")
		'html_content = chkString(Request.Form("message"),"message")
		html_content = replace(Request.Form("message"),"'","''")
		if intIsSuperAdmin then
		  asp_content = replace(Request.Form("asp_content"),"'","''")
		  if asp_content <> "" and bUseAspInput then
		    html_content = ""
		  else
		    asp_content = ""
		  end if
		else
		  asp_content = ""
		end if
		
		m_title = replace(request.Form("m_title"),"'","''")
		m_description = replace(request.Form("m_description"),"'","''")
		m_keywords = replace(request.Form("m_keywords"),"'","''")
		m_expires = request.Form("m_expires")
		m_rating = request.Form("m_rating")
		m_distribution = request.Form("m_distribution")
		m_robots = request.Form("m_robots")
		
		sSQL = "INSERT INTO PORTAL_PAGES ("
		sSQL = sSQL & "P_TITLE"
		sSQL = sSQL & ", P_NAME"
		sSQL = sSQL & ", P_INAME"
		sSQL = sSQL & ", P_LEFTCOL"
		sSQL = sSQL & ", P_MAINTOP"
		sSQL = sSQL & ", P_MAINBOTTOM"
		sSQL = sSQL & ", P_RIGHTCOL"
		sSQL = sSQL & ", P_CONTENT"
		sSQL = sSQL & ", P_ACONTENT"
		sSQL = sSQL & ", P_USE_PG_DISP"
		sSQL = sSQL & ", P_OTHER_URL"
		sSQL = sSQL & ", P_CAN_DELETE"
		sSQL = sSQL & ", P_META_TITLE"
		sSQL = sSQL & ", P_META_DESC"
		sSQL = sSQL & ", P_META_KEY"
		sSQL = sSQL & ", P_META_EXPIRES"
		sSQL = sSQL & ", P_META_RATING"
		sSQL = sSQL & ", P_META_DIST"
		sSQL = sSQL & ", P_META_ROBOTS"
		sSQL = sSQL & ") VALUES ("
		sSQL = sSQL & "'" & ChkString(Request.Form("pg_title"),"sqlstring") & "'"
		sSQL = sSQL & ", '" & ChkString(Request.Form("pg_name"),"sqlstring") & "'"
		sSQL = sSQL & ", '" & replace(Request.Form("pg_name")," ","_") & "'"
		sSQL = sSQL & ", '" & left_col & "'"
		sSQL = sSQL & ", '" & maintop_col & "'"
		sSQL = sSQL & ", '" & mainbottom_col & "'"
		sSQL = sSQL & ", '" & right_col & "'"
		sSQL = sSQL & ", '" & html_content & "'"
		sSQL = sSQL & ", '" & asp_content & "'"
		sSQL = sSQL & ", '" & Request.Form("pg_display") & "'"
		sSQL = sSQL & ", '" & ChkString(Request.Form("pg_otherurl"),"sqlstring") & "'"
		sSQL = sSQL & ", '" & pg_delete & "'"
		sSQL = sSQL & ", '" & m_title & "'"
		sSQL = sSQL & ", '" & m_description & "'"
		sSQL = sSQL & ", '" & m_keywords & "'"
		sSQL = sSQL & ", '" & m_expires & "'"
		sSQL = sSQL & ", '" & m_rating & "'"
		sSQL = sSQL & ", '" & m_distribution & "'"
		sSQL = sSQL & ", '" & m_robots & "'"
		sSQL = sSQL & ")"
		executeThis(sSQL)
	case 2 'update existing page
		if request.Form("pg_delete")="" then
			pg_delete=0
		else
			pg_delete=1
		end if
		left_col = request.Form("left_select")
		maintop_col = request.Form("maintop_select")
		mainbottom_col = request.Form("mainbottom_select")
		right_col = request.Form("right_select")
		html_content = replace(Request.Form("message"),"'","''")
		'asp_content = replace(Request.Form("asp_content"),"'","''")
		if intIsSuperAdmin then
		  asp_content = replace(replace(Request.Form("asp_content"),"'","''"),"< /textarea>","</textarea>")
		  if asp_content <> "" and bUseAspInput then
		    html_content = ""
		  else
		    asp_content = ""
		  end if
		else
		  asp_content = ""
		end if
		    'html_content = "manual adjust"
		  'asp_content = ""
		
		m_title = replace(request.Form("m_title"),"'","''")
		m_description = replace(request.Form("m_description"),"'","''")
		m_keywords = replace(request.Form("m_keywords"),"'","''")
		m_expires = request.Form("m_expires")
		m_rating = request.Form("m_rating")
		m_distribution = request.Form("m_distribution")
		m_robots = request.Form("m_robots")
		
		sSQL = "UPDATE PORTAL_PAGES SET"
		sSQL = sSQL & " P_TITLE = '" & ChkString(Request.Form("pg_title"),"sqlstring") & "'"
		sSQL = sSQL & ", P_NAME = '" & ChkString(Request.Form("pg_name"),"sqlstring") & "'"
		sSQL = sSQL & ", P_LEFTCOL = '" & left_col & "'"
		sSQL = sSQL & ", P_MAINTOP = '" & maintop_col & "'"
		sSQL = sSQL & ", P_MAINBOTTOM = '" & mainbottom_col & "'"
		sSQL = sSQL & ", P_RIGHTCOL = '" & right_col & "'"
		sSQL = sSQL & ", P_CONTENT = '" & html_content & "'"
		sSQL = sSQL & ", P_ACONTENT = '" & asp_content & "'"
		sSQL = sSQL & ", P_USE_PG_DISP = '" & Request.Form("pg_display") & "'"
		sSQL = sSQL & ", P_OTHER_URL = '" & ChkString(Request.Form("pg_otherurl"),"sqlstring") & "'"
		sSQL = sSQL & ", P_CAN_DELETE = '" & pg_delete & "'"
		sSQL = sSQL & ", P_META_TITLE = '" & m_title & "'"
		sSQL = sSQL & ", P_META_DESC = '" & m_description & "'"
		sSQL = sSQL & ", P_META_KEY = '" & m_keywords & "'"
		sSQL = sSQL & ", P_META_EXPIRES = '" & m_expires & "'"
		sSQL = sSQL & ", P_META_RATING = '" & m_rating & "'"
		sSQL = sSQL & ", P_META_DIST = '" & m_distribution & "'"
		sSQL = sSQL & ", P_META_ROBOTS = '" & m_robots & "'"
		sSQL = sSQL & " WHERE p_id = " & p_id
		executeThis(sSQL)
end select
else
	cmd = 4
end if %>

<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	fpConfigMenu("1")
  	response.Write("<hr />")
	menu_admin()
spThemeBlock1_close(intSkin) %>
      </td>
<td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
	if p_id = 0 then
		if cmd = 3 then
			bcarg3 = txtNewCustPg
		else
			bcarg3 = ""
		end if
	else
		bc3SQL = "select P_NAME from PORTAL_PAGES where p_id =" & p_id
		set rsBC3 = my_Conn.execute(bc3SQL)
		bcarg3 = "Editing " & rsBC3("P_NAME") & "|admin_config_cp.asp?cmd=1&p_id=" & p_id
		set rsBC3 = nothing
	end if

  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtCustPgCfg & "|admin_config_cp.asp"
  arg3 = bcarg3
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

if strFPmsg <> "" then
    call showMsgBlock(1,strFPmsg)
end if

select case cmd
	case 1 'edit existing page
		if Request("p_id") <> "" then
			editCPlayout()
		else
			closeAndGo("admin_config_cp.asp")
		end if
	case 0,2
		cpSelectPg()
		metaTags()
	case 3 'create new page
		editCPlayout()
	case 4 'copy existing page
		if Request("p_id") <> "" then
			editCPlayout()
		else
			closeAndGo("admin_config_cp.asp")
		end if
	case 5
		%>
		<p align=center><%= txtThereIsProb %></p>
			<table align=center border=0>
				<tr>
					<td>
						<ul><% =Err_Msg %></ul>
					</td>
 				</tr>
			</table>
		<p align=center><a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
		<%
end select

%>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %>
<%     Response.Redirect "admin_login.asp?target=admin_config_cp.asp" %>
<% end if 

' here are the page subroutines
sub editCPlayout()
	spThemeTitle= txtSkyPgCre
	spThemeBlock1_open(intSkin)

select case iMode
	case 1
		response.write "<table border=1><tr><td class='tSubTitle'><b>" & txtNewPgCrtd & "</b></td></tr></table>"
	case 2
		response.write "<table border=1><tr><td class='tSubTitle'><b>" & txtChgsSvd & "</b></td></tr></table>"
end select


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
set rsMM = nothing %>
<script type="text/javascript">
var block_descr = new Array();
<%= b_desc %>
</script>
  <%
  ' populate the select boxes with the default config
if p_id > 0 and (cmd = 1 or cmd = 4) then 
  'edit existing page
  mmSQL = "select * from PORTAL_PAGES where p_id =" & p_id
  set rsMM = my_Conn.execute(mmSQL)
  
  l_col = split(rsMM("p_leftcol"),",")
  maintop_col = split(rsMM("p_maintop"),",")
  mainbottom_col = split(rsMM("p_mainbottom"),",")
  r_col = split(rsMM("p_rightcol"),",")
  
  for cp = 0 to ubound(l_col)
	l_select = l_select & "<option value=""" & l_col(cp) & """>" & split(l_col(cp),":")(0) & "</option>" & vbcrlf
  next
  for cp = 0 to ubound(maintop_col)
	mt_select = mt_select & "<option value=""" & maintop_col(cp) & """>" & split(maintop_col(cp),":")(0) & "</option>" & vbcrlf
  next
  for cp = 0 to ubound(mainbottom_col)
	mb_select = mb_select & "<option value=""" & mainbottom_col(cp) & """>" & split(mainbottom_col(cp),":")(0) & "</option>" & vbcrlf
  next
  for cp = 0 to ubound(r_col)
	r_select = r_select & "<option value=""" & r_col(cp) & """>" & split(r_col(cp),":")(0) & "</option>" & vbcrlf
  next
  
  p_id = rsMM("P_ID")
  pg_title = replace(rsMM("P_TITLE"),"''","'")
  pg_name = replace(rsMM("P_NAME"),"''","'")
  pg_content = replace(rsMM("P_CONTENT"),"''","'")
  if len(rsMM("P_ACONTENT")) > 0 then
    pg_acontent = replace(replace(rsMM("P_ACONTENT"),"''","'"),"</textarea>","< /textarea>")
  else
    pg_acontent = ""
  end if
  'mnu_grp =  rsMM("P_MENU_GRP")
  'ckmnu_show = rsMM("P_MENU_SHOW")
  ckpg_display = rsMM("P_USE_PG_DISP")
  pg_otherurl = rsMM("P_OTHER_URL")
  ckpg_delete = rsMM("P_CAN_DELETE")
  
  ':: get meta tag info
		m_title = rsMM("P_META_TITLE")
		m_description = rsMM("P_META_DESC")
		m_keywords = rsMM("P_META_KEY")
		m_expires = rsMM("P_META_EXPIRES")
		m_rating = rsMM("P_META_RATING")
		m_distribution = rsMM("P_META_DIST")
		m_robots = rsMM("P_META_ROBOTS")

  set rsMM = nothing
end if %>
<div style="text-align:left;margin:10px;">
<script type="text/javascript">
var selectedtablink="";

function handlelink(aobject,tab){
//selectedtablink=aobject.href;
if (document.getElementById){
var tabobj=document.getElementById("tablist");
var tabobjlinks=tabobj.getElementsByTagName("A");
for (i=0; i<tabobjlinks.length; i++)
tabobjlinks[i].className=""
document.getElementById("" + tab + "").className="current";
$('pg_html', 'pg_metaTags', 'pg_layout', 'pg_asp').invoke('hide');
$(aobject).show();
return false;
}
else
return true;
}
</script>
<b><%= txtPgEdCr %></b><br /><br /><%= txtPgEdCr2 %>
<br /><br />
<%= txtPgEdCr3 %>
</div>
	<% if cmd = 3 then %>
		<form method="post" action="admin_config_cp.asp" onsubmit="return select_options();">
		<input type="hidden" name="mode" value="1" />
		<%pg_name = ""%>
		<%ckmnu_show = 0%>
		<%ckpg_display = 1%>
	<% end if %>
	<% if cmd = 1 then %>
		<form method="post" action="admin_config_cp.asp?cmd=1&p_id=<% = p_id %>" onsubmit="return select_options();">
		<input type="hidden" name="mode" value="2" />
	<% end if %>
	<% if cmd = 4 then %>
		<%pg_name = "Copy of " & pg_name%>
		<form method="post" action="admin_config_cp.asp?cmd=4" onsubmit="return select_options();">
		<input type="hidden" name="mode" value="1" />
	<% end if %>
<table border="0" cellspacing="0" cellpadding="3" width="100%">
	<tr class="tCellAlt2"> 
		<td align="right" width="30%"><b><%= txtPgNam %>:</b>&nbsp;</td>
		<td><input class="textbox" type="text" name="pg_name" size="50" maxlength="250" value="<%= pg_name %>"></td>
	</tr>
	<tr class="tCellAlt2"> 
		<td align="right"><b><%= txtPgTitle %>:</b>&nbsp;</td>
		<td><input class="textbox" type="text" name="pg_title" size="50" maxlength="250" value="<%= pg_title %>"></td>
	</tr>
	<tr class="tCellAlt1">
		<td align="right"><b><%= txtPgHowToDisp %>:</b>&nbsp; </td>
		<td><input type="radio" name="pg_display" value="1" <% if ckpg_display = "1" then Response.Write ("checked") end if %>> <%= txtPgUseGenCont %>: <i><%=strHomeURL%>SkyPage.asp</i><br />
						<input type="radio" name="pg_display" value="0" <% if ckpg_display = "0" then Response.Write ("checked") end if %>> <%= txtPgOtherPg %>: <input class="textbox" type="text" name="pg_otherurl" size="20" value="<% if pg_otherurl <> "" then Response.Write(pg_otherurl) end if %>"> (<%= txtPgExample %>)
		</td>
	</tr>
	<tr class="tCellAlt2">
		<td align="right"><b><%= txtPgCnBDel %></b>&nbsp; </td>
		<td><input class="button" type="checkbox" name="pg_delete" value="1" <% if ckpg_delete = 1 then Response.Write ("checked") end if %>> <%= txtPgCnBDel2 %></td>
	</tr>
	<tr class="tCellAlt1">
		<td colspan="2" align="center"><hr />		
		<span class="fSubTitle"><%= txtPgSelTabs %></span>
		<hr /></td>
	</tr>
	<tr> 
		<td colspan="2">
	  <ul id="tablist">
      <li><a id="tab1" class="current" href="javascript:;" onclick="handlelink('pg_html','tab1');"><%= txtPgHTML %></a></li>
	  <% If bUseAspInput Then %>
      <li><a id="tab2" class="" href="javascript:;" onclick="handlelink('pg_asp','tab2');show('pg_asp');"><%= txtPgASP %></a></li>
	  <% End If %>
	  <li><a id="tab3" class="" href="javascript:;" onclick="handlelink('pg_metaTags','tab3');"><%= txtPgMeta %></a></li>
      <li><a id="tab4" class="" href="javascript:;" onclick="handlelink('pg_layout','tab4');"><%= txtPgLayout %></a></li>
    </ul>
		  <div class="tabframe">
		  <%
		  htmleditor()
		  aspeditor()
		  pg_layout()
		  metaTags()
		  %>
		  </div>
		</td>
	</tr>
</table>
<% 
%>
<!-- end tabs -->
<br />
<center><input type="submit" class="button" value="<%= txtCFP23 %>" /></center>
</form><br />
<%
spThemeBlock1_close(intSkin)
end sub

sub aspeditor() %>
	<div id="pg_asp" style="display:none;">
	<% If bUseAspInput Then %>
	<table width="100%">
	<tr><td colspan="2"><p style="margin:10px;"><b><%= txtPgASPDir %></b></p></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2" align="center">
	<textarea name="asp_content" id="asp_content" cols="50" rows="15"><%= pg_acontent %></textarea>
	</td></tr>
	</table>
	<% End If %>
	</div>
	<%
end sub

sub htmleditor()
	response.Write("<div id=""pg_html"" style=""display:block;"">") %>
	<table width="100%">
	<tr><td colspan="2"><p style="margin:10px;"><b><%= txtHTMLDir %></b></p></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<%
	If strAllowHtml = 1 Then 				
	  displayHTMLeditor "Message", "", "" & pg_content & ""
	  'displayHTMLeditor "Message", "", ""
	else
	  displayPLAINeditor 1,"" & pg_content & ""
	end if
	response.Write("</table></div>")
end sub

sub pg_layout() %>
<div id="pg_layout" style="display:none;">
<p style="text-align:left;margin:10px;">
<%= txtPgLayoutDir %></p>
<table border="1" align="center">
<tr class="tTitle"><td valign="center" width="33%" align="center">
<b><%= txtCFP17 %></b></td>
<td valign="center" width="34%" align="center">
<b><%= txtCFP18 %></b></td>
<td valign="center" align="center">
<b><%= txtCFP19 %></b></td></tr>

<tr><td valign="center">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="left_select" name="left_select" size="10">
<%= l_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_up_block('left_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('left_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('left_select', 'right_select');" value=" <%= txtRight %> " /><br />
<input type="button" class="details1" onclick="remove_block('left_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="top">
<table><tr><td colspan="2" align="center"><b><u><%= txtPgLayAbv %></u></b></td></tr>
<tr><td><select multiple="multiple" style="text-align:left;" id="maintop_select" name="maintop_select" size="10">
<%= mt_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_up_block('maintop_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('maintop_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="remove_block('maintop_select');" value="<%= txtRemove %>" />
</td></tr>
<tr><td colspan="2"><select style="text-align:left;" id="maintop_add" name="maintop_add" onchange="show_description('maintop_add');">
<option value=""><%= txtAddMnCol %>...</option>
<%= m_options %>
</select><br />
<input type="button" class="details1" onclick="add_block('maintop_select', 'maintop_add');" value="<%= txtAdd %>" />
</td></tr>
<tr><td colspan="2" align="center"><b><hr /><%= txtPgContent %><hr /></b></td></tr>
<tr><td colspan="2" align="center"><b><u><%= txtPgLayBelo %></u></b></td></tr>
<tr><td><select multiple="multiple" style="text-align:left;" id="mainbottom_select" name="mainbottom_select" size="10">
<%= mb_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_up_block('mainbottom_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('mainbottom_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="remove_block('mainbottom_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td><td valign="center">
<table><tr><td><select multiple="multiple" style="text-align:left;" id="right_select" name="right_select" size="10">
<%= r_select %>
</select>
</td><td align="center">
<input type="button" class="details1" onclick="move_up_block('right_select');" value=" <%= txtUp %> " /><br />
<input type="button" class="details1" onclick="move_down_block('right_select');" value=" <%= txtDown %> " /><br />
<input type="button" class="details1" onclick="move_left_right_block('right_select', 'left_select');" value=" <%= txtLeft %> " /><br />
<input type="button" class="details1" onclick="remove_block('right_select');" value="<%= txtRemove %>" />
</td></tr></table>
</td></tr>
<tr><td>
<select style="text-align:left;" id="left_add" name="left_add" onchange="show_description('left_add');">
<option value=""><%= txtAddLftCol %>...</option>
<%= l_options %>
</select><br />
<input type="button" class="details1" onclick="add_block('left_select', 'left_add');" value="<%= txtAdd %>" />
</td><td>
<select style="text-align:left;" id="mainbottom_add" name="mainbottom_add" onchange="show_description('mainbottom_add');">
<option value=""><%= txtAddMnCol %>...</option>
<%= m_options %>
</select><br />
<input type="button" class="details1" onclick="add_block('mainbottom_select', 'mainbottom_add');" value="<%= txtAdd %>" />
</td><td>
<select style="text-align:left;" id="right_add" name="right_add" onchange="show_description('right_add');">
<option value=""><%= txtAddRtCol %>...</option>
<%= l_options %>
</select><br />
<input type="button" class="details1" onclick="add_block('right_select', 'right_add');" value="<%= txtAdd %>" />
</td></tr>
<tr><td colspan="3"><div id="instructions"></div>
</td></tr>
</table>
</div>
<%
end sub

sub metaTags() %>
<div id="pg_metaTags" style="display:none;padding:10px;">
<p style="text-align:center;margin-bottom:10px;margin-left:10px;margin-right:10px;margin-top:0px;width:300px;">
<h3><%= txtPgMTags %></h3></p>
  <table align="center" border="0" cellpadding="5" cellspacing="0" class="grid">
    <tr> 
      <td valign=top width="40%"><label for="m_title"><b><%= txtTitle %></b></label><br />
        <span class="fSmall"><%= txtPgMTitleDef %></span></td>
      <td> 
        <input type="text" id="m_title" name="m_title" maxlength="100" size="30" value="<%= m_title %>">
      </td>
    </tr>
    <tr> 
      <td valign=top><label for="m_description"><b><%= txtDesc %></b></label><br />
        <span class="fSmall"><%= txtPgMDescDef %></span><br /><br />
		<span id="charLeft1" class="fSmall">150 characters left&nbsp;</span>
		</td>
      <td> 
        <textarea rows="6" id="m_description" name="m_description" cols="30" onKeyUp="cntChar('m_description','charLeft1','{CHAR} characters left&nbsp;',150);"><%= m_description %></textarea>
      </td>
    </tr>
    <tr> 
      <td valign="top"><label for="m_keywords"><b><%= txtKeyWds %></b></label><br />
        <span class="fSmall"><%= txtPgMKeyWdDef %></span><br /><br />
		<span id="charLeft2" class="fSmall"><b>200 characters left</b>&nbsp;</span>
		</td>
      <td valign=top>
        <textarea rows="8" name="m_keywords" id="m_keywords" cols="30" onKeyUp="cntChar('m_keywords','charLeft2','<b>{CHAR} characters left</b>&nbsp;',200);"><%= m_keywords %></textarea>
      </td>
    </tr>
    <tr> 
      <td valign=top> 
        <label for="m_expires"><b><%= txtPgMExp %></b></label>
        <br />
        <span class="fSmall"><%= txtPgMExpDef %></span></td>
      <td valign=top> 
	  <% If m_expires = "" Then m_expires = "never" %>
        <input id="m_expires" type="text" name="m_expires" value="<%= m_expires %>">
      </td>
    </tr>
    <tr> 
      <td valign=top> 
        <label for="m_rating"><b><%= txtRating %></b></label>
        <br />
        <span class="fSmall"><%= txtPgMRateDef %></span></td>
      <td valign=top> 
        <select id="m_rating" name="m_rating">
          <option<% If m_rating = "" Then response.Write(" selected") %>>(<%= txtNone %>)</option>
          <option value="<%= txtPgMGen %>"<% If m_rating = txtPgMGen Then response.Write(" selected") %>><%= txtPgMGen %></option>
          <option value="<%= txtPgMMature %>"<% If m_rating = txtPgMMature Then response.Write(" selected") %>><%= txtPgMMature %></option>
          <option value="<%= txtPgMRestr %>"<% If m_rating = txtPgMRestr Then response.Write(" selected") %>><%= txtPgMRestr %></option>
        </select>
      </td>
    </tr>
    <tr> 
      <td valign=top> 
        <label for="m_distribution"><b><%= txtPgMDistr %></b></label>
        <br />
        <span class="fSmall"><%= txtPgMDistrDef %></span></td>
      <td valign=top> 
	  <% If m_distribution = "" Then m_distribution = txtPgMGlbl %>
        <select id="m_distribution" name="m_distribution">
          <option value="Global"<% If m_distribution = "Global" Then response.Write(" selected") %>><%= txtPgMGlbl %></option>
          <option value="Local"<% If m_distribution = "Local" Then response.Write(" selected") %>><%= txtPgMLocal %></option>
          <option value="Internal Use"<% If m_distribution = "Internal Use" Then response.Write(" selected") %>><%= txtPgMIntUse %></option>
        </select>
      </td>
    </tr>
    <tr> 
      <td valign=top> 
        <label for="m_robots"><b><%= txtPgMBots %></b></label>
        <br />
        <span class="fSmall"><%= txtPgMBotsDef %></span></td>
      <td valign=top> 
	  <% If m_robots = "" Then m_robots = "index,follow" %>
        <select id="m_robots" name="m_robots">
          <option value="index,follow"<% If m_robots = "index,follow" Then response.Write(" selected") %>><%= txtPgMBot1 %></option>
          <option value="index,nofollow"<% If m_robots = "index,nofollow" Then response.Write(" selected") %>><%= txtPgMBot2 %></option>
          <option value="noindex,follow"<% If m_robots = "noindex,follow" Then response.Write(" selected") %>><%= txtPgMBot3 %></option>
          <option value="noindex,nofollow"<% If m_robots = "noindex,nofollow" Then response.Write(" selected") %>><%= txtPgMBot4 %></option>
        </select>
      </td>
    </tr>
    <tr> 
      <td colspan=2 align=center valign=top><br />
      </td>
    </tr>
  </table>
  </div>
<%
end sub

sub cpSelectPg()
	spThemeTitle= txtSkyPgMan
	spThemeBlock1_open(intSkin)
%>
	<p align=center><font size=4><%= txtSkyPgMan %></font><br /><%= txtPgSelPg %></p>
	<% response.write "<p align=""center""><input class=""button"" type=""button"" value=""" & txtCreNewPg & """ id=""create"" name=""create"" onclick=""location.href='admin_config_cp.asp?cmd=3'""></p>" %>
	<table class="tPlain" width="550" align="center">
		<tr class="tSubTitle">
			<td align="center" ><b><%= ucase(txtPgOpts) %></b></td>
			<td align="center" ><b><%= ucase(txtID) %><span class="fAlert"> *</span></b></td>
			<td align="center" ><b><%= ucase(txtPgNam) %></b></td>
			<td align="center" ><b><%= ucase(txtPgUrl) %></b></td>
		</tr>
		<tr class="tCellAlt2"><td align="center">
		<a href="admin_config_fp.asp?cmd=3" title="<%= txtEdit %>"><%= icon(icnEdit,txtEdit,"","","") %></a>
		<a href="javascript:;" onclick="window.open('default.asp');" title="<%= txtView %>"><%= icon(icnBinoc,txtView,"","","") %></a></td>
		<td align="center"><b>-</b></td>
		<td align="center"><a href="admin_config_fp.asp?cmd=3" title="<%= txtEdit %>"><%= txtPgHomePg %></a></td>
		<td align=center><a href="<%= oClik %>" target="_blank" title="<%= txtView %>">default.asp</a></td>
		</tr>
<%	
	strSql1 = "SELECT * FROM PORTAL_PAGES ORDER BY P_NAME ASC"
	Set rs1 = my_Conn.Execute(strSql1)
		rColor = "tCellAlt2"
	do while not rs1.eof

		if rColor = "tCellAlt1" then 
			rColor = "tCellAlt2"
		else
			rColor = "tCellAlt1"
		end if

		%>
		<tr class="<%=rColor%>">
			<td align="center">
				<% if rs1("P_USE_PG_DISP")=1 then
				      oClik = "SkyPage.asp?pg=" & rs1("p_id") %>
				<% else
				 	  if trim(rs1("P_OTHER_URL")) = "" then
				        oClik = "SkyPage.asp?pg=" & rs1("p_id")
					  else
				        oClik = rs1("P_OTHER_URL")
					  end if %>
				<% end if %>
				<% if rs1("P_CAN_DELETE")=1 then %>
		<a href="admin_pages_delete.asp?p_id=<%=rs1("p_id")%>" title="<%= txtDel %>"><%= icon(icnDelete,txtDel,"","","") %></a>
				<% end if %>
		<a href="admin_config_cp.asp?cmd=1&p_id=<%=rs1("p_id")%>" title="<%= txtEdit %>"><%= icon(icnEdit,txtEdit,"","","") %></a>
		<a href="admin_config_cp.asp?cmd=4&p_id=<%=rs1("p_id")%>" title="<%= txtCopy %>"><%= icon(icnCopy,txtCopy,"","","") %></a>
		<a href="javascript:;" onclick="window.open('<%= oClik %>');" title="<%= txtView %>"><%= icon(icnBinoc,txtView,"","","") %></a>
			</td>
			<td align="center"><b><%=rs1("p_id")%></b></td>
			<td align="center"><a href="admin_config_cp.asp?cmd=1&p_id=<%=rs1("p_id")%>" title="<%= txtEdit %>"><%=rs1("p_name")%></a></td>
			<td align=center><a href="<%= oClik %>" target="_blank" title="<%= txtView %>"><%= oClik %></a>
			</td>
		</tr>
		<%
	rs1.movenext
	loop
	rs1.close
	set rs1=nothing
	response.write "</table><br />"
	Response.Write("<span class=""fAlert"">* </span>")
	Response.Write("Use this value as the [<b>sPage_id</b>] variable when creating a SkyPage template page")
	spThemeBlock1_close(intSkin)
end sub

%>