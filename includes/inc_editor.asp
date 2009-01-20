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
if hasEditor = true then
  if strAllowHTML = 1 then
	if intEditor = 5 then
	  if hasAccess(1) and intIsSuperAdmin then
	    if editorType = "tinymce" then
	      strEditBtn = "advanced"
		  editorAdmin = true
	      'strEditBtn = "default"
	    elseif editorType = "fckeditor" then
	      strEditBtn = "Admin"
	    end if
	  end if
	end if
	if hasAccess(intEditor) then
	  if editorType = "tinymce" then
	    strEditBtn = "advanced"
		editorAdmin = true
	    'strEditBtn = "default"
	  elseif editorType = "fckeditor" then
	    strEditBtn = "Admin"
	  end if
	else
	  if editorType = "tinymce" then
	    strEditBtn = "advanced"
		editorAdmin = false
	    'strEditBtn = "default"
	  elseif editorType = "fckeditor" then
	    strEditBtn = "Default"
	  end if
	end if
  End If

   	If editorType = "tinymce" and strAllowHtml = 1 Then
	  if strEditorType <> "" then
	    strEditBtn = strEditorType
	  end if %>
<!-- tinyMCE -->
<script language="javascript" type="text/javascript">
	tinyMCE.init({
		theme : "<%= strEditBtn %>",
		language : "<%= strLang %>",
		mode : "exact",
		elements : "<%= strEditorElements %>" ,
		//convert_newlines_to_brs : true ,
		//force_br_newlines : true ,
		//preformatted : false ,
		//content_css : "<%= strHomeURL %>themes/Frosty Strawberry/editor.css",
		//editor_css : "<%= strHomeURL %>themes/<%= strTheme %>/editor.css",		
		
		<% if editorAdmin then %>
		//theme_advanced_buttons3_add : "advhr,fullpage",
		theme_advanced_disable : "hr,anchor,visualaid",
		<% Else %>
		theme_advanced_disable : "hr,anchor,code,visualaid",
		<% end if %>
		
<% if editorFull = true then %>
		
		plugins : "style,table,advhr,advimage,advlink,emotions,iespell,preview,zoom,print,contextmenu,paste,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template",
		theme_advanced_buttons1 : "preview,print,newdocument,|,styleselect,formatselect,fontselect,fontsizeselect",
		theme_advanced_buttons2 : "cut,copy,paste,pastetext,pasteword,|,bullist,numlist,|,outdent,indent,|,undo,redo,|,link,unlink,anchor,|,forecolor,backcolor,|,iespell",
		theme_advanced_buttons3 : "tablecontrols,|,sub,sup,|,charmap,advhr,|,fullscreen",
		theme_advanced_buttons4 : "emotions,image,|,bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,|,styleprops,attribs,|,nonbreaking,removeformat,|,code,help",
		//template_external_list_url : "example_template_list.js"
<% Else %>
		plugins : "style,table,advhr,advimage,advlink,emotions,iespell,preview,zoom,print,contextmenu,paste,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template",
		theme_advanced_buttons1 : "newdocument,|,styleselect,formatselect,fontselect,fontsizeselect",
		theme_advanced_buttons2 : "cut,copy,paste,pastetext,pasteword,|,bullist,numlist,|,undo,redo,|,link,unlink,|,forecolor,backcolor,|,iespell,advhr",
		//theme_advanced_buttons2 : "tablecontrols,|,sub,sup,|,charmap,|,fullscreen",
		theme_advanced_buttons3 : "emotions,image,|,bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,|,styleprops,attribs,|,nonbreaking,removeformat,|,code,help",
<% end if %>
		
		theme_advanced_toolbar_location : "top",
		theme_advanced_toolbar_align : "center",
		theme_advanced_path_location : "bottom",
		theme_advanced_resize_horizontal : false,
		theme_advanced_resizing : true,
		content_css : "<%= strHomeURL %>themes/<%= strTheme %>/editor.css",
	    plugin_insertdate_dateFormat : "%m-%d-%Y",
	    plugin_insertdate_timeFormat : "%H:%M:%S",
		extended_valid_elements : "a[name|href|target|title|onclick],img[class|src|border=0|alt|title|hspace|vspace|width|height|align|onmouseover|onmouseout|name|id|onclick],hr[class|width|size|noshade],font[face|size|color|style|id],span[class|align|style|id],form[name|method|action|id],textarea[name|id|rows|cols|style|wrap|readonly],pre[id],br[id],input[type|value|id|name|class|src|border|style|onclick],script[type],iframe[name|id|style|width|height|src|frameborder|scrolling],fieldset[style|id]",
		//invalid_elements : "a",
		//theme_advanced_styles : "Quote=quote;Code=code;Header 1=header1;Header 2=header2;Header 3=header3;Table Row=tableRow1", // Theme specific setting CSS classes
		//theme_advanced_styles : "Quote=quote;Code=code;Header 1=fTitle;Header 2=fSubTitle",
		debug : false
	});
	
	function toggleEditor(id) {
	var elm = document.getElementById(id);

	if (tinyMCE.getInstanceById(id) == null)
		tinyMCE.execCommand('mceAddControl', false, id);
	else
		tinyMCE.execCommand('mceRemoveControl', false, id);
	}
	
	// Custom save callback, gets called when the contents is to be submitted
	function customSave(id, content) {
		alert(id + "=" + content);
	}
</script>
<!-- /tinyMCE -->
<% 	ElseIf editorType = "fckeditor" and strAllowHtml = 1 Then %>
	<script type="text/javascript" src="FCKeditor/fckeditor.js"></script>
<% 
    if instr(strEditorElements,",") > 0 then
	  strEditorElements = split(strEditorElements,",")(0)
	end if
	editorJS = "<script type=""text/javascript"">" & vbcrlf
	editorJS = editorJS & "<!-- FCKeditor Modified  for  SkyPortal  by  S k y D o gg -->" & vbcrlf
	editorJS = editorJS & "var oFCKeditor = new FCKeditor('" & strEditorElements & "','550','300','" & strEditBtn & "') ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.BasePath = 'fckeditor/' ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""EditorAreaCSS"" ] = """ & strHomeUrl & "Themes/" & strTheme & "/editor.css"" ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""SkinPath"" ] = """ & strHomeUrl & "Themes/" & strTheme & "/editor/"" ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""SmileyPath"" ] = '" & strHomeUrl & "images/smilies/' ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""UseBROnCarriageReturn"" ] = false ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""ToolbarCanCollapse"" ] = false ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""TabSpaces"" ] = 4 ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""ImageBrowser"" ] = false ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""LinkBrowser"" ] = false ;" & vbcrlf
	editorJS = editorJS & "//FCKeditor( instanceName[, width, height, toolbarSet, value] )" & vbcrlf
	editorJS = editorJS & "oFCKeditor.ReplaceTextarea() ;" & vbcrlf
	editorJS = editorJS & "</script>" & vbcrlf
	
	editorJS = "<script type=""text/javascript"">" & vbcrlf
	editorJS = editorJS & "<!-- FCKeditor Modified  for  SkyPortal  by  S k y D o gg" & vbcrlf
	editorJS = editorJS & "var oFCKeditor = new FCKeditor('" & strEditorElements & "','550','300','" & strEditBtn & "') ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.BasePath = 'fckeditor/' ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""AutoDetectLanguage"" ] = false ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.Config[ ""DefaultLanguage"" ] = """ & strLang & """ ;" & vbcrlf
	editorJS = editorJS & "oFCKeditor.ReplaceTextarea() ;" & vbcrlf
	editorJS = editorJS & "-->" & vbcrlf
	editorJS = editorJS & "</script>" & vbcrlf	
 	End If %>
<% 	if strAllowForumCode = 1 then
	  Response.Write "<script type=""text/javascript"" src=""includes/scripts/inc_jsfcode.js""></script>"	
 	End If
end if %>
<%	

' This function writes out the icon String
Function writeicon(title,icon,w,h,command,iname,offset)
   varoffset=offset*(w+3)
    Response.Write "<div unselectable=""on"" style=""position:absolute;margin-left:"&varoffset&"px;margin-right:1;width:"&w+3&";height:"&h&";"">"
    Response.Write "	     <div unselectable='on' style=""position:absolute;clip: rect(0 "&w&"px "&h&"px 0)"">"
	Response.Write "	       <img unselectable='on' src="""&icon&"""  style=""width:"&w&"px;height:"&h&"px;border:0px;z-index:100;position:absolute;top:-0"" title="""&title&""" alt="""&title&""" "
	Response.Write "	            onmouseover=""document."&iname&".style.top=-"&h&""" onmouseout=""document."&iname&".style.top=0"" "
	Response.Write "		        onmousedown=""document."&iname&".style.top=-"&h&""" onmouseup=""document."&iname&".style.top=0;"&command&""" />"
	Response.Write "	       <img name="""&iname&""" unselectable='on' src="""&strHomeUrl&"Themes/"&strTheme&"/edit_button.gif"" style=""z-index:10;position:absolute;top:-0;width:"&w&""" />"
	Response.Write "	     </div>"
	Response.Write "</div>"
End Function 

function showPostButtons()
  thmbtnw = 23                          ' Width of the themed editor button
  thmbtnh = 22                          ' Height of the themed editor button
'First row of buttons
   Dim MWPbuttons(11) 
if lcase(strIcons) = "1" then 
   MWPbuttons(10)="""Insert Smilie"","""&strHomeUrl&"images/editor_icons/icon_editor_smilie.gif"",23,22,""JavaScript:openWindow2('pop_portal.asp?cmd=9');"",""btnsmile"""
Else
   ReDim MWPbuttons(10) 
End If

MWPbuttons(0)="""Bold"","""&strHomeUrl&"images/editor_icons/icon_editor_bold.gif"",23,22,""Javascript:bold();"",""btnbold"""
MWPbuttons(1)="""Italicized"","""&strHomeUrl&"images/editor_icons/icon_editor_italicize.gif"",23,22,""Javascript:italicize();"",""btnitalic"""
MWPbuttons(2)="""Underline"","""&strHomeUrl&"images/editor_icons/icon_editor_underline.gif"",23,22,""Javascript:underline();"",""btnunderline"""
MWPbuttons(3)="""Strikethrough"","""&strHomeUrl&"images/editor_icons/icon_editor_strike.gif"",23,22,""Javascript:strike();"",""btnstrike"""
MWPbuttons(4)="""Insert Hyperlink"","""&strHomeUrl&"images/editor_icons/icon_editor_url.gif"",23,22,""Javascript:hyperlink();"",""btnhyper"""
MWPbuttons(5)="""Insert Email"","""&strHomeUrl&"images/editor_icons/icon_editor_email.gif"",23,22,""Javascript:email();"",""btnemail"""
MWPbuttons(6)="""Insert Image"","""&strHomeUrl&"images/editor_icons/icon_editor_image.gif"",23,22,""Javascript:image();"",""btnimage"""
MWPbuttons(7)="""Insert Code"","""&strHomeUrl&"images/editor_icons/icon_editor_code.gif"",23,22,""Javascript:showcode();"",""btncode"""
MWPbuttons(8)="""Insert Quote"","""&strHomeUrl&"images/editor_icons/icon_editor_quote.gif"",23,22,""Javascript:quote();"",""btnquote"""
MWPbuttons(9)="""Insert List"","""&strHomeUrl&"images/editor_icons/icon_editor_list.gif"",23,22,""Javascript:list();"",""btnlist"""

' This is the second row of buttons
Dim MWPbuttons2(10) 
MWPbuttons2(0)="""Superscript"","""&strHomeUrl&"images/editor_icons/icon_editor_sup.gif"",23,22,""Javascript:sup();"",""btnsup"""
MWPbuttons2(1)="""Subscript"","""&strHomeUrl&"images/editor_icons/icon_editor_sub.gif"",23,22,""Javascript:sub();"",""btnsub"""
MWPbuttons2(2)="""Align Left"","""&strHomeUrl&"images/editor_icons/icon_editor_left.gif"",23,22,""Javascript:aleft();"",""btnleft"""
MWPbuttons2(3)="""Centered"","""&strHomeUrl&"images/editor_icons/icon_editor_center.gif"",23,22,""Javascript:center();"",""btncenter"""
MWPbuttons2(4)="""Align Right"","""&strHomeUrl&"images/editor_icons/icon_editor_right.gif"",23,22,""Javascript:aright();"",""btnright"""
MWPbuttons2(5)="""Pre-formated"","""&strHomeUrl&"images/editor_icons/icon_editor_pre.gif"",23,22,""Javascript:pre();"",""btnpre"""
MWPbuttons2(6)="""Teletype"","""&strHomeUrl&"images/editor_icons/icon_editor_tt.gif"",23,22,""Javascript:tt();"",""btntt"""
MWPbuttons2(7)="""Moving Text"","""&strHomeUrl&"images/editor_icons/icon_editor_move.gif"",23,22,""Javascript:marquee();"",""btnmove"""
MWPbuttons2(8)="""Insert Horizontal Rule"","""&strHomeUrl&"images/editor_icons/icon_editor_hr.gif"",23,22,""Javascript:hr();"",""btnhr"""
MWPbuttons2(9)="""Highlight (Yellow)"","""&strHomeUrl&"images/editor_icons/icon_editor_hl.gif"",23,22,""Javascript:hl();"",""btnhl"""

%>
<tr>
<td align=right>
</td>
<td>
  <table width="200"><tr>
	<td width="120">
    <select name="font" onchange="showfont(this.options[this.selectedIndex].value)">
	<option value="Verdana" selected="selected"> -- Font Type -- </option>
	<option value="Andale Mono">Andale Mono</option>
	<option value="Arial">Arial</option>
	<option value="Arial Black">Arial Black</option>
	<option value="Book Antiqua">Book Antiqua</option>
	<option value="Century Gothic">Century Gothic</option>
	<option value="Comic Sans MS">Comic Sans MS</option>
	<option value="Courier New">Courier New</option>
	<option value="Georgia">Georgia</option>
	<option value="Impact">Impact</option>
	<option value="Tahoma">Tahoma</option>
	<option value="Times New Roman">Times New Roman</option>
	<option value="Trebuchet MS">Trebuchet MS</option>
	<option value="Script MT Bold">Script MT Bold</option>
	<option value="Stencil">Stencil</option>
	<option value="Verdana">Verdana</option>
	<option value="Lucida Console">Lucida Console</option>
</select></td><td align="left">
<select name="size" onchange="showsize(this.options[this.selectedIndex].value)">
	<option value="3" selected="selected"> Size </option>
	<option value="1">1</option>
	<option value="2">2</option>
	<option value="3">3</option>
	<option value="4">4</option>
	<option value="5">5</option>
	<option value="6">6</option>	
</select>
	</td>
	<td></td></tr></table></td>
</tr>
<tr>
<td align=left></td>
<td align=left>
&nbsp;<a href="Javascript:red();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_red.gif" width="10" height="22" alt="Red" border="0"></a>
<a href="Javascript:green();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_green.gif" width="10" height="22" alt="Green" border="0"></a>
<a href="Javascript:blue();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_blue.gif" width="10" height="22" alt="Blue" border="0"></a>
<a href="Javascript:white();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_white.gif" width="10" height="22" alt="White" border="0"></a>
<a href="Javascript:purple();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_purple.gif" width="10" height="22" alt="Purple" border="0"></a>
<a href="Javascript:yellow();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_yellow.gif" width="10" height="22" alt="Yellow" border="0"></a>
<a href="Javascript:violet();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_violet.gif" width="10" height="22" alt="Violet" border="0"></a>
<a href="Javascript:brown();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_brown.gif" width="10" height="22" alt="Brown" border="0"></a>
<a href="Javascript:black();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_black.gif" width="10" height="22" alt="Black" border="0"></a>
<a href="Javascript:pink();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_pink.gif" width="10" height="22" alt="Pink" border="0"></a>
<a href="Javascript:orange();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_orange.gif" width="10" height="22" alt="Orange" border="0"></a>
<a href="Javascript:gold();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_gold.gif" width="10" height="22" alt="Gold" border="0"></a>
<a href="Javascript:beige();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_beige.gif" width="10" height="22" alt="Beige" border="0"></a>
<a href="Javascript:teal();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_teal.gif" width="10" height="22" alt="Teal" border="0"></a>
<a href="Javascript:navy();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_navy.gif" width="10" height="22" alt="Navy" border="0"></a>
<a href="Javascript:maroon();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_maroon.gif" width="10" height="22" alt="Maroon" border="0"></a>
<a href="Javascript:limegreen();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_limegreen.gif" width="10" height="22" alt="Limegreen" border="0"></a></td>
</tr>
<tr>
<td align=right valign=top>
<b><%= txtFormat %>:&nbsp;</b>
</td>
<td height="60" align="left" valign="top">
<% If varBrowser = "firefox" Then %>
<div id="iconrow2" style="position:relative;top:2px;margin-left:5px;height=25px">
<% Else %>
<div id="iconrow2" style="position:relative;margin-left:5px;height=25px"> 
  <% End If %>
  <% numIcons=UBound(MWPbuttons2)-1
For i=0 To numIcons
  execute("writeicon " & MWPbuttons2(i)&","&i)
Next %>
</div>
<% If varBrowser = "ie" and isMac Then %>
<div id="iconrow1" style="position:absolute;margin-top:30px;margin-left:5px;height=25px">
<% Elseif varBrowser = "ie" then %>
<div id="iconrow1" style="position:absolute;margin-top:2px;margin-left:5px;height=25px">
<% Elseif varBrowser <> "opera" then %>
<div id="iconrow1" style="position:absolute;margin-top:28px;margin-left:5px;height=25px">
<% Else %>
<div id="iconrow1" style="position:absolute;margin-top:18px;margin-left:5px;height=25px">
<% End If %>
<% numIcons=UBound(MWPbuttons)-1
For i=0 To numIcons
  execute("writeicon " & MWPbuttons(i)&","&i)
Next %>
</div></td>
</tr>
<%
end function

function showPostButtonsShort()
thmbtnw = 23  ' Width of the themed editor button
thmbtnh = 22  ' Height of the themed editor button
%>
        <tr> 
          
    <td height="25" align="left" valign="middle" nowrap="nowrap">&nbsp;</td>
          <td height="25" width="100%">
            <select class="textbox" name="size" onchange="showsize(this.options[this.selectedIndex].value)">
              <option value="3" selected="selected"><%= txtSize %></option>
              <option value="1">1</option>
              <option value="2">2</option>
              <option value="3">3</option>
              <option value="4">4</option>
              <option value="5">5</option>
              <option value="6">6</option>
            </select>
<% 
' This is the list of  editor icons to generate
' Each item needs the following:
'    title - this is the text for the alt tag and the helptext EG "Bold"
'    img - this is the url to icon for the button eg "images/editor_icons/icon_editor_bold.gif"
'    w - this is the width of the button from the themes folder (in pixels) eg 23
'    h - this is the height of the button from the themes folder (in pixels) eg 22
'    command - this is the command that is executed on click eg "bold();"
'    iname - this is the name for the background button tag so it can be toggled on mouseover eg "btnbold"
Dim MWPbuttons(10) 
MWPbuttons(0)="""Bold"","""&strHomeUrl&"images/editor_icons/icon_editor_bold.gif"",23,22,""Javascript:bold();"",""btnbold"""
MWPbuttons(1)="""Italicized"","""&strHomeUrl&"images/editor_icons/icon_editor_italicize.gif"",23,22,""Javascript:italicize();"",""btnitalic"""
MWPbuttons(2)="""Underline"","""&strHomeUrl&"images/editor_icons/icon_editor_underline.gif"",23,22,""Javascript:underline();"",""btnunderline"""
MWPbuttons(3)="""Strikethrough"","""&strHomeUrl&"images/editor_icons/icon_editor_strike.gif"",23,22,""Javascript:strike();"",""btnstrike"""
MWPbuttons(4)="""Insert Horizontal Rule"","""&strHomeUrl&"images/editor_icons/icon_editor_hr.gif"",23,22,""Javascript:hr();"",""btnhr"""
MWPbuttons(5)="""Insert List"","""&strHomeUrl&"images/editor_icons/icon_editor_list.gif"",23,22,""Javascript:list();"",""btnlist"""
MWPbuttons(6)="""Insert Hyperlink"","""&strHomeUrl&"images/editor_icons/icon_editor_url.gif"",23,22,""Javascript:hyperlink();"",""btnhyper"""
MWPbuttons(7)="""Insert Image"","""&strHomeUrl&"images/editor_icons/icon_editor_image.gif"",23,22,""Javascript:image();"",""btnimage"""
MWPbuttons(8)="""Insert Quote"","""&strHomeUrl&"images/editor_icons/icon_editor_quote.gif"",23,22,""Javascript:quote();"",""btnquote"""
MWPbuttons(9)="""Highlight (Yellow)"","""&strHomeUrl&"images/editor_icons/icon_editor_hl.gif"",23,22,""Javascript:hl();"",""btnhl"""
%>
<span id="iconrow" style="position:absolute;margin-left:5">
<%
numIcons=UBound(MWPbuttons)-1
For i=0 To numIcons
  execute("writeicon " & MWPbuttons(i)&","&i)
Next
varleft=(UBound(MWPbuttons))*(thmbtnw+3)
%>
<div id="colors" style="position:absolute;margin-left:<% =varleft%>px;width:57px;">
&nbsp;<a href="Javascript:red();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_red.gif" width="10" height="22" alt="Red" border="0"></a>
      <a href="Javascript:green();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_green.gif" width="10" height="22" alt="Green" border="0"></a>
      <a href="Javascript:blue();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_blue.gif" width="10" height="22" alt="Blue" border="0" style="display:inline;"></a>
      <a href="Javascript:orange();"><img src="<%= strHomeUrl %>images/icon_color/icon_color_orange.gif" width="10" height="22" alt="Orange" border="0"></a>
</div></span>
    </td>
        </tr>
<%
End Function 


function incSmilies() %>

<script language="Javascript" type="text/javascript">
<!-- hide
function insertsmilie(smilieface) {
		if (document.PostTopic.Message.createTextRange && document.PostTopic.Message.caretPos) {
			var caretPos = document.PostTopic.Message.caretPos;
			caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? smilieface + ' ' : smilieface;
			document.PostTopic.Message.focus();
		} else {
			document.PostTopic.Message.value+=smilieface;
			document.PostTopic.Message.focus();
		}
}
// -->
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="2" align="center">
  <tr align="center">
    <td align="center" colspan="3"><span class="fSmall"><b><%= txtInsSmile %></b></span></td>
  </tr>
  <tr align="center" valign="middle">
    <td><a href="Javascript:insertsmilie('[:)]')"><img src="<%= strHomeUrl %>images/Smilies/smile.gif" width="15" height="15" border="0" alt="Smile [:)]" title="Smile [:)]"></a></td>
    <td><a href="Javascript:insertsmilie('[:D]')"><img src="<%= strHomeUrl %>images/Smilies/big.gif" width="15" height="15" border="0" alt="Big Smile [:D]" title="Big Smile [:D]"></a></td>
    <td><a href="Javascript:insertsmilie('[8D]')"><img src="<%= strHomeUrl %>images/Smilies/cool.gif" width="15" height="15" border="0" alt="Cool [8D]" title="Cool [8D]"></a></td>
  </tr>
  <tr align="center" valign="middle">
    <td><a href="Javascript:insertsmilie('[:p]')"><img src="<%= strHomeUrl %>images/Smilies/tongue.gif" width="15" height="15" border="0" alt="Tongue [:P]" title="Tongue [:P]"></a></td>
    <td><a href="Javascript:insertsmilie('[}:)]')"><img src="<%= strHomeUrl %>images/Smilies/evil.gif" width="15" height="15" border="0" alt="Evil [):]" title="Evil [):]"></a></td>
    <td><a href="Javascript:insertsmilie('[;)]')"><img src="<%= strHomeUrl %>images/Smilies/wink.gif" width="15" height="15" border="0" alt="Wink [;)]" title="Wink [;)]"></a></td>
  </tr>
  <tr align="center" valign="middle">
    <td><a href="Javascript:insertsmilie('[:0]')"><img src="<%= strHomeUrl %>images/Smilies/shock.gif" width="15" height="15" border="0" alt="Shocked [:0]" title="Shocked [:0]"></a></td>
    <td><a href="Javascript:insertsmilie('[xx(]')"><img src="<%= strHomeUrl %>images/Smilies/dead.gif" width="15" height="15" border="0" alt="Dead [xx(]" title="Dead [xx(]"></a></td>
    <td><a href="Javascript:insertsmilie('[?]')"><img src="<%= strHomeUrl %>images/Smilies/question.gif" width="15" height="15" border="0" alt="Question [?]" title="Question [?]"></a></td>
  </tr>
  <tr align="center">
    <td align="center" colspan="3"><a href="javascript:openWindow2('pop_portal.asp?cmd=9')"><span class="fSmall"><b><%= txtMorSmile %></b></span></a></td>
  </tr>
</table>
<%
end function

sub displayHTMLeditor(fmName,fmTitle,strM) %>
<tr>
<td align="right" valign="top"><br /><%= fmTitle %>&nbsp;</td>
<td align="left"><div style="width:200px;"><textarea id="<%= fmName %>" name="<%= fmName %>" rows="15" cols="85" style="width:100%;"><%= strM %></textarea></div><br />
<%	if strAllowHTML = 1 and trim(editorJS) <> "" then
  		response.Write(editorJS)
	End If %>
  </td>
</tr>
<%
end sub

sub displayPLAINeditor(typ, strM) %>
<tr>
	<td align="left" colspan="2" height="15"><%  %>&nbsp;
</td></tr>
<% 
if typ = 1 then
	showPostButtons()
else
	showPostButtonsShort()
end if
%>
<tr><td align="right" valign="top" nowrap="nowrap"><br />
	<a href="JavaScript:openWindow3('pop_help.asp?mode=1&place=1#mode')"><img src="images/icons/icon_smile_question.gif" alt="<%= txtHelp %>" title="<%= txtHelp %>" border="0" WIDTH="15" HEIGHT="15"></a>&nbsp;
		      <select name="font" onchange="thelp(this.options[this.selectedIndex].value)">
				<option selected value="0"><%= txtPrompt %>&nbsp;</option>	
				<option value="1"><%= txtHelp %>&nbsp;</option>
				<option value="2"><%= txtBasic %>&nbsp;</option>
	  		  </select>
<br /><br /><div style="width: 100px;"><% incSmilies() %></div>
</td>
<td align="left"><textarea name="Message" id="Message" rows="15" cols="60" onfocus="getActiveText(this)" onkeyup="getActiveText(this)" onselect="getActiveText(this)" onclick="getActiveText(this)" onchange="getActiveText(this)"><%= strM %></textarea><br /><%= strST %></td>
</tr>
<%
end sub
 %>