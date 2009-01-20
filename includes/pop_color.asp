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

%>
  
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<%
Dim tCount
%>
<head>
<script type="text/javascript">
function retinfo(V2)
{
    opnform=window.opener.document.forms['Form1'];
    opnform['<%= chkString(Request.QueryString("box"),"sqlstring") %>'].value=V2;
    opnform['<%= chkString(Request.QueryString("box"),"sqlstring") %>'].focus();
    self.close();
}
</script>
</head>
<body <%=spThemeShortBodyTag%>>
<table border="0" cellspacing="0" cellpadding="0" align=center>
 <tr>
 <td class="tCellAlt2">
 <table border="0" cellspacing="1" cellpadding="1">
<tr>
<td class="tTitle" colspan="16">
<b><%= txtClrSel %></b>
</td>
</tr>
<tr>
    <td class="tCellAlt1" colspan="16">
    <p>
    <%= txtClkColor %>.<br />&nbsp;
</td>
</tr>
<tr>
<td class="tTitle" colspan="16">
<b><%= txtNamClrsB %></b></font>
</td>
</tr>
<tr>
<%
	tCount = 0
	ShowColorBox "black"
	ShowColorBox "white"
	ShowColorBox "gray"
	ShowColorBox "silver"
	ShowColorBox "red"
	ShowColorBox "green"
	ShowColorBox "blue"
	ShowColorBox "yellow"
	ShowColorBox "purple"
	ShowColorBox "olive"
	ShowColorBox "navy"
	ShowColorBox "aqua"
	ShowColorBox "lime"
	ShowColorBox "maroon"
	ShowColorBox "teal"
	ShowColorBox "fuchsia"
	
	if tCount < 16 then
		Response.Write "<td colspan=""" & (16-tCount) & """ class=""tCellAlt0""></td>" & vbCrlf
	end if
%>
</tr>
<tr>
<td class="tTitle" colspan="16">
<b><%= txtNamClrsA %></b></font>
</td>
</tr>
<tr>
<%
	tCount = 0
	ShowColorBox "darkgreen"
	ShowColorBox "midnightblue"
	ShowColorBox "dimgray"
	ShowColorBox "slategray"
	ShowColorBox "lightslategray"
	ShowColorBox "darkblue"
	ShowColorBox "mediumblue"
	ShowColorBox "darkcyan"
	ShowColorBox "deepskyblue"
	ShowColorBox "darkturquoise"
	ShowColorBox "mediumspringgreen"
	ShowColorBox "springgreen"
	ShowColorBox "aqua"
	ShowColorBox "dodgerblue"
	ShowColorBox "lightseagreen"
	ShowColorBox "forestgreen"
	
	ShowColorBox "seagreen"
	ShowColorBox "darkslategray"
	ShowColorBox "limegreen"
	ShowColorBox "mediumseagreen"
	ShowColorBox "turquoise"
	ShowColorBox "royalblue"
	ShowColorBox "steelblue"
	ShowColorBox "darkslateblue"
	ShowColorBox "mediumturquoise"
	ShowColorBox "indigo"
	ShowColorBox "darkolivegreen"
	ShowColorBox "cadetblue"
	ShowColorBox "cornflowerblue"
	ShowColorBox "mediumaquamarine"
	ShowColorBox "slateblue"
	ShowColorBox "olivedrab"
	
	ShowColorBox "mediumslateblue"
	ShowColorBox "lawngreen"
	ShowColorBox "chartreuse"
	ShowColorBox "aquamarine"
	ShowColorBox "skyblue"
	ShowColorBox "lightskyblue"
	ShowColorBox "blueviolet"
	ShowColorBox "darkred"
	ShowColorBox "darkmagenta"
	ShowColorBox "saddlebrown"
	ShowColorBox "darkseagreen"
	ShowColorBox "lightgreen"
	ShowColorBox "mediumpurple"
	ShowColorBox "darkviolet"
	ShowColorBox "palegreen"
	ShowColorBox "darkorchid"

	ShowColorBox "yellowgreen"
	ShowColorBox "sienna"
	ShowColorBox "brown"
	ShowColorBox "darkgray"
	ShowColorBox "lightblue"
	ShowColorBox "greenyellow"
	ShowColorBox "paleturquoise"
	ShowColorBox "lightsteelblue"
	ShowColorBox "powderblue"
	ShowColorBox "firebrick"
	ShowColorBox "darkgoldenrod"
	ShowColorBox "mediumorchid"
	ShowColorBox "rosybrown"
	ShowColorBox "darkkhaki"
	ShowColorBox "mediumvioletred"
	ShowColorBox "indianred"

	ShowColorBox "peru"
	ShowColorBox "chocolate"
	ShowColorBox "tan"
	ShowColorBox "lightgrey"
	ShowColorBox "thistle"
	ShowColorBox "orchid"
	ShowColorBox "goldenrod"
	ShowColorBox "palevioletred"
	ShowColorBox "crimson"
	ShowColorBox "gainsboro"
	ShowColorBox "plum"
	ShowColorBox "burlywood"
	ShowColorBox "lightcyan"
	ShowColorBox "lavender"
	ShowColorBox "darksalmon"
	ShowColorBox "violet"

	ShowColorBox "palegoldenrod"
	ShowColorBox "lightcoral"
	ShowColorBox "khaki"
	ShowColorBox "aliceblue"
	ShowColorBox "honeydew"
	ShowColorBox "azure"
	ShowColorBox "sandybrown"
	ShowColorBox "wheat"
	ShowColorBox "beige"
	ShowColorBox "whitesmoke"
	ShowColorBox "mintcream"
	ShowColorBox "ghostwhite"
	ShowColorBox "salmon"
	ShowColorBox "antiquewhite"
	ShowColorBox "linen"
	ShowColorBox "lightgoldenrodyellow"

	ShowColorBox "oldlace"
	ShowColorBox "fuchsia"
	ShowColorBox "deeppink"
	ShowColorBox "orangered"
	ShowColorBox "tomato"
	ShowColorBox "hotpink"
	ShowColorBox "coral"
	ShowColorBox "darkorange"
	ShowColorBox "lightsalmon"
	ShowColorBox "orange"
	ShowColorBox "lightpink"
	ShowColorBox "pink"
	ShowColorBox "gold"
	ShowColorBox "peachpuff"
	ShowColorBox "navajowhite"
	ShowColorBox "moccasin"

	ShowColorBox "bisque"
	ShowColorBox "mistyrose"
	ShowColorBox "blanchedalmond"
	ShowColorBox "papayawhip"
	ShowColorBox "lavenderblush"
	ShowColorBox "seashell"
	ShowColorBox "cornsilk"
	ShowColorBox "lemonchiffon"
	ShowColorBox "floralwhite"
	ShowColorBox "snow"
	ShowColorBox "lightyellow"
	ShowColorBox "ivory"

	if tCount < 16 then
		Response.Write "<td colspan=""" & (16-tCount) & """ class=""tCellAlt0""></td>" & vbCrlf
	end if
%>
</tr>
<tr>
<td class="tTitle" colspan="16">
<b><%= txtSafPalette %></b></font>
</td>
</tr>
<tr>
<%
	tCount = 0
	ShowColorBox "#330099"
	ShowColorBox "#6633ff"
	ShowColorBox "#3300cc"
	ShowColorBox "#3300ff"
	ShowColorBox "#9999ff"
	ShowColorBox "#6633cc"
	ShowColorBox "#ffccff"
	ShowColorBox "#9966cc"
	ShowColorBox "#663399"
	ShowColorBox "#330066"
	ShowColorBox "#9933cc"
	ShowColorBox "#cc99ff"
	ShowColorBox "#cc66ff"
	ShowColorBox "#ff99ff"
	ShowColorBox "#9966ff"
	ShowColorBox "#6600cc"

	ShowColorBox "#6600ff"
	ShowColorBox "#9933ff"
	ShowColorBox "#9900ff"
	ShowColorBox "#cc33ff"
	ShowColorBox "#cc00ff"
	ShowColorBox "#ff66ff"
	ShowColorBox "#ff33ff"
	ShowColorBox "#ff00ff"
	ShowColorBox "#9900cc"
	ShowColorBox "#660099"
	ShowColorBox "#cc66cc"
	ShowColorBox "#cc33cc"
	ShowColorBox "#ff99cc"
	ShowColorBox "#ff66cc"
	ShowColorBox "#ff33cc"
	ShowColorBox "#993399"

	ShowColorBox "#cc00cc"
	ShowColorBox "#ff00cc"
	ShowColorBox "#cc0099"
	ShowColorBox "#990099"
	ShowColorBox "#cc99cc"
	ShowColorBox "#996699"
	ShowColorBox "#663366"
	ShowColorBox "#990066"
	ShowColorBox "#cc3399"
	ShowColorBox "#660066"
	ShowColorBox "#ff0099"
	ShowColorBox "#ff3399"
	ShowColorBox "#cc6699"
	ShowColorBox "#330033"
	ShowColorBox "#993366"
	ShowColorBox "#cc3366"

	ShowColorBox "#cc0066"
	ShowColorBox "#ff6699"
	ShowColorBox "#660033"
	ShowColorBox "#ff0066"
	ShowColorBox "#ff3366"
	ShowColorBox "#ffcccc"
	ShowColorBox "#ff9999"
	ShowColorBox "#cc9999"
	ShowColorBox "#cc6666"
	ShowColorBox "#ff6666"
	ShowColorBox "#996666"
	ShowColorBox "#663333"
	ShowColorBox "#993333"
	ShowColorBox "#990033"
	ShowColorBox "#cc0033"
	ShowColorBox "#ff0033"

	ShowColorBox "#ff3333"
	ShowColorBox "#cc3333"
	ShowColorBox "#ff6600"
	ShowColorBox "#ff3300"
	ShowColorBox "#ff6633"
	ShowColorBox "#cc6633"
	ShowColorBox "#660000"
	ShowColorBox "#330000"
	ShowColorBox "#ff0000"
	ShowColorBox "#990000"
	ShowColorBox "#cc3300"
	ShowColorBox "#cc0000"
	ShowColorBox "#996633"
	ShowColorBox "#cc6600"
	ShowColorBox "#ffcc99"
	ShowColorBox "#ff9966"

	ShowColorBox "#663300"
	ShowColorBox "#cc9966"
	ShowColorBox "#996600"
	ShowColorBox "#cc9933"
	ShowColorBox "#cc9900"
	ShowColorBox "#ffcc66"
	ShowColorBox "#ff9933"
	ShowColorBox "#993300"
	ShowColorBox "#ff9900"
	ShowColorBox "#ffcc33"
	ShowColorBox "#ffcc00"
	ShowColorBox "#ffff99"
	ShowColorBox "#ffff66"
	ShowColorBox "#ffff33"
	ShowColorBox "#ffff00"
	ShowColorBox "#cccc00"

	ShowColorBox "#999900"
	ShowColorBox "#999966"
	ShowColorBox "#cccc99"
	ShowColorBox "#ffffcc"
	ShowColorBox "#cccc33"
	ShowColorBox "#cccc66"
	ShowColorBox "#999933"
	ShowColorBox "#666633"
	ShowColorBox "#666600"
	ShowColorBox "#333300"
	ShowColorBox "#ccff00"
	ShowColorBox "#ccff33"
	ShowColorBox "#99cc33"
	ShowColorBox "#99cc00"
	ShowColorBox "#ccff66"
	ShowColorBox "#ccff99"

	ShowColorBox "#99ff00"
	ShowColorBox "#669933"
	ShowColorBox "#336600"
	ShowColorBox "#336633"
	ShowColorBox "#669966"
	ShowColorBox "#66cc66"
	ShowColorBox "#99ff99"
	ShowColorBox "#66ff66"
	ShowColorBox "#339933"
	ShowColorBox "#99cc99"
	ShowColorBox "#99ff66"
	ShowColorBox "#99ff33"
	ShowColorBox "#66cc33"
	ShowColorBox "#66cc00"
	ShowColorBox "#99cc66"
	ShowColorBox "#669900"

	ShowColorBox "#339900"
	ShowColorBox "#66ff33"
	ShowColorBox "#66ff00"
	ShowColorBox "#ccffcc"
	ShowColorBox "#99ffcc"
	ShowColorBox "#66ff99"
	ShowColorBox "#33ff99"
	ShowColorBox "#33ff00"
	ShowColorBox "#33ff33"
	ShowColorBox "#33cc00"
	ShowColorBox "#33cc33"
	ShowColorBox "#33ff66"
	ShowColorBox "#00ff00"
	ShowColorBox "#33cc66"
	ShowColorBox "#006600"
	ShowColorBox "#003300"

	ShowColorBox "#009900"
	ShowColorBox "#00ff33"
	ShowColorBox "#00ff66"
	ShowColorBox "#00ff99"
	ShowColorBox "#00cc66"
	ShowColorBox "#00cc00"
	ShowColorBox "#00cc33"
	ShowColorBox "#009933"
	ShowColorBox "#66cc99"
	ShowColorBox "#339966"
	ShowColorBox "#33cc99"
	ShowColorBox "#006633"
	ShowColorBox "#009966"
	ShowColorBox "#00cc99"
	ShowColorBox "#66ffcc"
	ShowColorBox "#33ffcc"

	ShowColorBox "#00ffcc"
	ShowColorBox "#009999"
	ShowColorBox "#00cccc"
	ShowColorBox "#33cccc"
	ShowColorBox "#003333"
	ShowColorBox "#006666"
	ShowColorBox "#339999"
	ShowColorBox "#66cccc"
	ShowColorBox "#336666"
	ShowColorBox "#669999"
	ShowColorBox "#99cccc"
	ShowColorBox "#ccffff"
	ShowColorBox "#99ffff"
	ShowColorBox "#66ffff"
	ShowColorBox "#33ffff"
	ShowColorBox "#00ffff"

	ShowColorBox "#00ccff"
	ShowColorBox "#66ccff"
	ShowColorBox "#33ccff"
	ShowColorBox "#3399cc"
	ShowColorBox "#006699"
	ShowColorBox "#0099cc"
	ShowColorBox "#0099ff"
	ShowColorBox "#0066cc"
	ShowColorBox "#003399"
	ShowColorBox "#3366cc"
	ShowColorBox "#003366"
	ShowColorBox "#6699ff"
	ShowColorBox "#3366ff"
	ShowColorBox "#3399ff"
	ShowColorBox "#0066ff"
	ShowColorBox "#0033cc"

	ShowColorBox "#336699"
	ShowColorBox "#000033"
	ShowColorBox "#333366"
	ShowColorBox "#666699"
	ShowColorBox "#9999cc"
	ShowColorBox "#333399"
	ShowColorBox "#6666cc"
	ShowColorBox "#ccccff"
	ShowColorBox "#3333ff"
	ShowColorBox "#3333cc"
	ShowColorBox "#6666ff"
	ShowColorBox "#000066"
	ShowColorBox "#000099"
	ShowColorBox "#0000cc"
	ShowColorBox "#0000ff"
	ShowColorBox "#0033ff"

	ShowColorBox "#6699cc"
	ShowColorBox "#99ccff"
	ShowColorBox "#ffffff"
	ShowColorBox "#cccccc"
	ShowColorBox "#999999"
	ShowColorBox "#666666"
	ShowColorBox "#333333"
	ShowColorBox "#000000"

	if tCount < 16 then
		Response.Write "<td colspan=""" & (16-tCount) & """ class=""tCellAlt0""></td>" & vbCrlf
	end if
%>
</tr>
</table>
 </td>
 </tr>
</table>
<%
Sub ShowColorBox(pValue)
	tCount = tCount + 1
	if tCount>16 then
		tCount=1
		Response.Write "</tr>" & vbCrLf
		Response.Write "<tr>" & vbCrLf
	end if
	Response.Write "<td style=""CURSOR:hand"" width=""25"" height=""20"" bgcolor=""" & pValue & """ title=""" & pValue & """ onclick=""JavaScript:retinfo('" & pValue & "')""></td>" & vbCrLf
End Sub
%>
<!--#include file="inc_footer_short.asp" -->
