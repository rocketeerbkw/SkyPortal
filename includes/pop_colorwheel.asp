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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<!--include virtual="config.asp" -->
<!--include virtual="inc_functions.asp" -->

<script type="text/javascript">
<!-- Original:  D10n (biab@iinet.net.au) -->
<!-- Web Site:  http://www.iinet.net.au/~biab -->

<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Begin
addary = new Array();           //red
addary[0] = new Array(0,1,0);   //red green
addary[1] = new Array(-1,0,0);  //green
addary[2] = new Array(0,0,1);   //green blue
addary[3] = new Array(0,-1,0);  //blue
addary[4] = new Array(1,0,0);   //red blue
addary[5] = new Array(0,0,-1);  //red
addary[6] = new Array(255,1,1);
clrary = new Array(360);
for(i = 0; i < 6; i++)
for(j = 0; j < 60; j++) {
clrary[60 * i + j] = new Array(3);
for(k = 0; k < 3; k++) {
clrary[60 * i + j][k] = addary[6][k];
addary[6][k] += (addary[i][k] * 4);
   }
}
function capture() {
if(document.layers) {
layobj = document.layers['cwheel'];
layobj.document.captureEvents(Event.MOUSEMOVE);
layobj.document.onmousemove = moved;
layobj.document.captureEvents(Event.CLICK);
layobj.document.onclick = clicker;
}
else {
layobj = document.all["cwheel"];
layobj.onmousemove = moved;
layobj.onclick = clicker;
   }
}
function clicker(e) {
y = 4 * ((document.layers)?e.layerX:event.offsetX);
x = 4 * ((document.layers)?e.layerY:event.offsetY);
sx = x - 512;
sy = y - 512;
qx = (sx < 0)?0:1;
qy = (sy < 0)?0:1;
q = 2 * qy + qx;
quad = new Array(-180,360,180,0);
xa = Math.abs(sx);
ya = Math.abs(sy);
d = ya * 45 / xa;
if(ya > xa) d = 90 - (xa * 45 / ya);
deg = Math.floor(Math.abs(quad[q] - d));
n = 0;
sx = Math.abs(x - 512);
sy = Math.abs(y - 512);
r = Math.sqrt((sx * sx) + (sy * sy));
if(x == 512 & y == 512) {
c = "000000";
}
else {
for(i = 0; i < 3; i++) {
r2 = clrary[deg][i] * r / 256;
if(r > 256) r2 += Math.floor(r - 256);
if(r2 > 255) r2 = 255;
n = 256 * n + Math.floor(r2);
}
c = n.toString(16);
while(c.length < 6) c = "0" + c;
}
if(document.layers) {
//document.layers["wheel"].document.f.t.value = "#" + c;
//document.layers["wheel"].bgColor = "#" + c;
    //opnform=window.opener.document.forms['Form1'];
    opnform=window.opener.document.forms['<%= Request.QueryString("form") %>'];
    opnform['<%= Request.QueryString("box") %>'].value= "#" + c;
    opnform['<%= Request.QueryString("box") %>'].focus();
    self.close();
}
else {
//document.all["wheel"].document.f.t.value = "#" + c;
//document.all["wheel"].style.backgroundColor = "#" + c;
    //opnform=window.opener.document.forms['Form1'];
    opnform=window.opener.document.forms['<%= Request.QueryString("form") %>'];
    opnform['<%= Request.QueryString("box") %>'].value= "#" + c;
    opnform['<%= Request.QueryString("box") %>'].focus();
    self.close();
}
return true;
}

function moved(e) {
y = 4 * ((document.layers)?e.layerX:event.offsetX);
x = 4 * ((document.layers)?e.layerY:event.offsetY);
sx = x - 512;
sy = y - 512;
qx = (sx < 0)?0:1;
qy = (sy < 0)?0:1;
q = 2 * qy + qx;
quad = new Array(-180,360,180,0);
xa = Math.abs(sx);
ya = Math.abs(sy);
d = ya * 45 / xa;
if(ya > xa) d = 90 - (xa * 45 / ya);
deg = Math.floor(Math.abs(quad[q] - d));
n = 0;
sx = Math.abs(x - 512);
sy = Math.abs(y - 512);
r = Math.sqrt((sx * sx) + (sy * sy));
if(x == 512 & y == 512) {
c = "000000";
}
else {
for(i = 0; i < 3; i++) {
r2 = clrary[deg][i] * r / 256;
if(r > 256) r2 += Math.floor(r - 256);
if(r2 > 255) r2 = 255;
n = 256 * n + Math.floor(r2);
}
c = n.toString(16);
while(c.length < 6) c = "0" + c;
}
if(document.layers) {
document.layers["wheel"].document.f.t.value = "#" + c;
document.layers["wheel"].bgColor = "#" + c;
}
else {
document.all["wheel"].document.f.t.value = "#" + c;
document.all["wheel"].style.backgroundColor = "#" + c;
}
return false;
}

function retinfo(V2)
{
}
//  End -->
</script>

</HEAD>
<BODY onLoad="capture()">
<div id=cwheel style="position:absolute; visibility:visible; top:40px; left:50px; background-image: url(colorwheel.jpg);">
<table width="256" height="256" border=0 cellpadding=0 cellspacing=0>
<tr>
<td>
</td>
</tr>
</table>
</div>
<div id=wheel style="position:absolute; visibility:visible; top:300px; left:50px;">
<table width="256" height="50" border=0 cellpadding=0 cellspacing=0>
<tr>
<td>

</td>
</tr>
<tr>
<td align="center">
<br />
<form name="f">
<input type="text" name="t" size=27>
</form>
</td>
</tr>
</table>
</div>
<!-- Free JavaScripts provided by The JavaScript Source -->
</BODY>
</HTML>
