<!--#include file="rand.asp"-->
<!--#include file="canvas.asp"-->
<!--#include file="inc_encryption.asp"-->
<%
'Make sure this page is not cached
Response.Expires = -10000
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"

Dim objCanvas	'Hold Object of Canvas Class
Dim strSecuritycode
strSecuritycode = RandomPW(5)
Session("secCode")=pEncrypt(strSecuritycode)
'Session("secCode2")=strSecuritycode
'Create Object of Canvas Class
Set objCanvas = New Canvas

objCanvas.GlobalColourTable(0) = RGB(255,255,255) ' White
objCanvas.GlobalColourTable(1) = RGB(0,0,0) ' Black
objCanvas.GlobalColourTable(2) = RGB(255,0,0) ' Red
objCanvas.GlobalColourTable(3) = RGB(0,255,0) ' Green
objCanvas.GlobalColourTable(4) = RGB(0,0,255) ' Blue
objCanvas.GlobalColourTable(5) = RGB(250,236,7) 'Yellow
objCanvas.GlobalColourTable(6) = RGB(12,50,151) 'Dark Bkue
objCanvas.GlobalColourTable(7) = RGB(211,205,205) 'Dark Green
objCanvas.BackgroundColourIndex = 1
objCanvas.Resize 105,35,false ' Resize without preserving the image
'objCanvas.LoadBMP("c:\inetpub\wwwroot\captcha\test.bmp")

'Line(X1,Y1,X2,Y2)
Randomize Time
objCanvas.ForegroundColourIndex = 2 ' Set the pen for lines
objCanvas.Line 0,0,Int(70*Rnd()+10),35
objCanvas.Line 25,0,Int(70*Rnd()+25),35
objCanvas.Line 50,0,Int(70*Rnd()+50),35
objCanvas.Line 105,10,0,Int(35*Rnd())
objCanvas.Line 105,30,0,Int(35*Rnd())

'DrawVectorTextWE(X,Y,Text,Scale)
'DrawVectorTextNS(X,Y,Text,Scale)
objCanvas.ForegroundColourIndex = 0 ' Set the pen for font
objCanvas.DrawVectorTextNS 1,1,"" & Mid(UCase(strSecuritycode),1,1) & "",5
objCanvas.DrawVectorTextNS 2,2,"" & Mid(UCase(strSecuritycode),1,1) & "",5
objCanvas.DrawVectorTextWE 21,10,"" & Mid(UCase(strSecuritycode),2,1) & "",4
objCanvas.DrawVectorTextWE 22,11,"" & Mid(UCase(strSecuritycode),2,1) & "",4
objCanvas.DrawVectorTextWE 41,4,"" & Mid(UCase(strSecuritycode),3,1) & "",5
objCanvas.DrawVectorTextWE 42,5,"" & Mid(UCase(strSecuritycode),3,1) & "",5
objCanvas.DrawVectorTextWE 61,1,"" & Mid(UCase(strSecuritycode),4,1) & "",4
objCanvas.DrawVectorTextWE 62,2,"" & Mid(UCase(strSecuritycode),4,1) & "",4
objCanvas.DrawVectorTextWE 81,6,"" & Mid(UCase(strSecuritycode),5,1) & "",5
objCanvas.DrawVectorTextWE 82,7,"" & Mid(UCase(strSecuritycode),5,1) & "",5

objCanvas.Write ' Write the image to the browser

%>
