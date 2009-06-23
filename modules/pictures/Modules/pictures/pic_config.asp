<%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'::::::::::::::::::::::::::::::::
'::: Config for MASS UPLOAD mod
'::::::::::::::::::::::::::::::::

'::: Paths from your SITE ROOT
strDirPath = "files/pictures"
strImgPath = "files/pictures/temp/"

':: path to your TEMP directory for uploads
':: this is the path from the Domain ROOT
BasePath = "/files/pictures/temp"
strUpDirPath = "/files/pictures/temp"

':: path to your SITE IMGES directory
':: this is the path from the Domain ROOT
strImgDir = "/images/icons/"

':: Select the image component for resizing.
'strImgComp = "aspjpeg"
'strImgComp = "aspimage"
strImgComp = "aspnet"
'strImgComp = "none"

':: Maxumim WIDTH, HEIGHT and SIZE for uploading to TEMP directory
iMaxWidth = 800
iMaxHeight = 800
iMaxSize = 2000000  ':: in bytes

physTmpDirPath = Server.MapPath(BasePath)
physDirPath = Server.MapPath(strDirPath) & "\"
%>