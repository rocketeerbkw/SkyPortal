
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
Response.Buffer = true
Session.Timeout = 20
Session.CodePage = "65001"
%>
<!-- #include file="includes/classes/clsMenu.asp" -->
<!-- #include file="includes/classes/includes.asp" -->
<!-- #include file="includes/classes/clsSPFS.asp" -->
<!-- #include file="includes/classes/clsData.asp" -->
<!-- include file="modules/rss/clsRSS.asp" -->
<%
dim startTime : startTime = timer
dim pageTimer, strDBType, strConnString, strTablePrefix, strMemberTablePrefix, strTheme, strWebMaster
dim intDisplay, bFso, bAspNet, strCharset
dim intAllowed, strCookieURL, strCurSymbol, showGames, showGold, showRep, sqlver
dim intMyMax, strUnicode, strUniqueID, intBookmarks, intSubscriptions

strCharset = "utf-8"
strPortalTimeZone = "EST"
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: SELECT YOUR DATABASE TYPE AND CONNECTION TYPE (access, sqlserver)
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
strDBType = "access"
'strDBType = "sqlserver"

	'## if you require unicode language support uncomment "YES" the line below
	'## and comment out the NO line. Unicode support is required for languages that use a different alphabet
	'## for more info see http://www.unicode.org/standard/WhatIsUnicode.html
	'## Access database is unicode by default and as such the variable will not be used
	strUnicode="NO"
	'strUnicode="YES"

'::: Provide the full path to your Access database here.
'::: Please rename your database. Do not use the database names below.
'strDBPath = "C:\Domains\your_folder\wwwroot\db\db_name.mdb"
strDBPath = server.MapPath("db/sp_db2k6.mdb")

'::: Choose one of the connection strings below
'::: The string directly below is for an Access 2000 DB. 
'::: Do nothing if you are using Access
strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&strDBPath

'::: If you are using SQL Server, Comment out the line above and uncomment
'::: the line below and fill in the correct connection variables
 
'strConnString = "Provider=SQLOLEDB;Data Source=SQL_server_name_or_IP;Initial Catalog=db_name_here;UID=db_user_name_here;PWD=db_password_here" 'SQL Server

'strConnString = "Driver={MySQL ODBC 3.51 Driver};Server=data.domain.com;Port=3306;Database=myDataBase;User=myUsername; Password=myPassword;Option=3;"


'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: strWebMaster is the list of Super Admins. Use lowercase member names.
':: names should always end with a comma
':: strWebMaster = "admin,skydogg,santa claus,"
':: This is also your initial login username.
strWebMaster = "administrator,"

':: Set the portal LCID 
':: the default of ENGLISH-USA is set
intPortalLCID = 1033

':: strUniqueID is your cookie name prefix. make it unique for your site. Keep it short.
':: If for some reason you need to "log everyone out", 
':: just change the variable below, save and upload.
strUniqueID = "SP10_"

':: If your site is in a VIRTUAL directory, 
':: or if you have problems logging in, uncomment
':: the line below and comment out the line under it.
'strCookieURL = "/"
strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))

':: Show page load time at bottom of all pages
pageTimer = 1	' 1 = yes; 0 = no   

':: Allow uploads?
':: This will override the database setting of "ON" for allowing uploads
intUploads = 1 ' 0 = OFF; 1 = ON

':: Allow subscriptions?
':: Having your EMAIL turned off will override this value
intSubscriptions = 1 ' 0=off 1=on

':: Allow bookmarks?
intBookmarks = 1 ' 0=off 1=on

':: Allow members to access the myMax feature:
':: If set to '0', only superadmin can arrange the front page
':: layout from using the myMax link in the members menu while
':: looking at the front page.
intMyMax = 0  '0=no 1=yes

'Access for GROUPS who can 'view source' on the editor
' You can add additional group ID's to the variable seperated with a comma
' 0 for super admin only
' 1 for all admins only
' 2 for all members
intEditor = 4 

':: HTML editor language.
':: use 2 letter abbreviation... lower case
strLang = "en"

'Currency Symbol
strCurSymbol = "$"

'What HTML editor is used.
'SkyPortal currently supports FCKeditor and tinyMCE editor
editorType = "tinymce" 
'editorType = "fckeditor" 

'Default member name glow (IE only)
def_glow = "#00FF00:#FFFFFF"

showGold = 0 	' 0=no; 1=yes;
showRep = 0 	' 0=no; 1=yes;
showGames = 0 	' 0=no; 1=yes;

':: installTheme is the default theme when you install the portal.
':: This is the folder name of the theme.
':: This variable also needs changed in site_setup.asp
installTheme = "sp_IceMan"

':: If your server sessions expire too often,
':: change the following vaiable to FALSE
bUseMemberSession = true

':: Set the following values to true or false
':: depending on if your server supports them or not.
':: You can check these by running detect2.asp
bFso = true
bAspNet = true
bXmlHttp = true

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':: Registration Configuration
'::
':: intMinUsernameLength = The minimum length for user names.
':: Cannot be less than 3 characters.
'::
':: intMinimumPasswordLength = The minimum length for passwords.
':: Cannot be less than 3 characters.
'::
':: strInvalidUsernameChars = String of characters not allowed
':: in the Usernames. This always includes some characters by default
'::        " ' ; : # *
'::
':: strInvalidIPs = A comma delimited list of IPs that are not
':: able to register. This will also track subnets and partial IPs.
':: Do not use wildcards!
'::
':: showRegisterLongForm = false:Short form   true:Long form
'::
intMinUsernameLength = 4
intMinimumPasswordLength = 5
strInvalidUsernameChars = "<,>,&,(,),{,},+,=,%,$,!,`,~,_,-,|,\,/,?,^"
strInvalidIPs = ""
showRegisterLongForm = false


'::<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
'::<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><
':: 
':: Set the variables below if you are going to use Active Directory
':: as the Authorization Type. For INTRANET use only.
'::

Const sADnetbiosDomain = ""		' The AD domain name
Const sADusername = ""		    ' The account that connects to AD
Const sADpassword = ""		    ' The password for the above account

%>
<!-- #include file="includes/inc_ADOVBS.asp" -->
<!-- #include file="config_core2.asp" -->