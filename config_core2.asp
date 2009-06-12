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

'###############################################################
'## Do Not Edit Below This Line - 
'## You could destroy your database and lose data
'###############################################################
strTablePrefix = "PORTAL_"
strMemberTablePrefix = "PORTAL_"
shoBlkTimer = false
Set oSpData = New SqlCache
oSpData.ConnString = strConnString
if bFso then
  set oSFS = New clsSFSO
end if

dim strSiteTitle, strCopyright, strTitleImage, strHomeURL, strWebSiteVersion
dim strImgComp, strUploadComp, strXmlHttpComp, strDotNetResizeURL, bOnlineUsers
dim strEmail,strEmailVal,strUniqueEmail,strMailMode
dim strMailServer,strSender
dim strMailServerLogon,strMailServerPassword,strMailServerPort
dim strAuthType, strIPLogging, strHeaderType
dim strDateType, strTimeAdjust, strTimeType, strCurDate
dim strCurDateAdjust, strCurDateString, strMTimeAdjust
dim strMTimeType, strMCurDateAdjust, strMCurDateString
dim strBadWordFilter, strBadWords
dim strMoveTopicMode, strPrivateForums, strShowModerators, strShowRank, strAllowForumCode, strAllowHTML
dim intHotTopicNum, strLockDown, strHotTopic
dim strIMGInPosts, strEditedByDate, strForumStatus
dim strHomepage, strICQ, strAIM, strIcons
dim strRankAdmin, strRankMod
dim strRankLevel0, strRankLevel1, strRankLevel2, strRankLevel3, strRankLevel4, strRankLevel5
dim intRankLevel0, intRankLevel1, intRankLevel2, intRankLevel3, intRankLevel4, intRankLevel5
dim strShowStatistics, strLogonForMail, strShowPaging, strPageSize, strPageNumberSize
dim strNTGroupsSTR, strPollCreate, strFeaturedPoll
dim strNewReg, pEnPrefix, blnSetup, my_Conn, strChkDate
dim counter, strDBNTSQLName, strDBNTFUserName, strDBNTUserName
dim strFloodCheck, strFloodCheckTime, strTimeLimit, strNavIcons, sysDebugMode
dim strMSN, strDefTheme, strAllowUploads, strPMtype
dim strQuickReply, strForumSubscription
dim StrIPGateBan ,StrIPGateLck ,StrIPGateCok ,StrIPGateMet ,StrIPGateMsg ,StrIPGateLog ,StrIPGateTyp
dim StrIPGateExp, StrIPGateCss, strIPGateVer, StrIPGateLkMsg, strIPGateNoAcMsg, StrIPGateWarnMsg
dim strYAHOO, strFullName, strPicture, stMx, strSex, strCity, strState, strAge, strCountry, strOccupation
dim strBio, strHobbies, strLNews, strQuote, strMarStatus, strFavLinks, strRecentTopics, strAllowHideEmail
dim strRankColorAdmin, strRankColorMod, strRankColor0, strRankColor1, strRankColor2, strRankColor3
dim strRankColor4, strRankColor5, strNTGroups, strAutoLogon, strVar1, strVar2, strVar3, strVar4, strZip
dim strLoginType, browserReq, varBrowser, memID, isMAC, SecImage, dbHits, intMemberLCID, intPortalLCID
dim mLev, strLoginStatus, strSiteOwner, strUserEmail, PMaccess, strUserMemberID
dim left_Col,maint_Col,mainb_Col,right_Col
dim cont,bLeft,bMaint,bMainb,bRight
Dim arg1, arg2, arg3, arg4, arg5, arg6 'page breadcrumb variables
dim arrCurOnline(), arrGroups(), arrAppPerms()
		
  cont = 0
  bLeft = false
  bMaint = false
  bMainb = false
  bRight = false
intSkin=1
intIsSuperAdmin = 0
sysDebugMode = 0
pr=1
stMx = "Sk"
pEnPrefix = ""
strWebMaster = lcase(strWebMaster)
bOnlineUsers = false

  ':: parse the invalid username characters.
  if strInvalidUsernameChars <> "" then
    if right(strInvalidUsernameChars,1) <> "," then
      strInvalidUsernameChars = strInvalidUsernameChars & ","
	end if
    if instr(strInvalidUsernameChars,"%") = 0 then
      strInvalidUsernameChars = strInvalidUsernameChars & "%" & ">,<%,"
    end if
  else
    strInvalidUsernameChars = "%" & ">,<%,"
  end if
  strInvalidUsernameChars = strInvalidUsernameChars & """,',;,:,#,*"

  strDotNetResizeURL = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
  LastPath = InStrRev(strDotNetResizeURL,"/")
  if LastPath > 0 then
    strDotNetResizeURL = left(strDotNetResizeURL,Lastpath)
  end if
  strDotNetResizeURL = strDotNetResizeURL & "includes/image_resizer.aspx"

Session.LCID = intPortalLCID
'Session.LCID = 1033
	on error resume next
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	my_Conn.Open strConnString
	'	Lets check to see if the strConnString or db path has changed
		if my_conn.Errors.Count <> 0 then 
		  'we can't connect, lets log the error
		  for counter = 0 to my_conn.Errors.Count -1
			ConnErrorNumber = my_conn.Errors(counter).Number
			ConnErrorDesc = my_conn.Errors(counter).Description
			if ConnErrorNumber <> 0 and ConnErrorNumber <> -2147217887 then 
			  writeToLog "Database","",ConnErrorNumber & " : " & ConnErrorDesc
			end if
		  next
		  if ConnErrorNumber <> -2147217887 then
			my_conn.Errors.Clear 
			set my_Conn = nothing
			on error goto 0
			Response.Redirect "site_setup.asp?RC=1"
		  end if
		end if
	on error goto 0

'blnSetup="N"
if request.QueryString("sky") = "dogg" then
  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end if

	'FileSystemObject check
	if bFso then
	 on error resume next
	 Err.Clear
	 set fso = Server.CreateObject("Scripting.FileSystemObject")
	 if err.number = 0 then
	   bFso = true
	 end if
	 set fso = nothing
	 Err.Clear
	 on error goto 0
	end if
	
	strXmlHttp = "none"
	if bXmlHttp then
	  strXmlHttp = DetectXmlHttp()
	end if
	'if bAspNet then
	  'bAspNet = CheckAspNet(strXmlHttp)
	'end if
	
if blnSetup<>"Y" then

  if stMx = "Sk" then 
	strSql = "SELECT C_STRSITETITLE "
	strSql = strSql & ", C_STRCOPYRIGHT "
	strSql = strSql & ", C_STRTITLEIMAGE "
	strSql = strSql & ", C_STRHOMEURL "
	strSql = strSql & ", C_STRAUTHTYPE "
	
	strSql = strSql & ", C_STREMAIL "
	strSql = strSql & ", C_STRUNIQUEEMAIL "
	strSql = strSql & ", C_STRMAILMODE "
	strSql = strSql & ", C_STRMAILSERVER "
	strSql = strSql & ", C_STREMAILUSERNAME "
	strSql = strSql & ", C_STREMAILPASSWORD "
	strSql = strSql & ", C_STREMAILPORT "
	strSql = strSql & ", C_STRSENDER "
	
	strSql = strSql & ", C_STRDATETYPE "
	strSql = strSql & ", C_STRTIMEADJUST "
	strSql = strSql & ", C_STRTIMETYPE "
	strSql = strSql & ", C_STRALLOWFORUMCODE "
	strSql = strSql & ", C_STRALLOWHTML "
	strSql = strSql & ", C_STRNTGROUPS"
	strSql = strSql & ", C_STRAUTOLOGON"
	strSql = strSql & ", C_STRBADWORDFILTER "
	strSql = strSql & ", C_STRBADWORDS "
	strSql = strSql & ", C_STRSIGNATURES "
	strSql = strSql & ", C_STRLOGONFORMAIL "
	strSql = strSql & ", C_STREMAILVAL"
	strSql = strSql & ", C_STRFLOODCHECK"
	strSql = strSql & ", C_STRFLOODCHECKTIME"
	strSql = strSql & ", C_STRNEWREG"
	strSql = strSql & ", C_STRDEFTHEME"
	strSql = strSql & ", C_ALLOWUPLOADS "
	strSql = strSql & ", C_PMTYPE"
	strSql = strSql & ", C_STRHEADERTYPE"
	strSql = strSql & ", C_STRLOGINTYPE"
	strSql = strSql & ", C_SECIMAGE"
	strSql = strSql & ", C_INTSUBSKIN"
	strSql = strSql & ", C_ONEADAYDATE"
	strSql = strSql & ", C_COMP_IMAGE"
	strSql = strSql & ", C_COMP_UPLOAD"
	
	strSql = strSql & ", C_STRMOVETOPICMODE "
	strSql = strSql & ", C_STRIPLOGGING "
	strSql = strSql & ", C_STRPRIVATEFORUMS "
	strSql = strSql & ", C_STRSHOWMODERATORS "
	strSql = strSql & ", C_STRHOTTOPIC "
	strSql = strSql & ", C_INTHOTTOPICNUM "
	strSql = strSql & ", C_STRIMGINPOSTS "
	strSql = strSql & ", C_STRICONS "
	strSql = strSql & ", C_STREDITEDBYDATE "
	strSql = strSql & ", C_STRSHOWSTATISTICS "
	strSql = strSql & ", C_STRSHOWPAGING "
	strSql = strSql & ", C_STRPAGESIZE "
	strSql = strSql & ", C_STRPAGENUMBERSIZE "
	strSql = strSql & ", C_STRLOCKDOWN"
	strSql = strSql & ", C_STRLNEWS"
	strSql = strSql & ", C_STRMARSTATUS"
	strSql = strSql & ", C_STRFAVLINKS"
	strSql = strSql & ", C_STRRECENTTOPICS"
	strSql = strSql & ", C_STRHOMEPAGE"
	strSql = strSql & ", C_FORUMSTATUS"
	strSql = strSql & ", C_POLLCREATE"
	strSql = strSql & ", C_FEATUREDPOLL"
	strSql = strSql & ", C_STRQUICKREPLY"
    strSql = strSql & ", C_FORUMSUBSCRIPTION"
	strSql = strSql & ", C_STRSHOWRANK "
	strSql = strSql & ", C_STRRANKADMIN "
	strSql = strSql & ", C_STRRANKMOD "
	strSql = strSql & ", C_STRRANKLEVEL0 "
	strSql = strSql & ", C_STRRANKLEVEL1 "
	strSql = strSql & ", C_STRRANKLEVEL2 "
	strSql = strSql & ", C_STRRANKLEVEL3 "
	strSql = strSql & ", C_STRRANKLEVEL4 "
	strSql = strSql & ", C_STRRANKLEVEL5 "
	strSql = strSql & ", C_STRRANKCOLORADMIN "
	strSql = strSql & ", C_STRRANKCOLORMOD "
	strSql = strSql & ", C_STRRANKCOLOR0 "
	strSql = strSql & ", C_STRRANKCOLOR1 "
	strSql = strSql & ", C_STRRANKCOLOR2 "
	strSql = strSql & ", C_STRRANKCOLOR3 "
	strSql = strSql & ", C_STRRANKCOLOR4 "
	strSql = strSql & ", C_STRRANKCOLOR5 "
	strSql = strSql & ", C_INTRANKLEVEL0 "
	strSql = strSql & ", C_INTRANKLEVEL1 "
	strSql = strSql & ", C_INTRANKLEVEL2 "
	strSql = strSql & ", C_INTRANKLEVEL3 "
	strSql = strSql & ", C_INTRANKLEVEL4 "
	strSql = strSql & ", C_INTRANKLEVEL5 "
	
	strSql = strSql & ", C_STRVAR1"
    strSql = strSql & ", C_STRVAR2"
    strSql = strSql & ", C_STRVAR3"
    strSql = strSql & ", C_STRVAR4"
	strSql = strSql & ", C_STRFULLNAME"
	strSql = strSql & ", C_STRPICTURE"
	strSql = strSql & ", C_STRSEX"
	strSql = strSql & ", C_STRCITY"
	strSql = strSql & ", C_STRSTATE"
	strSql = strSql & ", C_STRZIP"
	strSql = strSql & ", C_STRAGE"
	strSql = strSql & ", C_STRCOUNTRY"
	strSql = strSql & ", C_STROCCUPATION"
	strSql = strSql & ", C_STRBIO"
	strSql = strSql & ", C_STRHOBBIES"
	strSql = strSql & ", C_STRQUOTE"
	strSql = strSql & ", C_STRHOMEPAGE "
	strSql = strSql & ", C_STRICQ "
	strSql = strSql & ", C_STRYAHOO "
	strSql = strSql & ", C_STRAIM "
	strSql = strSql & ", C_STRMSN "
	
	' # added for IPGATE Mod
	strSql = strSql & ", C_STRIPGATEBAN"
	strSql = strSql & ", C_STRIPGATELCK"
	strSql = strSql & ", C_STRIPGATECOK"
	strSql = strSql & ", C_STRIPGATEMET"
	strSql = strSql & ", C_STRIPGATEMSG"
	strSql = strSql & ", C_STRIPGATELOG"
	strSql = strSql & ", C_STRIPGATETYP"
	strSql = strSql & ", C_STRIPGATEEXP"
	strSql = strSql & ", C_STRIPGATECSS"
	strSql = strSql & ", C_STRIPGATEVER"
	strSql = strSql & ", C_STRIPGATELKMSG"
	strSql = strSql & ", C_STRIPGATENOACMSG"
	strSql = strSql & ", C_STRIPGATEWARNMSG"
	strSql = strSql & " FROM " & strTablePrefix & "CONFIG "
	strSql = strSql & " WHERE CONFIG_ID = 1"
	
    if Application(strCookieURL & strUniqueID & "ConfigLoaded")= "" or IsNull(Application(strCookieURL & strUniqueID & "ConfigLoaded")) then
	  
	  oSpData.RemoveAll()	
      Application.Lock
	  application.Contents.RemoveAll()
	  Application(strCookieURL & strUniqueID & "ConfigLoaded")= "YES"
	  Application.UnLock
	  ':: force rebuild of the APPS
	  'tSql = "SELECT * FROM "& strTablePrefix & "APPS"
	  'oSpData.Remove(tSql)
    end if 
	
	on error resume next
	  err.clear()
	set rsCfg = my_Conn.Execute(strSql)
	'set rsCfg = oSpData.GetRecordset(strSql)
	'	Lets check to see if the strConnString or db path has changed
		if err.number <> 0 then
	  		err.clear()
			set my_Conn = nothing
			'Response.Redirect "site_setup.asp?err=no_config_table"		
		end if
	on error goto 0


	strSiteTitle = replace(rsCfg("C_STRSITETITLE"),"''","'")
	strCopyright = rsCfg("C_STRCOPYRIGHT")
	strTitleImage = rsCfg("C_STRTITLEIMAGE")
	strHomeURL = rsCfg("C_STRHOMEURL")
	strAuthType = rsCfg("C_STRAUTHTYPE")
	strUniqueEmail = rsCfg("C_STRUNIQUEEMAIL")
	strIPLogging = rsCfg("C_STRIPLOGGING")
	strIMGInPosts = rsCfg("C_STRIMGINPOSTS")
	strAllowHTML = rsCfg("C_STRALLOWHTML")
	strAllowForumCode = rsCfg("C_STRALLOWFORUMCODE")
	strBadWordFilter = rsCfg("C_STRBADWORDFILTER")
	strBadWords = rsCfg("C_STRBADWORDS")
	strLockDown = rsCfg("C_STRLOCKDOWN")
	
	strDateFormat = getDateFormat()
	strDateType = rsCfg("C_STRDATETYPE")
	strTimeAdjust = rsCfg("C_STRTIMEADJUST")
	strTimeType = rsCfg("C_STRTIMETYPE")
	
	strCurDateAdjust = DateAdd("h", strTimeAdjust , Now()) 'portal offset from server
	strCurDateString = DateToStr(strCurDateAdjust)
	strCurDate = ChkDate2(strCurDateString)
	
	strMTimeAdjust = 0
	strMTimeType = strDateType
	strMCurDateAdjust = strCurDateAdjust
	strMCurDateString = strCurDateString
	intMemberLCID = intPortalLCID
	
	strNTGroups = rsCfg("C_STRNTGROUPS")
	strAutoLogon = rsCfg("C_STRAUTOLOGON")
	
	strEmail = rsCfg("C_STREMAIL")
	strMailMode = rsCfg("C_STRMAILMODE")
	strMailServer = rsCfg("C_STRMAILSERVER")
	strMailServerLogon = rsCfg("C_STREMAILUSERNAME")
	strMailServerPassword = rsCfg("C_STREMAILPASSWORD")
	strMailServerPort = rsCfg("C_STREMAILPORT")
	strSender = rsCfg("C_STRSENDER")
	strLogonForMail = rsCfg("C_STRLOGONFORMAIL")
	strEmailVal = rsCfg("C_STREMAILVAL")
	
	strFloodCheck = rsCfg("C_STRFLOODCHECK")
	strFloodCheckTime = rsCfg("C_STRFLOODCHECKTIME")
	strNewReg = rsCfg("C_STRNEWREG")
	strDefTheme = rsCfg("C_STRDEFTHEME")
	strPMtype = rsCfg("C_PMTYPE")
	StrIPGateBan = rsCfg("C_STRIPGATEBAN")
	StrIPGateLck = rsCfg("C_STRIPGATELCK")
	StrIPGateCok = rsCfg("C_STRIPGATECOK")
	StrIPGateMet = rsCfg("C_STRIPGATEMET")
	StrIPGateMsg = rsCfg("C_STRIPGATEMSG")
	StrIPGateLog = rsCfg("C_STRIPGATELOG")
	StrIPGateTyp = rsCfg("C_STRIPGATETYP")
	StrIPGateExp = rsCfg("C_STRIPGATEEXP")
	StrIPGateCss = rsCfg("C_STRIPGATECSS")
	strIPGateVer = rsCfg("C_STRIPGATEVER")
	StrIPGateLkMsg = rsCfg("C_STRIPGATELKMSG")
	strIPGateNoAcMsg = rsCfg("C_STRIPGATENOACMSG")
	StrIPGateWarnMsg = rsCfg("C_STRIPGATEWARNMSG")
	
	strAllowHideEmail = "1"
	stWb = "yPor"
	strHeaderType = rsCfg("C_STRHEADERTYPE")
	strLoginType = rsCfg("C_STRLOGINTYPE")
	SecImage = rsCfg("C_SECIMAGE")
	intSubSkin = rsCfg("C_INTSUBSKIN")
	strChkDate = rsCfg("C_ONEADAYDATE")
	strImgComp = rsCfg("C_COMP_IMAGE")
	strAllowUploads = rsCfg("C_ALLOWUPLOADS")
	strUploadComp = rsCfg("C_COMP_UPLOAD")
	
	strXmlHttpComp = strXmlHttp
	':: build array for app configuration
	bldArrAppAccess()
	'Application(strCookieURL & strUniqueID & "strAppVars") = bldAppAccess()	
	
	'forums
	strIcons = rsCfg("C_STRICONS")
	strMoveTopicMode = rsCfg("C_STRMOVETOPICMODE")
	strPrivateForums = rsCfg("C_STRPRIVATEFORUMS")
	strShowModerators = rsCfg("C_STRSHOWMODERATORS")
	strHotTopic = rsCfg("C_STRHOTTOPIC")
	intHotTopicNum = rsCfg("C_INTHOTTOPICNUM")
	strEditedByDate = rsCfg("C_STREDITEDBYDATE")
	strShowRank = rsCfg("C_STRSHOWRANK")
	strRankAdmin = rsCfg("C_STRRANKADMIN")
	strRankMod = rsCfg("C_STRRANKMOD")
	strRankLevel0 = rsCfg("C_STRRANKLEVEL0")
	strRankLevel1 = rsCfg("C_STRRANKLEVEL1")
	strRankLevel2 = rsCfg("C_STRRANKLEVEL2")
	strRankLevel3 = rsCfg("C_STRRANKLEVEL3")
	strRankLevel4 = rsCfg("C_STRRANKLEVEL4")
	strRankLevel5 = rsCfg("C_STRRANKLEVEL5")
	strRankColorAdmin = rsCfg("C_STRRANKCOLORADMIN")
	strRankColorMod = rsCfg("C_STRRANKCOLORMOD")
	strRankColor0 = rsCfg("C_STRRANKCOLOR0")
	strRankColor1 = rsCfg("C_STRRANKCOLOR1")
	strRankColor2 = rsCfg("C_STRRANKCOLOR2")
	strRankColor3 = rsCfg("C_STRRANKCOLOR3")
	strRankColor4 = rsCfg("C_STRRANKCOLOR4")
	strRankColor5 = rsCfg("C_STRRANKCOLOR5")
	intRankLevel0 = rsCfg("C_INTRANKLEVEL0")
	intRankLevel1 = rsCfg("C_INTRANKLEVEL1")
	intRankLevel2 = rsCfg("C_INTRANKLEVEL2")
	intRankLevel3 = rsCfg("C_INTRANKLEVEL3")
	intRankLevel4 = rsCfg("C_INTRANKLEVEL4")
	intRankLevel5 = rsCfg("C_INTRANKLEVEL5")
	strShowStatistics = rsCfg("C_STRSHOWSTATISTICS")
	strShowPaging = rsCfg("C_STRSHOWPAGING")
	strPageSize = rsCfg("C_STRPAGESIZE")
	strPageNumberSize = rsCfg("C_STRPAGENUMBERSIZE")
	strForumStatus = rsCfg("C_FORUMSTATUS")
	strPollCreate = rsCfg("C_POLLCREATE")
	strFeaturedPoll = rsCfg("C_FEATUREDPOLL")
	strQuickReply = rsCfg("C_STRQUICKREPLY")
	strForumSubscription = rsCfg("C_FORUMSUBSCRIPTION")
	
	'member stuff
	strFullName = rsCfg("C_STRFULLNAME")
	strPicture = rsCfg("C_STRPICTURE")
	strMarStatus = rsCfg("C_STRMARSTATUS")
	strAge = rsCfg("C_STRAGE")
	strSex = rsCfg("C_STRSEX")
	strCity= rsCfg("C_STRCITY")
	strState = rsCfg("C_STRSTATE")
	strZip = rsCfg("C_STRZIP")
	strCountry = rsCfg("C_STRCOUNTRY")
	strICQ = rsCfg("C_STRICQ")
	strYAHOO = rsCfg("C_STRYAHOO")
	strAIM = rsCfg("C_STRAIM")
	strMSN = rsCfg("C_STRMSN")
	strHomepage = rsCfg("C_STRHOMEPAGE")
	strOccupation = rsCfg("C_STROCCUPATION")
	strBio = rsCfg("C_STRBIO")
	strHobbies = rsCfg("C_STRHOBBIES") 
	strLNews = rsCfg("C_STRLNEWS") 
	strQuote = rsCfg("C_STRQUOTE")
	strFavLinks = rsCfg("C_STRFAVLINKS")
	strRecentTopics = rsCfg("C_STRRECENTTOPICS")
	strVar1 = rsCfg("C_STRVAR1")
	strVar2 = rsCfg("C_STRVAR2")
	strVar3 = rsCfg("C_STRVAR3")
	strVar4 = rsCfg("C_STRVAR4")
	intSubSkin = rsCfg("C_INTSUBSKIN")
	'intSubSkin = 0
	set rsCfg = nothing
		

	':: config maintenance
	if not bFso then
		intUploads = 0
  		strAllowUploads = 0
		bAspNet = false
		strUploadComp = "none"
	end if
	if strUploadComp = "none" then
		'intUploads = 0
  		strAllowUploads = 0
	end if
	if intUploads = 0 then
  		strAllowUploads = 0
		strUploadComp = "none"
	end if
	if not bXmlHttp then
	  strXmlHttpComp = "none"
	end if
	if strXmlHttpComp = "none" then
	  bAspNet = false
	end if
	if strEmail = 0 then
	  intSubscriptions = 0
	end if

	if strAuthType = "db" then
		strDBNTSQLName = "M_NAME"
		strAutoLogon ="0"
		strNTGroups  ="0"
	else
		strDBNTSQLName = "M_USERNAME"
	end if

'::::::::::::: browser sniffer code. ::::::::::::::::::::::::
	browserReq = request.ServerVariables("HTTP_USER_AGENT")
	varBrowser = ""
   		if instr(lcase(browserReq),"msie") <> 0 then ' Is MSIE browser
			varBrowser = "ie"
		elseif instr(lcase(browserReq),"firefox") <> 0 then ' FireFox
			varBrowser = "firefox"
		elseif instr(lcase(browserReq),"opera") <> 0 then ' Opera
			varBrowser = "opera"
		elseif instr(lcase(browserReq),"firebird") <> 0 then ' Firebird
			varBrowser = "firebird"
		elseif instr(lcase(browserReq),"safari") <> 0 then ' Safari
			varBrowser = "safari"
		elseif instr(lcase(browserReq),"lynx") <> 0 then ' lynx
			varBrowser = "lynx"
		elseif instr(lcase(browserReq),"camino") <> 0 then ' Safari
			varBrowser = "camino"
		elseif instr(lcase(browserReq),"gecko") <> 0 then ' Netscape
			varBrowser = "netscape"
		else
			varBrowser = "other"	
		end if
			isMAC = false
			stPl = "tal."
		
		':: This code detects if the editor is browser compatable
		':: If not, then change settings to show the default [forum code] editor.
		'if (instr(lcase(browserReq),"mac") <> 0 and varBrowser <> "firefox") or varBrowser = "opera" then
		if instr(lcase(browserReq),"mac") <> 0 then
		  if varBrowser = "ie" then
			strAllowHtml = 0
			strAllowForumCode = 1
			strIMGInPosts = 1
		  end if
			isMAC = true
		end if

	':: Check the default theme to use on the site
	If trim(strDefTheme) = "" or isNull(strDefTheme) Then
  		strDefTheme = installTheme
	end if
	strTheme = strDefTheme
  end if 
end if
'response.Write("blnSetup: " & blnSetup)

if blnSetup <> "Y" then

 if not uploadPg then
  if request("lang") <> "" then
    strLang = request("lang")
  end if
 end if
'if bFso then
  'include server.mappath("lang/" & strLang & "/core.asp")
'else 
'end if

  ':: set last here date
  sLastHereDate = readSession("last_here_date")
  if IsEmpty(sLastHereDate) or IsNull(sLastHereDate) or trim(sLastHereDate) = "" then
    lhdName = ChkString(Request.Cookies(strUniqueID & "User")("Name"), "SQLString")
    Call setSession("last_here_date",ReadLastHereDate(lhdName))
  end if
sLastHereDate = readSession("last_here_date")
end if
  %><!-- #include file="lang/en/core.asp" --><%

  sScript = request.ServerVariables("SCRIPT_NAME")
  if instr(sScript,"/") > 0 then
    sScript = mid(sScript,instrrev(sScript,"/")+1,len(sScript))
  end if

':::: these functions execute with each page load :::::::
  'check to see if it is a new day and run the once per day routine if needed
  OncePerDayChecks()
  'custom_functions pageload call
  eachPageLoad()
':::: end page load functions :::::::::::::::::::::::::::

function DetectXmlHttp()
  Dim sTempComponent
	sTempComponent = "none"	
	if CheckXmlHttpComp("Microsoft.XMLHTTP") then
	  'sTempComponent = "DOTNET3"
	  sTempComponent = "Microsoft.XMLHTTP"
	else
	  if CheckXmlHttpComp("Msxml2.ServerXMLHTTP") then
		'sTempComponent = "DOTNET2"
		sTempComponent = "Msxml2.ServerXMLHTTP"
	  else
		if CheckXmlHttpComp("Msxml2.ServerXMLHTTP.4.0") then
		  'sTempComponent = "DOTNET1"
		  sTempComponent = "Msxml2.ServerXMLHTTP.4.0"
		else
		  'sTempComponent = "NOT FOUND: ASP.NET Server Component<br>"
		end if
	  end if
	end if
	DetectXmlHttp = sTempComponent
end function

function CheckAspNet(XmlHttpObj)
  dim objHttp, Detected
	Detected = false
  on error resume next
  err.clear
  Set objHttp = Server.CreateObject(XmlHttpObj)
  if err.number = 0 then
    objHttp.open "GET", strDotNetResizeURL, false
	if err.number = 0 then
      objHttp.Send ""
	  if (objHttp.status <> 200 ) then
		sErr = "An error has accured with ASP.NET component " & XmlHttpObj & vbcrlf
		sErr = sErr & "Error: " & objHttp.responseText
		writeToLog "AspNet","",sErr
		'Response.End
	  end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detected = true
      end if
	end if
  End if
  Set objHttp = nothing
  on error goto 0
  CheckAspNet = Detected
end function

function CheckXmlHttpComp(XmlHttpObj)
  dim objHttp, Detected
	Detected = false
  on error resume next
  err.clear
  Set objHttp = Server.CreateObject(XmlHttpObj)
  if err.number = 0 then
	Detected = true
  End if
  Set objHttp = nothing
  on error goto 0
  CheckXmlHttpComp = Detected
end function

	if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - startTime),3)
	  response.Write("config_core bottom: " & blkLoadTime & "<br>")
	end if
%>