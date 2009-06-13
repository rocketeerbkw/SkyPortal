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
dim strDateType, strTimeAdjust, strTimeType, strCurDateAdjust, strCurDate
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
dim counter, strDBNTSQLName
dim strFloodCheck, strFloodCheckTime, strTimeLimit, strNavIcons, sysDebugMode
dim strMSN, strDefTheme, strAllowUploads,  strPMtype, strDBNTUserName
dim strQuickReply, strForumSubscription
dim StrIPGateBan ,StrIPGateLck ,StrIPGateCok ,StrIPGateMet ,StrIPGateMsg ,StrIPGateLog ,StrIPGateTyp
dim StrIPGateExp, StrIPGateCss, strIPGateVer, StrIPGateLkMsg, strIPGateNoAcMsg, StrIPGateWarnMsg
dim strYAHOO, strFullName, strPicture, stMx, strSex, strCity, strState, strAge, strCountry, strOccupation
dim strBio, strHobbies, strLNews, strQuote, strMarStatus, strFavLinks, strRecentTopics, strAllowHideEmail
dim strRankColorAdmin, strRankColorMod, strRankColor0, strRankColor1, strRankColor2, strRankColor3
dim strRankColor4, strRankColor5, strNTGroups, strAutoLogon, strVar1, strVar2, strVar3, strVar4, strZip
dim strLoginType, browserReq, varBrowser, memID, isMAC, SecImage, dbHits, intMemberLCID, intPortalLCID
dim strCurDateAdjust, strCurDateString, strMTimeAdjust, strMTimeType, strMCurDateAdjust, strMCurDateString
dim mLev, strLoginStatus, strSiteOwner, strUserEmail, PMaccess, strUserMemberID
Dim arg1, arg2, arg3, arg4, arg5, arg6 'page breadcrumb variables
dim arrCurOnline(), arrGroups(), arrAppPerms()
		
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
	'on error resume next
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	my_Conn.Open strConnString
	'	Lets check to see if the strConnString or db path has changed
		if my_conn.Errors.Count <> 0 then 
		  'we can't connect, lets log the error
		  for counter = 0 to my_conn.Errors.Count -1
			ConnErrorNumber = my_conn.Errors(counter).Number
			ConnErrorDesc = my_conn.Errors(counter).Description
			if ConnErrorNumber <> 0 then 
			  writeToLog "Database","",ConnErrorNumber & " : " & ConnErrorDesc
			end if
		  next
			my_conn.Errors.Clear 
			set my_Conn = nothing
			on error goto 0
			Response.Redirect "site_setup.asp?RC=1"
		end if
	'on error goto 0

'blnSetup="N"
if request.QueryString("sky") = "dogg" then
  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end if

if blnSetup<>"Y" then
  if Application(strCookieURL & strUniqueID & "ConfigLoaded")= "" or IsNull(Application(strCookieURL & strUniqueID & "ConfigLoaded")) then 
	'## if the config variables aren't loaded into the Application object
	'## or after the admin has changed the configuration
	'## the variables get (re)loaded 
	
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
	strSql = strSql & ", C_STRRECENTTOPICS"
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
	strSql = strSql & ", C_STRMARSTATUS"
	strSql = strSql & ", C_STRFAVLINKS"
	strSql = strSql & ", C_STRHOMEPAGE"
	strSql = strSql & ", C_STRICQ "
	strSql = strSql & ", C_STRYAHOO "
	strSql = strSql & ", C_STRAIM "
	strSql = strSql & ", C_STRMSN "
	
	
	':: IPGATE Mod
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
	
	
	on error resume next
	'set rsCfg = my_Conn.Execute(strSql)
	set rsCfg = oSpData.GetRecordset(strSql)
	'	Lets check to see if the strConnString or db path has changed
		if err.number <> 0 then
			set my_Conn = nothing
			Response.Redirect "site_setup.asp?err=no_config_table"		
		end if
	on error goto 0

	Application.Lock
	'application.Contents.RemoveAll()
	Application(strCookieURL & strUniqueID & "strSiteTitle") = rsCfg("C_STRSITETITLE")
	Application(strCookieURL & strUniqueID & "strCopyright") = rsCfg("C_STRCOPYRIGHT")
	Application(strCookieURL & strUniqueID & "strTitleImage") = rsCfg("C_STRTITLEIMAGE")
	Application(strCookieURL & strUniqueID & "strHomeURL") = rsCfg("C_STRHOMEURL")
	Application(strCookieURL & strUniqueID & "strAuthType") = rsCfg("C_STRAUTHTYPE")
	Application(strCookieURL & strUniqueID & "strEmail") = rsCfg("C_STREMAIL")
	Application(strCookieURL & strUniqueID & "strUniqueEmail") = rsCfg("C_STRUNIQUEEMAIL")
	Application(strCookieURL & strUniqueID & "strMailMode") = rsCfg("C_STRMAILMODE")
	Application(strCookieURL & strUniqueID & "strMailServer") = rsCfg("C_STRMAILSERVER")
	Application(strCookieURL & strUniqueID & "strMailServerLogon") = rsCfg("C_STREMAILUSERNAME")
	Application(strCookieURL & strUniqueID & "strMailServerPassword") = rsCfg("C_STREMAILPASSWORD")
	Application(strCookieURL & strUniqueID & "strMailServerPort") = rsCfg("C_STREMAILPORT")
	Application(strCookieURL & strUniqueID & "strSender") = rsCfg("C_STRSENDER")
	Application(strCookieURL & strUniqueID & "strDateType") = rsCfg("C_STRDATETYPE")
	Application(strCookieURL & strUniqueID & "strTimeAdjust") = rsCfg("C_STRTIMEADJUST")
	Application(strCookieURL & strUniqueID & "strTimeType") = rsCfg("C_STRTIMETYPE")
	Application(strCookieURL & strUniqueID & "strIPLogging") = rsCfg("C_STRIPLOGGING")
	Application(strCookieURL & strUniqueID & "strAllowForumCode") = rsCfg("C_STRALLOWFORUMCODE")
	Application(strCookieURL & strUniqueID & "strIMGInPosts") = rsCfg("C_STRIMGINPOSTS")
	Application(strCookieURL & strUniqueID & "strAllowHTML") = rsCfg("C_STRALLOWHTML")
	Application(strCookieURL & strUniqueID & "strLockDown") = rsCfg("C_STRLOCKDOWN")
	Application(strCookieURL & strUniqueID & "strIcons") = rsCfg("C_STRICONS")
	Application(strCookieURL & strUniqueID & "strBadWordFilter") = rsCfg("C_STRBADWORDFILTER")
	Application(strCookieURL & strUniqueID & "strBadWords") = rsCfg("C_STRBADWORDS")
	Application(strCookieURL & strUniqueID & "strLogonForMail") = rsCfg("C_STRLOGONFORMAIL")
	Application(strCookieURL & strUniqueID & "strEmailVal") = rsCfg("C_STREMAILVAL")
	Application(strCookieURL & strUniqueID & "strFloodCheck") = rsCfg("C_STRFLOODCHECK")
	Application(strCookieURL & strUniqueID & "strFloodCheckTime") = rsCfg("C_STRFLOODCHECKTIME")
	Application(strCookieURL & strUniqueID & "strNewReg") = rsCfg("C_STRNEWREG")
	Application(strCookieURL & strUniqueID & "strDefTheme") = rsCfg("C_STRDEFTHEME")
	Application(strCookieURL & strUniqueID & "strAllowUploads") = rsCfg("C_ALLOWUPLOADS")
	Application(strCookieURL & strUniqueID & "strPMtype") = rsCfg("C_PMTYPE")
	Application(strCookieURL & strUniqueID & "STRIPGATEBAN")= rsCfg("C_STRIPGATEBAN")
	Application(strCookieURL & strUniqueID & "STRIPGATELCK")= rsCfg("C_STRIPGATELCK")
	Application(strCookieURL & strUniqueID & "STRIPGATECOK")= rsCfg("C_STRIPGATECOK")
	Application(strCookieURL & strUniqueID & "STRIPGATEMET")= rsCfg("C_STRIPGATEMET")
	Application(strCookieURL & strUniqueID & "STRIPGATEMSG")= rsCfg("C_STRIPGATEMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATELOG")= rsCfg("C_STRIPGATELOG")
	Application(strCookieURL & strUniqueID & "STRIPGATETYP")= rsCfg("C_STRIPGATETYP")
	Application(strCookieURL & strUniqueID & "STRIPGATEEXP")= rsCfg("C_STRIPGATEEXP")
	Application(strCookieURL & strUniqueID & "STRIPGATECSS")= rsCfg("C_STRIPGATECSS")
	Application(strCookieURL & strUniqueID & "STRIPGATEVER")= rsCfg("C_STRIPGATEVER")
	Application(strCookieURL & strUniqueID & "STRIPGATELKMSG")= rsCfg("C_STRIPGATELKMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATENOACMSG")= rsCfg("C_STRIPGATENOACMSG")
	Application(strCookieURL & strUniqueID & "STRIPGATEWARNMSG")= rsCfg("C_STRIPGATEWARNMSG")
	Application(strCookieURL & strUniqueID & "strHeaderType")= rsCfg("C_STRHEADERTYPE") 
	Application(strCookieURL & strUniqueID & "strLoginType")= rsCfg("C_STRLOGINTYPE")
	'Application(strCookieURL & strUniqueID & "bFso")= bFso 
	Application(strCookieURL & strUniqueID & "SECIMAGE")= rsCfg("C_SECIMAGE")
	Application(strCookieURL & strUniqueID & "intSubSkin")= rsCfg("C_INTSUBSKIN")
	Application(strCookieURL & strUniqueID & "strChkDate")= rsCfg("C_ONEADAYDATE")
	Application(strCookieURL & strUniqueID & "strImgComp")= rsCfg("C_COMP_IMAGE")
	Application(strCookieURL & strUniqueID & "strUploadComp")= rsCfg("C_COMP_UPLOAD")
	Application(strCookieURL & strUniqueID & "strXmlHttpComp")= strXmlHttp
	Application(strCookieURL & strUniqueID & "strAppVars") = bldAppAccess()	
	
	Application(strCookieURL & strUniqueID & "STRNTGROUPS") = rsCfg("C_STRNTGROUPS")
	Application(strCookieURL & strUniqueID & "STRAUTOLOGON") = rsCfg("C_STRAUTOLOGON")
	
	Application(strCookieURL & strUniqueID & "strMoveTopicMode") = rsCfg("C_STRMOVETOPICMODE")
	Application(strCookieURL & strUniqueID & "strPrivateForums") = rsCfg("C_STRPRIVATEFORUMS")
	Application(strCookieURL & strUniqueID & "strShowModerators") = rsCfg("C_STRSHOWMODERATORS")
	Application(strCookieURL & strUniqueID & "strHotTopic") = rsCfg("C_STRHOTTOPIC")
	Application(strCookieURL & strUniqueID & "intHotTopicNum") = rsCfg("C_INTHOTTOPICNUM")
	Application(strCookieURL & strUniqueID & "strShowRank") = rsCfg("C_STRSHOWRANK")
	Application(strCookieURL & strUniqueID & "strRankAdmin") = rsCfg("C_STRRANKADMIN")
	Application(strCookieURL & strUniqueID & "strRankMod") = rsCfg("C_STRRANKMOD")
	Application(strCookieURL & strUniqueID & "strRankLevel0") = rsCfg("C_STRRANKLEVEL0")
	Application(strCookieURL & strUniqueID & "strRankLevel1") = rsCfg("C_STRRANKLEVEL1")
	Application(strCookieURL & strUniqueID & "strRankLevel2") = rsCfg("C_STRRANKLEVEL2")
	Application(strCookieURL & strUniqueID & "strRankLevel3") = rsCfg("C_STRRANKLEVEL3")
	Application(strCookieURL & strUniqueID & "strRankLevel4") = rsCfg("C_STRRANKLEVEL4")
	Application(strCookieURL & strUniqueID & "strRankLevel5") = rsCfg("C_STRRANKLEVEL5")
	Application(strCookieURL & strUniqueID & "strRankColorAdmin") = rsCfg("C_STRRANKCOLORADMIN")
	Application(strCookieURL & strUniqueID & "strRankColorMod") = rsCfg("C_STRRANKCOLORMOD")
	Application(strCookieURL & strUniqueID & "strRankColor0") = rsCfg("C_STRRANKCOLOR0")
	Application(strCookieURL & strUniqueID & "strRankColor1") = rsCfg("C_STRRANKCOLOR1")
	Application(strCookieURL & strUniqueID & "strRankColor2") = rsCfg("C_STRRANKCOLOR2")
	Application(strCookieURL & strUniqueID & "strRankColor3") = rsCfg("C_STRRANKCOLOR3")
	Application(strCookieURL & strUniqueID & "strRankColor4") = rsCfg("C_STRRANKCOLOR4")
	Application(strCookieURL & strUniqueID & "strRankColor5") = rsCfg("C_STRRANKCOLOR5")
	Application(strCookieURL & strUniqueID & "intRankLevel0") = rsCfg("C_INTRANKLEVEL0")
	Application(strCookieURL & strUniqueID & "intRankLevel1") = rsCfg("C_INTRANKLEVEL1")
	Application(strCookieURL & strUniqueID & "intRankLevel2") = rsCfg("C_INTRANKLEVEL2")
	Application(strCookieURL & strUniqueID & "intRankLevel3") = rsCfg("C_INTRANKLEVEL3")
	Application(strCookieURL & strUniqueID & "intRankLevel4") = rsCfg("C_INTRANKLEVEL4")
	Application(strCookieURL & strUniqueID & "intRankLevel5") = rsCfg("C_INTRANKLEVEL5")
	Application(strCookieURL & strUniqueID & "strShowStatistics") = rsCfg("C_STRSHOWSTATISTICS")
	Application(strCookieURL & strUniqueID & "strShowPaging") = rsCfg("C_STRSHOWPAGING")
	Application(strCookieURL & strUniqueID & "strPageSize") = rsCfg("C_STRPAGESIZE")
	Application(strCookieURL & strUniqueID & "strPageNumberSize") = rsCfg("C_STRPAGENUMBERSIZE")
	Application(strCookieURL & strUniqueID & "strForumStatus") = rsCfg("C_FORUMSTATUS")
	Application(strCookieURL & strUniqueID & "strPollCreate") = rsCfg("C_POLLCREATE")
	Application(strCookieURL & strUniqueID & "strFeaturedPoll") = rsCfg("C_FEATUREDPOLL")
	Application(strCookieURL & strUniqueID & "strQuickReply") = rsCfg("C_STRQUICKREPLY")
	Application(strCookieURL & strUniqueID & "strForumSubscription") = rsCfg("C_FORUMSUBSCRIPTION")
	Application(strCookieURL & strUniqueID & "strEditedByDate") = rsCfg("C_STREDITEDBYDATE")
	Application(strCookieURL & strUniqueID & "strRecentTopics") = rsCfg("C_STRRECENTTOPICS")
	
	Application(strCookieURL & strUniqueID & "strICQ") = rsCfg("C_STRICQ")
	Application(strCookieURL & strUniqueID & "strYAHOO") = rsCfg("C_STRYAHOO")
	Application(strCookieURL & strUniqueID & "strAIM") = rsCfg("C_STRAIM")
	Application(strCookieURL & strUniqueID & "strMSN") = rsCfg("C_STRMSN")
	Application(strCookieURL & strUniqueID & "strHomepage") = rsCfg("C_STRHOMEPAGE")
	Application(strCookieURL & strUniqueID & "strFullName") = rsCfg("C_STRFULLNAME")
	Application(strCookieURL & strUniqueID & "strPicture") = rsCfg("C_STRPICTURE")
	Application(strCookieURL & strUniqueID & "strSex") = rsCfg("C_STRSEX")
	Application(strCookieURL & strUniqueID & "strAge") = rsCfg("C_STRAGE")
	Application(strCookieURL & strUniqueID & "strMarStatus") = rsCfg("C_STRMARSTATUS")
	Application(strCookieURL & strUniqueID & "strCity") = rsCfg("C_STRCITY")
	Application(strCookieURL & strUniqueID & "strState") = rsCfg("C_STRSTATE")
	Application(strCookieURL & strUniqueID & "strZip") = rsCfg("C_STRZIP")
	Application(strCookieURL & strUniqueID & "strCountry") = rsCfg("C_STRCOUNTRY")
	Application(strCookieURL & strUniqueID & "strOccupation") = rsCfg("C_STROCCUPATION")
	Application(strCookieURL & strUniqueID & "strFavLinks") = rsCfg("C_STRFAVLINKS")
	Application(strCookieURL & strUniqueID & "strVar1") = rsCfg("C_STRVAR1")
	Application(strCookieURL & strUniqueID & "strBio") = rsCfg("C_STRBIO")
	Application(strCookieURL & strUniqueID & "strVar2") = rsCfg("C_STRVAR2")
	Application(strCookieURL & strUniqueID & "strHobbies") = rsCfg("C_STRHOBBIES")
	Application(strCookieURL & strUniqueID & "strVar3") = rsCfg("C_STRVAR3")
	Application(strCookieURL & strUniqueID & "strLNews") = rsCfg("C_STRLNEWS")
	Application(strCookieURL & strUniqueID & "strVar4") = rsCfg("C_STRVAR4")
	Application(strCookieURL & strUniqueID & "strQuote") = rsCfg("C_STRQUOTE")
	
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= "YES"

	Application.UnLock
	set rsCfg = nothing
  end if 
end if
okoame = 1
if blnSetup <> "Y" and stMx = "Sk" then 
	strSiteTitle = replace(Application(strCookieURL & strUniqueID & "strSiteTitle"),"''","'")
	strCopyright = Application(strCookieURL & strUniqueID & "strCopyright")
	strTitleImage = Application(strCookieURL & strUniqueID & "strTitleImage")
	strHomeURL = Application(strCookieURL & strUniqueID & "strHomeURL")
	strAuthType = Application(strCookieURL & strUniqueID & "strAuthType")
	strUniqueEmail = Application(strCookieURL & strUniqueID & "strUniqueEmail")
	strIPLogging = Application(strCookieURL & strUniqueID & "strIPLogging")
	strIMGInPosts = Application(strCookieURL & strUniqueID & "strIMGInPosts")
	strAllowHTML = Application(strCookieURL & strUniqueID & "strAllowHTML")
	strAllowForumCode = Application(strCookieURL & strUniqueID & "strAllowForumCode")
	strBadWordFilter = Application(strCookieURL & strUniqueID & "strBadWordFilter")
	strBadWords = Application(strCookieURL & strUniqueID & "strBadWords")
	strLockDown = Application(strCookieURL & strUniqueID & "strLockDown")
	
	strDateFormat = getDateFormat()
	strDateType = Application(strCookieURL & strUniqueID & "strDateType")
	strTimeAdjust = Application(strCookieURL & strUniqueID & "strTimeAdjust")
	strTimeType = Application(strCookieURL & strUniqueID & "strTimeType")
	
	strCurDateAdjust = DateAdd("h", strTimeAdjust , Now()) 'portal offset from server
	strCurDateString = DateToStr(strCurDateAdjust)
	strCurDateAdjust = strCurDateAdjust
	strCurDate = ChkDate2(strCurDateString)
	
	strMTimeAdjust = 0
	strMTimeType = strDateType
	strMCurDateAdjust = strCurDateAdjust
	strMCurDateString = strCurDateString
	intMemberLCID = intPortalLCID
	
	strNTGroups = Application(strCookieURL & strUniqueID & "STRNTGROUPS")
	strAutoLogon = Application(strCookieURL & strUniqueID & "STRAUTOLOGON")
	
	strEmail = Application(strCookieURL & strUniqueID & "strEmail")
	strMailMode = Application(strCookieURL & strUniqueID & "strMailMode")
	strMailServer = Application(strCookieURL & strUniqueID & "strMailServer")
	strMailServerLogon = Application(strCookieURL & strUniqueID & "strMailServerLogon")
	strMailServerPassword = Application(strCookieURL & strUniqueID & "strMailServerPassword")
	strMailServerPort = Application(strCookieURL & strUniqueID & "strMailServerPort")
	strSender = Application(strCookieURL & strUniqueID & "strSender")
	strLogonForMail = Application(strCookieURL & strUniqueID & "strLogonForMail")
	strEmailVal = Application(strCookieURL & strUniqueID & "STREMAILVAL")
	
	strFloodCheck = Application(strCookieURL & strUniqueID & "STRFLOODCHECK")
	strFloodCheckTime = Application(strCookieURL & strUniqueID & "STRFLOODCHECKTIME")
	strNewReg = Application(strCookieURL & strUniqueID & "STRNEWREG")
	strDefTheme = Application(strCookieURL & strUniqueID & "strDefTheme")
	strPMtype = Application(strCookieURL & strUniqueID & "strPMtype")
	StrIPGateBan = Application(strCookieURL & strUniqueID & "STRIPGATEBAN")
	StrIPGateLck = Application(strCookieURL & strUniqueID & "STRIPGATELCK")
	StrIPGateCok = Application(strCookieURL & strUniqueID & "STRIPGATECOK")
	StrIPGateMet = Application(strCookieURL & strUniqueID & "STRIPGATEMET")
	StrIPGateMsg = Application(strCookieURL & strUniqueID & "STRIPGATEMSG")
	StrIPGateLog = Application(strCookieURL & strUniqueID & "STRIPGATELOG")
	StrIPGateTyp = Application(strCookieURL & strUniqueID & "STRIPGATETYP")
	StrIPGateExp = Application(strCookieURL & strUniqueID & "STRIPGATEEXP")
	StrIPGateCss = Application(strCookieURL & strUniqueID & "STRIPGATECSS")
	strIPGateVer = Application(strCookieURL & strUniqueID & "STRIPGATEVER")
	StrIPGateLkMsg = Application(strCookieURL & strUniqueID & "STRIPGATELKMSG")
	strIPGateNoAcMsg = Application(strCookieURL & strUniqueID & "STRIPGATENOACMSG")
	StrIPGateWarnMsg = Application(strCookieURL & strUniqueID & "STRIPGATEWARNMSG") 
	strAllowHideEmail = "1"
	stWb = "yPor"
	strIcons = Application(strCookieURL & strUniqueID & "strIcons")
	strHeaderType = Application(strCookieURL & strUniqueID & "strHeaderType") 
	strLoginType = Application(strCookieURL & strUniqueID & "strLoginType") 
	'bFso = Application(strCookieURL & strUniqueID & "bFso")
	SecImage = Application(strCookieURL & strUniqueID & "SecImage")
	intSubSkin = Application(strCookieURL & strUniqueID & "intSubSkin")
	strChkDate = Application(strCookieURL & strUniqueID & "strChkDate")
	strImgComp = Application(strCookieURL & strUniqueID & "strImgComp")
	strAllowUploads = Application(strCookieURL & strUniqueID & "strAllowUploads")
	strUploadComp = Application(strCookieURL & strUniqueID & "strUploadComp")
	strXmlHttpComp = Application(strCookieURL & strUniqueID & "strXmlHttpComp")
	':: build array for app configuration
	bldArrAppAccess()
	
	'forums
	strMoveTopicMode = Application(strCookieURL & strUniqueID & "strMoveTopicMode")
	strPrivateForums = Application(strCookieURL & strUniqueID & "strPrivateForums")
	strShowModerators = Application(strCookieURL & strUniqueID & "strShowModerators")
	strHotTopic = Application(strCookieURL & strUniqueID & "strHotTopic")
	intHotTopicNum = Application(strCookieURL & strUniqueID & "intHotTopicNum")
	strEditedByDate = Application(strCookieURL & strUniqueID & "strEditedByDate")
	strShowRank = Application(strCookieURL & strUniqueID & "strShowRank")
	strRankAdmin = Application(strCookieURL & strUniqueID & "strRankAdmin")
	strRankMod = Application(strCookieURL & strUniqueID & "strRankMod")
	strRankLevel0 = Application(strCookieURL & strUniqueID & "strRankLevel0")
	strRankLevel1 = Application(strCookieURL & strUniqueID & "strRankLevel1")
	strRankLevel2 = Application(strCookieURL & strUniqueID & "strRankLevel2")
	strRankLevel3 = Application(strCookieURL & strUniqueID & "strRankLevel3")
	strRankLevel4 = Application(strCookieURL & strUniqueID & "strRankLevel4")
	strRankLevel5 = Application(strCookieURL & strUniqueID & "strRankLevel5")
	strRankColorAdmin = Application(strCookieURL & strUniqueID & "strRankColorAdmin")
	strRankColorMod = Application(strCookieURL & strUniqueID & "strRankColorMod")
	strRankColor0 = Application(strCookieURL & strUniqueID & "strRankColor0")
	strRankColor1 = Application(strCookieURL & strUniqueID & "strRankColor1")
	strRankColor2 = Application(strCookieURL & strUniqueID & "strRankColor2")
	strRankColor3 = Application(strCookieURL & strUniqueID & "strRankColor3")
	strRankColor4 = Application(strCookieURL & strUniqueID & "strRankColor4")
	strRankColor5 = Application(strCookieURL & strUniqueID & "strRankColor5")
	intRankLevel0 = Application(strCookieURL & strUniqueID & "intRankLevel0")
	intRankLevel1 = Application(strCookieURL & strUniqueID & "intRankLevel1")
	intRankLevel2 = Application(strCookieURL & strUniqueID & "intRankLevel2")
	intRankLevel3 = Application(strCookieURL & strUniqueID & "intRankLevel3")
	intRankLevel4 = Application(strCookieURL & strUniqueID & "intRankLevel4")
	intRankLevel5 = Application(strCookieURL & strUniqueID & "intRankLevel5")
	strShowStatistics = Application(strCookieURL & strUniqueID & "strShowStatistics")
	strShowPaging = Application(strCookieURL & strUniqueID & "strShowPaging")
	strPageSize = Application(strCookieURL & strUniqueID & "strPageSize")
	strPageNumberSize = Application(strCookieURL & strUniqueID & "strPageNumberSize")
	strForumStatus = Application(strCookieURL & strUniqueID & "strForumStatus") 
	strPollCreate = Application(strCookieURL & strUniqueID & "STRPOLLCREATE")
	strFeaturedPoll = Application(strCookieURL & strUniqueID & "STRFEATUREDPOLL")
	strQuickReply = Application(strCookieURL & strUniqueID & "STRQUICKREPLY")
	strForumSubscription = Application(strCookieURL & strUniqueID & "strForumSubscription") 
	'member stuff
	strFullName = Application(strCookieURL & strUniqueID & "strFullName")
	strPicture = Application(strCookieURL & strUniqueID & "strPicture")
	strMarStatus = Application(strCookieURL & strUniqueID & "strMarStatus")
	strAge = Application(strCookieURL & strUniqueID & "strAge")
	strSex = Application(strCookieURL & strUniqueID & "strSex")
	strCity= Application(strCookieURL & strUniqueID & "strCity")
	strState = Application(strCookieURL & strUniqueID & "strState")
	strZip = Application(strCookieURL & strUniqueID & "strZip")
	strCountry = Application(strCookieURL & strUniqueID & "strCountry") 
	strICQ = Application(strCookieURL & strUniqueID & "strICQ")
	strYAHOO = Application(strCookieURL & strUniqueID & "strYAHOO")
	strAIM = Application(strCookieURL & strUniqueID & "strAIM")
	strMSN = Application(strCookieURL & strUniqueID & "strMSN")
	strHomepage = Application(strCookieURL & strUniqueID & "strHomepage")
	strOccupation = Application(strCookieURL & strUniqueID & "strOccupation")
	strBio = Application(strCookieURL & strUniqueID & "strBio") 
	strHobbies = Application(strCookieURL & strUniqueID & "strHobbies") 
	strLNews = 	Application(strCookieURL & strUniqueID & "strLNews") 
	strQuote = Application(strCookieURL & strUniqueID & "strQuote") 
	strFavLinks = Application(strCookieURL & strUniqueID & "strFavLinks")
	strRecentTopics = Application(strCookieURL & strUniqueID & "strRecentTopics") 
	strVar1 = Application(strCookieURL & strUniqueID & "strVar1")
	strVar2 = Application(strCookieURL & strUniqueID & "strVar2")
	strVar3 = Application(strCookieURL & strUniqueID & "strVar3")
	strVar4 = Application(strCookieURL & strUniqueID & "strVar4")
	
	intSubSkin = Application(strCookieURL & strUniqueID & "intSubSkin")
	'intSubSkin = 0	

	':: config maintenance
	if not bFso then
		bFso = false
		intUploads = 0
  		strAllowUploads = 0
		strUploadComp = "none"
		bAspNet = false
	end if
	if strUploadComp = "none" then
		'intUploads = 0
  		strAllowUploads = 0
	end if
	if intUploads = 0 then
  		strAllowUploads = 0
		strUploadComp = "none"
	end if
	if not bAspNet then
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

'::::::::::::::::::::::: browser sniffer code. ::::::::::::::::::::::::::::::
	browserReq = request.ServerVariables("HTTP_USER_AGENT")
	varBrowser = ""
   		if instr(lcase(browserReq),"opera") <> 0 then
			varBrowser = "opera"
		elseif instr(lcase(browserReq),"firefox") <> 0 then ' Is FireFox browser
			varBrowser = "firefox"
		elseif instr(lcase(browserReq),"firebird") <> 0 then ' Is Firebird browser
			varBrowser = "firebird"
		elseif instr(lcase(browserReq),"safari") <> 0 then ' Is Safari browser
			varBrowser = "safari"
		elseif instr(lcase(browserReq),"lynx") <> 0 then ' Is lynx browser
			varBrowser = "lynx"
		elseif instr(lcase(browserReq),"camino") <> 0 then ' Is Safari browser
			varBrowser = "camino"
		elseif instr(lcase(browserReq),"msie") <> 0 then ' Is MSIE browser
			varBrowser = "ie"
		elseif instr(lcase(browserReq),"gecko") <> 0 then ' Is Netscape browser
			varBrowser = "netscape"
		else
			varBrowser = "other"	
		end if
			isMAC = false
			stPl = "tal."
		
		':: This code detects if the editor is browser compatable
		':: If not, then change settings to show the default [forum code] browser.
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
'response.Write("blnSetup: " & blnSetup)
%>
<% if blnSetup <> "Y" then %>
<%
if not uploadPg then
  if request("lang") <> "" then
    strLang = request("lang")
  end if
end if
'if bFso then
  'include server.mappath("lang/" & strLang & "/core.asp")
'else 
  %><!-- #include file="lang/en/core.asp" --><%
'end if
 end if

function DetectXmlHttp()
  Dim sTempComponent
	sTempComponent = "none"
	on error resume next	
	if CheckXmlHttpComp("Microsoft.XMLHTTP") = true then
	  'sTempComponent = "DOTNET3"
	  sTempComponent = "Microsoft.XMLHTTP"
	else
	  if CheckXmlHttpComp("Msxml2.ServerXMLHTTP") = true then
		'sTempComponent = "DOTNET2"
		sTempComponent = "Msxml2.ServerXMLHTTP"
	  else
		if CheckXmlHttpComp("Msxml2.ServerXMLHTTP.4.0") = true then
		  'sTempComponent = "DOTNET1"
		  sTempComponent = "Msxml2.ServerXMLHTTP.4.0"
		else
		  'sTempComponent = "NOT FOUND: ASP.NET Server Component<br>"
		end if
	  end if
	end if
	on error goto 0
	DetectXmlHttp = sTempComponent
end function

function CheckXmlHttpComp(XmlHttpObj)
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
		erHttp = "An error has accured with ASP.NET component: " & XmlHttpObj
		erHttp = erHttp & " - Error:" & objHttp.responseText
		writeToLog "Config","","[ERROR!] (config_core CheckXmlHttpComp)" & erHttp
		'Response.End
	  else
        if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
          Detected = true
        end if
	  end if
	end if
  End if
  Set objHttp = nothing
  on error goto 0
  CheckXmlHttpComp = Detected
end function
%>