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

sub chkLogInOut()
  select case Request.Form("Method_Type")
	case "login"
	    if not chkValidUserName(Request.Form("Name")) then
		  closeAndGo("error.asp?type=luser")
		end if
		strDBNTFUserName = chkString(Request.Form("Name"),"sqlstring")
		isMbr = chkIsMbr(strDBNTFUserName, pEncrypt(pEnPrefix & Request.Form("Password")))
		if isMbr = 1 then
		  fSecCode=Ucase(Request.Form("SecCode"))
		  If (SecImage<2) OR (SecImage>1 AND DoSecImage(fSecCode)) Then
			Call DoCookies(Request.Form("SavePassword"))
			if request.QueryString() <> "" then
			  sScript = sScript & "?" & request.QueryString()
			end if
			if instr(sScript,"error.asp") > 0 then
			  closeAndGo(strHomeUrl)
			else
			  closeAndGo(sScript)
			end if
		  Else
			closeAndGo("error.asp?type=lsec")
		  End If
		else
			closeAndGo("error.asp?type=luser")
		end if
	case "logout"
		Call ClearCookies()
		Session.Contents.RemoveAll()
		'Session.Abandon()
		strSql = "DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP='" & request.ServerVariables("REMOTE_ADDR") & "'"
		executeThis(strSql)
		closeAndGo("default.asp")
  end select
end sub

function chkLoginStatus()
  if strAuthType = "ad" then
	call NTauthenticate()
	if ChkAccountReg() = "1" then
		call NTUser()
	else
	  'call regNTuser()
	  call NTUser()
	end if
	'NTdebug()
	strDBNTUserName = readSession("userID")
	strDBNTFUserName = readSession("userID")
	
  elseif strAuthType = "nt" then
	call NTauthenticate()
	if ChkAccountReg() = "1" then
		call NTUser()
	else
	  'call regNTuser()
	  call NTUser()
	end if
	'NTdebug()
	strDBNTUserName = readSession("userID")
	strDBNTFUserName = readSession("userID")
	
  elseif strAuthType = "db" then
    
	if (readMultiCookie("User","Name") <> "" and readMultiCookie("User","PWord") <> "") then
	 cName = readMultiCookie("User","Name")
	 cPass = readMultiCookie("User","PWord")
	 if not bUseMemberSession then
	   setMemberVars_old cName,cPass
	 else
	  if IsEmpty(readSession("userID")) or IsNull(readSession("userID")) or trim(readSession("userID")) = "" then
	    Call setMemberSessVars(cName,cPass)
		setMemberVars()
	  else
		setMemberVars()
	  end if
	 end if
		
	else
		strDBNTUserName = ""
		strUserMemberID = 0
		strUserEmail = ""
		strMTimeAdjust = 0
		mLev = 0
		PMaccess = 0
	end if
  end if
end function

function setMemberVars()
	strDBNTUserName = readSession("username")
	strUserMemberID = clng(readSession("userID"))
	strUserEmail = readSession("useremail")
	mLev = readSession("usermlev")
	intIsSuperAdmin = chkIsSuperAdmin(2,strDBNTUsername)
	strMTimeAdjust = readSession("usertimeadjust")
	strMTimeType = readSession("usertimetype")
	intMemberLCID = readSession("userlcid")
	if len(intMemberLCID) = 4 or len(intMemberLCID) = 5 then
	  Session.LCID = intMemberLCID
	  strDateFormat = getDateFormat()
	end if
			
	strTimeType = strMTimeType
	strMCurDateAdjust = DateAdd("h", (strTimeAdjust + strMTimeAdjust) , now())
	strMCurDateString = DateToStr(strMCurDateAdjust)
	strCurDate = ChkDate2(strMCurDateString)
			
	PMaccess = cint(readSession("userpmaccess"))
end function

function setMemberSessVars(mname,mpass)
  strSql = "SELECT MEMBER_ID, M_NAME, M_USERNAME, M_LEVEL, M_EMAIL, M_PASSWORD, M_PMSTATUS"
  strSql = strSql & ", M_PMRECEIVE, M_TIME_OFFSET, M_TIME_TYPE, M_LCID, M_AGE"
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
  strSql = strSql & " WHERE M_NAME = '" & mname & "'"
  if strAuthType = "db" then
	strSql = strSql & " AND M_PASSWORD = '" & mpass &"'"
  end if
  strSql = strSql & " and M_STATUS=1"
  'response.Write(strSql & "<br /><br />")
  Set rsCheck = my_Conn.Execute(strSql)
  if rsCheck.BOF and rsCheck.EOF then
	Call ClearCookies()
	strDBNTUserName = ""
	strUserMemberID = 0
	strUserEmail = ""
	strMTimeAdjust = 0
	mLev = 0
	PMaccess = 0
  else
    if strAuthType = "db" then
	  Call setSession("username",rsCheck("M_NAME"))
	else
	  Call setSession("username",rsCheck("M_USERNAME"))
    end if
	  Call setSession("userID",clng(rsCheck("MEMBER_ID")))
	  Call setSession("useremail",rsCheck("M_EMAIL"))
	  Call setSession("usermlev",rsCheck("M_LEVEL")+1)
	  Call setSession("usertimeadjust",rsCheck("M_TIME_OFFSET"))
	  Call setSession("usertimetype",rsCheck("M_TIME_TYPE"))
	  Call setSession("userlcid",rsCheck("M_LCID"))
	  Call setSession("userpmaccess",rsCheck("M_PMSTATUS"))
	if rsCheck("M_PMRECEIVE") = 0 then
	  Call setSession("userpmaccess",0)
	end if
  end if
  set rsCheck = nothing
end function

function setMemberVars_old(mname,mpass)
  strSql = "SELECT MEMBER_ID, M_NAME, M_USERNAME, M_LEVEL, M_EMAIL, M_PASSWORD, M_PMSTATUS"
  strSql = strSql & ", M_PMRECEIVE, M_TIME_OFFSET, M_TIME_TYPE, M_LCID, M_AGE"
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
  strSql = strSql & " WHERE M_NAME = '" & mname & "'"
  if strAuthType = "db" then
	strSql = strSql & " AND M_PASSWORD = '" & mpass &"'"
  end if
  strSql = strSql & " and M_STATUS=1"
  'response.Write(strSql & "<br /><br />")
  Set rsCheck = my_Conn.Execute(strSql)
  if rsCheck.BOF and rsCheck.EOF then
	Call ClearCookies()
	strDBNTUserName = ""
	strUserMemberID = 0
	strUserEmail = ""
	mLev = 0
	PMaccess = 0
  else
    if strAuthType <> "db" then
	  Session(strUniqueID & "username") = rsCheck("M_USERNAME")
    end if
	strDBNTUserName = rsCheck("M_NAME")
	strUserMemberID = clng(rsCheck("MEMBER_ID"))
	strUserEmail = rsCheck("M_EMAIL")
	mLev = rsCheck("M_LEVEL")+1
	intIsSuperAdmin = chkIsSuperAdmin(2,strDBNTUsername)
	strMBirthday = rsCheck("M_AGE")
	strMTimeAdjust = rsCheck("M_TIME_OFFSET")
	strMTimeType = rsCheck("M_TIME_TYPE")
	intMemberLCID = rsCheck("M_LCID")
	if len(intMemberLCID) = 4 or len(intMemberLCID) = 5 then
	  Session.LCID = intMemberLCID
	  strDateFormat = getDateFormat()
	end if
			
	strTimeType = strMTimeType
	strMCurDateAdjust = DateAdd("h", (strTimeAdjust + strMTimeAdjust) , now())
	strMCurDateString = DateToStr(strMCurDateAdjust)
	strCurDate = ChkDate2(strMCurDateString)
			
	PMaccess = rsCheck("M_PMSTATUS")
	if rsCheck("M_PMRECEIVE") = 0 then
	  PMaccess = 0
	else
	end if
  end if
  set rsCheck = nothing
end function

function bldArrAppAccess_old()
	dim tmpAppID, tmpAppActive, tmpAppGroupsR, tmpAppGroupsW, tmpAppGroupsF, tmpAppIName, bHasAccess
	dim tmpAppSubsc, tmpAppBkMk
	'bHasAccess = true
	
	    tmpApp = split(Application(strCookieURL & strUniqueID & "strAppVars"),";")
	    tmpAppID = tmpApp(0)
	    tmpAppIName = tmpApp(1)
		tmpAppActive = tmpApp(2)
		tmpAppGroupsR = tmpApp(3)
		tmpAppGroupsW = tmpApp(4)
		tmpAppGroupsF = tmpApp(5)
		tmpAppSubsc = tmpApp(6)
		tmpAppBkMk = tmpApp(7)
		tmpAppSecCode = tmpApp(8)
		tmpiData1 = tmpApp(9)
		tmpiData2 = tmpApp(10)
		tmpiData3 = tmpApp(11)
		tmpiData4 = tmpApp(12)
		tmpiData5 = tmpApp(13)
		tmpiData6 = tmpApp(14)
		tmpiData7 = tmpApp(15)
		tmpiData8 = tmpApp(16)
		tmpiData9 = tmpApp(17)
		tmpiData10 = tmpApp(18)
		tmptData1 = tmpApp(19)
		tmptData2 = tmpApp(20)
		tmptData3 = tmpApp(21)
		tmptData4 = tmpApp(22)
		tmptData5 = tmpApp(23)
	  
	  if tmpAppID <> "" then
		tmpAppID1 = split(tmpAppID,"@")
		tmpAppIName1 = split(tmpAppIName,"@")
		tmpAppActive1 = split(tmpAppActive,"@")
		tmpAppGroupsR1 = split(tmpAppGroupsR,"@")
		tmpAppGroupsW1 = split(tmpAppGroupsW,"@")
		tmpAppGroupsF1 = split(tmpAppGroupsF,"@")
		tmpAppSubsc1 = split(tmpAppSubsc,"@")
		tmpAppBkMk1 = split(tmpAppBkMk,"@")
		tmpAppSecCode1 = split(tmpAppSecCode,"@")
		tmpiData11 = split(tmpiData1,"@")
		tmpiData12 = split(tmpiData2,"@")
		tmpiData13 = split(tmpiData3,"@")
		tmpiData14 = split(tmpiData4,"@")
		tmpiData15 = split(tmpiData5,"@")
		tmpiData16 = split(tmpiData6,"@")
		tmpiData17 = split(tmpiData7,"@")
		tmpiData18 = split(tmpiData8,"@")
		tmpiData19 = split(tmpiData9,"@")
		tmpiData110 = split(tmpiData10,"@")
		tmptData11 = split(tmptData1,"@")
		tmptData12 = split(tmptData2,"@")
		tmptData13 = split(tmptData3,"@")
		tmptData14 = split(tmptData4,"@")
		tmptData15 = split(tmptData5,"@")
		acnt = ubound(tmpAppID1)-1
		redim arrAppPerms(acnt,23)
		for ag = 0 to acnt
		  arrAppPerms(ag,0) = tmpAppID1(ag)
		  arrAppPerms(ag,1) = tmpAppIName1(ag)
		  arrAppPerms(ag,2) = tmpAppActive1(ag)
		  arrAppPerms(ag,3) = tmpAppGroupsR1(ag)
		  arrAppPerms(ag,4) = tmpAppGroupsW1(ag)
		  arrAppPerms(ag,5) = tmpAppGroupsF1(ag)
		  arrAppPerms(ag,6) = tmpAppSubsc1(ag)
		  arrAppPerms(ag,7) = tmpAppBkMk1(ag)
		  arrAppPerms(ag,8) = tmpAppSecCode1(ag)
		  arrAppPerms(ag,9) = tmpiData11(ag)
		  arrAppPerms(ag,10) = tmpiData12(ag)
		  arrAppPerms(ag,11) = tmpiData13(ag)
		  arrAppPerms(ag,12) = tmpiData14(ag)
		  arrAppPerms(ag,13) = tmpiData15(ag)
		  arrAppPerms(ag,14) = tmpiData16(ag)
		  arrAppPerms(ag,15) = tmpiData17(ag)
		  arrAppPerms(ag,16) = tmpiData18(ag)
		  arrAppPerms(ag,17) = tmpiData19(ag)
		  arrAppPerms(ag,18) = tmpiData110(ag)
		  arrAppPerms(ag,19) = tmptData11(ag)
		  arrAppPerms(ag,20) = tmptData12(ag)
		  arrAppPerms(ag,21) = tmptData13(ag)
		  arrAppPerms(ag,22) = tmptData14(ag)
		  arrAppPerms(ag,23) = tmptData15(ag)
		next
	  end if
end function

function bldArrAppAccess()
	dim tmpAppID, tmpAppActive, tmpAppGroupsR, tmpAppGroupsW, tmpAppGroupsF, tmpAppIName, bHasAccess
	dim tmpAppSubsc, tmpAppBkMk
	'bHasAccess = true
	sSql = "SELECT * FROM "& strTablePrefix & "APPS"
	'set rsA = my_Conn.execute(sSql)
	set rsA = oSpData.GetRecordset(sSql)
	if not rsA.eof then
	  do until rsA.eof
	    tmpAppID = tmpAppID & rsA("APP_ID") & "|"
	    tmpAppIName = tmpAppIName & rsA("APP_INAME") & "|"
		tmpAppActive = tmpAppActive & rsA("APP_ACTIVE") & "|"
		tmpAppGroupsR = tmpAppGroupsR & rsA("APP_GROUPS_USERS") & "|"
		tmpAppGroupsW = tmpAppGroupsW & rsA("APP_GROUPS_WRITE") & "|"
		tmpAppGroupsF = tmpAppGroupsF & rsA("APP_GROUPS_FULL") & "|"
		tmpAppSubsc = tmpAppSubsc & rsA("APP_SUBSCRIPTIONS") & "|"
		tmpAppBkMk = tmpAppBkMk & rsA("APP_BOOKMARKS") & "|"
		tmpAppSecCode = tmpAppSecCode & rsA("APP_SUBSEC") & "|"
		tmpiData1 = tmpiData1 & rsA("APP_iData1") & "|"
		tmpiData2 = tmpiData2 & rsA("APP_iData2") & "|"
		tmpiData3 = tmpiData3 & rsA("APP_iData3") & "|"
		tmpiData4 = tmpiData4 & rsA("APP_iData4") & "|"
		tmpiData5 = tmpiData5 & rsA("APP_iData5") & "|"
		tmpiData6 = tmpiData6 & rsA("APP_iData6") & "|"
		tmpiData7 = tmpiData7 & rsA("APP_iData7") & "|"
		tmpiData8 = tmpiData8 & rsA("APP_iData8") & "|"
		tmpiData9 = tmpiData9 & rsA("APP_iData9") & "|"
		tmpiData10 = tmpiData10 & rsA("APP_iData10") & "|"
		tmptData1 = tmptData1 & rsA("APP_tData1") & "|"
		tmptData2 = tmptData2 & rsA("APP_tData2") & "|"
		tmptData3 = tmptData3 & rsA("APP_tData3") & "|"
		tmptData4 = tmptData4 & rsA("APP_tData4") & "|"
		tmptData5 = tmptData5 & rsA("APP_tData5") & "|"
		rsA.movenext
	  loop
	  if tmpAppID <> "" then
		tmpAppID1 = split(tmpAppID,"|")
		tmpAppIName1 = split(tmpAppIName,"|")
		tmpAppActive1 = split(tmpAppActive,"|")
		tmpAppGroupsR1 = split(tmpAppGroupsR,"|")
		tmpAppGroupsW1 = split(tmpAppGroupsW,"|")
		tmpAppGroupsF1 = split(tmpAppGroupsF,"|")
		tmpAppSubsc1 = split(tmpAppSubsc,"|")
		tmpAppBkMk1 = split(tmpAppBkMk,"|")
		tmpAppSecCode1 = split(tmpAppSecCode,"|")
		tmpiData11 = split(tmpiData1,"|")
		tmpiData12 = split(tmpiData2,"|")
		tmpiData13 = split(tmpiData3,"|")
		tmpiData14 = split(tmpiData4,"|")
		tmpiData15 = split(tmpiData5,"|")
		tmpiData16 = split(tmpiData6,"|")
		tmpiData17 = split(tmpiData7,"|")
		tmpiData18 = split(tmpiData8,"|")
		tmpiData19 = split(tmpiData9,"|")
		tmpiData110 = split(tmpiData10,"|")
		tmptData11 = split(tmptData1,"|")
		tmptData12 = split(tmptData2,"|")
		tmptData13 = split(tmptData3,"|")
		tmptData14 = split(tmptData4,"|")
		tmptData15 = split(tmptData5,"|")
		acnt = ubound(tmpAppID1)-1
		redim arrAppPerms(acnt,23)
		for ag = 0 to acnt
		  arrAppPerms(ag,0) = tmpAppID1(ag)
		  arrAppPerms(ag,1) = tmpAppIName1(ag)
		  arrAppPerms(ag,2) = tmpAppActive1(ag)
		  arrAppPerms(ag,3) = tmpAppGroupsR1(ag)
		  arrAppPerms(ag,4) = tmpAppGroupsW1(ag)
		  arrAppPerms(ag,5) = tmpAppGroupsF1(ag)
		  arrAppPerms(ag,6) = tmpAppSubsc1(ag)
		  arrAppPerms(ag,7) = tmpAppBkMk1(ag)
		  arrAppPerms(ag,8) = tmpAppSecCode1(ag)
		  arrAppPerms(ag,9) = tmpiData11(ag)
		  arrAppPerms(ag,10) = tmpiData12(ag)
		  arrAppPerms(ag,11) = tmpiData13(ag)
		  arrAppPerms(ag,12) = tmpiData14(ag)
		  arrAppPerms(ag,13) = tmpiData15(ag)
		  arrAppPerms(ag,14) = tmpiData16(ag)
		  arrAppPerms(ag,15) = tmpiData17(ag)
		  arrAppPerms(ag,16) = tmpiData18(ag)
		  arrAppPerms(ag,17) = tmpiData19(ag)
		  arrAppPerms(ag,18) = tmpiData110(ag)
		  arrAppPerms(ag,19) = tmptData11(ag)
		  arrAppPerms(ag,20) = tmptData12(ag)
		  arrAppPerms(ag,21) = tmptData13(ag)
		  arrAppPerms(ag,22) = tmptData14(ag)
		  arrAppPerms(ag,23) = tmptData15(ag)
		next
	  end if
	else
	end if
	set rsA = nothing
end function

function bldAppAccess()
	dim tmpAppID, tmpAppActive, tmpAppGroupsR, tmpAppGroupsW, tmpAppGroupsF, tmpAppIName, bHasAccess
	dim tmpAppSubsc, tmpAppBkMk
	dim tmpApp
	tmpApp = ""
	'bHasAccess = true
	sSql = "SELECT * FROM "& strTablePrefix & "APPS"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  do until rsA.eof
	    tmpAppID = tmpAppID & rsA("APP_ID") & "@"
	    tmpAppIName = tmpAppIName & rsA("APP_INAME") & "@"
		tmpAppActive = tmpAppActive & rsA("APP_ACTIVE") & "@"
		tmpAppGroupsR = tmpAppGroupsR & rsA("APP_GROUPS_USERS") & "@"
		tmpAppGroupsW = tmpAppGroupsW & rsA("APP_GROUPS_WRITE") & "@"
		tmpAppGroupsF = tmpAppGroupsF & rsA("APP_GROUPS_FULL") & "@"
		tmpAppSubsc = tmpAppSubsc & rsA("APP_SUBSCRIPTIONS") & "@"
		tmpAppBkMk = tmpAppBkMk & rsA("APP_BOOKMARKS") & "@"
		tmpAppSecCode = tmpAppSecCode & rsA("APP_SUBSEC") & "@"
		tmpiData1 = tmpiData1 & rsA("APP_iData1") & "@"
		tmpiData2 = tmpiData2 & rsA("APP_iData2") & "@"
		tmpiData3 = tmpiData3 & rsA("APP_iData3") & "@"
		tmpiData4 = tmpiData4 & rsA("APP_iData4") & "@"
		tmpiData5 = tmpiData5 & rsA("APP_iData5") & "@"
		tmpiData6 = tmpiData6 & rsA("APP_iData6") & "@"
		tmpiData7 = tmpiData7 & rsA("APP_iData7") & "@"
		tmpiData8 = tmpiData8 & rsA("APP_iData8") & "@"
		tmpiData9 = tmpiData9 & rsA("APP_iData9") & "@"
		tmpiData10 = tmpiData10 & rsA("APP_iData10") & "@"
		tmptData1 = tmptData1 & rsA("APP_tData1") & "@"
		tmptData2 = tmptData2 & rsA("APP_tData2") & "@"
		tmptData3 = tmptData3 & rsA("APP_tData3") & "@"
		tmptData4 = tmptData4 & rsA("APP_tData4") & "@"
		tmptData5 = tmptData5 & rsA("APP_tData5") & "@"
		rsA.movenext
	  loop
	  
	  tmpApp = tmpAppID & ";" & tmpAppIName & ";" & tmpAppActive & ";" & tmpAppGroupsR & ";" & tmpAppGroupsW & ";" & tmpAppGroupsF & ";" & tmpAppSubsc & ";" & tmpAppBkMk & ";" & tmpAppSecCode & ";" & tmpiData1 & ";" & tmpiData2 & ";" & tmpiData3 & ";" & tmpiData4 & ";" & tmpiData5 & ";" & tmpiData6 & ";" & tmpiData7 & ";" & tmpiData8 & ";" & tmpiData9 & ";" & tmpiData10 & ";" & tmptData1 & ";" & tmptData2 & ";" & tmptData3 & ";" & tmptData4 & ";" & tmptData5
	  
	end if
	set rsA = nothing
	
	bldAppAccess = tmpApp
end function



'##############################################
'##            NT Authentication             ##
'##############################################
sub regNTuser()
		strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & "(M_NAME"
		strSql = strSql & ", M_USERNAME"
		strSql = strSql & ", M_PASSWORD"
		strSql = strSql & ", M_EMAIL"
		strSql = strSql & ", M_KEY"
		strSql = strSql & ", M_LEVEL"
		strSql = strSql & ", M_DATE"
		strSql = strSql & ", M_LASTHEREDATE"
		strSql = strSql & ", M_IP" 
		strSql = strSql & ", M_RNAME"
		strSql = strSql & ", M_STATUS"
		strSql = strSql & ", M_GLOW"
		strSql = strSql & ", THEME_ID"
		strSql = strSql & ", M_RECMAIL"
		strSql = strSql & ", M_HIDE_EMAIL"
		strSql = strSql & ", M_TIME_TYPE"
		strSql = strSql & ", M_TIME_OFFSET"
		strSql = strSql & ", M_LCID"
		strsql = strsql & ", M_PHOTO_URL"
		strsql = strsql & ", M_AVATAR_URL"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & Session(strUniqueID & "userID") & "'"
		strSql = strSql & ", " & "'" & Session(strUniqueID & "userID") & "'"
		strSql = strSql & ", " & "'" & pEncrypt(pEnPrefix & Session(strUniqueID & "strNTUserFullName")) & "'"
		strSql = strSql & ", " & "'" & Request.Form("Email") & "'"
		strSql = strSql & ", " & "'" & actkey & "'"
		strSql = strSql & ", 1"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", '" & Request.ServerVariables("REMOTE_HOST") & "'"	
		strSql = strSql & ", 'x'"
		strSql = strSql & ", 1"
		strSql = strSql & ", ''"
		strSql = strSql & ", '" & strDefTheme & "'"
		strsql = strsql & ", '0'"
		strSql = strSql & ", 1"	
		strSql = strSql & ", '" & strTimeType & "'"	
		strSql = strSql & ", " & strTimeAdjust & ""	
		strSql = strSql & ", " & intPortalLCID & ""	
		strSql = strSql & ", 'images/no_photo.gif'"	
		strSql = strSql & ", 'files/avatars/noavatar.gif'"						
		strSql = strSql & ")"
		executeThis(strSql)
		
	'## Updates the member count by 1
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & "SET U_COUNT = (U_COUNT+1) WHERE ID = 1"
	executeThis(strSql)
end sub

sub NTUser()
		Call setMemberSessVars(Session(strUniqueID & "username"),"")
		setMemberVars()
		'if hasAccess(1) then 
		if chkIsAdmin(strUserMemberID) then 
		  Session(strCookieURL & "Approval") = "256697926329"
		end if
end sub

function ChkAccountReg()
  if Session(strUniqueID & "userID") = "" then
	ChkAccountReg = "0"
  else
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_USERNAME "
	strSql = strSql & "FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & "WHERE " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strUniqueID & "userID") & "' " 
	strSql = strSql & "AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"

	set rs_chk = my_conn.Execute(strSql)

	if rs_chk.BOF and rs_chk.EOF then
		'ChkAccountReg = "0"
		call regNTuser()
		ChkAccountReg = "1"
	else
		ChkAccountReg = "1"
	end if
	
	set rs_chk = nothing
  end if
end function

sub NTAuthenticate()
	dim strUser, strNTUser, checkNT
	strNTUser = Request.ServerVariables("AUTH_USER") 
	strNTUser = replace(strNTUser, "\", "/")
	if Session(strUniqueID & "username") = "" then
		strUser = Mid(strNTUser,(instr(1,strNTUser,"/")+1),len(strNTUser))
		Session(strUniqueID & "username") = strUser
	else
		Session(strUniqueID & "username") = Session(strUniqueID & "username")
	end if
	if strNTGroups="1" then
		strNTGroupsSTR = Session(strUniqueID & "strNTGroupsSTR")
		if trim(strNTGroupsSTR) = "" then
			Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
			For Each strNTUserInfoGroup in strNTUserInfo.Groups
				strNTGroupsSTR=strNTGroupsSTR+", "+strNTUserInfoGroup.name
			NEXT
			Session(strUniqueID & "strNTGroupsSTR") = strNTGroupsSTR
		end if
	end if

	strNTUserFullName = Session(strUniqueID & "strNTUserFullName")
	if Session(strUniqueID & "strNTUserFullName") = "" then
	  Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
	  strNTUserFullName=strNTUserInfo.FullName
	  Session(strUniqueID & "strNTUserFullName") = strNTUserFullName
	end if
end sub

sub NTdebug()
	strNTUser = Request.ServerVariables("AUTH_USER") 
	strNTUser = replace(strNTUser, "\", "/")
    'Set strNTUserInfo = GetObject("LDAP://ldapservername/RootDSE")
	Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
	For Each strNTUserInfoGroup in strNTUserInfo.Groups
		strNTGroupsSTR=strNTGroupsSTR+", "+strNTUserInfoGroup.name
	NEXT
	strNTUserFullName=strNTUserInfo.FullName
	
  Response.Write("AUTH_USER: " & Request.ServerVariables("AUTH_USER") & "<br />")
  'Response.Write("userid: " & Session(strUniqueID & "userID") & "<br />")
  'Response.Write("username: " & Session(strUniqueID & "username") & "<br />")
  Response.Write("strNTUserFullName: " & strNTUserFullName & "<br />")
  Response.Write("strNTGroupsSTR: " & strNTGroupsSTR & "<br />")
  'Response.Write("ChkAccountReg: " & ChkAccountReg & "<br />")
  'Response.Write(" " &  & "<br />")
  Response.End()
  'Response.Write(": " &  & "<br />")
end sub

'##############################################
'##         END - NT Authentication          ##
'##############################################
 %>
