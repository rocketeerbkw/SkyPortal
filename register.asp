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
CurPageType = "register"
minPassLen = intMinimumPasswordLength
sInvalid = strInvalidUsernameChars
'strInvalidIPs = "72.158.217.34,65.5.209.138,65.27.4.,"
'intMinUsernameLength
if strInvalidIPs <> "" then
  if right(strInvalidIPs,1) <> "," then
    strInvalidIPs = strInvalidIPs & ","
  end if
  arrInvalidIP = split(strInvalidIPs,",")
  curIP = request.ServerVariables("REMOTE_ADDR")
  for ip = 0 to ubound(arrInvalidIP)-1
    if left(curIP,len(arrInvalidIP(ip))) = arrInvalidIP(ip) then
      closeAndGo("default.asp")
    end if
  next
end if
  
 %>
<!--#include file="inc_functions.asp" -->
<%
CurPageInfoChk = "1"
dogbug=false
actkey = ""
function CurPageInfo ()
	PageName = txtSiteReg
	PageAction = txtViewing & "<br />" 
	PageLocation = "register.asp"
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"
end function

dim newUserEmail, newUser, Err_Msg

function DoCookies2(fName,fPass)
	':: New User - delete any existing cookies and sessions
	Call ClearCookies()
	strSql = "DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP='" & request.ServerVariables("REMOTE_ADDR") & "'"
	executeThis(strSql)
		
	Response.Cookies(strUniqueID & "User").Path = strCookieURL
	Response.Cookies(strUniqueID & "User")("Name") = fName
	Response.Cookies(strUniqueID & "User")("Pword") = fPass
	Response.Cookies(strUniqueID & "User")("Cookies") = "1"
	Response.Cookies(strUniqueID & "User").Expires = dateAdd("d", 30, now())

	Response.Cookies(strUniqueID & "hide").Path = strCookieURL
	Response.Cookies(strUniqueID & "hide")("Name") = fName
	Response.Cookies(strUniqueID & "hide").Expires = dateAdd("d", 30, now())
end function
%>
<!--#include file="inc_top.asp" -->
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<tr>
<td width="200" class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)
Menu_fp()
affiliateBanners()
%>
<script type="text/javascript">
<!--//
function OpenSPreview()
{
	var curCookie = "strSignaturePreview=" + escape(document.form1.Sig.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_portal.asp?cmd=6', 'preview_page', 'scrollbars=yes,width=450,height=250')	
}
//-->
</script>
</td>
<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtSiteRul & "|policy.asp"
  arg2 = strSiteTitle & "&nbsp;" & txtSiteReg & "|register.asp?mode=register"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if strNewReg = "0" and not hasAccess(1) then
	closeAndGo("default.asp")
  elseif Request.QueryString("actkey") <> "" and strEmail = 1 then
	'they submitted an activation key, lets process it
	call processKey()
  else
    select case Request.QueryString("mode")
      case "reqEmail"
	    ':: request email of validation key
		if strEmail = 1 then
	      call emailKey()
		end if
      case "DoIt"
		' they submitted a registration form. lets process it.
		call processRegFormSubmission()
      case "Register"
		'Lets show the registration form
		call ShowForm()
      case else
		'Lets show the registration form
		'call ShowForm()
		closeAndGo("default.asp")
    end select
  end if
  
'if Request.QueryString("mode") = "reqEmail" then
	':::::::::: request email of lost key ::::::::::::::
	'call emailKey()
'elseif Request.QueryString("mode") <> "DoIt" and Request.QueryString("actkey") = "" then
	'if strNewReg = "0" and not hasAccess(1) then
		'closeAndGo("default.asp")
	'else
		'Lets show the registration form
		'call ShowForm()
	'end if		
'elseif Request.QueryString("actkey") <> "" and lcase(strEmail) = 1 then
	'they submitted an activation key, lets process it
	'call processKey()
'else 
	' they submitted a registration form. lets process it.
	'call processRegFormSubmission()
'end if %>
</td></tr></table>
<!--#include file="inc_footer.asp" -->
<% 
function chkNewUserName(n)
	tMsg = ""
	if len(n & "x") = 1 then 
	  tMsg = tMsg & "<li>" & txtChoseUsrNam & "</li>"
	end if
	
	if len(n) < intMinUsernameLength then 
	  tMsg = tMsg & "<li>" & replace(txtLongerUsrName,"[%min%]",intMinUsernameLength) & "</li>"
	end if
	
	if not chkValidUserName(n) then
		tMsg = tMsg & "<li>" & txtCharsNotAllow
		tMsg = tMsg & "<br/><b><span class=""fAlert"">"
		tMsg = tMsg & replace(strInvalidUsernameChars,",","&nbsp;")
		tMsg = tMsg & "&nbsp;," & "</b></span></li>"
	else
	  if tMsg = "" then
		strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS WHERE M_NAME = '" & Trim(chkString(n,"sqlstring")) &"'"
		set rs = my_Conn.Execute(strSql)
		if not rs.EOF then
			tMsg = tMsg & "<li>" & txtChsAnother & "</li>"
		end if
		rs.close
		set rs = nothing
	
		if len(Request.Form("Referrer")) > 0 then
			strSql = "SELECT " & strDBNTSQLName & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(Trim(Request.Form("Referrer")), "refer") &"'"
			set rs = my_Conn.Execute(strSql)	
			if rs.BOF and rs.EOF then 
				tMsg = tMsg & "<li>" & txtBadRefer & ".</li>"
			end if
			rs.close
			set rs = nothing
			if Trim(n) = Trim(Request.Form("Referrer")) then
				tMsg = tMsg & "<li>" & txtReferNotU & ".</li>"
			end if
		end if
	  end if
	end if
	chkNewUserName = tMsg
end function

sub processRegFormSubmission()
	Err_Msg = ""
	
	Err_Msg = chkNewUserName(trim(Request.Form("Name")))
	
	'## NT authentication no additional password needed
	if strAuthType = "db" then
		if trim(Request.Form("Password")) = "" then 
			Err_Msg = Err_Msg &  "<li>" & txtChosPswd & "</li>"
		end if
		if Len(Request.Form("Password")) > 25 or Len(Request.Form("Password")) < minPassLen then 
			Err_Msg = Err_Msg & "<li>" & replace(txtUPassLen,"[%min%]",minPassLen) & "</li>" 
		end if
		if Request.Form("Password") <> Request.Form("Password2") then 
			Err_Msg = Err_Msg & "<li>" & txtPassNoMatch & ".</li>"
		end if
		if (Instr(Request.Form("Password"), ">") > 0 ) or (Instr(Request.Form("Password"), "<") > 0) or (Instr(Request.Form("Password"), ",") > 0) or (Instr(Request.Form("Password"), "&") > 0) or (Instr(Request.Form("Password"), "#") > 0) or (Instr(Request.Form("Password"), "'") > 0) then
			Err_Msg = Err_Msg & "<li>" & txtCharsNoAllow & "&nbsp;" & txtPass & "</li>"
		end if
	end if

	If strAutoLogon <> 1 then
		if Request.Form("Email") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtErNoEmlAdd & "</li>"
		end if
		if EmailField(Request.Form("Email")) = 0 then 
			Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
		end if
	end if

	if (Request.Form("Email") <> Request.Form("Email2")) and strEmail = 1 and (strEmailVal = 2 or strEmailVal = 4 or strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 8) then
		Err_Msg = Err_Msg & "<li>" & txtEmlNoMatch & "</li>"
	end if
	
	'iEmlVal = 
	if not validate_email(Request.Form("Email")) then
		Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
	end if
	
	if strAuthType <> "db" then
	  if ChkAccountReg = 1 then
		Err_Msg = Err_Msg & "<li>" & txtNTusrReg & ".</li>"
	  end if
	end if

	if strUniqueEmail = 1 and Err_Msg = "" then
		strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE M_EMAIL = '" & Trim(chkString(Request.Form("Email"),"sqlstring")) &"'"
		set rs = my_Conn.Execute (strSql)
		if rs.BOF and rs.EOF then 
			'## Do Nothing
		else 
			Err_Msg = Err_Msg & "<li>" & txtEmlInUse & "</li>"
		end if
		rs.close
		set rs = nothing
		if lcase(strEmail) = "1" then
			'
			strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
			strSql = strSql & " WHERE M_EMAIL = '" & Trim(chkString(Request.Form("Email"),"sqlstring")) & "'"
			set rs = my_Conn.Execute (strSql)
			if rs.BOF and rs.EOF then 
				'## Do Nothing
			else
				Err_Msg = Err_Msg & "<li>" & txtEmlInUse & "</li>"
			end if
			rs.close
			set rs = nothing
		end if
	end if
	
	
	
  if showRegisterLongForm then
    '########## check for valid date entry in 'age' (if applicable) ##########
    if strAge = 1 then
	 formbirthdate = " "
      If trim(Request.Form("B_Month")) <> "" and trim(Request.Form("B_Day")) <> "" and trim(Request.Form("B_Year")) <> "" then
        formbirthdate = chkString(Request.Form("B_Month"),"sqlstring") & "/" & chkString(Request.Form("B_Day"),"sqlstring") & "/" & chkString(Request.Form("B_Year"),"sqlstring")
         ' Check to see if birthdate is a valid date
	    If NOT IsDate(formbirthdate) Then
		  Err_Msg = Err_Msg & "<li>" & txtValBday & ".</li>"
	    End If 
	    If IsDate(formbirthdate) then
          if CDate(formbirthdate) > CDate(strCurDateAdjust) then
          Err_Msg = Err_Msg & "<li>" & txtBdayPrior & "</li>"          
          end if
        end if
      end if
    end if
    '############## End Validate Age ###############

	if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
		Err_Msg = Err_Msg & "<li>" & txtPrefixUrl & "</li>"
	end if
	if Len(Request.Form("Sig")) > 255 then
		Err_Msg = Err_Msg & "<li>" & txtSigTooLng & "<br />"
		Err_Msg = Err_Msg & "" & txtLenIs & "&nbsp;<b>" & Len(Request.Form("Sig")) & "</b>.</li>"
	end if
  end if 'show long form
  
	if SecImage > 0  then
	  if not DoSecImage(Ucase(request.form("SecCode"))) Then
	    Err_Msg = Err_Msg & "<li>" & txtBadSecCode & "</li>"
	  end if
	end if
	
	if Err_Msg <> "" then
		call showErrMsg()
	else
		if Request.Form("Homepage") <> "" and lcase(Request.Form("Homepage")) <> "http://" and lcase(Request.Form("Homepage")) <> "https://" then
			regHomepage = chkString(Request.Form("Homepage"),"sqlstring")
		else
			regHomepage = " "
		end if
		actkey = GetKey("none")

	  if not showRegisterLongForm then
		strSql = "INSERT INTO " & strMemberTablePrefix 
		if (strEmail = 1 and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8)) or ((Request.Form("reservation") = "yes" or Request.Form("active") = 0 or Request.Form("sendinvite") = "yes") and hasAccess(1))  then
			strSql = strSql & "MEMBERS_PENDING "
		else
			strSql = strSql & "MEMBERS "
		end if
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
		strSql = strSql & "'" & ChkString(Request.Form("Name"),"sqlstring") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Name"),"sqlstring") & "'"
		strSql = strSql & ", " & "'" & pEncrypt(pEnPrefix & chkString(Request.Form("Password"),"sqlstring")) & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Email"),"email") & "'"
		strSql = strSql & ", " & "'" & actkey & "'"
		if lcase(strEmail) = "1" and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) and Request.Form("reservation") <> "yes" then
			strSql = strSql & ", " & "-1"
		elseif (Request.Form("reservation") = "yes" or Request.Form("sendinvite") = "yes") and hasAccess(1) then
			strSql = strSql & ", -7"
		else
			strSql = strSql & ", 1"
		end if	
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", '" & Request.ServerVariables("REMOTE_HOST") & "'"	
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Referrer"),"sqlstring") & "'"
		if (Request.Form("reservation") = "yes" and hasAccess(1)) then
			strSql = strSql & ", " & "0"
		else
			strSql = strSql & ", 1"
		end if
		strSql = strSql & ", ''"
		strSql = strSql & ", '" & strDefTheme & "'"
		if isadmin = true and sendinviteemail = true then
		  strsql = strsql & ", '1'"
		else
		  strsql = strsql & ", '0'"
		end if	
		strSql = strSql & ", 1"	
		strSql = strSql & ", '" & strTimeType & "'"	
		strSql = strSql & ", 0"	
		strSql = strSql & ", " & intPortalLCID & ""	
		strSql = strSql & ", 'images/no_photo.gif'"	
		strSql = strSql & ", 'files/avatars/noavatar.gif'"						
		strSql = strSql & ")"
			
	  else
		strSql = "INSERT INTO " & strMemberTablePrefix 
		if (lcase(strEmail) = "1" and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8)) or ((Request.Form("reservation") = "yes" or Request.Form("sendinvite") = "yes") and hasAccess(1))  then
			strSql = strSql & "MEMBERS_PENDING "
		else
			strSql = strSql & "MEMBERS "
		end if
		strSql = strSql & "(M_NAME"
		strSql = strSql & ", M_USERNAME"
		strSql = strSql & ", M_KEY"
		strSql = strSql & ", M_LEVEL"
		strSql = strSql & ", M_PASSWORD"
		strSql = strSql & ", M_EMAIL"
		strSql = strSql & ", M_HIDE_EMAIL"
		strSql = strSql & ", M_DATE"
		strSql = strSql & ", M_COUNTRY"
		strSql = strSql & ", M_SIG"
		strSql = strSql & ", M_YAHOO"
		strSql = strSql & ", M_ICQ"
		strSql = strSql & ", M_AIM"
		strSql = strSql & ", M_POSTS"
		strSql = strSql & ", M_HOMEPAGE"
		strSql = strSql & ", M_LASTHEREDATE"
		strSql = strSql & ", M_STATUS"
		strSql = strSql & ", M_IP" 
		strSql = strSql & ", M_FIRSTNAME" 
		strSql = strSql & ", M_LASTNAME"
		strsql = strsql & ", M_CITY" '#20
		strsql = strsql & ", M_STATE"
		strsql = strsql & ", M_PHOTO_URL"
		strsql = strsql & ", M_AVATAR_URL"		
		strsql = strsql & ", M_LINK1" 
		strSql = strSql & ", M_LINK2"
		strSql = strsql & ", M_AGE"
		strSql = strSql & ", M_MARSTATUS"
		strSql = strsql & ", M_SEX"
		strSql = strSql & ", M_OCCUPATION" 
		strSql = strSql & ", M_BIO" '#30
		strSql = strSql & ", M_HOBBIES"
		strsql = strsql & ", M_LNEWS"
		strSql = strSql & ", M_QUOTE"
		strSql = strSql & ", M_RECMAIL"
		strSql = strSql & ", M_RNAME"
		strSql = strSql & ", M_MSN"
		strSql = strSql & ", M_ZIP"
		strSql = strSql & ", M_GLOW"
		strSql = strSql & ", THEME_ID"
		strSql = strSql & ", M_TIME_TYPE"
		strSql = strSql & ", M_TIME_OFFSET"
		strSql = strSql & ", M_LCID"
		strSql = strSql & ") VALUES ("
		If strAutoLogon = 1 then
			strSql = strSql & "'" & Session(strUniqueID & "strNTUserFullName") & "'"
		Else
			strSql = strSql & "'" & ChkString(Request.Form("Name"),"sqlstring") & "'"
		end if
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Name"),"sqlstring") & "'"
		strSql = strSql & ", " & "'" & chkString(actkey,"") & "'"
		if lcase(strEmail) = "1" and (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) and Request.Form("reservation") <> "yes" then
			strSql = strSql & ", " & "-1"
		elseif (Request.Form("reservation") = "yes" or Request.Form("sendinvite") = "yes") and hasAccess(1) then
			strSql = strSql & ", -7"
		else
			strSql = strSql & ", 1"
		end if
		strSql = strSql & ", " & "'" & pEncrypt(pEnPrefix & chkString(Request.Form("Password"),"sqlstring")) & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Email"),"email") & "'"
		strSql = strSql & ", " & ChkString(Request.Form("HideMail"), "sqlstring")
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Country"),"sqlstring") & " '"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Sig"),"message") & " '"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("YAHOO"),"sqlstring") & " '"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("ICQ"),"sqlstring") & " '"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("AIM"),"sqlstring") & " '"
		strSql = strSql & ", " & "0"
		strSql = strSql & ", " & "'" & ChkString(htmlencode(regHomepage),"display") & " '"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		if (Request.Form("reservation") = "yes" and hasAccess(1)) then
			strSql = strSql & ", " & "0"
		else
			strSql = strSql & ", " & "1"
		end if
			strSql = strSql & ", '" & Request.ServerVariables("REMOTE_HOST") & "'"
		if strfullName = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("FirstName"),"sqlstring") & "'" 
			strSql = strSql & ", '" & ChkString(Request.Form("LastName"),"sqlstring") & "'"  
		else
			strSql = strSql & ", ''" 
			strSql = strSql & ", ''"  
		end if
		if strCity = "1" then '#20
			strsql = strsql & ", '" & ChkString(Request.Form("City"),"sqlstring") & "'"    
		else
			strsql = strsql & ", ''"
		end if
		if strState = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("State"),"sqlstring") & "'" 
		else
			strsql = strsql & ", ''" 
		end if
		if strPicture = "1" then
			strsql = strsql & ", '" & ChkString(htmlencode(Request.Form("Photo_URL")),"display") & "'"  
		else
			strsql = strsql & ", ''"  
		end if
		strsql = strsql & ", '" & ChkString(htmlencode(Request.Form("Avatar_URL")),"display") & "'"  		
		if strFavLinks = "1" then
			strsql = strsql & ", '" & ChkString(htmlencode(Request.Form("LINK1")),"display") & "'"
			strSql = strSql & ", '" & ChkString(htmlencode(Request.Form("LINK2")),"display") & "'"
		else
			strsql = strsql & ", ''"
			strSql = strSql & ", ''"  
		end if
		if strAge = 1 then
			 strSql = strsql & ", '" & formbirthdate & "'"
		else
			 strSql = strsql & ", ''"
		end if
		if strMarStatus = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("MarStatus"),"sqlstring") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strSex = "1" then
			strSql = strsql & ", '" & ChkString(Request.Form("Sex"),"sqlstring") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strOccupation = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Occupation"),"sqlstring") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strBio = "1" then '#30
			strSql = strSql & ", '" & ChkString(Request.Form("Bio"),"sqlstring") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strHobbies = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Hobbies"),"sqlstring") & "'" 
		else
			strSql = strSql & ", ''" 
		end if
		if strLNews = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("LNews"),"sqlstring") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strQuote = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Quote"),"sqlstring") & "'"
		else 
			strSql = strSql & ", ''" 
		end if
		
		strSql = strSql & ", " & "'" & ChkString(Request.Form("recmail"),"sqlstring") & "'"		
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Referrer"),"sqlstring") & "'"
		if strMSN = "1" then
			if Request.Form("MSN") <> "" then
			strSql = strSql & ", '" & ChkString(Request.Form("MSN"),"sqlstring") & "'"
			else
			strSql = strSql & ", ''"
			end if
		else 
			strSql = strSql & ", ''" 
		end if	
		if strZip = "1" then
		  if trim(Request.Form("Zipcode")) <> "" then
			strsql = strsql & ", '" & ChkString(Request.Form("Zipcode"),"sqlstring") & "'" 
		  else
			strsql = strsql & ", ' '" 
		  end if
		else
		  strsql = strsql & ", ' '" 
		end if	
		strSql = strSql & ", ''"
		strSql = strSql & ", '" & strTheme & "'"	
		strSql = strSql & ", '" & strTimeType & "'"	
		strSql = strSql & ", 0"	
		strSql = strSql & ", " & intPortalLCID & ""						
		strSql = strSql & ")"


	  end if 'end if long or short form
'		on error resume next
		'response.Write(strSql)
		'response.End()
		executeThis(strSql)
		
		if trim(Request.Form("Referrer")) <> "" then
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strsql = strsql & "SET M_RTOTAL = M_RTOTAL + 1 "
		strsql = strsql & ",    M_GOLD = M_GOLD + 5 "		
		strsql = strsql & ",    M_REP = M_REP + 5 "		
		strSql = strSql & " WHERE MEMBER_ID=" & getMemberID(ChkString(Request.Form("Referrer"),"sqlstring"))
	'	on error resume next
		executeThis(strSql)
		end if
	  
	  'send the emails as needed	
		select case strEmailVal
		  case 1
			docount
		  case 2
			sendWelcomeEmail("r")
			docount
		  case 3
		  	sendAdminNewMemberEmail(chkString(Request.Form("name"),"sqlstring"))
			docount
		  case 4
			sendWelcomeEmail("r")
		  	sendAdminNewMemberEmail(chkString(Request.Form("name"),"sqlstring"))
			docount
		  case 5
			sendWelcomeEmail("r") 'validate key
		  case 6
			sendWelcomeEmail("r") 'validate key
		  case 7
		  	sendWelcomeEmail("r")
			sendAdminNewMemberEmail(chkString(Request.Form("name"),"sqlstring"))
		  case 8
			sendWelcomeEmail("r") 'validate key
		end select
		
		if strEmail = "1" and Request.Form("sendinvite") = "yes" and hasAccess(1) then
			call sendAdminInviteEmail()
		end if
		
	end if 'end err_msg check
	
	call postRegTextPM()

		if strAuthType = "db" and not hasAccess(2) then
		  select case chkIsMbr(chkString(Request.Form("Name"),"sqlstring"), chkString(pEncrypt(pEnPrefix & Request.Form("Password")),"sqlstring"))
			case 1
				Call DoCookies("true")
			case else
				':: do nothing
		  end select
		end if
		
		If strAutoLogon = 1 then
  			closeAndGo("default.asp")
		Else
		  if hasAccess(1) then %>
			<p align="center"><a href="register.asp?mode=register">Add another?</a></p>
<%		  else %>
			<meta http-equiv="Refresh" content="2; URL=default.asp">
		  <%
		  end if
		End if %>
    <p align="center">
	<a href="default.asp"><%= txtHome %></a></p>
<% 
end sub

function displayEmail(e)
  displayEmail = replace(e,"@",icon(icnAt,"","","","align=""middle"""),1,-1,1)
end function

sub DoCount
	'## Updates the member count by 1
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & "SET U_COUNT = U_COUNT + 1"
	executeThis(strSql)
end sub

sub ShowForm()
  spThemeTitle = strSiteTitle & " " & txtSiteReg
  spThemeBlock1_open(intSkin)
	if strAuthType <> "db" then
	  if ChkAccountReg = "1" then 
		'the NT user is already in the database %>
		<p align="center"><b><%= txtNTRegNoNec %>.</b></p>
		<table align="center">
  		  <tr>
    		<td>
    		  <ul>
      			<li><%= txtNTusrReg %>.</li>
    		  </ul>
    		</td>
  		  </tr>
		</table>
		<% 
		End If %>
<%  else %>
	<form action="register.asp?mode=DoIt" method="post" id="form1" name="Form1">
	 <table width="100%" border="0" align="center">
 	  <tr><td><input name="Refer" type="hidden" value="<% =chkString(Request.Form("Refer"), "display") %>" />
	  <% if showRegisterLongForm then %>
	   <!--#include file="includes/inc_profile.asp" -->
	  <% Else
	  		showRegisterShortForm()
	     End If %>
	   </td></tr>
	 </table>
	</form>
<%	end if
  spThemeBlock1_close(intSkin)
end sub

sub showRegisterShortForm()%>
	<table border="0" width="100%" cellspacing="2" cellpadding="0">
	  <tr>
	    <td align="center" colspan="2">
		<p><b><%= txtReg1a %>&nbsp;<span class="fAlert"><b>*</b></span>&nbsp;<%= txtReg1b %></b>
<%		if lcase(strEmail) = "1" And (strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 or strEmailVal = 8) then
			If Request.Querystring("mode") = "Register" Then %>
				<br /><span class="fAlert"><%= txtReg2a %></span>.</p><p><%= txtReg3a %>&nbsp;<%= displayEmail(strSender) %>.<br /><%= txtReg3b %>&nbsp;"[no-spam]"<br /><%= txtReg3c %>.<br /><br /></p>
<%			else %>
				<br /><%= txtReg2b %>.</p><p><%= txtReg3a %>&nbsp;<%= displayEmail(strSender) %>.<br /><%= txtReg3b %>&nbsp;"[no-spam]"<br /><%= txtReg3c %>.<br /><br /></p>
<%      		
			end if
		end if%><!-- S k y D o g g - S k y P o r t a l - is here - december 2006-->
	    </td>
	  </tr>
  	  <tr>
		<td colspan="2" valign="top">&nbsp;</td>
  	  </tr>
<%
'<!-- :::::::::::::: start BASICS info ::::::::::::::::: --> %>
      <tr> 
        <td valign="top" align="center" colspan="2" class="tSubTitle"><b><%= txtSiteReg %></b></td>
      </tr>
  	  <tr>
		<td colspan="2" valign="top">&nbsp;</td>
  	  </tr>
      <% if trim(Request.QueryString("rname") <> "") then %>
        <tr> 
          <td width="40%" align="right" nowrap="nowrap"><b><%= txtRefer %>:&nbsp;</b></td>
          <td align="left" nowrap="nowrap"> 
	  <%  if hasAccess(1) then %>
            <input name="Referrer" size="25" maxlength="90" value="" />
      <%  else %>
            <%= ChkString(Request.Querystring("rname"), "sqlstring") %>
          <input type="hidden" name="Referrer" value="<%= ChkString(Request.Querystring("rname"),"sqlstring") %>" />
          </td>
        </tr>
	  <%  end if 
		end if%>
        <tr> 
          <td align="right" nowrap="nowrap" width="40%" class="fNorm">
		  <b><span class="fAlert">*</span><%= txtUsrNam %>:&nbsp;</b></td>
          <td>
		  <input name="Name" size="25" maxlength="90"  value="" />
		</td></tr>
      <%
		if strAuthType <> "db" then %>
          <tr> 
            <td align="right" valign="top" class="fNorm" nowrap="nowrap">
		  	  <b><span class="fAlert">*</span><%= txtUAcct %>:&nbsp;</b></td>
          	<td> 
            <%if hasAccess(1) then %>
            <input name="Account" value="<%= Session(strUniqueID & "userID") %>" size="20" />
            <%else %>
            <%=Session(strUniqueID & "userID")%> 
            <input type="hidden" name="Account" value="<%= Session(strUniqueID & "userID") %>" />
            <%end if %>
            </td>
          </tr>
        <%
		else %>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap"><b><span class="fAlert"><b>*</b></span> 
            <%= txtPass %>:&nbsp;</b><br /><%= replace(txtMinChars,"[#]",minPassLen) %></td>
          <td> 
            <input name="password" type="password" size="25" maxlength="25" value="" />
            </td>
        </tr>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap"><b><span class="fAlert"><b>*</b></span> 
            <%= txtPassAgn %>:&nbsp;</b></td>
          <td> 
            <input name="password2" type="password" value="" size="25" />
            </td>
        </tr>
        <%
		end if 
		%>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap"><span class="fAlert"><b>*</b></span><b><%= txtEmlAdd %>:&nbsp;</b></td>
          <td>
            <input name="Email" size="25" maxlength="90" value="" />
            </td>
        </tr>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap"><b><span class="fAlert">*</span><%= txtCfmEml %>:&nbsp;</b></td>
          <td>
            <input name="Email2" size="25" maxlength="90" value="" />
            </td>
        </tr>
		<%
		if hasAccess(1) then %>
          <td align="right" class="fNorm" nowrap="nowrap"><b><%= txtActive %>:&nbsp;</b></td>
          <td> 
            <select name="active">
              <option value="1" selected="selected"><%= txtActive %></option>
              <option value="0"><%= txtPend %></option>
            </select>
          </td>
        </tr>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap">
		  <b><%= txtResUNam %>:</b>&nbsp;</td>
          <td>
		  <input type="checkbox" Value="yes" name="reservation" />
          </td>
        </tr>
        <tr> 
          <td align="right" class="fNorm" nowrap="nowrap"><b><%= txtEmlNewUsr %>:</b>&nbsp;</td>
          <td><input type="checkbox" Value="yes" name="sendinvite" />
          </td>
        </tr>
        <%
		end if
	  %>
	<tr><td nowrap="nowrap" class="fNorm" align="center" valign="middle" colspan="2">&nbsp;
<%  If SecImage > 0 Then %>
	<br /><%= txtEntrSecImg %><br />
	<img align="absolute" src="<%= strHomeUrl %>includes/securelog/image.asp" /><br />
	<input type="text" name="secCode" size="8" maxLength="8" value="" />
<%  end if %>
	</td>
	</tr>
	<tr><td align="center" valign="middle" colspan="2">
        <p><input type="submit" value="  <%= txtSubmit %>  " name="Submit1" class="button" /></p>
	</td>
	</tr>
	</table><%
end sub


sub emailKey()
	'::::::::::::::::: request email of lost key ::::::::::::::::::::::::::
	Err_Msg = ""
	reqEmailAddress = chkString(Request.Form("reqEmailAddress"),"sqlstring")
	if reqEmailAddress = "" then 
		Err_Msg = Err_Msg & "<li>" & txtErNoEmlAdd & "</li>"
	end if
	if EmailField(reqEmailAddress) = 0 then 
		Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
	end if

	if Err_Msg = "" then
	  strSql = "SELECT * " & _
		" FROM " & strMemberTablePrefix & "MEMBERS_PENDING " & _
		" WHERE M_EMAIL = '" & reqEmailAddress & "'"
	  set rsReqM = my_Conn.Execute (strSql)
	  if rsReqM.EOF or rsReqM.BOF then
		Err_Msg = Err_Msg & "<li>" & txtEmlNoExist & ".</li>"
	  else

		if strEmail = 1 then
			strRecipientsName = rsReqM("M_NAME")
			strRecipients = rsReqM("M_EMAIL")
			strFrom = strSender
			strFromName = strSiteTitle
			strsubject = strSiteTitle & " " & txtRegActReq
			strMessage = chkString(Request.Form("name"),"sqlstring") & vbCrLf & vbCrLf
			strMessage = strMessage & replace(replace(txtEmlVal2,"[%sitetitle%]",strSiteTitle),"[%siteurl%]",strHomeURL) & vbCrLf & vbCrLf
			if strAuthType="db" then
				if strEmailVal = "1" then
					strMessage = strMessage & txtEmlVal3 & "." & vbNewline
					strMessage = strMessage & strHomeURL & "register.asp?actkey=" & rsReqM("M_KEY") & vbNewline & vbNewline
				else
					'strMessage = strMessage & "Password: " & Request.Form("Password") & vbCrLf & vbCrLf
				end if
			end if
			strMessage = strMessage & txtEmlVal4 & "." & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal5
			sendOutEmail strRecipients,strSubject,strMessage,2,0
  	    end if %>
		
		<p align="center"><span class="fTitle">Your Activation Email has been sent</span></p>
		<p align="center"><%= replace(txtEmlVal6,"[%email%]",rsReqM("M_EMAIL")) %>.</p>
		<p align="center"><a href="default.asp"><%= txtHome %></a></p>

<%		end if  'rsReqM.EOF or rsReqM.BOF
	
		rsReqM.close
		set rsReqM = nothing
	  else '(Err_Msg <> "") %>
		<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
		<table align="center" border="0">
	  	 <tr>
	      <td>
			<ul>
			  <% =Err_Msg %>
			</ul>
	       </td>
	  	  </tr>
		 </table>
		 <p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBackData %></a></p>
<%	end if  '(Err_Msg <> "")  %>
<!--#include file="inc_footer.asp" -->
<%
response.end
end sub 'request email of actkey

'############################################################################################
sub processKey()
	'They clicked the link from the email validate, Site email is turned ON.
	'get the activation key
	key = chkString(Request.QueryString("actkey"),"sqlstring")
	if len(key) <> 10 then
	  response.Write("<span class=""fTitle"">Not a valid activation key</span>")
	  closeAndGo("stop")
	end if
	'check the activation key
	strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE M_KEY = '" & key & "'"
	set rsKey = my_Conn.Execute (strSql)
	'response.Write(key)
	
	if rsKey.EOF then	' Key was not found in MEMBERS_PENDING table
		'Check if member has already been validated
		strSql = "SELECT MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_KEY = '" & key & "'"
		set rsKey2 = my_Conn.Execute (strSql)
		if rsKey2.EOF then 	'Key was not found in MEMBERS table %>
		<form method="POST" action="register.asp?mode=reqEmail">
			<p align="center"><span class="fTitle"><b><%= txtRegKey2 %></b></span></p>
			<p align="center"><%= txtRegKey3 %><br />
			<%= txtRegKey4 %><br />
			 <br /><input class="textbox" name="reqEmailAddress" size="30" maxlength="90" />&nbsp;<input type="submit" value="<%= txtReqActEml %>" name="Submit" class="button" /><br /><br />
			 <%= replace(replace(txtRegKey5,"[%a%]","<a href=""mailto:" & strSender & """>"),"[%/a%]","</a>") %></p>
		</form>
			<p align="center"><a href="default.asp"><%= txtHome %></a></p>
<%		else 'Key was found in MEMBERS table - account already activated %>
			<p align="center"><span class="fTitle"><b><%= txtRegKey6 %></b></span></p>
			<p align="center"><%= txtRegKey7 %></p>
    		<p align="center"><a href="default.asp"><%= txtHome %></a></p>
<%  	end if 'end - key check in MEMBERS table
		rsKey2.close
		set rsKey2 = nothing
	elseif strComp(key, rsKey("M_KEY")) <> 0 then 'Key was found in the MEMBERS PENDING table, lets check to see if they match.
		'Keys didn't match %>
		<form method="POST" action="register.asp?mode=reqEmail">
		<p align="center"><%= txtRegKey8 %><br />
		<br /><input class="textbox" name="reqEmailAddress" size="30" maxlength="90" /> <input type="submit" value="Request Activation Email" name="Submit" class="button" /><br /><br />
		<p align="center"><a href="default.asp"><%= txtHome %></a></p>
		</form>
<% 		rsKey.close
		set rsKey = nothing
	 else 'Keys match
		newUser = ChkString(rsKey("M_NAME"),"name")				
		newPass = ChkString(rsKey("M_PASSWORD"),"password")
		newUserEmail = ChkString(rsKey("M_EMAIL"),"email")
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::    add in admin authorize checking here
':::::::	member has recieved registration email and clicked on the activation link.
':::::::	member has validated, we are ready to move them into the MEMBERS table.
':::::::	2 options... 
':::::::	Send admin an email for admin validate or insert member into MEMBERS table
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	 'response.Write("strEmailVal: " & strEmailVal)
	  if strEmailVal = 8 then
	  	'update the MEMBERS_PENDING table to show the member has validated their email
		'and are now awaiting admon approval
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " SET M_LEVEL = -2"
		strSql = strSql & " WHERE M_KEY = '" & key & "'"

		my_Conn.Execute (strSql)
		%>
		<p align="center"><span class="fTitle"><b><%= txtRegKey9 %></b></span><br /><%= txtRegKey10 %></p>
		<% 
	'	sSql = "select M_NAME from " & strMemberTablePrefix & "MEMBERS_PENDING where M_KEY = '" & key & "'"
	'	set rsName = my_Conn.Execute (sSql)
	'	   mName = rsName("M_NAME")
	'	set rsName = nothing
		 sendWelcomeEmail("v")
		 sendAdminNewMemberEmail(newUser)
		
	  else
	 'response.Write("key: " & key)
		'## Move member info to MEMBERS table
		strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & "(M_NAME"
		strSql = strSql & ", M_USERNAME"
		strSql = strSql & ", M_PASSWORD"
		strSql = strSql & ", M_EMAIL"
		strSql = strSql & ", M_HIDE_EMAIL"
		strSql = strSql & ", M_DATE"
		strSql = strSql & ", M_COUNTRY"
		strSql = strSql & ", M_SIG"
		strSql = strSql & ", M_YAHOO"
		strSql = strSql & ", M_ICQ"
		strSql = strSql & ", M_AIM"
		strSql = strSql & ", M_POSTS"
		strSql = strSql & ", M_HOMEPAGE"
		strSql = strSql & ", M_LASTHEREDATE"
		strSql = strSql & ", M_STATUS"
		strSql = strSql & ", M_IP"
		strSql = strSql & ", M_FIRSTNAME" 
		strSql = strSql & ", M_LASTNAME"
		strsql = strsql & ", M_CITY"
		strsql = strsql & ", M_STATE"
		strsql = strsql & ", M_PHOTO_URL"
		strsql = strsql & ", M_AVATAR_URL"		
		strsql = strsql & ", M_LINK1" 
		strSql = strSql & ", M_LINK2"
		strSql = strsql & ", M_AGE"
		strSql = strSql & ", M_MARSTATUS"
		strSql = strsql & ", M_SEX"
		strSql = strSql & ", M_OCCUPATION" 
		strSql = strSql & ", M_BIO"
		strSql = strSql & ", M_HOBBIES"
		strsql = strsql & ", M_LNEWS"
		strSql = strSql & ", M_QUOTE"
		strSql = strSql & ", M_RECMAIL"
		strSql = strSql & ", M_RNAME"
		strSql = strSql & ", M_MSN"
		strSql = strSql & ", M_ZIP"
		strSql = strSql & ", M_GLOW"
		strSql = strSql & ", M_TIME_TYPE"
		strSql = strSql & ", M_TIME_OFFSET"
		strSql = strSql & ", M_LCID"
		strSql = strSql & ") "
		strSql = strSql & " VALUES ("
		strSql = strSql & "'" & ChkString(rsKey("M_NAME"),"name") & "'"
		strSql = strSql & ", " & "'" & ChkString(rsKey("M_NAME"),"name") & "'"
		strSql = strSql & ", " & "'" & ChkString(rsKey("M_PASSWORD"),"password") & "'"
		strSql = strSql & ", " & "'" & ChkString(rsKey("M_EMAIL"),"email") & "'"
		strSql = strSql & ", " & "'" & rsKey("M_HIDE_EMAIL") & "'"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & "'" & rsKey("M_COUNTRY") & " '"
		strSql = strSql & ", " & "'" & rsKey("M_SIG") & " '"
		strSql = strSql & ", " & "'" & rsKey("M_YAHOO") & " '"
		strSql = strSql & ", " & "'" & rsKey("M_ICQ") & " '"
		strSql = strSql & ", " & "'" & rsKey("M_AIM") & " '"
		strSql = strSql & ", " & "0"
		strSql = strSql & ", " & "'" & rsKey("M_HOMEPAGE") & " '"
		strSql = strSql & ", " & "'" & strCurDateString & "'"
		strSql = strSql & ", " & "1"
		strSql = strSql & ", '" & Request.ServerVariables("REMOTE_HOST") & "'" 
		strSql = strSql & ", '" & rsKey("M_FIRSTNAME") & "'" 
		strSql = strSql & ", '" & rsKey("M_LASTNAME") & "'"  
		strsql = strsql & ", '" & rsKey("M_CITY") & "'"    
		strsql = strsql & ", '" & rsKey("M_STATE") & "'" 
		strsql = strsql & ", '" & rsKey("M_PHOTO_URL") & "'"  
		strsql = strsql & ", '" & rsKey("M_AVATAR_URL") & "'"  	
		strsql = strsql & ", '" & rsKey("M_LINK1") & "'"
		strSql = strSql & ", '" & rsKey("M_LINK2") & "'"
		strSql = strsql & ", '" & rsKey("M_AGE") & "'" 
		strSql = strSql & ", '" & rsKey("M_MARSTATUS") & "'"
		strSql = strsql & ", '" & rsKey("M_SEX") & "'"
		strSql = strSql & ", '" & rsKey("M_OCCUPATION") & "'"
		strSql = strSql & ", '" & rsKey("M_BIO") & "'"
		strSql = strSql & ", '" & rsKey("M_HOBBIES") & "'" 
		strsql = strsql & ", '" & rsKey("M_LNEWS") & "'"
		strSql = strSql & ", '" & rsKey("M_QUOTE") & "'"
		strSql = strSql & ", '" & rsKey("M_RECMAIL") & "'"		
		strSql = strSql & ", '" & rsKey("M_RNAME") & "'"		
		strSql = strSql & ", '" & rsKey("M_MSN") & "'"		
		strSql = strSql & ", '" & rsKey("M_ZIP") & "'"		
		strSql = strSql & ", ''"		
		strSql = strSql & ", '" & strTimeType & "'"	
		strSql = strSql & ", 0"	
		strSql = strSql & ", " & intPortalLCID & ""
						
		strSql = strSql & ")"
'	on error resume next
		'response.Write(strSql)	
'response.End()
		executeThis(strSql)

		' - Delete the Member
		strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " WHERE M_KEY = '" & key & "'"

		executeThis(strSql)
%>
	<p align="center"><span class="fTitle"><b><%= txtRegKey11 %></b></span><br /></p>
<%		
		
		':: New User - 
		DoCookies2 newUser,newPass
		Call sendPMtoNewUser(newUser)
		Call DoCount

		strLoginStatus = 1
		if strEmailVal = 5 or strEmailVal = 6 or strEmailVal = 7 then
		  sendWelcomeEmail("v")
		end if
		if strEmailVal = 6 then
		  sendAdminNewMemberEmail(newUser)
		end if
	  end if ' end check for - strEmailVal = 8
		rsKey.close
		set rsKey = nothing

	  If strAutoLogon = 1 then
  		closeAndGo("default.asp")
	  Else %>
		<meta http-equiv="Refresh" content="2; URL=default.asp">
<% 	  End if %>
    <p align="center"><a href="default.asp"><%= txtHome %></a></p>
<%
	end if ' end auth key EOF check
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub showErrMsg()
%>
			<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
			<table align="center" border="0">
			  <tr>
			  	<td>
				  <ul>
				    <% =Err_Msg %>
				  </ul>
				</td>
			  </tr>
			</table>
			<p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBackData %></a></p>
    		<!--#include file="inc_footer.asp" --><%
		Response.End
end sub

sub postRegTextPM()
	if lcase(strEmail) = "0" then %>
	  <p align="center"><span class="fTitle"><%=txtRegKey11%></span></p>
	  <% newUser = ChkString(Request.Form("Name"),"sqlstring")
	  call sendPMtoNewUser(newUser)
 	else
		  if strEmailVal = "1" or strEmailVal = "2" or  strEmailVal = "3" or  strEmailVal = "4" and Request.Form("reservation") <> "yes" then %>
			<p align="center"><span class="fTitle"><%= txtRegKey11 %></span></p>
		<%end if
		  if strEmailVal = "7" and Request.Form("reservation") <> "yes" then %>
			<p align="center"><span class="fTitle"><%= txtRegKey9 %></span></p>
			<p align="center"><%= txtRegKey13 %>.</p>
		<%end if
		  if strEmailVal = "5" or strEmailVal = "6" or strEmailVal = "8" and Request.Form("reservation") <> "yes" then %>
			<p align="center"><span class="fTitle"><%= txtRegKey9 %></span></p>
			<p align="center">
			<%= replace(txtEmlVal6,"[%email%]",ChkString(Request.Form("Email"),"sqlstring")) %></p>
<%		  else	
			newUser = ChkString(Request.Form("Name"),"sqlstring")
		    call sendPMtoNewUser(newUser)
		  end if
	end if
	if (Request.Form("reservation") = "yes" and hasAccess(1)) then %>
	  <p align="center"><span class="fTitle"><%= txtRegKey14 %></span></p>
 <% end if
end sub

sub sendWelcomeEmail(typ) ' R=register; V=validated email
	If lcase(strEmail) = "1" Then
			strRecipientsName = chkString(Request.Form("Name"),"sqlstring")
			strRecipients = chkString(Request.Form("Email"),"sqlstring")
			if strRecipientsName = "" then
			  strRecipientsName = newUser
			end if
			if strRecipients = "" then
			  strRecipients = newUserEmail
			end if
			strFrom = strSender
			strFromName = strSiteTitle
			strsubject = strSiteTitle & " " & txtRegistration
			strMessage = txtHello & " " & chkString(Request.Form("name"),"sqlstring") & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal7 & " " & strSiteTitle & "." & vbCrLf & vbCrLf
			if strAuthType="db" then
				if (strEmailVal = 5 or strEmailVal = 6 or  strEmailVal = 8) and typ = "r" then 'they need to validate their email
					strMessage = strMessage & txtEmlVal3 & vbNewline
					strMessage = strMessage & strHomeURL & "register.asp?actkey=" & actkey & vbNewline & vbNewline
				end if
				if strEmailVal = 7 and typ = "r" then 'they just registered and need admin approval
					strMessage = strMessage & replace(txtEmlVal8,"[%sitetitle%]",strSiteTitle)
				end if
				if strEmailVal = 8 then
				  if typ = "r" then
				  strMessage = strMessage & txtEmlVal9
				  end if
				  if typ = "v" then 'they just validated their email,
					strMessage = strMessage & replace(txtEmlVal10,"[%sitetitle%]",strSiteTitle)
				  end if
				end if
				if (strEmailVal = 7 and typ = "r") or (strEmailVal = 8 and typ = "v") then
					strMessage = strMessage & txtEmlVal11 & vbNewline & vbNewline
				end if
			end if
			strMessage = strMessage & txtEmlVal4 & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal5
			sendOutEmail strRecipients,strSubject,strMessage,2,0
	end if
end sub

sub sendAdminNewMemberEmail(nam)
	If lcase(strEmail) = "1" Then
			'if instr(strWebMaster,lcase(strDBNTUserName)&",") <> 0 then
			strRecipientsName = split(strWebMaster,",")(0)
			strRecipients = strSender
			strFrom = strSender
			strFromName = strSiteTitle
			strsubject = strSiteTitle & " " & txtRegistration
			strMessage = txtHello & " " & strRecipientsName & vbCrLf & vbCrLf
			strMessage = strMessage & replace(replace(replace(txtEmlVal12,"[%sitetitle%]",strSiteTitle),"[%siteurl%]",strHomeURL),"[%member%]",nam) & vbCrLf & vbCrLf
			if strAuthType="db" then
				if strEmailVal = 7 or strEmailVal = 8 then
				  if strEmailVal = 8 then
				  strMessage = strMessage & txtEmlVal13 & vbNewline
				  end if
				  strMessage = strMessage & replace(txtEmlVal14,"[%siteurl%]",strHomeURL) & vbNewline & vbNewline
				end if
			end if
			sendOutEmail strRecipients,strSubject,strMessage,2,0
	end if
end sub

sub sendAdminInviteEmail()
	If lcase(strEmail) = "1" Then
		strRecipientsName = chkString(Request.Form("Name"),"sqlstring")
		strRecipients = chkString(Request.Form("Email"),"sqlstring")
		strFrom = strSender
		strFromName = strSiteTitle
		strsubject = strSiteTitle & " " & txtEmlVal15
		strMessage = txtHello & " " & chkString(Request.Form("name"),"sqlstring") & vbCrLf & vbCrLf
		strMessage = strMessage & replace(replace(txtEmlVal16,"[%sitetitle%]",strSiteTitle),"[%sender%]",strSender) & vbCrLf & vbCrLf
		strMessage = strMessage & txtEmlVal3 & vbCrLf
		strMessage = strMessage & strHomeURL & "register.asp?actkey=" & actkey & vbCrLf & vbCrLf
		strMessage = strMessage & txtEmlVal17 & vbCrLf
		strMessage = strMessage & txtLogin & ": " & chkString(Request.Form("Name"),"sqlstring") & vbCrLf
		strMessage = strMessage & txtPass & ": " & chkString(Request.Form("Password"),"sqlstring") & vbCrLf & vbCrLf
		strMessage = strMessage & txtEmlVal4 & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal5 & vbCrLf & strSender & vbCrLf & vbCrLf
			sendOutEmail strRecipients,strSubject,strMessage,2,0
	end if
end sub
 %>