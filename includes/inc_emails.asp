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

function GetKey(action)
  intNumChars = 62
  keyArray = Array("a","A","b","B","c","C","d","D","e","E","f","F","g","G","h","H","i","I","j","J","k","K","l","L", _
                "m","M","n","N","o","O","p","P","q","Q","r","R","s","S","t","T","u","U","v","V","w","W","x","X", _
                "y","Y","z","Z","1","2","3","4","5","6","7","8","9","0")
Randomize
key1 = (Int(((intNumChars - 1) * Rnd) + 1))
key2 = (Int(((intNumChars - 1) * Rnd) + 1))
key3 = (Int(((intNumChars - 1) * Rnd) + 1))
key4 = (Int(((intNumChars - 1) * Rnd) + 1))
key5 = (Int(((intNumChars - 1) * Rnd) + 1))
key6 = (Int(((intNumChars - 1) * Rnd) + 1))
key7 = (Int(((intNumChars - 1) * Rnd) + 1))
key8 = (Int(((intNumChars - 1) * Rnd) + 1))
key9 = (Int(((intNumChars - 1) * Rnd) + 1))
key10 = (Int(((intNumChars - 1) * Rnd) + 1))

strKey = keyArray(key1) & keyArray(key2) & keyArray(key3) & keyArray(key4) & _
         keyArray(key5) & keyArray(key6) & keyArray(key7) & keyArray(key8) & keyArray(key9) & keyArray(key10)

GetKey = strKey
	browserIP = request.ServerVariables("REMOTE_ADDR")

if action = "sendemail" then
		strRName = chkString(Request.Form("Name"),"display")
		stEmail = replace(chkString(Request.Form("Email"),"email"),"[no-spam]@","@")
		'strFrom = strSender
		strFromName = strSiteTitle
		strsub = strRName & " - " & txtEmlChgd & " - " & strSiteTitle
		stMsg = strRName & vbCrLf & vbCrLf
		if chkString(Request.QueryString("mode"),"urlpath") <> "EditIt" then
			stMsg = stMsg & txtRecFrom & " " & strSiteTitle & " " & txtChgEmlAdd  & vbCrLf & vbCrLf
		else
			stMsg = stMsg & txtRecFrom & " " & strSiteTitle & " " & txtChgEmlAdd2 & vbCrLf & vbCrLf
		end if
		stMsg = stMsg & txtChgEmlAdd3 & ":" & vbCrLf
		stMsg = stMsg & strHomeURL & "cp_main.asp?verkey=" & strKey & vbCrLf & vbCrLf
	sendOutEmail stEmail,strsub,stMsg,2,0
end if
if action = "passemail" then
		strRName = memName
		if memEmail <> "" then
		  stEmail = memEmail
		else
		  stEmail = chkString(Request.Form("Email"),"sqlstring")
		end if
		strFrom = strSender
		strFromName = strSiteTitle
		strsub = txtResetPass & " - " & strSiteTitle
		strMsg = strMsg & strRName & "," & vbCrLf
		strMsg = strMsg & txtRecFrom & " " & strSiteTitle & " " & txtReqChgPass1 & browserIP & " " & txtReqChgPass2 &  ". " & vbCrLf & vbCrLf
		strMsg = strMsg & txtCompPassChg & ":" & vbCrLf
		strMsg = strMsg & strHomeURL & "password.asp?mode=validateEmail&actkey=" & strKey & vbCrLf & vbCrLf
	sendOutEmail stEmail,strsub,strMsg,2,0
  end if
end function

sub DoPmEmail(pmName,pmEmail,pmTitle)
	strRecipientsName = pmName
	strRecipients = pmEmail
	strSubject = strSiteTitle & " - " & txtNewPM
	strMessage = pmName & "," & vbCrLf & vbCrLf
	strMessage = strMessage & replace(replace(txtMemSntPM1,"[%member%]",strDBNTUserName),"[%url%]",strSiteTitle) & vbCrLf
	strMessage = strMessage & replace(txtMemSntPM2,"[%title%]",pmTitle) & vbCrLf & vbCrLf
	strMessage = strMessage & replace(txtMemSntPM3,"[%url%]",strHomeUrl) & vbCrLf

	sendOutEmail strRecipients,strSubject,strMessage,2,0
end sub

function checkevents()
end function
function checkevents_old()
'######## Added Series ###############
strSQL = "SELECT REMINDER_ID, EVENT_ID, EVENT_START, MEMBER_ID, REMINDER_INC FROM " &strTablePrefix & "EVENTS_REMINDERS ORDER BY EVENT_START ASC"
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.open  strSql, my_conn
    
    do while not rs.EOF
	
	EDate = strtodate(rs("EVENT_START"))
	RemindDiff = rs("REMINDER_INC")
	M_ID = rs("MEMBER_ID")
	Remind_ID = rs("REMINDER_ID")
	todayis = CDate(strCurDateAdjust)
	cdEdate = CDate(Edate)
	adjdate = cdEdate - RemindDiff
	
	if CDate(strCurDateAdjust) >= CDate(EDate) - RemindDiff then
	  if lcase(strEmail) = "1" then
	  'Add in Birthday Text
	    if RemindDiff = 0 then
	      strSubject = Subject & "Happy Birthday from " & strSiteTitle
	      strMessage = "Greetings from " & strSiteTitle & "!" & vbCrLf & vbCrLf
		  strMessage = strMessage & "We hope you have a very Happy Birthday." & vbCrLf
	      strMessage = strMessage & "We wish you many more!" & vbCrLf & vbCrLf
		  strMessage = strMessage & "You may now add your next birthday to our Calendar "
	      strMessage = strMessage & "by visting our site, choosing Control Panel and "
		  strMessage = strMessage & "selecting Personal Settings." & vbCrLf & vbCrLf
	      
	      'Set their control panel variable back to zero
	      EstrSQL = "UPDATE " & strTablePrefix & "MEMBERS SET "
	      EstrSQL = EstrSQL & "M_SHOW_BIRTHDAY='0' "
	      EstrSQL = EstrSQL & "WHERE MEMBER_ID=" & M_ID
	      executeThis(EstrSQL)
	    else 
	      strSubject = Subject & "Event Reminder from " & strSiteTitle
	      strMessage = "Hello! " & vbCrLf & vbCrLf
	      strMessage = strMessage & "An event that you set a reminder for will occur in " & RemindDiff & " Day(s). " & vbCrLf & vbCrLf
		  strMessage = strMessage & "The Start Date for the Event is - " & EDate & "." & vbCrLf & vbCrLf
		  strMessage = strMessage & "You can view the Event details at " & strHomeURL & "events.asp?date=" & EDate & vbCrLf
		end if
		'get subscribed members
		strRecipients = ""
		sql = "select M_EMAIL from " & strMemberTablePrefix & "members where member_id=" & M_ID
		set Eventrs = my_Conn.execute(sql)
	
		strRecipients = strRecipients & Eventrs(0)
		Eventrs.close
     	set Eventrs = Nothing
			if strRecipients<>"" then
			    'response.write "Sending Email to " & strRecipients
				'SubscriptionsForumsMod = "1"
				sendOutEmail strRecipients,strSubject,strMessage,2,0
				'SubscriptionsForumsMod = "0"
			end if
		'response.write("about to send PM ")	
		'###########  Send a PM too  ############
		    intWebId = split(strWebmaster,",")(0)	
			strSql = "INSERT INTO " & strTablePrefix & "PM ("
			strSql = strSql & " M_SUBJECT"
			strSql = strSql & ", M_MESSAGE"
			strSql = strSql & ", M_TO"
			strSql = strSql & ", M_FROM"
			strSql = strSql & ", M_SENT"
			strSql = strSql & ", M_MAIL"
   			strSql = strSql & ", M_READ"
			strSql = strSql & ", M_OUTBOX"
   			strSql = strSql & ") VALUES ("
			strSql = strSql & " '" & strSubject & "'"
			strSql = strSql & ", '" & strMessage & "'"
			strSql = strSql & ", " & M_ID
			strSql = strSql & ", " & getmemberid(intWebId)
			strSql = strSql & ", '" & strCurDateString & "'"
			strSql = strSql & ", " & "1"
			strSql = strSql & ", " & "0"
			if request.cookies(strCookieURL & "PmOutBox") = "1" then
				strSql = strSql & ", '1')"
			else
				strSql = strSql & ", '0')"
			end if

			executeThis(strSql)
		'response.write("PM Sent - at least it shouod be")
        '###########  End PM routine ############	
        
		rstrSQL = "DELETE FROM " &strTablePrefix & "EVENTS_REMINDERS WHERE REMINDER_ID=" & Remind_ID
		        executeThis(rstrSql)
		'response.write "The selected reminder has been deleted."
	
      end if
	end if
	rs.movenext
	loop
rs.Close
set rs = nothing
end function

function getEmailFromMemberTxt()
	' called from pop_mail.asp
	strSubject = "Sent From " & strSiteTitle & " by " & chkString(Request.Form("YName"),"sqlstring")
	strMessage = "Hello " & chkString(Request.Form("Name"),"sqlstring") & vbNewline & vbNewline
	strMessage = strMessage & "You received the following message from : " & chkString(Request.Form("YName"),"sqlstring") & " (" & chkString(Request.Form("YEmail"),"sqlstring") & ") " & vbNewline & vbNewline 
	strMessage = strMessage & "At: " & strHomeURL & vbNewline & vbNewline
	strMessage = strMessage & replace(replace(replace(chkString(Request.Form("Msg"),"sqlstring"),"&59;",";"),"&quot;",""""),"&rsquo;","'") & vbNewline & vbNewline
end function
%>
