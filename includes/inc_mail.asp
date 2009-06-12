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

On Error Resume Next
Err.Clear
'strMailMode = "smtp"
select case lcase(strMailMode) 
	case "cdosys"
	  Set objNewMail = Server.CreateObject("CDO.Message")
	  Set iConf = Server.CreateObject ("CDO.Configuration")
      Set Flds = iConf.Fields 

	  'Set and update fields properties
      Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
	  Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
	  if strMailServerPassword <> "" and strMailServerLogon <> "" then
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = strMailServerLogon
		Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strMailServerPassword
	  end if
	  if isnumeric(strMailServerPort) and strMailServerPort <> 0 then
		Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strMailServerPort
	  end if
	  'Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

      Flds.Update
      Set objNewMail.Configuration = iConf
	  if strUnicode="YES" then
		objNewMail.Bodypart.ContentTransferEncoding = "8bit"
	  end if
	  objNewMail.HTMLBodypart.Charset = strCharset
	  objNewMail.Bodypart.Charset = strCharset
	  'Format and send message
	  objNewMail.From = strSender
	  objNewMail.To = strRecipients
	  'objNewMail.Bcc = strRecipients
	  objNewMail.Subject = str_Subj
	  if mTyp = 1 then
        objNewMail.HTMLBody = str_Msg
	  else
        objNewMail.TextBody = str_Msg
	  end if
        
	  if attachment = 1 then
      '	objNewMail.AddAttachment strIcsLocation
      end if
	  objNewMail.Send
	  Set iConf = Nothing
	case "jmail"
		Set objNewMail = Server.CreateObject("Jmail.smtpmail")
		objNewMail.ServerAddress = strMailServer
	  	if strUnicode="YES" then
		end if
		  objNewMail.ISOEncodeHeaders = true
		  objNewMail.Charset = strCharset
		If mTyp = 1 THEN
		  objNewMail.ContentType = "text/html"
		End If
		objNewMail.AddRecipient strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = str_Subj
		objNewMail.body = str_Msg
		objNewMail.priority = 3
		objNewMail.execute
	case "aspmail"    'Dundas.Mailer
		'create instance of Mailer control
		Set objNewMail = Server.CreateObject("Dundas.Mailer")

		'set Mailer control properties and collection items
		objNewMail.SMTPRelayServers.Add strMailServer
		objNewMail.FromAddress = strSender
		objNewMail.TOs.Add strRecipients
		'objNewMail.CCs.Add strRecipients
		'objNewMail.BCCs.Add strRecipients
		objNewMail.Subject = str_Subj
		If mTyp = 1  Then
		  'objNewMail.ContentType = "text/html"
		  objNewMail.HTMLBody = str_Msg
		Else
		  'objNewMail.CharSet = 2
		  objNewMail.Body = str_Msg
		End If

		'send email
		objNewMail.SendMail
		
	case "aspemail"
		Set objNewMail = Server.CreateObject("Persits.MailSender")
		objNewMail.Host = strMailServer
		if strMailServerPassword <> "" and strMailServerLogon <> "" then
		  objNewMail.Username = strMailServerLogon
		  objNewMail.Password = strMailServerPassword
		end if
		objNewMail.FromName = strFromName
		objNewMail.From = strSender
		objNewMail.AddReplyTo strSender
		objNewMail.AddAddress strRecipients, strRecipientsName
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		objNewMail.Send
	case "aspqmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.QMessage = 1
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.BodyText = str_Msg
		objNewMail.SendMail
	case "cdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		objNewMail.From = strSender
		objNewMail.To = strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg

		If mTyp = 1 Then
		  objNewMail.Bodyformat=0  
		  objNewMail.Mailformat=0
		else
		  objNewMail.BodyFormat = 1
		  objNewMail.MailFormat = 0
		End If
		if attachment = 1 then
            objNewMail.AttachFile strIcsLocation '## Un-Comment for CDONTS
        end if
		objNewMail.Send
	case "chilicdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		objNewMail.Send strSender, strRecipients, str_Subj, str_Msg
	case "dkqmail"
		Set objNewMail = Server.CreateObject("dkQmail.Qmail")
		objNewMail.FromEmail = strSender
		objNewMail.ToEmail = strRecipients
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		objNewMail.CC = ""
		objNewMail.MessageType = "TEXT"
		objNewMail.SendMail()
	case "geocel"
		set objNewMail = Server.CreateObject("Geocel.Mailer")
		objNewMail.AddServer strMailServer, 25
		objNewMail.AddRecipient strRecipients, strRecipientsName
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strFrom
		objNewMail.Subject = str_Subj
		objNewMail.Body = str_Msg
		objNewMail.Send()
	case "iismail"
		Set objNewMail = Server.CreateObject("iismail.iismail.1")
		MailServer = strMailServer
		objNewMail.Server = strMailServer
		objNewMail.addRecipient(strRecipients)
		objNewMail.From = strSender
		objNewMail.Subject = str_Subj
		objNewMail.body = str_Msg
		objNewMail.Send
	case "smtp"
		Set objNewMail = Server.CreateObject("SmtpMail.SmtpMail.1")
		objNewMail.MailServer = strMailServer
		objNewMail.Recipients = strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = str_Subj
		objNewMail.Message = str_Msg
		objNewMail.SendMail2
end select

If Err.Number <> 0 Then 
  Err_Msg = Err.Description
  writeToLog "Email","",Err_Msg
End if
on error goto 0
Err.Clear
Set objNewMail = Nothing
%>
