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
dim curpagetype
curpagetype = "PM"
%>
<!--#include file="inc_functions.asp" -->
<!--#include file="includes/inc_emails.asp" -->
<!--#include file="modules/pvt_msg/pm_functions.asp" -->
<%
dim iMode, iCmd, cid
iMode = 0
iCmd = 0
icid = 0
cid = 0
memID = 0
intSkin=1
boolPOP = true

if Request("mode") <> "" and Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("stop")
	end if
end if

if Request("cmd") <> "" and Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iCmd = cLng(Request("cmd"))
	else
		closeAndGo("stop")
	end if
end if

if Request("cid") <> "" and Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cid = cLng(Request("cid"))
	end if
end if

if Request("sid") <> "" and Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		memID = cLng(Request("sid"))
	end if
end if

if iMode = 2 then
  hasEditor = true
  strEditorElements = "Message"
  strEditorType = "simple"
end if

' If strLockDown then only allow the lockDownOverride if:
' pm_pop.asp has been called with no parameters (defaults to contact admin) or
' if the pm to admin function is posting to this page to process the contact admin pm
lockDownOverRide = ""
if len(trim(Request("mode") & Request("cmd") & Request("cid") & Request("sid"))) = 0 or _
   (request("Method_Type")="Contact" _
    and len(trim(request("sendto"))) <> 0 _
 and instr(1,strWebMaster,request("sendto") & ",") <> 0) then
  lockDownOverRide = "1"
end if
%>
<!--#include file="inc_top_short.asp" -->
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
function pmmembers() { var MainWindow = window.open ("<%= strHomeUrl %>pop_memberlist.asp?pageMode=pm", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=250,height=500,top=100,left=100,resizeable=yes,status=no");
}
// done hiding -->
</script>
<%
	sSql = "SELECT * FROM "& strTablePrefix & "APPS WHERE APP_INAME = 'PM'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	  iAutoDelete = rsA("APP_iDATA1")
	  iAutoDelDays = rsA("APP_iDATA2")
	  iMaxInbox = rsA("APP_iDATA3")
	  iMemberBlacklist = rsA("APP_iDATA4")
	  iPMsaveFolder = rsA("APP_iDATA5")
	  sPMfolderAccess = rsA("APP_tDATA1")
	else
	  intAppID = 0
	end if
	
  if iMode > 0 and cid >= 0 and hasAccess(2) then
	  select case iMode
	    case 1
		  popRead()
		case 2
		  pmid = cid
		  composePM(iCmd)
		case 3
		  pmDelete()
		case 4
		  pmid = cid
		  sendPM(iCmd)
		case 5 
		  pmToAdmin()
		case 6
		  pmDelete2()
	  end select
  else
    if iMode = 5 then
	  pmToAdmin()
	else
  	  contactAdmin()
	end if
  end if
 %>
<!--#include file="inc_footer_short.asp" -->
<% 

sub pmToAdmin()
	boolSend = True
	tMessage = ChkString(replace(Request.Form("Message"),vbcrlf,"<br />"),"message")
	tSubject = ChkString(Request.Form("Subject"),"sqlstring")
	tName = ChkString(Request.Form("senderName"),"sqlstring")
	tEmail = ChkString(Request.Form("senderEmail"),"sqlstring")
	tIP = ChkString(Request.Form("senderIP"),"sqlstring")
	if not DoSecImage(Ucase(request.form("secCode"))) then
      strErr = strErr & "<li>" & txtSecCodeBad & "</li>"
      boolSend = false
	end if
	if trim(tEmail) = "" then
		strErr = strErr & "<li>" & txtErNoEmlAdd & ".</li>"
		boolSend = false
	end if

	if trim(tName) = "" then
		strErr = strErr & "<li>" & txtErNoNam & ".</li>"
		boolSend = false
	end if

	if trim(tSubject) = "" then
		strErr = strErr & "<li>" & txtErNoSubj & ".</li>"
		boolSend = false
	end if

	if trim(tMessage) = "" then
		strErr = strErr & "<li>" & txtErNoMsg & ".</li>"
		boolSend = false
	end if
	
	tSubject = "[Contact us] " & tSubject

	if strDBNTUserName <> "" then
		tAuthor = getmemberid(strDBNTUserName)
	else
		tAuthor = 0
	end if
		
	if boolSend then
	  gotOne = false
	  tLngMessage = "<span class=""quote"">" & txtSndrNam & ": " & tName & "<br />" & txtEmail & ": <a href=""mailto:" & tEmail & """>" & tEmail & "</a><br />" & txtIP & ": " & tIP & "</span><hr />" & tMessage
	  
	  tempArr = split(strWebMaster,",")
	  for cu = 0 to ubound(tempArr)
		strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE M_NAME = '" & tempArr(cu) & "' "
		set rsName = my_Conn.Execute (strSql)

		if rsName.EOF then 
			'strMessage = strMessage & ""
		else
		    if tAuthor = 0 then
			  tAuthor = rsName("MEMBER_ID")
			end if
	
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
			strSql = strSql & " '" & tSubject & "'"
			strSql = strSql & ", '" & tLngMessage & "'"
			strSql = strSql & ", " & rsName("MEMBER_ID")
			strSql = strSql & ", " & tAuthor
			strSql = strSql & ", '" & strCurDateString & "'"
			strSql = strSql & ", 0"
			strSql = strSql & ", 0"
			strSql = strSql & ", 0)"

			executeThis(strSql)
				
			if strEmail = 1 then
				DoPmEmail rsName("M_NAME"),rsName("M_EMAIL"),tSubject
			end if
			gotOne = true
			set rsName = nothing
		end if
	  next
	  if Err.description <> "" then 
		txMessage = txMessage & txtWasErr & " = " & Err.description
		boolSend = false
	  Elseif gotOne then
		txMessage = txMessage & txtEmlSntAdmins
	  else
		txMessage = txMessage & txtEmlNoSent & "."
		boolSend = false
	  end if
	end if
	
	if not boolSend then
  	  tTxt = txtThereIsProb
	  tmpMsg = strErr & txMessage
	  txMessage = tmpMsg
	end if
	
	spThemeBlock1_open(intskin) %>
	  <p align="center"><br />
	  <span class="fSubTitle"><%= tTxt %>&nbsp;</span></p>
	  <div class="fSubTitle" style="text-align:center;"><ul><%= txMessage %></ul></div>
	  <hr /><p align="center"><%= tMessage %></p>
	  <% 
	  if not boolSend then %>
	    <p align="center">
	    <a href="JavaScript:history.go(-1)"><%= txtGoBack %>.</a></p>
<% 	  end if
	  response.Write("<p>&nbsp;</p>")
  	  spThemeBlock1_close(intskin)
end sub

sub contactAdmin() %>
<p align="center">&nbsp;</p>
<form action="pm_pop.asp?mode=5&amp;cmd=0" method="post" name="PostTopic">
<input name="Method_Type" type="hidden" value="Contact" ID="Method_Type">
<input name="sendto" type=hidden value="<%= split(strWebMaster,",")(0) %>">
<input name="senderIP" type=hidden value="<%= request.ServerVariables("REMOTE_ADDR") %>">
<%
spThemeTitle = "<center>" & txtCnctAdmin & ".</center>"
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
      <tr>
        <td colspan="2">
<p align="center"><b>Please do not ask for support using this tool, they will not be answered.<br>Please post support questions in our forums.</b></p><br />
                  <p align="center"><%= txtErValEml %>.<br />
                    <%= txtAllFldsReq %>.</p>
                  <p align="center"><%= txtYrIP %>: <span class="fAlert"><b><%= request.ServerVariables("REMOTE_ADDR") %></b></span><br />
                    <%= txtIsIncl %>.</p>
                </td>
      </tr>
	  <tr>
        <td colspan=2 height="10" valign="middle"></td>
      </tr>
      <tr>
        <td noWrap vAlign="top" align="right"><b><%= txtUNam %>:&nbsp;</b></td>
        <td><input class="textbox" name="senderName" value="" size="40"></td>
      </tr>
      <tr>
        <td noWrap vAlign="top" align="right"><b><%= txtUEml %>:&nbsp;</b></td>
        <td><input class="textbox" name="senderEmail" value="" size="40"></td>
      </tr>
      <tr>
        <td noWrap vAlign="top" align="right"><b><%= txtSubject %>:&nbsp;</b></td>
        <td><input class="textbox" maxLength="40" name="Subject" value="" size="40"></td>
      </tr>
	  <tr>
        <td noWrap vAlign="top" align="right"><b><%= txtMsg %>:&nbsp;</b></td>
        <td>
        <textarea class="textbox" cols="40" name="Message" rows="8"></textarea></td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td noWrap valign="top" align="center" colspan="2">
          <img align="absolute" src="<%= strHomeUrl %>includes/securelog/image.asp" /><br />
          <%= txtEntrSecImg %><br />
          <input CLASS="textbox" type="text" name="secCode" size="8" maxLength="8" value="" onFocus="javascript:this.value='';">
        </td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" align="center"><input class="button" name="<%= txtSubmit %>" type="submit" value="Send Message"></td>
      </tr>
<%Response.Write("</table>")
spThemeBlock1_close(intSkin)%>
</form>
<%
end sub

sub popRead()
	'mark PM as read
	strSql = "UPDATE " & strTablePrefix & "PM "
	strSql = strSql & " SET " & strTablePrefix & "PM.M_READ = 1 "
	strSql = strSql & " WHERE ((" & strTablePrefix & "PM.M_ID = " & cid & ") AND (" & strTablePrefix & "PM.M_TO = " & strUserMemberID & "));"
	executeThis(strSql)
	
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & cid
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

	Set rsMessage = my_Conn.Execute(strSql)

	if rsMessage.BOF or rsMessage.EOF then
	   closeAndGo("stop")
	end if

	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_CITY, " & strMemberTablePrefix & "MEMBERS.M_STATE, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_SUBJECT, " & strMemberTablePrefix & "MEMBERS.M_GLOW"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & cid

	Set rs = my_Conn.Execute(strSql) %>
	<p>&nbsp;</p>
<table border="0" width="95%">
  <tr>
    <td  height="40">&nbsp;</td>
	<td align="left" valign="top" width="380" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm_pop.asp?mode=2&amp;cmd=1&amp;cid=<%= cid %>">
	     <img  unselectable='on' src="<%= strHomeURL %>images/icons/pmreply.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtRplyMsg %>" alt="<%= txtRplyMsg %>"
	          onmouseover="document.btnreply.style.top=-37" onmouseout="document.btnreply.style.top=0"
	          onmousedown="document.btnreply.style.top=-37" onmouseup="document.btnreply.style.top=0"; /></A>
         <img name="btnreply" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	<div unselectable='on' style="position:absolute;margin-left:93;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm_pop.asp?mode=2&amp;cmd=2&amp;cid=<%= cid %>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmreplyQ.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtRplyQtMsg %>" alt="<%= txtRplyQtMsg %>"
	          onmouseover="document.btnreplyQ.style.top=-37" onmouseout="document.btnreplyQ.style.top=0"
	          onmousedown="document.btnreplyQ.style.top=-37" onmouseup="document.btnreplyQ.style.top=0"; /></A>
         <img name="btnreplyQ" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	<div unselectable='on' style="position:absolute;margin-left:186;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm_pop.asp?mode=2&amp;cmd=3&amp;cid=<%= cid %>">
	     <img  unselectable='on' src="<%= strHomeURL %>images/icons/pmforward.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtFwdMsg %>" alt="<%= txtFwdMsg %>"
	          onmouseover="document.btnforward.style.top=-37" onmouseout="document.btnforward.style.top=0"
	          onmousedown="document.btnforward.style.top=-37" onmouseup="document.btnforward.style.top=0"; /></A>
         <img name="btnforward" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	<div unselectable='on' style="position:absolute;margin-left:279;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm_pop.asp?mode=3&amp;cid=<%= cid %>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmdelete.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtDelMsg %>" alt="<%= txtDelMsg %>"
	          onmouseover="document.btndelete.style.top=-37" onmouseout="document.btndelete.style.top=0"
	          onmousedown="document.btndelete.style.top=-37" onmouseup="document.btndelete.style.top=0"; /></A>
         <img name="btndelete" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
	 </tr>
</table>
<% spThemeBlock1_open(intSkin) %>
<table width="100%" cellpadding="3">
  <tr>
    <td align="center" class="tCellAlt0" height="20" width="25%"><span class="fSubTitle"><b><%= txtSubject %>:</b></span></td>
    <td align="left" class="tCellAlt0"><span class="fSubTitle"><b>&nbsp;<% =rsMessage("M_SUBJECT") %></b></span></td>
  </tr>
  <tr>
    <td align="center" class="tCellAlt0" height="20"><span class="fSubTitle"><b><%= txtFrom %>:</b></span></td>
    <td align="left" class="tCellAlt0"><b>&nbsp;<%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt2" colspan="2" valign="top">
      <img src="<%= strHomeURL %>images/icons/icon_posticon.gif" border="0" hspace="3"><%= txtSent %>&nbsp;-&nbsp;<% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %>
      <hr />
    </td>
  </tr>
  <tr>
    <td class="tCellAlt2" colspan="2" valign="top"><%= Replace(formatStr(Replace(Replace(rsMessage("M_MESSAGE"),"tiny_mce/", strHomeUrl & "tiny_mce/"),"''","'")),"images/Smilies/", strHomeUrl & "images/Smilies/") %>
    </td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)

end sub

sub pmDelete()
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & chkString(Request.Cookies(strUniqueID & "User")("PWord"),"sqlstring") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & cid

	Set rsMessage = my_Conn.Execute(strSql)

	if rsMessage.EOF or rsMessage.BOF then
  		response.Write("<br /><br />")
  		spThemeBlock1_open(intSkin) %>
        	<center><br /><br />
	        <p><%= txtNoPermDel %>.</p>
	        <p>&nbsp;</p><br /><br /></center>
		<%  	
		response.Write("<br /><br />")
		spThemeBlock1_close(intSkin)
		set rsMessage = nothing
	else
		set rsMessage = nothing
		strSql = "DELETE FROM " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strTablePrefix & "PM.M_ID = " & cid
		executeThis(strSql)
  		response.Write("<br /><br />")
  		spThemeBlock1_open(intSkin) %>
	        <center>
	        <br />
	        <br />
	        <P><span class="fTitle"><%= txtMsgDel %>!</span></p>
	        <P>&nbsp;</p>
	        <br />
	        <br />
	        </center>
	        <script type="text/javascript">
			opener.document.location.reload();
			window.close();
			</script><%  
		response.Write("<br /><br />")
		spThemeBlock1_close(intSkin)
	end if
end sub
 %>