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

function forumMess()
	'this shows on fhome.asp page
	if chkApp("PM","USERS") then
%>
        <tr>
          <td class="tSubTitle" colspan="<% if strShowModerators = "1" or hasAccess(1) then Response.Write("7") else Response.Write("6")%>">
            <b><%= txtPvtMessgs %>: </b></td>
        </tr>
        <tr>
          <td align="center" valign="middle">
            <% if not hasAccess(2) Then %>
              <IMG SRC="images/icons/icon_pmdead.gif" alt="" />
            <% else
                 if pmCount = 0 then %>
                 <IMG SRC="images/icons/icon_pm.gif" alt="" />
            <%   end if
                 if pmCount >= 1 then
                   response.Write(pmImage)
                 end if %>
            <% end if %>
          </td>
          <td valign=top colspan="<% if (strShowModerators = "1") or hasAccess(1) then Response.Write("7") else Response.Write("5")%>" class="fNorm"><A HREF="pm.asp"><%= txtPMinbox %></A><br />
          <% if strDBNTUserName = "" Then %>
           &nbsp;-&nbsp;<%= txtLogChkPM %>
          <% else %>
            <b><% =strDBNTUserName %></b>
			&nbsp;-&nbsp;<%= replace(txtCntUnread,"[%count%]",clng(pmcount)) %>
		  <% end if %>
          </td>
        </tr>
<%	end if
end function

function pmSavedCount()
  tmpCnt = 0
  If strDBType = "access" then
	strSqL = "SELECT count(M_TO) as [scount] " 
  else
	strSqL = "SELECT count(M_TO) as scount " 
  end if
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO"
  strSql = strSql & " AND " & strTablePrefix & "PM.M_SAVED = 1"

  Set rsSavebox = my_Conn.Execute(strSql)
  tmpCnt = rsSavebox("scount")
  set rsSavebox = nothing
  pmSavedCount = tmpCnt
end function

function pmInCount(m_id)
  tmpCnt = 0
  If strDBType = "access" then
	strSqL = "SELECT count(M_TO) as [incount] " 
  else
	strSqL = "SELECT count(M_TO) as incount " 
  end if
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & m_id & ""
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
  strSql = strSql & " AND " & strTablePrefix & "PM.M_SAVED = 0"

  Set rsOutbox = my_Conn.Execute(strSql)
  tmpCnt = rsOutbox("incount")
  set rsOutbox = nothing
  pmInCount = tmpCnt
end function

function pmInboxCount()
  tmpCnt = 0
  If strDBType = "access" then
	strSqL = "SELECT count(M_TO) as [incount] " 
  else
	strSqL = "SELECT count(M_TO) as incount " 
  end if
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
  strSql = strSql & " AND " & strTablePrefix & "PM.M_SAVED = 0"

  Set rsOutbox = my_Conn.Execute(strSql)
  tmpCnt = rsOutbox("incount")
  set rsOutbox = nothing
  pmInboxCount = tmpCnt
end function

function pmSentCount()
  tmpCnt = 0
  If strDBType = "access" then
	strSqL = "SELECT count(M_TO) as [outcount] " 
  else
	strSqL = "SELECT count(M_TO) as outcount " 
  end if
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
  strSql = strSql & " AND " & strTablePrefix & "PM.M_OUTBOX = 1" 
  Set rsOutbox = my_Conn.Execute(strSql)
  tmpCnt = rsOutbox("outcount")
  set rsOutbox = nothing
  pmSentCount = tmpCnt
end function

function pmSentUnreadCount()
  tmpCnt = 0
  If strDBType = "access" then
	strSqL = "SELECT count(M_TO) as [outcount] " 
  else
	strSqL = "SELECT count(M_TO) as outcount " 
  end if
  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
  strSql = strSql & " AND " & strTablePrefix & "PM.M_OUTBOX = 1" 
  strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0"

  Set rsOutbox = my_Conn.Execute(strSql)
  tmpCnt = rsOutbox("outcount")
  set rsOutbox = nothing
  pmSentUnreadCount = tmpCnt
end function

function GetPMCOUNT()
  dim rs, intFemales, intMales 
  set rs = server.CreateObject("adodb.recordset")
  strSql = "SELECT COUNT(M_ID) AS PMCount FROM " & strTablePrefix & "PM"
  strSql = strSql & " WHERE M_SAVED = 0"
  rs.Open strSql, my_Conn
  GetPMCount = rs("PMCOUNT")
  rs.Close
  set rs = nothing
end function

function GetPMToday()
  dim rs, intFemales, intMales 
  set rs = server.CreateObject("adodb.recordset")
  strSql = "SELECT COUNT(M_ID) AS PMCount FROM " & strTablePrefix & "PM WHERE m_sent > '" & DateToStr(DateAdd("h",-24,DateAdd("h", strTimeAdjust , Now()))) & "'"
  rs.Open strSql, my_Conn
  GetPMToday = rs("PMCOUNT")
  rs.Close
  set rs = nothing
end function

function pop_pmToast()
'check to see if the 'Toast' PM type is selected to show
If (strPMtype = 1 or strPMtype = 2) and chkApp("PM","USERS") then  
  MessengerTitle = txtSitMes
  toastSpeed = 30 'how fast the toasts animate... a smaller number = faster
  toastActive = 7 'number of seconds to show toast
  pmcnt = pmcount
  '************************************************************
  If pmcnt > 1 and request.Cookies("PMnotify") = "" then %>
	<DIV ID="toast" class="spThemeToast" STYLE="display:none; z-index:105;">
	<Table cellpadding=0 cellspacing=0 border=0 width="100%" ID="Table1">
	  <tr>
		<td width=5 class="spThemeToast_header_left">
			<div class="spThemeToast_header_img_left"></div>
		</td>
		<td class="spThemeToast_title" valign=top>
			<%= MessengerTitle%>
		</td>
		<td class="spThemeToast_header_right">
			<div class="spThemeToast_header_img_right" onclick="document.all.toast.style.display='none'"></div>
		</td>
	  </tr>
	  <tr>
		<td valign=top colspan=3>
		<table width="100%" cellpadding=0 cellspacing=0 height=90 ID="Table2"><tr>
		<td class="spThemeToast_content_left" valign=top><img src="<%= strHomeUrl %>images/clear.gif" width=1 height=1 border=0>
		</td>
		<td class="spThemeToast_content" width="100%" onclick="document.location.href='pm.asp'" valign=center align=center>
		<%= replace(txtPMUnread,"[%count%]",pmcnt) %>
		</td>
		<td class="spThemeToast_content_right" valign=top><img src="<%= strHomeUrl %>images/clear.gif" width=1 height=1 border=0></td></tr></table>
		</td>
	  </tr>
	</table>
	</DIV>
	<% 'response.Cookies("pmnotify") = "done"
	dojava=true
  '***************************************************
  elseif pmcnt = 1 then
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "PM.M_ID, " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0"
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

	Set rsMessage = my_Conn.Execute(strSql)
	if not (rsmessage.eof and rsmessage.bof) then
		PMID = rsMessage("M_ID")
		PMFrom = rsmessage("M_FROM")
		rsmessage.close
		Set rsmessage=nothing
		strSql = "SELECT M_NAME FROM " & strmembertableprefix & "MEMBERS WHERE MEMBER_ID = " & pmfrom
		Set rsMessage = my_Conn.Execute(strSql)
		pmName = rsmessage("M_NAME")
		rsmessage.close
		Set rsmessage=nothing
		if request.Cookies("pmnotify" & pmid) = "" then%>
		<DIV ID="toast" class="spThemeToast" STYLE="display:none; z-axis:105;">
		<Table cellpadding=0 cellspacing=0 border=0 width="100%" ID="Table1">
		  <tr>
			<td width=5 class="spThemeToast_header_left">
			<div class="spThemeToast_header_img_left"></div>
			</td>
			<td class="spThemeToast_title" valign=top>
			<%= MessengerTitle%>
			</td>
			<td class="spThemeToast_header_right">
			<div class="spThemeToast_header_img_right" onclick="document.all.toast.style.display='none'"></div>
			</td>
		  </tr>
		  <tr>
			<td valign=top colspan=3>
			<table width="100%" cellpadding=0 cellspacing=0 height=90 ID="Table2"><tr>
			  <td class="spThemeToast_content_left" valign=top><img src="<%= strHomeUrl %>images/clear.gif" width=1 height=1 border=0>
			  </td>
			  <td class="spThemeToast_content" width="100%" onclick="document.location.href='pm.asp?cmd=4&amp;pmid=<%= PMID %>'" valign=center align=center>
		<%= txtPMnewFrm %>&nbsp;<%=pmname%>			
			  </td>
			  <td class="spThemeToast_content_right" valign=top><img src="<%= strHomeUrl %>images/clear.gif" width=1 height=1 border=0></td></tr>
		    </table>
			</td></tr>
	    </table>
		</DIV><% 'response.Cookies("pmnotify" & pmid) = "done"
		dojava = true
		end if
  	end if
'	rsmessage.close
	Set rsmessage=nothing
  end if
  '**************************************************************
  if instr(lcase(request.ServerVariables("script_name")),"private") <> 0 or instr(lcase(request.ServerVariables("script_name")),"pm.asp") <> 0 then
	dojava=false
  end if
  '**********************************************************
  if pmcnt > 0 and dojava then %>
<script>
var ypos= -120

function hidePanel(){
	ypos=ypos-5;
	if(ypos<=-120){ypos=-120}

if(document.layers){
	document.toast.bottom=ypos;
}
if(document.all && navigator.userAgent.indexOf('Opera') == -1){
	document.all.toast.style.bottom=ypos;
}

if(document.all && ! (navigator.userAgent.indexOf('Opera') == -1)){ 
	document.all.toast.style.top = (document.body.clientHeight - ypos) -111;
}
if(!document.all && document.getElementById){
	document.getElementById("toast").style.bottom=ypos+"px";
}
if (ypos<=-120){
	window.clearTimeout(Id);Id=0;
	if(document.layers){
		document.toast.display='none';
	}
	if(document.all){
		document.all.toast.style.display='none';
	}
	if(!document.all && document.getElementById){
		document.getElementById("toast").style.display='none';
	}	
}else{
	Id = setTimeout("hidePanel();",toastSpeed);	
}
					}
function showPanel(){
	if(document.layers){
		document.toast.display='';
	}
	if(document.all){
		document.all.toast.style.display='';
	}
	if(!document.all && document.getElementById){
		document.getElementById("toast").style.display='';
	}
ypos=ypos+5;
if(ypos>=0){ypos=0}
if(document.layers){
	document.toast.bottom=ypos;
}

if(document.all && navigator.userAgent.indexOf('Opera') == -1){
	document.all.toast.style.bottom=ypos;
}

if(document.all && ! (navigator.userAgent.indexOf('Opera') == -1)){ 
	document.all.toast.style.top = (document.body.clientHeight - ypos) -111;
}

if(!document.all && document.getElementById){
	document.getElementById("toast").style.bottom=ypos+"px";
}
if (ypos>=0){
	window.clearTimeout(Id);Id=0;
	Id = setTimeout("hidePanel();",toastActive);
}else{
	Id = window.setTimeout("showPanel();",toastSpeed);
}
}
function doNotify() {
	window.setTimeout("showPanel();",300) // delay while sound loads
	document.write ('<EMBED SRC=Themes/<%= strTheme %>/newpm.wav WIDTH=1 HEIGHT=1 HIDDEN=true AUTOSTART=true LOOP=false volume=100></EMBED>')
}

toastActive = <%=toastActive * 1000%>
toastSpeed = <%=toastSpeed%>
doNotify();
</script>
<% 
  end if
end if
end function

sub composePM(typ)
  ' typ reference
  ' 0 = New message
  ' 1 = reply to message
  ' 2 = reply with quote
  ' 3 = forward message

  rly = "/"
  if typ = 1 or typ = 2 or typ = 3 then
	strSql = "SELECT * FROM " & strTablePrefix & "PM "
	'strSql = strSql & " WHERE " & strTablePrefix & "PM.M_ID = " & pmid & 
	strSql = strSql & " WHERE ((" & strTablePrefix & "PM.M_ID = " & pmid & ") AND ((" & strTablePrefix & "PM.M_TO = " & strUserMemberID & ") or (" & strTablePrefix & "PM.M_FROM = " & strUserMemberID & ")));"
	set rs = my_Conn.Execute(strSql)
	if rs.eof then
	  'Response.Write "Hello5:<br>" & strSql
	  closeAndGo("stop")
	else
	strAuthor = rs("M_FROM")
	strSubject = rs("M_SUBJECT")
	strPMmsg = ""
	rly = "readonly=""readonly"""
		if typ = 2 then
		  if strAllowHtml <> 1 then
			strPMmsg = "[quote]" & vbCrLf
			strPMmsg = strPMmsg & rs("M_MESSAGE") & vbCrLf
			strPMmsg = strPMmsg & "[/quote]"
			strPMmsg = CleanCode(strPMmsg)
		  else
			strPMmsg = "<span class=quote><i>" & getMemberName(strAuthor) & "&nbsp;" & txtWrote & " :</i><br />" & rs("M_MESSAGE") & "</span><br />" & vbCrLf
			strPMmsg = chkString(strPMmsg,"message")
		  end if
		end if
		if typ = 3 then
		  if strAllowHtml <> 1 then
			strPMmsg = vbCrLf & vbCrLf & "---------- " & txtFwddMsg & " ----------" & vbCrLf
			strPMmsg = strPMmsg & rs("M_MESSAGE") & vbCrLf
			strPMmsg = strPMmsg & "-------------------------------------------------"
			strPMmsg = CleanCode(strPMmsg)
		  else
			strPMmsg = "---------- " & txtFwddMsg & " ----------<br />" & vbCrLf
			strPMmsg = strPMmsg & rs("M_MESSAGE") & vbCrLf
			strPMmsg = strPMmsg & "<br /><br />-------------------------------------------------"
		  end if
		end if
	strPMmsg = replace(strPMmsg,"''","'")
 	end if
  end if
  if memID > 0 then
    memName = getMemberName(memID)
  end if
  
  if not boolPOP then
  %>
<form action="pm.asp?cmd=3&amp;mode=<%= typ %>" method="post" name="PostTopic">

<input name="mode" type="hidden" value="<%= typ %>">
<input name="pmid" type="hidden" value="<%= pmid %>">
<%
  else
  %>
<form action="pm_pop.asp?mode=4&amp;cmd=<%= typ %>" method="post" name="PostTopic">

<input name="cmd" type="hidden" value="<%= typ %>">
<input name="cid" type="hidden" value="<%= pmid %>">
<%end if

spThemeBlock1_open(intSkin)
Response.Write("<br /><table class=""tPlain"">")

if typ = 0 or typ = 3 then %>
	  <tr><td></td><td class="fNorm" vAlign="top" align="left">(<%= txtSepWthComm %>)<br />&nbsp;</td></tr>
      <tr>
        <td nowrap valign="top" align="right" class="fNorm"><b><%= txtSndTo %>:</b></td>
        <td><input name="sendto" value="<%= memName %>" size="45">&nbsp;&nbsp;<a href="JavaScript:pmmembers();"><span class="fNorm"><%= txtMbrLst %></span></a>
	<br /><%if hasAccess(1) then %><input type="checkbox" name="allmem" value="true"><span class="fNorm"><%= txtSndPMAllMem %></span><% Else %><input type="hidden" name="allmem" value="false"><%end if%>
        <br />&nbsp;</td>
      </tr>
<% else %>
	  <tr><td></td><td  vAlign="top" align="left"><input type="hidden" name="sendto" value="<%= getMemberName(strAuthor) %>"></td></tr>
<% end if %>
      <tr>
        <td vAlign="top" align="right" class="fNorm" nowrap>
		<b><%= txtSubject %>:</b></td>
        <td>
		<input maxLength="90" name="Subject" value="<%= strSubject %>" size="45" <%= rly %>></td>
      </tr>
<% 

  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<b>" & txtMsg & ":</b>", strPMmsg
  else
    if not boolPOP then
      displayPLAINeditor 1, strPMmsg
    else %>
      <tr>
        <td vAlign="top" align="right" class="fNorm" nowrap>
		<b><%= txtMsg %>:</b></td>
        <td>
      <br /><textarea tabindex="1" name="message" rows="20" cols="60" class="textbox"><%= strPMmsg %></textarea></td>
      </tr>
<%	end if
  end if
  %>
	<tr>
        <td>&nbsp;</td>
        <td class="fNorm">
        <input name="Sig" type="checkbox" value="yes" <% =Chked(Request.Cookies("User")("sig")) %>><%= txtClkForSig %><br />
        </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input name="Submit" type="submit" value="<%= txtSndMsg %>" accesskey="s" title="Shortcut Key: Alt+S" class="button">&nbsp;<input name="Reset" type="reset" value="<%= txtReset %>" class="button"></td>
      </tr></table>
  <%
  spThemeBlock1_close(intSkin)
if typ > 0 then
  spThemeTitle = txtMsg
  spThemeBlock1_open(intSkin)  %>
  <table class="tPlain">
  <%
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "PM.M_MESSAGE " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM AND "
	strSql = strSql & "       " & strTablePrefix & "PM.M_ID = " & pmid

	set rs = my_Conn.Execute (strSql) 

	Response.Write "      <tr class=""tCellAlt0"">" & vbCrLf
	Response.Write "        <td valign=""top"" width=""150"" class=""fNorm"" nowrap=""nowrap"">"
	Response.Write "><b>" & rs("M_NAME") & "</b></td>" & vbCrLf
	Response.Write "        <td class=""fNorm"" valign=""top"">"
	Response.Write Replace(formatStr(Replace(Replace(rs("M_MESSAGE"),"tiny_mce/", strHomeUrl & "tiny_mce/"),"''","'")),"images/Smilies/", strHomeUrl & "images/Smilies/") & "</td>" & vbCrLf
	Response.Write "      </tr>" & vbCrLf
	Response.Write "    </TABLE>" & vbCrLf
	spThemeBlock1_close(intSkin)
  end if
end sub

function sendPMtoMember(iTo,iFrom,sSubj,sMsg,iOutBox,errMsg)
  boolSend = True
  bMsg = False
  errMsg = ""
  
  if len(sSubj & "x") = 1 then
    errMsg = errMsg & "<li>" & txtMstPstSubj & "</li>"
	boolSend = False
  end if
  if len(sMsg & "x") = 1 then
    errMsg = errMsg & "<li>" & txtMstPstMsg & "</li>"
	boolSend = False
  end if
  
  if boolSend then
    strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_PMEMAIL, " & strMemberTablePrefix & "MEMBERS.M_RECEIVE_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_RECMAIL, " & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, " & strMemberTablePrefix & "MEMBERS.M_PMBLACKLIST, " & strMemberTablePrefix & "CP_CONFIG.PM_OUTBOX"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS INNER JOIN " & strMemberTablePrefix & "CP_CONFIG ON " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strMemberTablePrefix & "CP_CONFIG.MEMBER_ID"
	strSql = strSql & " WHERE (((" & strMemberTablePrefix & "MEMBERS.MEMBER_ID)=" & iTo & "));"
	set rsName = my_Conn.Execute(strSql)
	if rsName.BOF or rsName.EOF then '::  no one registered
	  errMsg = errMsg & "<li>"&txtSryMemNoFnd&"</li>"
	  boolSend = False
	else
	  ':: check if sTo has PM access and they have PMs turned ON		
	  if rsName("M_PMRECEIVE") = "0" or rsName("M_PMSTATUS") = 0 then
		errMsg = errMsg & "<li>" & replace(txtSryNoRecPM,"[%member%]",rsName("M_NAME")) & "</li>"
		boolSend = False
	  end if
	  
	  ':: Check if sTo has full PM box
	  if iMaxInbox > 0 and pmInCount(rsName("MEMBER_ID")) >= iMaxInbox and not hasAccess(1) then
		errMsg = errMsg & "<li>" & replace(txtSryInbxFul,"[%member%]",rsName("M_NAME")) & "</li>"
		boolSend = False
	  end if
	  
	  ':: check member blacklist if present
	  if iMemberBlacklist then
		if rsName("M_PMBLACKLIST") <> "" then
		  if instr(rsName("M_PMBLACKLIST"),",") > 0 then
			arrTemp = split(rsName("M_PMBLACKLIST"),",")
			for bl = 0 to ubound(arrTemp)
			  if cLng(strUserMemberID) = cLng(arrTemp(bl)) then
			   	errMsg = errMsg & "<li>" & replace(txtSryNoRecPM,"[%member%]",rsName("M_NAME")) & "</li>"
				boolSend = False
			  end if
			next
		  else
			if strUserMemberID = cLng(rsName("M_PMBLACKLIST")) then
			  errMsg = errMsg & "<li>" & replace(txtSryNoRecPM,"[%member%]",rsName("M_NAME")) & "</li>"
			  boolSend = False
			end if
		  end if
		end if
	  end if
	  ':: end blacklist check
	  if boolSend then
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
		strSql = strSql & " '" & sSubj & "'"
		strSql = strSql & ", '" & sMsg & "'"
		strSql = strSql & ", " & iTo
		strSql = strSql & ", " & iFrom
		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", 0"
		strSql = strSql & ", 0"
		strSql = strSql & ", '" & iOutBox & "')"
		
		on error resume next
		my_Conn.execute(strSql)

		if Err.description <> "" then 
		  errMsg = errMsg & "<li>"& txtWasErr &"<br /> = " & Err.description
		Else
		  errMsg = errMsg & "<li>"& txtMsgSntTo &" " & rsName("M_NAME") & "</li>"
		  bMsg = True
		end if
			
		if strEmail = "1" and rsName("M_PMEMAIL") = "1" then
		  DoPmEmail rsName("M_NAME"),rsName("M_EMAIL"),strSubject
		end if
		on error goto 0
  	  end if
    end if
    set rsName = nothing
  end if
  sendPMtoMember = bMsg
end function

Function sendPMtoNewUser(reciev)
  if chkAPP("PM","USERS") then
	strSql = "SELECT AUTOPM_ON, AUTOPM_SUBJECTLINE, AUTOPM_MESSAGE "
	strSql = strSql & "FROM " & strTablePrefix & "CONFIG "
	strSql = strSql & "WHERE CONFIG_ID = 1"
	Set rs = my_Conn.Execute (strSql)
 if not rs.eof then
  if rs("AUTOPM_ON") = 1 then

	welcomeMessage = rs("AUTOPM_MESSAGE")
	msgSubject = rs("AUTOPM_SUBJECTLINE")
	senderName = split(strWebMaster,",") 'you can edit this but MUST be a valid username
	rs.close
	set rs = nothing

	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & senderName(0) & "'" 
	set rsP = my_Conn.Execute (strSql)
	adminId = rsP(0)
	set rsP = nothing
	
	sendThePM = true
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & ChkString(reciev,"SQLString") & "'" 
	set rsP = my_Conn.Execute (strSql)
	if not rsP.eof then
	  newuserId = rsP(0)
	else
	  sendThePM = false
	end if
	set rsP = nothing
	
	if sendThePM then
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
	strSql = strSql & " '" & ChkString(msgSubject,"SQLString") & "'"
	strSql = strSql & ", '" & ChkString(welcomeMessage,"SQLString") & "'"
	strSql = strSql & ", " & newUserId
	strSql = strSql & ", " & adminId
	strSql = strSql & ", '" & strCurDateString & "'"
	strSql = strSql & ", " & "0"
	strSql = strSql & ", " & "0"
	if request.cookies(strCookieURL & "PmOutBox") = "1" then
	strSql = strSql & ", '1')"
	else
	strSql = strSql & ", '0')"
	end if
	executeThis(strSql)
	end if
  end if
 end if	
 set rs = Nothing
  end if
end function

Function sendPMtoUser_old(pto,pfrm,psubj,pmsg)
  if chkApp("PM","USERS") then

	welcomeMessage = rs("AUTOPM_MESSAGE")
	msgSubject = rs("AUTOPM_SUBJECTLINE")
	senderName = split(strWebMaster,",") 'you can edit this but MUST be a valid username

	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & senderName(0) & "'" 
	set rsP = my_Conn.Execute (strSql)
	adminId = rsP(0)
	set rsP = nothing
	
	sendThePM = false
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & ChkString(reciev,"SQLString") & "'" 
	set rsP = my_Conn.Execute (strSql)
	if not rsP.eof then
	  newuserId = rsP(0)
	  sendThePM = true
	end if
	set rsP = nothing
	
  if sendThePM then
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
	strSql = strSql & " '" & ChkString(msgSubject,"SQLString") & "'"
	strSql = strSql & ", '" & ChkString(welcomeMessage,"SQLString") & "'"
	strSql = strSql & ", " & newUserId
	strSql = strSql & ", " & adminId
	strSql = strSql & ", '" & strCurDateString & "'"
	strSql = strSql & ", " & "0"
	strSql = strSql & ", " & "0"
	if request.cookies(strCookieURL & "PmOutBox") = "1" then
	strSql = strSql & ", '1')"
	else
	strSql = strSql & ", '0')"
	end if
	executeThis(strSql)
  end if
 end if	
 set rs = Nothing
end function

sub sendPM(typ)
  ' typ reference
  ' 0 = New message
  ' 1 = reply to message
  ' 2 = reply with quote
  ' 3 = forward message
  dim arrNames, strMessage2, rsName
  strErr = ""
  boolSend = true
  strMess = ChkString(Request.Form("Message"),"message")
  sSubject = ChkString(Request.Form("Subject"),"message")
  'sSubject = replace(ChkString(Request.Form("Subject"),"sqlstring"),"'","''")
  if Request.form("allmem") = false then
    if trim(Request.Form("sendto")) = "" or Request.Form("sendto") = " " then
  	  strErr = "<li>" & txtMstSplyMemNam & "</li>"
	  boolSend = False
    end if
  end if
  if strMess = " " or trim(strMess) = "" then
    strErr = strErr & "<li>" & txtMstPstMsg & "</li>"
	boolSend = False
  end if
  if Request.Form("sig") = "yes" and boolSend = true then
    strMess = strMess & vbCrLf & vbCrLf & ChkString(GetSig(strDBNTusername), "message" )
  end if
  
  if typ = 1 or typ = 2 or typ = 3 then
	strSql = "SELECT * FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE ((" & strTablePrefix & "PM.M_ID = " & pmid & ") AND ((" & strTablePrefix & "PM.M_TO = " & strUserMemberID & ") or (" & strTablePrefix & "PM.M_FROM = " & strUserMemberID & ")));"
	set rs = my_Conn.Execute(strSql)
	if rs.eof then
      strErr = "<li>" & txtMsgNotFnd & "</li>"
	  boolSend = False	  
	else
	  strAuthor = rs("M_FROM")
	  sSubject = rs("M_SUBJECT")
	  sSubject = Replace(sSubject, txtRE & " ", "")
	  sSubject = Replace(sSubject, txtFWD & " ", "")
	  sSubject = replace(sSubject,"'","''")
	end if
	set rs = nothing
  else
    'sSubject = replace(ChkString(Request.Form("Subject"),"sqlstring"),"'","''")
  end if
  
  if typ = 0 and boolSend = true then
  	  if sSubject = " " or trim(sSubject) = "" then
    	strErr = "<li>" & txtMstPstSubj & "</li>"
		boolSend = False
	  else
	    strSubject = sSubject
  	  end if
  elseif typ = 1 or typ = 2 and boolSend = true then
    strSubject = txtRE & " " & sSubject
  elseif typ = 3 and boolSend = true then
   	strSubject = txtFWD & " " & sSubject
  else
    boolSend = false
  end if
		
	if Request.form("allmem") = "true" and hasAccess(1) and typ = 0 and boolSend = true then
		set rsName = server.CreateObject("adodb.recordset")
		strSql = "SELECT M_NAME"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE M_PMRECEIVE = 1 AND M_PMSTATUS = 1"
		strSQL = strSQL & " ORDER BY M_NAME Asc"
		set rsName = my_Conn.Execute (strSql)
		rsName.movefirst
		do while not rsName.eof
			arrAllNames = arrAllNames & rsName("M_NAME")
			rsName.moveNext
			if not rsName.eof then
				 arrAllNames = arrAllNames & ","
			end if
		loop
		rsName.close
		set rsName = nothing
		arrNames = split(arrAllNames, ",")
		'############## End PM all members ###################
	else
		arrNames = split(replace(Request.Form("sendto"),";",""), ",")
	end if

	  if boolSend then
	    for i = 0 to ubound(arrNames)
		  boolSend = true
		  strSql = "SELECT MEMBER_ID, M_NAME, M_PMRECEIVE, M_PMEMAIL"
		  strSql = strSql & ", M_EMAIL, M_PMSTATUS, M_LEVEL, M_PMBLACKLIST"
		  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		  strSql = strSql & " WHERE M_NAME='"& trim(arrNames(i)) &"'"
		  set rsName = my_Conn.Execute (strSql)

		  if rsName.BOF or rsName.EOF then '::  no one registered
		    if request.form("allmem") <> "true" then
			  strMessage2 = strMessage2 & "<li>" & txtSryMemNoFnd & " " & arrNames(i) & "</li>"
		    end if
		    boolSend = False
		  else		
		    if rsName("M_PMRECEIVE") = "0" or rsName("M_PMSTATUS") = 0 then
			  if request.form("allmem") <> "true" then
			    strMessage2 = strMessage2 & "<li>" & replace(txtSryNoRecPM,"[%member%]",arrNames(i)) & "</li>"
		      end if
			  boolSend = False
		    end if		'
		    if iMaxInbox > 0 and pmInCount(rsName("MEMBER_ID")) >= iMaxInbox and rsName("M_LEVEL") < 3 then
			  if request.form("allmem") <> "true" then
			    strMessage2 = strMessage2 & "<li>" & replace(txtSryInbxFul,"[%member%]",arrNames(i)) & "</li>"
		      end if
			  boolSend = False
		    end if
			':: check member blacklist if present
			if iMemberBlacklist then
			    if rsName("M_PMBLACKLIST") <> "" then
				  if instr(rsName("M_PMBLACKLIST"),",") > 0 then
				    arrTemp = split(rsName("M_PMBLACKLIST"),",")
					for bl = 0 to ubound(arrTemp)
					  if cLng(strUserMemberID) = cLng(arrTemp(bl)) then
			    		strMessage2 = strMessage2 & "<li>" & replace(txtSryNoRecPM,"[%member%]",arrNames(i)) & "</li>"
						boolSend = False
					  end if
					next
				  else
				    if strUserMemberID = cLng(rsName("M_PMBLACKLIST")) then
			    		strMessage2 = strMessage2 & "<li>" & replace(txtSryNoRecPM,"[%member%]",arrNames(i)) & "</li>"
						boolSend = False
					end if
				  end if
				end if
			end if
			':: end blacklist check
		  end if
		  if boolSend then
	
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
			strSql = strSql & ", '" & strMess & "'"
			strSql = strSql & ", " & rsName("MEMBER_ID")
			strSql = strSql & ", " & strUserMemberID
			strSql = strSql & ", '" & strCurDateString & "'"
			strSql = strSql & ", 0"
			strSql = strSql & ", 0"
			if request.cookies(strCookieURL & "PmOutBox") = "1" then
				strSql = strSql & ", '1')"
			else
				strSql = strSql & ", '0')"
			end if

			executeThis(strSql)

			'###### PM all members Mod - Take out e-mail notification #####
			if request.form("allmem") and request.form("allmem") <> "true" then	
				if strEmail = "1" and rsName("M_PMEMAIL") = "1" then
					DoPmEmail arrNames(i),rsName("M_EMAIL"),strSubject
				end if
			end if

			if Err.description <> "" then 
				strMessage2 = strMessage2 & "<li>" & txtWasErr & " = " & Err.description
			Elseif request.form("allmem") <> "true" then
				strMessage2 = strMessage2 & "<li>" & txtMsgSntTo & " " & arrNames(i) & "</li>"
			end if
		  end if
		  rsName.close
		  set rsName = nothing
	    next
	    if request.form("allmem") = "true" then
	      strMessage2 = strMessage2 & "<li>" & replace(txtMsgSntToCnt,"[%count%]",ubound(arrNames)+1) & "</li>"
	    end if
	  end if
	  
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
   '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

  if boolSend then
		select case typ
			case 0
				tTxt = txtMsgSent
			case 1
				tTxt = txtRplySent
			case 2
				tTxt = txtRplySent
			case 3
				tTxt = txtMsgFwdd
			case else
				tTxt = txtHvNicDay
		end select
  else
  	tTxt = txtWasErr
	tmpMsg = strErr & strMessage2
	strMessage2 = tmpMsg
  end if
  
  	  spThemeBlock1_open(intskin) %>
	  <table class="tPlain" style="width:100%;" width="100%"><tr><td width="100%">
	  <br />
	  <div class="fSubTitle" style="text-align:center;"><%= tTxt %></div>
	  <div class="fSubTitle" style="text-align:center;"><ul><%= strMessage2 %></ul></div>
	  <% if not boolPOP then %>
	  <p align="center"><a href="pm.asp"><%= txtGoBack %></a></p>
	  <% end if
	  	 if not boolPOP and boolSend then %>
	  	   <meta http-equiv="REFRESH" content="6; url=<%= strHomeUrl %>pm.asp">
	  <% end if
	  if not boolSend then %>
	    <p align="center">
	    <a href="JavaScript:history.go(-1)"><%= txtGoBack %></a></p>
<% 	  end if
	  response.Write("<p>&nbsp;</p></td></tr></table>")
  	  spThemeBlock1_close(intskin)
end sub

sub config_pm() %>
<script type="text/javascript">
function pmBanMembers() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=pmBan", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no");
}

function selectBanUsers()
{
	if (document.PostTopic.BlockedUsers.length == 1)
	{
		document.PostTopic.BlockedUsers.options[0].value = "";
		return;
	}
	if (document.PostTopic.BlockedUsers.length == 2)
		document.PostTopic.BlockedUsers.options[0].selected = true
	else
	for (x = 0;x < document.PostTopic.BlockedUsers.length - 1 ;x++)
		document.PostTopic.BlockedUsers.options[x].selected = true;
}

function DeleteBanSelection()
{
	var user,mText;
	var count,finished;

		finished = false;
		count = 0;
		count = document.PostTopic.BlockedUsers.length - 1;
		if (count<1) {
			return;
		}
		do //remove from source
		{	
			if (document.PostTopic.BlockedUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.BlockedUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.BlockedUsers.length-1;z++)
				{	
					document.PostTopic.BlockedUsers.options[z].value = document.PostTopic.BlockedUsers.options[z+1].value;	
					document.PostTopic.BlockedUsers.options[z].text = document.PostTopic.BlockedUsers.options[z+1].text;
				}
				document.PostTopic.BlockedUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
}
</script>
  <tr valign="middle"> 
    <td class="tSubTitle" align="center" colspan="2"><%= txtConfig %></td>
  </tr>
  <tr><td align="right" style="padding:2px;" class="fNorm">
      <b><%= txtAutoDel %> </b>
	</td>
	<td align="left" style="padding:2px;">
	<% 'auto delete on/off %>
	  <select name="iDATA1">
	    <option value="1"<%= CheckSelected(iDATA1,1) %>><%= txtOn %></option>
	    <option value="0"<%= CheckSelected(iDATA1,0) %>><%= txtOff %></option>
	  </select>
	  &nbsp;<span class="fNorm">(<%= txtAdminExemp %>)</span>
    </td>
  </tr>
  <tr><td align="right" style="padding:2px;" class="fNorm">
      <b><%= txtDaysTilAutoDel %> </b>
	</td>
	<td align="left" style="padding:2px;">
	<% 'days before auto delete %>
	  <select name="iDATA2">
	    <option value="10"<%= CheckSelected(iDATA2,10) %>>10</option>
	    <option value="15"<%= CheckSelected(iDATA2,15) %>>15</option>
	    <option value="30"<%= CheckSelected(iDATA2,30) %>>30</option>
	    <option value="60"<%= CheckSelected(iDATA2,60) %>>60</option>
	    <option value="90"<%= CheckSelected(iDATA2,90) %>>90</option>
	    <option value="120"<%= CheckSelected(iDATA2,120) %>>120</option>
	  </select>
    </td>
  </tr>
  <tr><td align="right" style="padding:2px;" class="fNorm">
      <b><%= txtMaxInbxSiz %> </b>
	</td>
	<td align="left" style="padding:2px;">
	  <% 'inbox size %>
	  <select name="iDATA3">
	    <option value="25"<%= CheckSelected(iDATA3,25) %>>25</option>
	    <option value="50"<%= CheckSelected(iDATA3,50) %>>50</option>
	    <option value="75"<%= CheckSelected(iDATA3,75) %>>75</option>
	    <option value="100"<%= CheckSelected(iDATA3,100) %>>100</option>
	    <option value="150"<%= CheckSelected(iDATA3,150) %>>150</option>
	    <option value="200"<%= CheckSelected(iDATA3,200) %>>200</option>
	    <option value="250"<%= CheckSelected(iDATA3,250) %>>250</option>
	    <option value="300"<%= CheckSelected(iDATA3,300) %>>300</option>
	    <option value="0"<%= CheckSelected(iDATA3,0) %>>Unlimited</option>
	  </select>
	  &nbsp;<span class="fNorm">(<%= txtAdminUnlim %>)</span>
    </td>
  </tr>
  <tr><td align="right" style="padding:2px;" class="fNorm">
      <b><%= txtMemIgnLst %> </b>
	</td>
	<td align="left" style="padding:2px;">
	<% 'allow member blacklist %>
	  <select name="iDATA4">
	    <option value="1"<%= CheckSelected(iDATA4,1) %>><%= txtOn %></option>
	    <option value="0"<%= CheckSelected(iDATA4,0) %>><%= txtOff %></option>
	  </select>
    </td>
  </tr>
  <tr><td align="center" style="padding:2px;" colspan="2">&nbsp;</td></tr>
  <tr><td align="right" style="padding:2px;" class="fNorm">
      <b><%= txtPMSavFolder %> </b>
	</td>
	<td align="left" style="padding:2px;">
	<% 'allow 'save' folder %>
	  <select name="iDATA5">
	    <option value="1"<%= CheckSelected(iDATA5,1) %>><%= txtOn %></option>
	    <option value="0"<%= CheckSelected(iDATA5,0) %>><%= txtOff %></option>
	  </select>
    </td>
  </tr>
  <tr><td align="center" style="padding:2px;" colspan="2" class="fNorm">
      <b><%= txtWhoHvSvdFldr %> </b>
	</td>
  </tr>
  	  <tr><td align="right" valign="middle" width="50%" class="fNorm" nowrap="nowrap">
  		<a href="JavaScript:allowgroups('PostTopic','tDATA1','<%= grpRead %>');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:removeGroup('PostTopic','tDATA1');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;
		</td>
          <td align="left"><p>
            <select size="5" name="tDATA1" style="width:120;" multiple>
			  <% 'if gRead <> "" then
			  		getOptGroups(tDATA1)
				 'end if %>
			  <option value="0"></option>
            </select></p>
          </td>
        </tr>
  <tr valign="middle"> 
    <td class="tSubTitle" align="center" colspan="2" style="padding:2px;"><%= txtPMaccess %></td>
  </tr>
  <tr valign="middle"> 
    <td align="center" colspan="2" style="padding:2px;" class="fNorm"><%= txtMemNoPMAcces %></td>
  </tr>
        <tr><td align="center" valign="middle" class="fNorm"><a href="JavaScript:pmBanMembers();" title="Add member to blacklist"><b><%= txtAddMem %></b></a><br />
				<a href="JavaScript:DeleteBanSelection();" title="Remove selected members(s) from this blacklist"><b><%= txtRemMem %></b></a></td>
          <td align="center"><p><b><%= txtBlkdMem %></b></p>
            <select size="5" name="BlockedUsers" style="width:170;" multiple>
			  <%= getBlockedUsers() %>
			  <option value="0"></option>
            </select>
            <input type="hidden" name="hasBlockedUsers" value="1"><br />&nbsp;<br />
          </td>
        </tr>
<%
end sub

function getBlockedUsers()
  tmpLst = ""
  bSQL = "select M_NAME, MEMBER_ID FROM " & strTablePrefix & "MEMBERS WHERE M_PMSTATUS=0"
  set pmban = my_Conn.execute(bSQL)
  if not pmban.eof then
    do until pmban.eof
      tmpLst = tmpLst & "<option value=""" & pmban("MEMBER_ID") & """>" & pmban("M_NAME") & "</option>" & vbcrlf
	  pmban.movenext
	loop
  end if
  set pmban = nothing
  getBlockedUsers = tmpLst
end function

function chkPMaccess(uid)
  tmpResult = false
  pSQL = "SELECT M_PMSTATUS FROM " & strMemberTablePrefix & "MEMBERS WHERE MEMBER_ID=" & uid
  set pmChk = my_Conn.execute(pSQL)
  if not pmChk.eof then
    if pmChk("M_PMSTATUS") = 1 then
	  tmpResult = true
	end if
  end if
  chkPMaccess = tmpResult
end function

function addToBlackList(mem_id)
	blUpdate = true
	strSQL = "SELECT M_LEVEL FROM " & strTablePrefix & "MEMBERS WHERE MEMBER_ID = " & mem_id
	set rsB = my_Conn.execute(strSQL)
	if not rsB.eof then
	  if blID = strUserMemberID then
		stMessage = txtNoAddSelf
		'stMessage = stMessage & "Don't hack the querystring... :)"
		blUpdate = false
	  elseif rsB("M_LEVEL") = 3 then
		stMessage = txtNoBLadmin
		blUpdate = false
	  else
	    strSQL = "SELECT M_PMBLACKLIST FROM " & strTablePrefix & "MEMBERS WHERE MEMBER_ID = " & strUserMemberID
		set rsBB = my_Conn.execute(strSQL)
		  tmpList = mem_id
		  if rsBB("M_PMBLACKLIST") <> "" and rsBB("M_PMBLACKLIST") <> "0" then
			if instr(rsBB("M_PMBLACKLIST"),",") > 0 then
			  arrTemp = split(rsBB("M_PMBLACKLIST"),",")
			  for bl = 0 to ubound(arrTemp)
			    if cLng(blID) = cLng(arrTemp(bl)) then
				  stMessage = txtMemAlrInBL
				  blUpdate = false
			    else
				  tmpList = tmpList & "," & arrTemp(bl)
			    end if
			  next
			else
			  if cLng(blID) = cLng(rsBB("M_PMBLACKLIST")) then
			    stMessage = txtMemAlrInBL
				blUpdate = false
			  else
				tmpList = mem_id & "," & rsBB("M_PMBLACKLIST")
			  end if
			end if
		  else
			tmpList = mem_id
		  end if
		  set rsBB = nothing
		  'insert list into db
		  if blUpdate then
			bSQL = "UPDATE " & strTablePrefix & "MEMBERS SET M_PMBLACKLIST = '" & tmpList & "' WHERE MEMBER_ID = " & strUserMemberID
			executeThis(bSQL)
			stMessage = txtMemAddBL
		  end if
	  end if
	else
	  stMessage = txtMemNoFnd
	end if
	set rsB = nothing
end function

function pmBlacklist(membID)
  if iMemberBlacklist then
  spThemeTitle= txtPMIgLst
  spThemeBlock1_open(intSkin)
  response.Write("<table border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">")%>
  <tr><td valign="top" width="100%" class="fNorm"><%= txtMemNotAllPM %><hr align="center"></td></tr>
<%
  blSQL = "select M_PMBLACKLIST from " & strTablePrefix & "MEMBERS where MEMBER_ID = " & membID
  set rsBL = my_Conn.execute(blSQL) 
  if not rsBL.EOF then
    'tmpBLKlist = "1,2,3,4"
    tmpBLKlist = rsBL("M_PMBLACKLIST")
	if trim(tmpBLKlist) <> "" then
	  if instr(tmpBLKlist,",") > 0 then
	    arrBlkLst = split(tmpBLKlist,",")
	    for bl = 0 to ubound(arrBlkLst)
	    blMemName = getMemberName(arrBlkLst(bl)) %>
  <tr> 
    <td  valign="top" width="100%" class="fNorm">&nbsp;
	<% if chkIsOnline(blMemName,0) then %>
	<img border="0" src="themes/<%= strTheme %>/icons/icon_block.gif" align="middle" alt="" /> 
	<% else %>
	<img border="0" src="themes/<%= strTheme %>/icons/icon_block.gif" align="middle" alt="" /> 
	<% end if %>
      <a href="pm.asp?unblock=<%= arrBlkLst(bl) %>"><b><%= blMemName %></b></a></td>
  </tr>
<%
	    next
	  else
	    arrBlkLst = tmpBLKlist %>
  <tr> 
    <td  valign="top" width="100%" class="fNorm">&nbsp;
	<% if chkIsOnline(blMemName,0) then %>
	<img border="0" src="themes/<%= strTheme %>/icons/icon_block.gif" align="middle" alt="" /> 
	<% else %>
	<img border="0" src="themes/<%= strTheme %>/icons/icon_block.gif" align="middle" alt="" /> 
	<% end if %>
      <a href="pm.asp?unblock=<%= arrBlkLst %>"><b><%= getMemberName(arrBlkLst) %></b></a></td>
  </tr>
<%
	  end if
	else%>
  <tr> 
    <td  valign="top" width="100%" align="center" class="fNorm">&nbsp;<b><%= txtEmpty %></b></td>
  </tr>
<%
	end if
  end if
  set rsBL = nothing%>
  <tr> 
    <td  valign="top" width="100%" align="center" class="fNorm"><hr /><b><%= txtClkMemToRem %></b></td>
  </tr>
<%
  response.Write("</table>")
spThemeBlock1_close(intSkin)
  end if
end function

function menu_PM()
	  'iAutoDelete
	  'iAutoDelDays
	  'iMaxInbox
	  'iMemberBlacklist
	  'iPMsaveFolder
	  'iPMfolderAccess
spThemeTitle= txtPMmenu
spThemeBlock1_open(intSkin)
  inBxCnt = cLng(pmInboxCount())
%>
<table border="0" cellpadding="0" cellspacing="1" width="100%">
  <tr> 
    <td  valign="top" width="100%" class="fNorm">&nbsp;<img border="0" src="images/icons/icon_sent_items2.gif" align="middle" width="15" height="15"> 
      <a href="pm.asp?cmd=2"><%= txtSndComp %></a></td>
  </tr>
  <tr> 
    <td  valign="top" width="100%">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="fNorm">&nbsp;<img border="0" src="images/icons/icon_inbox2.gif" align="middle" width="15" height="15">&nbsp;<a href="pm.asp"><%= txtPMinbox %></a>&nbsp;(<span class="fAlert"><%= clng(pmCount) %></span>)
          </td>
        </tr>
        <tr> 
          <td class="fNorm"> 
            <img border="0" src="images/clear.gif" align="middle" width="20" height="8"><img border="0" src="images/icons/icon_bar.gif" align="middle" width="15" height="15"> 
           &nbsp;<%= txtTotal %>&nbsp;(<span class="fAlert"><%= inBxCnt %></span>)</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top" width="100%"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <% If pm_out = 1 Then %>
        <tr> 
          <td class="fNorm">&nbsp;<img border="0" src="images/icons/icon_outbox2.gif" align="middle" width="15" height="15">&nbsp;<a href="pm.asp?cmd=1"><%= txtPMoutBx %> </a> 
            (<span class="fAlert"><%= pmSentCount() %></span>) 
          </td>
        </tr>
	  <% End If %>
        <tr> 
          <td class="fNorm">
            <img border="0" src="images/clear.gif" align="middle" width="20" height="8"><img border="0" src="images/icons/icon_bar.gif" align="middle" width="15" height="15"> 
           &nbsp;<%= txtUnread %>&nbsp;(<span class="fAlert"><%= pmSentUnreadCount() %></span>) </td>
        </tr>
      </table>
    </td>
  </tr>
<% if iPMsaveFolder and hasAccess(sPMfolderAccess) then %>
  <tr> 
    <td  valign="top" width="100%">
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td class="fNorm">&nbsp;<img border="0" src="images/icons/icon_inbox2.gif" align="middle" width="15" height="15">&nbsp;<a href="pm.asp?cmd=7"><%= txtSvdItems %></a>&nbsp;(<span class="fAlert"><%= clng(pmSavedCount) %></span>)
          </td>
        </tr>
      </table>
    </td>
  </tr>
<% end if %>
<% if iMaxInbox > 0 then
	  intPercent = cInt(Round((inBxCnt / iMaxInbox)*100, 2))
	  if intPercent < 75 then
	    bgimg = "themes/" & strTheme & "/icons/bar_gr.gif"
	  elseif intPercent >=75 and intPercent < 90 then
	    bgimg = "themes/" & strTheme & "/icons/bar_y.gif"
	  elseif intPercent >=90 and intPercent < 100 then
	    bgimg = "themes/" & strTheme & "/icons/bar_r.gif"
	  elseif intPercent >= 100 then
	    intPercent = 100
	    bgimg = "themes/" & strTheme & "/icons/bar_rg.gif"
	  end if
      %>
  <tr> 
    <td valign="top" align="center" width="100%"><hr />
	<div class="fSmall" style="margin-bottom:3px;"><%= txtInbQuota %>&nbsp;<% =inBxCnt %>&nbsp;<%= txtof %>&nbsp;<% =iMaxInbox %></div>
	<table cellpadding="1" border="0" cellspacing="0" width="100%" style="background-color:white; border: #104a7b 1px solid;">
	  <tr>
	     <td valign="middle">
		 <div class="fSmall" style="height:3px; width:<% =intPercent %>%;"><img src="<%= bgimg %>" width="100%" height="3" border="0"></div>
	    </td>
	  </tr>
	</table><hr />
    </td>
  </tr>
<% end if %>
</table>
<% 
dim boolPM
boolPM = 1

if boolPM = 1 and hasAccess(1) then %>
<table border="0" cellpadding="0" cellspacing="1" width="100%">
<tr>
<td height="18" class="tAltSubTitle" align="center" valign="top" width="100%"><b><%= txtPMstats %></b></td> </tr>
<tr>
<td valign="top" align="center" width="100%" class="fNorm"><br />
<span class="fAlert"><b><%= GetPMCount%></b></span>&nbsp;<%= txtPMinSys %>.<br />
<span class="fAlert"><b><%= GetPMToday%></b></span>&nbsp;<%= txtPMLst24Hrs %>.
<br /></td>
</tr>
</table>
<% end if %>
<%
spThemeBlock1_close(intSkin)
end function
%>
