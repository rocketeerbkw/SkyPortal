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
CurPageType = "core"
%>
<!-- #include file="inc_functions.asp" -->
<!-- #include file="Modules/pvt_msg/pm_custom.asp" -->
<%
bOnlineUsers = true
  Dim iPgType, iMode, pmid, pm_mid, stMessage
  
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = txtPvtMessgs
	PageAction = txtViewing & "<br />" 
	PageLocation = "pm.asp"
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"
end function
  
  intAppID = 0
  iPgType = 0
  iMode = 0
  pmid = 0
  
  if Request("cmd") <> "" or  Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iPgType = cLng(Request("cmd"))
	else
		closeAndGo("default.asp")
	end if
  end if
  if Request("mode") <> "" or  Request("mode") <> " " then
	if IsNumeric(Request("mode")) = True then
		iMode = cLng(Request("mode"))
	else
		closeAndGo("default.asp")
	end if
  end if
  if Request("pmid") <> "" or  Request("pmid") <> " " then
	if IsNumeric(Request("pmid")) = True then
		pmid = cLng(Request("pmid"))
	else
		closeAndGo("default.asp")
	end if
  end if

if iPgType = 2 then
  hasEditor = true
  strEditorElements = "Message"
end if
%>
<!-- #include file="inc_top.asp" -->
<% 
if not hasAccess(2) Then
	closeAndGo("default.asp")
elseif chkAPP("PM","USERS") and chkPMaccess(strUserMemberID) then
	'strUserMemberID = getMemberID(strDBNTUserName)
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
  arg1 = txtPM & "|pm.asp" 'this is for the page breadcrumb	
		strPMContent = ""
		strPMTitle = ""
		intArticleID = 0		
		strPostDate = ""
		dateSince=""
		boolPOP = false

      if Request.QueryString("save") = "1" then
		if iPMsaveFolder and hasAccess(sPMfolderAccess) then
	     strSql = "UPDATE " & strTablePrefix & "PM"
	     strSql = strSql & " SET " & strTablePrefix & "PM.M_SAVED = 1"
	     strSql = strSql & " WHERE " & strTablePrefix & "PM.M_ID = " & cLng(Request.QueryString("id")) & " AND " & strTablePrefix & "PM.M_TO = " & strUserMemberID
	     executeThis(strSql)
	    end if
      end if

      if Request("unblock") <> "" then
		if iMemberBlacklist then
		  blID = cLng(Request("unblock"))
		  blUpdate = false
		  tmpList = ""
		  stMessage = txtNotBlkLst
		  strSQL = "SELECT M_PMBLACKLIST FROM " & strTablePrefix & "MEMBERS WHERE MEMBER_ID = " & strUserMemberID
		  set rsUB = my_Conn.execute(strSQL)
		  if rsUB("M_PMBLACKLIST") <> "" and rsUB("M_PMBLACKLIST") <> "0" then
			if instr(rsUB("M_PMBLACKLIST"),",") > 0 then
			  arrTemp = split(rsUB("M_PMBLACKLIST"),",")
			  for bl = 0 to ubound(arrTemp)
			    if cLng(blID) = cLng(arrTemp(bl)) then
				  'Member matched
				  stMessage = txtRemBlkLst
				  blUpdate = true
			    else
				  tmpList = tmpList & arrTemp(bl) & ","
			    end if
			  next
			else
			  if cLng(blID) = cLng(rsUB("M_PMBLACKLIST")) then
			    'Member match
				stMessage = txtRemBlkLst
				blUpdate = true
			  else
				tmpList = rsUB("M_PMBLACKLIST")
			  end if
			end if
		  else
		    stMessage = txtBlkLstMT
		  end if
		  set rsUB = nothing
		  if blUpdate then
		    if right(tmpList,1) = "," then
			  tmpList = left(tmpList,len(tmpList)-1)
			end if
		    bSQL = "UPDATE " & strTablePrefix & "MEMBERS SET M_PMBLACKLIST = '" & tmpList & "' WHERE MEMBER_ID = " & strUserMemberID
		    executeThis(bSQL)
		    'stMessage = "Member removed from your Block List"
		  end if
		end if
	  end if

      if Request("block") <> "" then
		if iMemberBlacklist then
		 blID = cLng(Request("block"))
		 addToBlackList(blID)	     
	    end if
      end if

    sSql = "SELECT PM_OUTBOX FROM " & strTablePrefix & "CP_CONFIG WHERE MEMBER_ID = " & strUserMemberID
	set rx = my_Conn.execute(sSql)
	if rx.eof then
	  sSql = "INSERT INTO " & strTablePrefix & "CP_CONFIG (MEMBER_ID) values (" & strUserMemberID & ")"
	  executeThis(sSql)
	  pm_out = 0
	  'closeAndGo("cp_main.asp")
	else
	  pm_out = rx("PM_OUTBOX")
	end if
	set rx = nothing
	
  if Request.Cookies(strCookieURL & "PmOutBox") = "" then
	  Response.Cookies(strCookieURL & "PmOutBox").Path = strCookieURL
	  Response.Cookies(strCookieURL & "PmOutBox") = "" & pm_out & ""
	  Response.Cookies(strCookieURL & "PmOutBox").Expires = dateAdd("d", 360, strCurDateAdjust)
  end if
  
  if iMode=7 and iPgType = 7 then
	Dim strDeleteList
	strDeleteList = chkString(Request("Delete"),"sqlstring")
	
	if trim(strDeleteList) = "" then
		stMessage = txtNoMsgSel
	Else
		'Now, use the SQL set notation to move all of the records
		'specified by strDeleteList
		strSqL = "UPDATE " & strTablePrefix & "PM "
		strSql = strSql & "SET M_SAVED = 0 WHERE M_ID IN (" & strDeleteList & ")"
		executeThis(strSql)
		Mcnt = Request("Delete").Count
		stMessage = Mcnt & " " & txtMsgMvdInbx
	End If
  end if
  
  if (iMode=1 or iMode=0) and iPgType = 6 then
    pmDelete2()
    if iMode = 1 then
      iPgType = 1
  	else
      iPgType = 0
  	end if
  end if
%>
<script type="text/javascript">
<!--
function CheckAll2(Remove) {
	for( i=0; i<document.RemoveTopic.elements.length; i++) {
		if (document.RemoveTopic.elements[i].name==Remove) {
			document.RemoveTopic.elements[i].checked=document.RemoveTopic.markall2.checked;
		}
	}
}
function CheckAll(DELETE) {
	for( i=0; i<document.DeleteTopic.elements.length; i++) {
		if (document.DeleteTopic.elements[i].name==DELETE) {
		document.DeleteTopic.elements[i].checked=document.DeleteTopic.markall.checked;
		}
	}
}

function pmmembers() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=pm", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no");
}
//-->
</SCRIPT>

<table width="100%" cellspacing="0" style="border-collapse: collapse;">
<tr><td class="leftPgCol"><% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
</td>
<td class="mainPgCol">
<%
intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtPvtMessgs & "|pm.asp"
  arg2 = ""
  arg3 = ""
  'arg4 = ""
  'arg5 = ""
  'arg6 = ""
%>
<% 
  select case iPgType
	case 1
  	  arg2 = txtSntMsgs & "|pm.asp?cmd=1"
	  showSentItems()
	case 2
  	  arg2 = txtCreatMsg & "|pm.asp?cmd=2"
  	  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
	  composePM(iMode)
	case 3
  	  arg2 = ""
	  sendPM(iMode)
	case 4
  	  arg2 = txtRdMsg & "|pm.asp?cmd=4"
	  readPM()
	case 5
  	  arg2 = txtRdSntMsg & "|pm.asp?cmd=1"
	  readSent()
	case 6
  	  arg2 = txtDelMsg & "|pm.asp"
	  pmDelete2()
	case 7
  	  arg2 = txtSvdMsgs & "|pm.asp?cmd=7"
	  showSavedBox()
	case else
  	  arg2 = txtRecMsgs & "|pm.asp"
	  showInbox()
  end select  %>
  <% app_MainColumn_bottom() %>
  <div class="clsSpacer"></div>
</td>
</tr>
</table>
<% app_footer()
else
	closeAndGo("default.asp")
end if %>
<!-- #include file="inc_footer.asp" -->
<%

sub pmDelete2()
  boolSend = true
  If Request.Form("RemoveTopic") = "1" then 'remove from 'sent items'
	'The list that needs
	'to be deleted are in a comma-delimited list...
	Dim strRemoveList
	strRemoveList = chkString(Request("Remove"),"sqlstring")

	if trim(strRemoveList) = "" then
		strErr = strErr & "<li>" & txtNoMsgSel & ".</li>"
		boolSend = false
	Else
		'Now, use the SQL set notation to Remove all of the records
		'specified by strRemoveList
		Dim strSqL
		strSqL = "UPDATE " & strTablePrefix & "PM "
		strSql = strSql & "SET " & strTablePrefix & "PM.M_OUTBOX = 0 " & _
				 "WHERE M_ID IN (" & strRemoveList & ")"
		executeThis(strSql)
		
		Mcnt = Request("Remove").Count
		stMessage = stMessage & "<li>" & Mcnt & " " & txtMsgRem & "</li>"
	End If
  else 'remove from inbox
	Dim strDeleteList
	strDeleteList = chkString(Request("Delete"),"sqlstring")
	
	if trim(strDeleteList) = "" then
		strErr = strErr & "<li>" & txtNoMsgSel & ".</li>"
		boolSend = false
	Else
		'Now, use the SQL set notation to delete all of the records
		'specified by strDeleteList
		strSQL = "DELETE FROM " & strTablePrefix & "PM " & _
				 "WHERE M_ID IN (" & strDeleteList & ") AND M_TO = " & strUserMemberID

		executeThis(strSql)
		
		Mcnt = Request("Delete").Count
		stMessage = stMessage & "<li>" & Mcnt & " " & txtMsgDel & "</li>"
	End If
  end if
	
  if not boolSend then
  	tTxt = "<br /><span class=""fSubTitle"">" & txtThereIsProb & "</span>"
	tmpMsg = strErr & stMessage
	stMessage = tmpMsg
  end if
end sub

sub readSent()
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_MESSAGE " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & pmid
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

	Set rsMessage = my_Conn.Execute(strSql)

	if rsMessage.BOF or rsMessage.EOF then
	   Response.Redirect("pm.asp")
	end if

	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_CITY, " & strMemberTablePrefix & "MEMBERS.M_STATE, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strMemberTablePrefix & "MEMBERS.M_MSN "    
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & pmid

	Set rs = my_Conn.Execute(strSql)

spThemeTitle = txtSntMsgs
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tCellAlt1"" width=""100%"" cellpadding=""4"">")
%>
  <tr>
    <td align="center" class="tSubTitle" width="150" nowrap><b><%= txtSntTo %>:</b></td>
    <td align="left" class="tSubTitle"><b><%= txtSubject %>:&nbsp;&nbsp; <% =ChkString(rsMessage("M_SUBJECT"),"display") %></b></td>
  </tr>
  <tr>
    <td class="fNorm" valign="top">
      
      <%
	  	  strIMmsg = txtView & " " & ChkString(rs("M_NAME"),"display") & "'s " & txtProfile %>
   	<a href="cp_main.asp?cmd=8&member=<% =ChkString(rs("M_TO"),"displayimage") %>" title="<%= strIMmsg %>"><b><%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %></b></a>
	  <% 
	    if strShowRank = 2 or strShowRank = 3 then %>
        <% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %><br />
<%    end if
		  dnrLvl = getDonor_Level(rs("MEMBER_ID"))
		  if dnrLvl <> "" then
		  response.Write(dnrLvl)
		  end if
		 %>
<%    if strShowRank = 1 or strShowRank = 3 then %>
        <br /><small><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></small>
<%    end if %>
      <br />
      <br /><small><% =ChkString(rs("M_COUNTRY"),"display") %></small>
      <br /><small><% =ChkString(rs("M_POSTS"),"display") %>&nbsp;<%= txtPosts %></small>
    </td>
    <td class="fNorm" valign="top" width="100%">
      <img src="images/icons/icon_posticon.gif" border="0" hspace="3"><%= txtSent %>&nbsp;-&nbsp;<% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %>
      	&nbsp;<a href="cp_main.asp?cmd=8&member=<% =ChkString(rs("MEMBER_ID"),"displayimage") %>"><img src="images/icons/icon_profile.gif" height="15" width="15" title="<%= txtViewProf %>" alt="<%= txtViewProf %>" border="0" align="absmiddle" hspace="6"></a>
      &nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =ChkString(rs("M_TO"),"displayimage") %>')"><img src="images/icons/icon_email.gif" height="15" width="15" title="<%= txtEmlMbrs %>" alt="<%= txtEmlMbrs %>" border="0" align="absmiddle" hspace="6"></a>
      &nbsp;<a href="pm.asp?cmd=2&amp;mode=3&amp;pmid=<% =pmid%>"><img src="images/icons/icon_reply_topic.gif" height="15" width="15" title="<%= txtFwdMsg %>" alt="<%= txtFwdMsg %>" border="0" align="absmiddle" hspace="6"></a>
<%
	    if strMSN = "1" then
	    	if trim(rs("M_MSN")) <> "" then %>
          &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=3&msn=<% =ChkString(replace(rs("M_MSN"),"@","[no-spam]@"),"urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="images/icons/icon_msn.gif" height=15 width=15 alt="" border="0" align="absmiddle" hspace="6"></a>
<%		end if 
		end if 
	    if strICQ = "1" then
	    	if trim(rs("M_ICQ")) <> "" then %>
          &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=1&ICQ=<% = ChkString(rs("M_ICQ"),"urlpath")  %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="http://web.icq.com/whitepages/online?icq=<% = ChkString(rs("M_ICQ"),"display")  %>&img=5" alt="" border="0" align="absmiddle" hspace="3"></a>
<%		end if 
		end if 
	    if strYAHOO = "1" then
	      if rs("M_YAHOO") <> " " then %>
          &nbsp;<a href="JavaScript:openWindow('http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rs("M_YAHOO"),"urlpath") %>&.src=pg')"><img src="images/icons/icon_yahoo.gif" height=15 width=15 alt="" border="0" align="absmiddle" hspace="6"></a>
<%		end if 
		end if 
	    if (strAIM = "1") then
	    	if rs("M_AIM") <> " " then %>
          &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rs("M_AIM"),"urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="images/icons/icon_aim.gif" height=15 width=15 alt="" border="0" align="absmiddle" hspace="6"></a>
<%		end if 
		end if %>
      <hr noshade size=1>
      
	  <% if strAllowHtml = 1 then %>
	  <%= ReplaceUrls(rsMessage("M_MESSAGE")) %>
	  <% Else %>
	  <%= formatStr(rsMessage("M_MESSAGE")) %>
	  <% End If %>
    </td>
  </tr>
<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)

end sub

sub readPM()
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
	strSql = "UPDATE " & strTablePrefix & "PM "
	strSql = strSql & " SET " & strTablePrefix & "PM.M_READ = 1 "
	strSql = strSql & " WHERE ((" & strTablePrefix & "PM.M_ID = " & pmid & ") AND (" & strTablePrefix & "PM.M_TO = " & strUserMemberID & "));"
	executeThis(strSql)

	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & pmid
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

	Set rsMessage = my_Conn.Execute(strSql)

	if rsMessage.BOF or rsMessage.EOF then
	   Response.Redirect("pm.asp")
	end if

	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_CITY, " & strMemberTablePrefix & "MEMBERS.M_STATE, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_SUBJECT, " & strMemberTablePrefix & "MEMBERS.M_MSN "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_ID =  " & pmid

	Set rs = my_Conn.Execute(strSql)
	pm_mid = rs("MEMBER_ID")
	
pm_buttons()

spThemeBlock1_open(intSkin)
%><table class="tCellAlt1" width="100%">
  <tr>
    <td>
 <table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td align="center" class="tSubTitle" width="125" nowrap><b><%= txtFrom %></b></td>
    <td align="left" class="tSubTitle"><b><%= txtSubject %>:&nbsp;&nbsp; <% =rsMessage("M_SUBJECT") %></b></td>
  </tr>
  <tr>
    <td class="fNorm" valign="top" align="center"><% 
	  		strIMmsg = txtView & " " & ChkString(rs("M_NAME"),"display") & "'s " & txtProfile %>
      	<a href="cp_main.asp?cmd=8&member=<% =rs("M_FROM") %>" title="<%= strIMmsg %>">
      	<span class="fSubTitle">
	  <%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %>
	  	</span></a>
<%	    if strShowRank = 2 or strShowRank = 3 then %><br />
        <% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %><br />
<%      end if
		  dnrLvl = getDonor_Level(rs("MEMBER_ID"))
		  if dnrLvl <> "" then
		  response.Write(dnrLvl)
		  end if
		 %>
<%	if strShowRank = 1 or strShowRank = 3 then %><br /><small><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></small>
<%  	end if %>
          <br />
          <br /><small><% =rs("M_COUNTRY") %></small>
          <br /><small><% =rs("M_POSTS") %>&nbsp;<%= txtPosts %></small>
    </td>
    <td class="fNorm" valign="top">
      <img src="images/icons/icon_posticon.gif" border="0" hspace="3"><%= txtSent %>&nbsp;-&nbsp;<% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %>
      <hr />
	  <%= formatStr(replace(rsMessage("M_MESSAGE"),"''","'")) %>
    </td>
  </tr>
    </table></td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)%>
<% 'end if	
end sub

sub showSentItems()
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
   %>
<form action="pm.asp?cmd=6&amp;mode=1" method="post" name="RemoveTopic">
<input name="RemoveTopic" type="hidden" value="1">

<%
spThemeTitle= strDBNTUserName &"'s " & txtPMOutBx 
spThemeBlock1_open(intSkin)

if strDBType = "access" then
strSqL = "SELECT count(M_TO) as [oboxcount] " 
else
strSqL = "SELECT count(M_TO) as oboxcount " 
end if
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
strSql = strSql & " AND " & strTablePrefix & "PM.M_OUTBOX = 1" 

Set rsOboxcount = my_Conn.Execute(strSql)
oboxcount = rsOboxcount("oboxcount")
rsOboxcount.close
set rsOboxcount = nothing
%> 
<table class="tPlain" width="100%">
  <%
  if stMessage <> "" then %>
  <tr><td height="30" valign="middle" align="center"><h3><%= stMessage %></h3></td></tr>
  <%
  end if %>
<tr><td height="30">
&nbsp;&nbsp;&nbsp;<img src="themes/<%= strTheme %>/icon_pmold.gif" align="middle" border="0" width="24" height="18">&nbsp;<span class="fNorm"><%= replace(txtSntMsgCnt,"[%count%]",oboxcount) %>:</span>
</td></tr>
<tr><td>
<table cellspacing="1" cellpadding="4" width="100%">
  <tr class="tSubTitle">
    <td valign="middle">&nbsp;</td>
    <td valign="middle"><B><%= txtSubject %></B></td>
    <td valign="middle" nowrap><B><%= txtSntTo %></B></td>
    <td valign="middle" nowrap><B><%= txtDtSnt %></B></td>
    <td valign="middle"><B><%= txtRemove %>?</B></td>
  </tr>
<%
	'#PM SQL get private messages
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ, " & strTablePrefix & "PM.M_OUTBOX "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_OUTBOX = 1" 
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

	Set rsFMessage = my_Conn.Execute(strSql)

	if rsFMessage.EOF or rsFMessage.BOF then  '## No Private Messages found in DB
%>
  <tr>
    <td>&nbsp;</td>
    <td colspan="4" class="fSubTitle"><b><%= txtNoSentPMs %></b></td>
  </tr>
<%	else
	        i = 0
		do Until rsFMessage.EOF

		'#PM SQL get Message MemberName
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strTablePrefix & "PM.M_ID  "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsFMessage("M_TO") & ""

		Set rsTo = my_Conn.Execute(strSql)

		if i = 1 then
			CColor = "tCellAlt1"
		else
			CColor = "tCellAlt2"
		end if
	if rsTo.EOF or rsTo.BOF then
%>	
  <tr>
    <td colspan="5" align=center class="<% =CColor %>"><span class="fNorm"><%= txtMemDel %></span></td>
  </tr>
<%	
	else
%>
  <tr class="<% =CColor %>">
    <td align="center" width="20">
      <% if rsFMessage("M_READ") = "0" then %>
        <img title="<%= txtMsgNotRead %>" alt="<%= txtMsgNotRead %>" src="themes/<%= strTheme %>/icon_pmnew.gif" width="24" height="18">
      <% else %>
        <img title="<%= txtOldMsg %>" alt="<%= txtOldMsg %>" src="themes/<%= strTheme %>/icon_pmread.gif" width="24" height="18">
      <% end if %>
    </td>
    <td class="fNorm"><a href="pm.asp?cmd=5&amp;pmid=<% =rsFMessage("M_ID") %>"><% =chkString(rsFMessage("M_SUBJECT"),"display") %></a></td>
    <td width="100" class="fNorm">
         <a href="cp_main.asp?cmd=8&member=<% =chkString(rsFMessage("M_TO"),"displayimage") %>"><%= displayName(ChkString(rsTo("M_NAME"),"display"),rsTo("M_GLOW")) %></a>
	  </td>
    <td nowrap width="175" class="fNorm"><% =ChkDate(rsFMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsFMessage("M_SENT")) %></td>
    <td align="center" width="60" class="fNorm"><input type="checkbox" name="Remove" value="<% =rsFMessage("M_ID") %>"></td>
  </tr>
<%            end if
		rsFMessage.MoveNext
		i = i + 1
		if i = 2 then i = 0
	    Loop
	end if
%>
</table>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="left" width="50%">
          <p> 
          &nbsp;&nbsp;&nbsp;<img title="<%= txtMsgNotRead %>" alt="<%= txtMsgNotRead %>" src="themes/<%= strTheme %>/icon_pmnew.gif" width="24" height="18">&nbsp;<%= txtMsgNotRead %>.<br />
          &nbsp;&nbsp;&nbsp;<img title="<%= txtMsgIsRead %>" alt="<%= txtMsgIsRead %>" src="themes/<%= strTheme %>/icon_pmread.gif" width="24" height="18">&nbsp;<%= txtMsgIsRead %>.</p>
        </td>
        <td align="right">
		<p><%= txtSelAll %> &nbsp;&nbsp;
		<input type="checkbox" name="markall2" onclick="CheckAll2('Remove');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><br />
        <input class="button" type="submit" value="<%= txtRemChkItms %>">&nbsp;&nbsp;<br />
		<img alt="" src="images/clear.gif" width="1" height="6"></p>
        </td>
      </tr>
    </table>
</td></tr></table>
<%
spThemeBlock1_close(intSkin) %>
</form>
<%
end sub

sub showInbox()
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
   %>
<form action="pm.asp?cmd=6&amp;mode=0" method="post" name="DeleteTopic">
<%
spThemeTitle= strDBNTUserName &"'s " & txtPMInBx
spThemeBlock1_open(intSkin)
Response.Write("<table width=""100%"" cellpadding=""3"" cellspacing=""0"">")

if strDBType = "access" then
strSqL = "SELECT count(M_TO) as [pmcnts] " 
else
strSqL = "SELECT count(M_TO) as pmcnts " 
end if

strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 

Set rsMPcount = my_Conn.Execute(strSql)
pmcot = rsMPcount("pmcnts")
rsMPcount.close
set rsMPcount = nothing

  if stMessage <> "" then %>
  <tr><td height="30" valign="middle" align="center" class="fNorm"><h3><%= stMessage %></h3></td></tr>
  <%
  end if
   If iAutoDelete then %>
<tr><td height="30" align="center" valign="middle" class="fNorm"><b><%= replace(txtDelPMDays,"[%days%]",iAutoDelDays) %></b></td></tr>
<% end if %>
<tr><td height="30" valign="middle" class="fNorm">&nbsp;&nbsp;&nbsp;<img src="themes/<%= strTheme %>/icon_pmnew.gif" align="middle" border="0" width="24" height="18">&nbsp;<%= replace(txtCntUnread,"[%count%]",pmcot) %>:</td></tr>
<tr><td>
<table cellspacing="0" cellpadding="2" width="100%">
      <tr class="tSubTitle">
        <td>
		<% if iPMsaveFolder and hasAccess(sPMfolderAccess) then %>
			<%= txtSave %>
		<% else %>
		    &nbsp;
		<% end if %>
		</td>
        <td><B><%= txtSubject %></B></td>
        <td nowrap><B><%= txtFrom %></B></td>
        <td nowrap><B><%= txtDtSnt %></B></td>
        <td><B><%= txtRemove %>?</B></td>
      </tr>
<%
      if Request.QueryString("marknew") = "1" then
	     strSql = "UPDATE " & strTablePrefix & "PM "
	     strSql = strSql & " SET " & strTablePrefix & "PM.M_READ = 0 "
	     strSql = strSql & " WHERE (" & strTablePrefix & "PM.M_ID = " & chkString(Request.QueryString("id"),"numeric") & " AND " & strTablePrefix & "PM.M_TO = " & strUserMemberID & ");"
	     executeThis(strSql)
      end if
      Response.Flush

      strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
      strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
      strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
      strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
      strSql = strSql & " AND " & strTablePrefix & "PM.M_SAVED = 0"
      strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

      Set rsMessage = my_Conn.Execute(strSql)

   if rsMessage.EOF or rsMessage.BOF then  '## No Private Messages found in DB
%>
      <tr>
        <td>&nbsp;</td>
        <td colspan="4" class="fSubTitle"><b><%= txtNoPMs %></b></td>
      </tr>
<% else
        i = 0
      	do Until rsMessage.EOF

	  strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW,  " & strTablePrefix & "PM.M_ID  "
	  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsMessage("M_FROM") & ""

          Set rsFrom = my_Conn.Execute(strSql)	

	    if i = 1 then 
	    	CColor = "tCellAlt1"
	    else
	    	CColor = "tCellAlt0"
	    end if
%>

      <tr class="<% =CColor %>">
        <td align="center" width="20" class="fNorm">
        <% if rsMessage("M_READ") = "0" then %>
            <img alt="New Message" src="themes/<%= strTheme %>/icon_pmnew.gif">
        <% else %>
			<% if iPMsaveFolder and hasAccess(sPMfolderAccess) then %>
            	<a href="pm.asp?save=1&id=<% =rsMessage("M_ID") %>"><img alt="<%= txtSave %>" title="<%= txtMvSaved %>" src="themes/<%= strTheme %>/icon_pmread.gif" border="0"></a>
          	<% else %>
            	<a href="pm.asp?marknew=1&id=<% =rsMessage("M_ID") %>"><img title="<%= txtMkNew %>" alt="<%= txtMkNew %>" src="themes/<%= strTheme %>/icon_pmread.gif" border="0"></a>
          	<% end if %>
		<% end if %>
        </td>
        <td class="fNorm"><a href="pm.asp?cmd=4&amp;pmid=<% =chkString(rsMessage("M_ID"),"display") %>"><% =replace(chkString(rsMessage("M_SUBJECT"),"display"),"''","'") %></a></td>
        <td class="fNorm" width="100">
	  <a href="cp_main.asp?cmd=8&member=<% =chkString(rsMessage("M_FROM"),"display") %>"><%= displayName(ChkString(rsFrom("M_NAME"),"display"),rsFrom("M_GLOW")) %></a>
        </td>
		<td class="fNorm" width="175" nowrap><% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %></td>
        <td align="center" class="fNorm" width="50"><input type="checkbox" name="DELETE" value="<% =chkString(rsMessage("M_ID"),"display") %>"></td>
      </tr>
<%	
	    rsMessage.MoveNext
	    i = i + 1
	    if i = 2 then i = 0
	  Loop
  end if
%>
</table>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="left" width="50%">
          <p> 
          &nbsp;&nbsp;&nbsp;<img title="<%= txtNewMsg %>" alt="<%= txtNewMsg %>" src="themes/<%= strTheme %>/icon_pmnew.gif" width="24" height="18">&nbsp;<%= txtNewMsg %><br />
          &nbsp;&nbsp;&nbsp;<img title="<%= txtOldMsg %>" alt="<%= txtOldMsg %>" src="themes/<%= strTheme %>/icon_pmread.gif" width="24" height="18">&nbsp;<%= txtOldMsg %></p>
        </td>
        <td align="right">
		<p><%= txtSelAll %> &nbsp;&nbsp;
		<input type="checkbox" name="markall" onclick="CheckAll('DELETE');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><br />
        <input class="button" type="submit" value="<%= txtRemChkItms %>">&nbsp;&nbsp;<br />
		<img alt="" src="images/clear.gif" width="1" height="6"></p>
        </td>
      </tr>
    </table>
</td></tr>
<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
</form>
<%
end sub

sub showSavedBox()
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
   %>
<form action="pm.asp?cmd=7&amp;mode=7" method="post" name="DeleteTopic">
<%
spThemeTitle= strDBNTUserName &"'s Private Messages Saved Box"
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"" width=""100%"">")

if strDBType = "access" then
strSqL = "SELECT count(M_TO) as [pmcnts] " 
else
strSqL = "SELECT count(M_TO) as pmcnts " 
end if

strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 

Set rsMPcount = my_Conn.Execute(strSql)
pmcot = rsMPcount("pmcnts")
rsMPcount.close
set rsMPcount = nothing

if stMessage <> "" then %>
<tr><td height="30" valign="middle" class="fNorm"><%= stMessage %></td></tr>
<%
end if %>
<tr><td height="30" valign="middle" class="fNorm">&nbsp;&nbsp;&nbsp;<img src="themes/<%= strTheme %>/icon_pmnew.gif" align="middle" border="0" width="24" height="18">&nbsp;&nbsp;<%= replace(txtCntUnread,"[%count%]",pmcot) %>:</td></tr>
<tr><td>
<table cellspacing="1" cellpadding="2" width="100%">
      <tr class="tSubTitle">
        <td>
		<% if iPMsaveFolder and hasAccess(sPMfolderAccess) then %>
			<%= txtSave %>
		<% else %>
		    &nbsp;
		<% end if %>
		</td>
        <td><B><%= txtSubject %></B></td>
        <td nowrap><B><%= txtFrom %></B></td>
        <td nowrap><B><%= txtDtSnt %></B></td>
        <td><B><%= txtRemove %>?</B></td>
      </tr>
<%
      strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
      strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
      strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & strDBNTUserName & "'"
      strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
      strSql = strSql & " AND " & strTablePrefix & "PM.M_SAVED = 1"
      strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 

      Set rsMessage = my_Conn.Execute(strSql)

   if rsMessage.EOF or rsMessage.BOF then  '## No Private Messages found in DB
%>
      <tr>
        <td>&nbsp;</td>
        <td colspan="4" class="fSubTitle"><b><%= txtNoSvdPMs %></b></td>
      </tr>
<% else
        i = 0
      	do Until rsMessage.EOF

	  strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW,  " & strTablePrefix & "PM.M_ID  "
	  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsMessage("M_FROM") & ""

          Set rsFrom = my_Conn.Execute(strSql)	

	    if i = 1 then 
	    	CColor = "tCellAlt1"
	    else
	    	CColor = "tCellAlt2"
	    end if
%>

      <tr class="<% =CColor %>">
        <td align="center" width="20" class="fNorm">
           <img title="<%= txtSvdMsg %>" alt="<%= txtSvdMsg %>" src="themes/<%= strTheme %>/icon_pmread.gif" border="0">
        </td>
        <td class="fNorm"><a href="pm.asp?cmd=4&amp;pmid=<% =chkString(rsMessage("M_ID"),"display") %>"><% =replace(chkString(rsMessage("M_SUBJECT"),"display"),"''","'") %></a></td>
        <td class="fNorm" width="100">
	  <a href="cp_main.asp?cmd=8&member=<% =chkString(rsMessage("M_FROM"),"display") %>"><%= displayName(ChkString(rsFrom("M_NAME"),"display"),rsFrom("M_GLOW")) %></a>
        </td>
		<td nowrap class="fNorm" width="175"><% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %></td>
        <td align="center" width="50"><input type="checkbox" name="DELETE" value="<% =chkString(rsMessage("M_ID"),"display") %>"></td>
      </tr>
<%	
	    rsMessage.MoveNext
	    i = i + 1
	    if i = 2 then i = 0
	  Loop
  end if
%>
</table>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="left" width="50%" class="fNorm"><b>&nbsp;<%= txtRemRetInbox %>.</b>
        </td>
        <td align="right">
		<br /><span class="fNorm"><%= txtSelAll %></span> &nbsp;&nbsp;
		<input type="checkbox" name="markall" onclick="CheckAll('DELETE');">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br /><br />
        <input class="button" type="submit" value="<%= txtRemChkItms %>">&nbsp;&nbsp;<br />
		<img alt="" src="images/clear.gif" width="1" height="6">
        </td>
      </tr>
    </table>
</td></tr>
<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
</form>
<%
end sub

function pm_buttons() %>
<table border="0" width="95%">
  <tr>
    <td  height="40">&nbsp;</td>
		
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm.asp">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmreceive.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtChkNewMsgs %>" alt="<%= txtChkNewMsgs %>"
	          onmouseover="document.btnreveive.style.top=-37" onmouseout="document.btnreveive.style.top=0"
	          onmousedown="document.btnreveive.style.top=-37" onmouseup="document.btnreveive.style.top=0"; /></A>
         <img name="btnreveive" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm.asp?cmd=2&amp;mode=1&amp;pmid=<% =pmid%>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmreply.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtRplyMsg %>" alt="<%= txtRplyMsg %>" onmouseover="document.btnreply.style.top=-37" onmouseout="document.btnreply.style.top=0" onmousedown="document.btnreply.style.top=-37" onmouseup="document.btnreply.style.top=0"; /></A>
         <img name="btnreply" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm.asp?cmd=2&amp;mode=2&amp;pmid=<% =pmid%>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmreplyQ.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtRplyQtMsg %>" alt="<%= txtRplyQtMsg %>"
	          onmouseover="document.btnreplyQ.style.top=-37" onmouseout="document.btnreplyQ.style.top=0"
	          onmousedown="document.btnreplyQ.style.top=-37" onmouseup="document.btnreplyQ.style.top=0"; /></A>
         <img name="btnreplyQ" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm.asp?cmd=2&amp;mode=3&amp;pmid=<% =pmid%>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmforward.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtFwdMsg %>" alt="<%= txtFwdMsg %>"
	          onmouseover="document.btnforward.style.top=-37" onmouseout="document.btnforward.style.top=0"
	          onmousedown="document.btnforward.style.top=-37" onmouseup="document.btnforward.style.top=0"; /></A>
         <img name="btnforward" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
<% if iPMsaveFolder and hasAccess(sPMfolderAccess) then %>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><A HREF="pm.asp?save=1&id=<% =pmid%>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmsave.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtMvSaved %>" alt="<%= txtMvSaved %>"
	          onmouseover="document.btnsave.style.top=-37" onmouseout="document.btnsave.style.top=0"
	          onmousedown="document.btnsave.style.top=-37" onmouseup="document.btnsave.style.top=0"; /></A>
         <img name="btnsave" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
<% end if %>
<% if iMemberBlacklist and pm_mid <> strUserMemberID then %>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><a href="pm.asp?block=<%= pm_mid %>">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmblocklist.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtAddBlkLst %>" alt="<%= txtAddBlkLst %>" onmouseover="document.btnblock.style.top=-37" onmouseout="document.btnblock.style.top=0" onmousedown="document.btnblock.style.top=-37" onmouseup="document.btnblock.style.top=0"; /></a>
         <img name="btnblock" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90" />
	   </div>
	</div>
	</td>
<% end if %>
	<td align="left" valign="top" width="95" height="40">
	<div unselectable='on' style="position:absolute;margin-left:1;margin-right:1;width:93;height:37";>
       <div unselectable='on' style="position:absolute;clip: rect(0 90px 37px 0)"><a href="javascript:;" onclick="JavaScript:popUpWind('pm_pop.asp?mode=3&amp;cid=<%=pmid%>','pm','50','50','yes','yes');">
	     <img unselectable='on' src="<%= strHomeURL %>images/icons/pmdelete.gif"  style="width:90px;height:37px;border:0px;z-index:100;position:absolute;top:-0" title="<%= txtDelMsg %>" alt="<%= txtDelMsg %>" onmouseover="document.btndelete.style.top=-37" onmouseout="document.btndelete.style.top=0" onmousedown="document.btndelete.style.top=-37" onmouseup="document.btndelete.style.top=0"; /></a>
         <img name="btndelete" unselectable='on' src="<%= strHomeURL %>Themes/<%= strTheme %>/btn_pm.gif" style="z-index:10;position:absolute;top:-0;width:90">
	   </div>
	</div>
	</td>
	 </tr>
</table>
<%
end function
%>