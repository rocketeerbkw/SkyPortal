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
'<>
'<> Do not ML the strDebug variables
'<>
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
dim deBug, strResult, strDebug, daysOld, pgType
pgType = "memberConfig"
strDebug = ""
strResult = ""
deBug = false
daysOld = "15" 'days old that get shown in bold
%>

<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<script type="text/javascript">
function ChangePage(n){
	if (n == 1) {
    	document.Forms.PageNum1.submit();
    }
    else {
    	document.forms.PageNum2.submit();
    }
}

function cbValidate(cb,frm) {
	for (j = 1; j < 3; j++) {
		if (document[frm][cb].checked == true) {
			document[frm][cb].checked = false;
      	}
   	}
}
</script>
<%
'strEmail = 1
'strEmailVal = 7
if request("dog") = "1" then
 deBug = true
end if
catID = Request.Form("cat")
if catID <> "" and hasAccess(1) then
    if strEmail = 1 then oo = "ON" else oo = "OFF" end if
	strDebug = strDebug & "<li><b>Site Email is turned " & oo & "!</b></li>"  
	strDebug = strDebug & "<li><b>Registration is # " & strEmailVal & "!</b></li>" 
	strDebug = strDebug & "<li>--------------------</li>" 
	strDebug = strDebug & "<li>Doing it...</li><ul>" 
	recCt = Request.Form("recCount")
	adApp = Request.Form("approve")
	adDeny = Request.Form("deny")
	
  if strEmail = 0 and (strEmailVal = 5 or strEmailVal = 7 or strEmailVal = 8) and not deBug then
	strResult = strResult & "<p><span class=""fTitle"">" & txtEmlTrnOff & "</span><br />"
	strResult = strResult & txtActPngErr2 & "</p>"	
  else
	if adApp = 1 then
	  doAllApp()
	elseif adDeny = 1 then
	  doAllDeny()
	else
	  for ex = 1 to recCt
		act = Request.Form("action" & ex)
		actID = Request.Form("id" & ex)
		select case act
		  case 1 'approve
			if strEmailVal = 7 or strEmailVal = 8 then
			  sendApproval(actID)
			end if
		  	doApproval(actID)
		  case 2 'deny
			if strEmailVal = 7 or strEmailVal = 8 then
			  sendDeny(actID)
			end if
		    doDenial(actID)
		  case 3 'resend key
		  	doResend(actID)
		  case else
		end select
	  next
	end if
  end if
	strDebug = strDebug & "</ul><li>Finished Doing it...</li>"
end if

mypage = request.querystring("whichpage")

if mypage = "" then
	mypage = 1
end if

' - Find all records with the search criteria in them
strSql = "SELECT M_NAME, M_EMAIL, MEMBER_ID, M_DATE, M_KEY, M_LEVEL "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE M_LEVEL = -1"
strSql = strSql & " ORDER BY MEMBER_ID ASC;"
'strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS"

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		OffSet = CInt((mypage - 1) * strPageSize)
		strSql = strSql & " LIMIT " & OffSet & ", " & strPageSize & " "
	end if

	' - Get the total pagecount
	strSql2 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING " 
	strSql2 = strSql2 & " WHERE M_LEVEL = -1" 

	set rsCount = my_Conn.Execute(strSql2)
	if not rsCount.eof then
		maxpages = (rsCount("PAGECOUNT")/strPageSize )
		if rsCount("PAGECOUNT") mod strPageSize <> 0 then
			maxpages = maxpages + 1
		end if
		maxRecs = cint(strPageSize) * maxPages
	else
		maxpages = 0
	end if 

	rsCount.close
	set rsCount = nothing

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open  strSql, my_Conn, 3

	if not (rs.EOF or rs.BOF) then
		rs.movefirst
	end if

else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.RecordSet")
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3
	
	if not (rs.EOF or rs.BOF) then  '## Members found in DB
		rs.movefirst
		rs.pagesize = strPageSize
		rs.cacheSize = strPageSize
		maxPages = cint(rs.pageCount)
		maxRecs = cint(rs.pageSize)
		rs.absolutePage = myPage
		maxpages = cint(rs.pagecount)
	end if
end if
%>
<% 	if deBug then
	  response.Write("<p><ul><b>Debugger...</b>" & strDebug & "</ul></p>")
	end if %>

<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	menu_admin()
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtMembPend & "|admin_accounts_pending.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<% spThemeBlock1_open(intSkin) %>
<%  if maxpages > 1 then %>
		<table border="0" align="right">
			<tr><% Call DropDownPaging(1) %></tr>
		</table>
<%  end if %>
<% if mypage = 1 then %>
	<table align="center" width="100%" cellspacing="0" cellpadding="4" border="0">
	  <tr>
    	<td class="tTitle" align="center"><b><%= txtMembPendMgr %></b>
		</td>
	  </tr>
	</table><br />
<% end if %>
<% If strResult <> "" Then %>
<%= "<center>" & strResult & "</center><br />" %>
<% End If %>
<%if mypage = 1 then
strSqlR = "SELECT M_NAME, M_EMAIL, MEMBER_ID, M_DATE, M_KEY "
strSqlR = strSqlR & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
strSqlR = strSqlR & " WHERE M_LEVEL = " & -7
strSqlR = strSqlR & " ORDER BY M_NAME ASC;"

set rsr = Server.CreateObject("ADODB.Recordset")
rsr.open  strSqlR, my_Conn, 3
%>
<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr>
  	<td><b><%= txtLstResNams %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt2">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td class="tSubTitle" align="center"><b><%= txtID %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtUsrNam %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtEmlAdd %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtRegistered %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtAction %></b></td>
      </tr>
<%if rsr.EOF or rsr.BOF then  '## No members found in DB %>
      <tr>
        <td class="tCellAlt1" colspan="5"><b><%= txtNoMemFnd %></b></td>
      </tr>
<%else 
do until rsr.EOF
%>
      <tr>
<td class="tCellAlt1" align="center"><% =rsr("MEMBER_ID") %></td>
<td class="tCellAlt1" align="center"><% =rsr("M_NAME") %></td>
<td class="tCellAlt1" align="center"><% =rsr("M_EMAIL") %></td>
<td class="tCellAlt1" align="center"><% =ChkDate2(rsr("M_DATE")) %></td>
<td class="tCellAlt1" align="center"><a href="register.asp?actkey=<% =rsr("M_KEY") %>"><%= txtRelUsrNam %></a></td>
      </tr>

<%
rsr.MoveNext
loop%>
<tr>
<td colspan="5" class="tCellAlt1" align="center"><%= txtRelPendStr1 %></td>
</tr>
<%end if%>      
    </table>
    </td></tr></table><br />
<%
rsr.close
set rsr = nothing

  if strEmailVal = 8 then %>
    <form name="valForm" id="form2" action="admin_accounts_pending.asp" method="post"><%
	showValidatedMembers()
	%></form><%
	'rs.movefirst
  end if

end if
%>

<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr>
  	<td><b><%= txtPendMem2 %></b></td>
  </tr>
  <tr>
    <td class="tCellAlt2">
	  <form name="penForm" id="form1" action="admin_accounts_pending.asp" method="post">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td class="tSubTitle" align="center"><b><%= txtUsrNam %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtEmlAdd %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtRegistered %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtDaysSnce %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtAction %></b></td>
      </tr>
<%if rs.EOF or rs.BOF then  '## No members found in DB %>
      <tr>
        <td class="tCellAlt1" colspan="5"><b><%= txtNoMemFnd %></b></td>
      </tr>
<%else %>
<%	intI = 0
	howManyRecs = 0
	rec = 1
	do until rs.EOF or rec = (strPageSize + 1)
	  if rs("M_LEVEL") = -1 then
		days = DateDiff("d",  ChkDate2(rs("M_DATE")),  strCurDateAdjust)
		if days >= 15 then
			days2 = "<b>" & days & "</b>"
		else
			days2 = days
		end if
%>
      <tr>
        <td class="tCellAlt1" align="center"><% =rs("M_NAME") %></td>
        <td class="tCellAlt1" align="center"><% =rs("M_EMAIL") %></td>
		<td class="tCellAlt1" align="center"><% =ChkDate2(rs("M_DATE")) %></td>
		<td class="tCellAlt1" align="center"><span<% if days >= 7 then Response.Write(" class=""fAlert""") end if %>><% =days2 %></span></td>
        <td class="tCellAlt1" align="center">
  <select name="action<%= rec %>">
    <option value="0" selected="selected">--[<%= txtSelect %>]--</option>
	<% If not strEmailVal = 8 Then %>
    <option value="1"><%= txtApprove %></option>
	<% End If %>
    <option value="2"><%= txtDeny %></option>
	<% if strEmailVal = 5 or  strEmailVal = 6 or  strEmailVal = 7 or  strEmailVal = 8 then %>
    <option value="3"><%= txtRsndKey %></option>
	<% End If %>
  </select>
  <input type="hidden" name="id<%= rec %>" value="<% =rs("MEMBER_ID") %>">
		</td>
      </tr>
<%		rs.MoveNext
		intI = intI + 1
		if intI = 2 then
			intI = 0
		end if
		rec = rec + 1
	  else
	    rs.MoveNext
	  end if
	loop %>
      <tr>
        <td class="tCellAlt1" colspan="5">
		  <table border="0" class="tCellAlt1" width="100%" cellpadding="0" cellspacing="0">
		    <tr><td><b>
			<% If not strEmailVal = 8 Then %>&nbsp;&nbsp;&nbsp;<%= txtApprvAll %>:&nbsp;<input name="approve" type="checkbox" id="approve" value="1" onclick="javascript:cbValidate('deny','penForm')">
			&nbsp;&nbsp;&nbsp;&nbsp;<%= txtDelDenAll %>:&nbsp;<input name="deny" type="checkbox" id="deny" value="1" onclick="javascript:cbValidate('approve','penForm')"><% Else %>&nbsp;&nbsp;&nbsp;&nbsp;<%= txtDelDenAll %>:&nbsp;<input name="deny" type="checkbox" id="deny" value="1"><% End If %></b>
			</td><td align="right">
			<input type="hidden" name="cat" id="cat" value="1">
			<input type="hidden" name="recCount" value="<%= rec %>">
			<input type="hidden" name="dbug" value="<%= chkstring(request("dog"),display) %>">			
			<input class="button" name="submit" type="submit" id="submit" value="<%= txtSubmit %>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</td></tr>
		  </table>
		</td>
      </tr>
<% end if %>
    </table>
	  </form>
<% if maxpages > 1 then %>
<table border="0">
	<tr>
<% 	Call DropDownPaging(2) %>
	</tr>
</table><br />
<% end if %>
    </td>
  </tr>
</table><br />

<%
rs.Close
set rs = nothing
spThemeBlock1_close(intSkin) %>
</td></tr></table>
<!--#include file="inc_footer.asp" -->
<%Else
	Response.Redirect "admin_login.asp?target=admin_accounts_pending.asp"
End IF

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub DropDownPaging(fnm)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write	"<td nowrap=""nowrap"">" & vbNewLine
		Response.write "<form id=""PageNum" & fnm & """ name=""PageNum" & fnm & """ action=""admin_accounts_pending.asp"" method=""get"">" & vbNewLine
		if fnm = 1 then
			Response.Write("<b>" & txtPage & ": </b><select name=""whichpage"" onchange=""ChangePage(" & fnm & ");"">" & vbNewLine)
		else
		  Response.Write("<b>" & replace(txtPendMem3,"[%count%]",maxpages) & ": </b>")
		  Response.Write("<select id=""whichpage"" name=""whichpage"" onchange=""submit();"">" & vbNewLine)
		end if
		for counter = 1 to maxpages
			if counter <> cint(pge) then   
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			else
				Response.Write "<option value=""" & counter &  """ selected=""selected"">" & counter & "</option>" & vbNewLine
			end if
		next
		if fnm = 1 then
			Response.Write("</select><b> " & txtOf & " " & maxPages & "</b>" & vbNewLine)
		else
			Response.Write("</select>" & vbNewLine)
		end if
		Response.Write("</form>" & vbNewLine)
		Response.Write("</td>" & vbNewLine)
	end if
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub showValidatedMembers()
strSql = "SELECT M_NAME, M_EMAIL, MEMBER_ID, M_DATE, M_KEY "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
strSql = strSql & " WHERE M_LEVEL = -2"
strSql = strSql & " ORDER BY MEMBER_ID;"
set rsVPM = my_Conn.execute(strSql)
 %>
<table border="0" width="750" cellspacing="0" cellpadding="0" align="center">
  <tr>
  	<td><b><%= txtPendMem4 %>:</b></td>
  </tr>
  <tr>
    <td class="tCellAlt2">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td class="tSubTitle" align="center"><b><%= txtUsrNam %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtEmlAdd %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtRegistered %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtDaysSnce %></b></td>
        <td class="tSubTitle" align="center"><b><%= txtAction %></b></td>
      </tr>
<%if rsVPM.EOF then  '## No members found in DB %>
      <tr>
        <td class="tCellAlt1" colspan="5"><b><%= txtNoMemFnd %></b></td>
      </tr>
<%else %>
<%	'intI = 0
	'howManyRecs = 0
	rec = 0
	do until rsVPM.EOF
	  'if rsVPM("M_LEVEL") = -2 then
		rec = rec + 1
		days = DateDiff("d",  ChkDate2(rsVPM("M_DATE")),  strCurDateAdjust)
		if days >= 15 then
			days2 = "<b>" & days & "</b>"
		else
			days2 = days
		end if
%>
      <tr>
        <td class="tCellAlt1" align="center"><% =rsVPM("M_NAME") %></td>
        <td class="tCellAlt1" align="center"><% =rsVPM("M_EMAIL") %></td>
		<td class="tCellAlt1" align="center"><% =ChkDate(rsVPM("M_DATE")) %></td>
		<td class="tCellAlt1" align="center"><span <% if days >= 7 then Response.Write("class=""fAlert""") end if %>><% =days2 %></span></td>
        <td class="tCellAlt1" align="center">
  		<select name="action<%= rec %>">
    		<option value="0" selected="selected">--[<%= txtSelect %>]--</option>
    		<option value="1"> <%= txtApprove %></option>
    		<option value="2"> <%= txtDeny %></option>
  		</select>
  		<input type="hidden" name="id<%= rec %>" value="<% =rsVPM("MEMBER_ID") %>">
		</td>
      </tr>
<%		
	  'end if
		rsVPM.MoveNext
	loop %>
<% end if %>
<% If rec = 0 Then %>
<!--       <tr>
        <td class="tCellAlt1" colspan="5"><b>No Members Found</b></td>
      </tr> -->
<% Else %>
      <tr>
        <td class="tCellAlt1" colspan="5">
		  <table border="0" class="tCellAlt1" width="100%" cellpadding="0" cellspacing="0">
		    <tr><td><b>
			&nbsp;&nbsp;&nbsp;<%= txtDelDenAll %>:&nbsp;<input name="deny" type="checkbox" id="deny" value="1" onclick="javascript:cbValidate('approve','valForm')">
			&nbsp;&nbsp;&nbsp;&nbsp;<%= txtApprvAll %>:&nbsp;<input name="approve" type="checkbox" id="approve" value="1" onclick="javascript:cbValidate('deny','valForm')"></b>
			</td><td align="right">
			<input type="hidden" name="cat" id="cat" value="2">
			<input type="hidden" name="recCount" value="<%= rec %>">			
			<input class="button" name="submit" type="submit" id="submit" value="<%= txtSubmit %>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			</td></tr>
		  </table>
		</td>
      </tr>
<% End If %>
    </table>
    </td>
  </tr>
</table><br />
<%
set rsVPM = nothing
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub doDenial(did)
	set rsKey = nothing
	strDebug = strDebug & "<li>Delete member... id:"&did&"</li><ul>"
	if strEmail = 0 and (strEmailVal = 7 or strEmailVal = 8) then
	  'cant delete member because email is neened to notify them
	  strDebug = strDebug & "<li>Member NOT Deleted... id:"&did&"</li></ul>"
	  strResult = strResult & "<p align=center>" & txtPendErr2 & " <b> " & did & "</b><br />" & txtPendErr3 & "</p>"
	else	  
	  strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
	  strSql = strSql & " WHERE MEMBER_ID = " & did
	  'response.Write(strSql)
	  if not deBug then
	    executeThis(strSql)
	  end if
	  strDebug = strDebug & "<li>Member Deleted... id:"&did&"</li></ul>"
	  if not (strEmailVal = 7 or strEmailVal = 8) and act = 2 then
	    strResult = strResult & "<b>" & txtPendMem5 & "</b><br />"
	  end if
	end if	
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub doAllDeny()
	Dim strMem
	strMem = ""
  	strDebug = strDebug & "<li>Start 'Deny All' in '" & catID & "'!</li><ul>"
	if strEmailVal = 7 or strEmailVal = 8 then 'send denial emails
	  strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
	  strSql = strSql & " WHERE M_LEVEL=-" & catID
	    set rsD = my_Conn.execute(strSql)
	   if not rsD.eof then
	    do until rsD.eof
		  sendDeny(rsD("MEMBER_ID"))
		  strMem = strMem & rsD("MEMBER_ID") & ":"
		  rsD.movenext
		loop
	    set rsD = nothing
	    ids = split(strMem,":")
	    for xl = 0 to ubound(ids)-1
	      doDenial(ids(xl))
	    next
	   else
	    strDebug = strDebug & "<li>No members found to Deny/Delete</li>"
	  strResult = strResult & "<p align=center><b>" & txtPendErr4 & "</b></p>"
	    set rsD = nothing
	   end if
	else 'no emails need sent, just delete them
	  rec = request.Form("recCount")
	    'response.Write("rec: " & rec & "<br />")
      strDebug = strDebug & "No emails, All Denied in '" & catID & "'!<br />"
	  for x = 1 to rec - 1
	    sMem = request.Form("id" & x)
	    strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
	    strSql = strSql & " WHERE MEMBER_ID = " & sMem & ""
	    'response.Write("id" & x & "<br />")
	    'response.Write("sMem: " & sMem & "<br />")
	    if not deBug then
	      executeThis(strSql)
	    end if
	  next
	  strResult = strResult & "<p align=center><b>" & txtPendMem6 & "</b></p>"
	end if
  	strDebug = strDebug & "</ul><li>Finished 'All Denied' in '" & catID & "'!</li>"
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub doAllApp()
	Dim strMem, ids, xl, strSql
	strMem = ""
  	strDebug = strDebug & "<li>Start 'Approve All' in '" & catID & "'!</li><ul>"
	  strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
	  strSql = strSql & " WHERE M_LEVEL=-" & catID
	  set rsA = my_Conn.execute(strSql)
	  if not rsA.eof then
	    do until rsA.eof
		  if strEmailVal = 7 or strEmailVal = 8 then 'send approval emails
		    sendApproval(rsA("MEMBER_ID"))
		  end if
		  strMem = strMem & rsA("MEMBER_ID") & ":"
		  rsA.movenext
		loop
	    set rsA = nothing
	    ids = split(strMem,":")
	    for xl = 0 to ubound(ids)-1
	      doApproval(ids(xl))
	    next
	  else
	    strDebug = strDebug & "<li>No Approval members found</li>"
	    set rsA = nothing
	  end if
  	strDebug = strDebug & "</ul><li>Finished 'Approve All' in '" & catID & "'!</li>"
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	'::::::::::::::::: request email of lost key ::::::::::::::::::::::::::
sub doResend(rid)
  	strDebug = strDebug & "<li>Resending Validation key... " & rid & "</li><ul>"
	  strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE MEMBER_ID = " & rid
	  set rsReqM = my_Conn.Execute (strSql)
	  if rsReqM.EOF or rsReqM.BOF then
	    strDebug = strDebug & "<li>Member ID not found in MEMBERS_PENDING</li>"
	  else
	    strDebug = strDebug & "<li>Member found... "& rsReqM("M_NAME") &"</li>"
		if lcase(strEmail) = "1" then
			strRecipientsName = rsReqM("M_NAME")
			strRecipients = rsReqM("M_EMAIL")
			strFrom = strSender
			strFromName = strSiteTitle
			strSubject = strSiteTitle & " " & txtPendMem7 & " "
			strMessage = strRecipientsName & vbCrLf & vbCrLf
			strMessage = strMessage & replace(replace(txtEmlVal2,"[%sitetitle%]",strSiteTitle),"[%siteurl%]",strHomeURL) & vbCrLf & vbCrLf
			if strAuthType="db" then
				if strEmailVal = 5 or  strEmailVal = 6 or  strEmailVal = 7 or  strEmailVal = 8 then
					strMessage = strMessage & txtEmlVal3 & vbNewline
					strMessage = strMessage & strHomeURL & "register.asp?actkey=" & rsReqM("M_KEY") & vbNewline & vbNewline
				end if
			end if
			strMessage = strMessage & txtEmlVal4 & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal5
			if not deBug then
				sendOutEmail strRecipients,strSubject,strMessage,2,0
			end if
  			strDebug = strDebug & "<li>Validation key sent to... " & strRecipientsName & "</li>"
			strResult = strResult & txtPendMem8 & " '<b>" & strRecipientsName & "</b>'<br />"
		else
			strResult = strResult & "<p align=center>" & txtPendMem9 & "</p>"
  			strDebug = strDebug & "<li>Validation key <b>NOT</b> sent.</li>"
			strDebug = strDebug & "<li><b>Email is turned OFF.</b></li>"
  	    end if		
	  end if  'rsReqM.EOF or rsReqM.BOF
'	  rsReqM.close
	  set rsReqM = nothing
  	  strDebug = strDebug & "</ul><li>Finished 'Resend Validation Key'.</li>"
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub sendApproval(id)
  strDebug = strDebug & "<li>Sending Approval Email... id: "&id&"</li><ul>"
  strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE MEMBER_ID=" & id
  set rsS = my_Conn.execute(strSql)
	  if rsS.EOF or rsS.BOF then
	    strDebug = strDebug & "<li>Member ID not found in MEMBERS_PENDING</li>"
		strResult = strResult & "<p align=center>" & txtPendMem10 & ": " & rsS("M_NAME") & "</p>"
	  else
	    strDebug = strDebug & "<li>Member found... "& rsS("M_NAME") &"</li>"
		if strEmail = 1 then
			strRecipientsName = rsS("M_NAME")
			strRecipients = rsS("M_EMAIL")
			strFrom = strSender
			strFromName = strSiteTitle
			strsubject = strSiteTitle & " Membership"
			strMessage = strRecipientsName & vbCrLf & vbCrLf
			strMessage = strMessage & replace(txtEmlVal18,"[%sitetitle%]",strSiteTitle) & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal4 & vbCrLf & vbCrLf
			strMessage = strMessage & txtEmlVal5 & vbCrLf & vbCrLf
			if not deBug then
				sendOutEmail strRecipients,strSubject,strMessage,2,0
			end if
  			strDebug = strDebug & "<li>Approval Email sent to... " & strRecipientsName & "</li>"
			strResult = strResult & txtPendMem11 & " <b>" & strRecipientsName & "</b>'<br />"
		else
			strResult = strResult & "<p align=center>" & txtPendMem12 & "</p>"
  			strDebug = strDebug & "<li>Approval Email NOT sent to: " & rsS("M_NAME") & "</li>"
			strDebug = strDebug & "<li><b>Email is turned OFF.</b></li>"
  	    end if		
	  end if  'rsS.EOF or rsS.BOF
  set rsS = nothing
  strDebug = strDebug & "</ul>"
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub sendDeny(id)
  strDebug = strDebug & "<li>Sending denial email... id: "&id&"</li><ul>"
  strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE MEMBER_ID=" & id
  set rsS = my_Conn.execute(strSql)
	  if rsS.EOF or rsS.BOF then
	    strDebug = strDebug & "<li>Member ID not found in MEMBERS_PENDING</li>"
		strResult = strResult & "<p align=center>" & txtPendMem10 & ": " & rsS("M_NAME") & "</p>"
	  else
	    strDebug = strDebug & "<li>Member found... "& rsS("M_NAME") &"</li>"
		if strEmail = 1 then
			strRecipientsName = rsS("M_NAME")
			strRecipients = rsS("M_EMAIL")
			strFrom = strSender
			strFromName = strSiteTitle
			strsubject = strSiteTitle & " " & txtMbrshp
			strMessage = strRecipientsName & vbCrLf & vbCrLf
			strMessage = strMessage & replace(txtPendMem13,"[%sitetitle%]",strSiteTitle) & vbCrLf & vbCrLf
			strMessage = strMessage & txtPendMem14
			if not deBug then
			    sendOutEmail strRecipients,strSubject,strMessage,2,0
			end if
  			strDebug = strDebug & "<li>Denial Email sent to... " & strRecipientsName & "</li>"
			strResult = strResult & txtPendMem15 & " '<b>" & strRecipientsName & "</b>'<br />"
		else
			strResult = strResult & "<p align=center>" & txtPendMem16 & "</p>"
  			strDebug = strDebug & "<li>Denial Email <b>NOT</b> sent to " & rsS("M_NAME") & "</li>"
			strDebug = strDebug & "<li><b>Email is turned OFF.</b></li>"
  	    end if		
	  end if  'rsS.EOF or rsS.BOF
  set rsS = nothing
  strDebug = strDebug & "</ul>"
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub doApproval(aid)
  strDebug = strDebug & "<li>Approving Member... " & aid & "</li>"
  strDebug = strDebug & "<li>Moving to MEMBER table...</li>"
	
	strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE MEMBER_ID = " & aid
	  'Response.Write "<br>" & strSql & "<br>"
	set rsKey = my_Conn.Execute(strSql)
	
	if not rsKey.eof then
  		strDebug = strDebug & "<li>Member found: " & rsKey("M_NAME") & "</li>"
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
		strSql = strSql & ", " & "'" & rsKey("M_SIG") & "'"
		strSql = strSql & ", " & "'" & rsKey("M_YAHOO") & "'"
		strSql = strSql & ", " & "'" & rsKey("M_ICQ") & "'"
		strSql = strSql & ", " & "'" & rsKey("M_AIM") & "'"
		strSql = strSql & ", " & "0"
		strSql = strSql & ", " & "'" & rsKey("M_HOMEPAGE") & "'"
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
		'response.Write(strSql & "<br />aid: " & aid)
		if not deBug then
		  executeThis(strSql)
		end if
		'Response.Write "Member moved<br>"
  		strDebug = strDebug & "<li>Member moved: " & rsKey("M_NAME") & "</li></ul>"
		if not debug then
		  sendPMtoNewUser(rsKey("M_NAME"))
		end if
		'Response.Write "pm sent<br>"
  		strDebug = strDebug & "<li>Welcome PM attempted</li>"
		if not debug then
		  DoCount
		end if
		'Response.Write "count updated<br>"
  		strDebug = strDebug & "<li>Member count updated</li>"
		
  		doDenial(aid)
		'Response.Write "Member removed from pending_members<br>"
  
  
  		if not (strEmailVal = 7 or strEmailVal = 8) then
  			strResult = strResult & "<p align=center><b>Member Approved!</b></p>"
  		end if
  		strDebug = strDebug & "<li>" & txtPendMem17 & "</li></ul>"
	else
  		strResult = strResult & "<p align=center><b>" & txtMemNoFnd & "!</b></p>"
  		strDebug = strDebug & "<ul><li>Member not found... id:" & aid & "</li></ul>"
	end if
	set rsKey = nothing
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
sub DoCount
	'## Updates the member count by 1
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	if strDBType = "access" then
	  strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT = U_COUNT + 1"
	else
	  strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT = " & strTablePrefix & "TOTALS.U_COUNT + 1"
	end if
	executeThis(strSql)
end sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
%>