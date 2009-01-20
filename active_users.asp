<!--#include file="config.asp" -->
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
CurPageType = "core"
%>
<!--#include file="inc_functions.asp" -->
<%
PageTitle = txtActvUsrs

CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = txtActvUsrs
	PageAction = txtViewing & "<br />" 
	PageLocation = "active_users.asp" 
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"
end function 
%>
<!--#include file="inc_top.asp" -->
<%
if not hasAccess("1,2") then
  closeAndGo("default.asp")
end if

if request.Cookies(strCookieURL & strUniqueID & "Reload") <> "" then
	nRefreshTime = cLng(request.Cookies(strCookieURL & strUniqueID & "Reload"))
end if

if Request.form("cookie") = "1" then
    Response.Cookies(strCookieURL & strUniqueID & "Reload").Path = strCookieUrl
	Response.Cookies(strCookieURL & strUniqueID & "Reload") = chkString(request.Form("RefreshTime"),"sqlstring")
	Response.Cookies(strCookieURL & strUniqueID & "Reload").expires = strCurDateAdjust + 365
	nRefreshTime = chkString(request.Form("RefreshTime"),"sqlstring")
end if

if nRefreshTime = "" then
	nRefreshTime = 0
end if

'07-01-05 JayMonster
'Added To Allow Selection of Specific Views of who is online
If Request.Item("SelectedViewType") <> "" then
Response.Cookies("ViewTypeSelected") = Request.Item("SelectedViewType")
ViewType = Request.Item("SelectedViewType")
else
	If Request.Cookies("ViewTypeSelected") <> "" then
		ViewType = Request.Cookies("ViewTypeSelected")
	end if
end if

ActiveSince = Request.Cookies(strCookieURL & strUniqueID & "ActiveSince")
mypage = trim(chkString(request("whichpage"),"sqlstring"))

	If  mypage = "" then
	   mypage = 1
	end if

	mypagesize = 15

	If trim(mypagesize) = "" then
	   mypagesize = 15
	end if
%>
<script type="text/javascript">
<!--
function autoReload(){
	document.ReloadFrm.submit()
}
function SetLastDate(){
	document.LastDateFrm.submit()
}
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td class="leftPgCol" align="center" valign="top">
	<%
	intSkin = getSkin(intSubSkin,1)
	menu_fp() %>
	</td>
	<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtActvUsrs & " (" & txtLstUpdd & "&nbsp;" & strCurDateAdjust & ")|active_users.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="50%" align="center">
    <%
    '07-01-05 JayMonster
	'Added To Allow Selection of Specific Views of who is online
	'Form Allows Users to Selected Between
	' 0 - (Default) Members Only
	' 1 - Guests Only
	' 2 - All 
	%>
    <form id="frmSelectViewType" name="frmSelectViewType" action="active_users.asp" method="post"><br />
    <select name="SelectedViewType" size="1" onchange="document.frmSelectViewType.submit();">
		<option value="0" <% if ViewType="0" then Response.Write(" selected=""selected""")%>><%= txtViewMem %> Only</option>
		<option value="1" <% if ViewType="1" then Response.Write(" selected=""selected""")%>><%= txtViewGst %> Only</option>
		<option value="2" <% if ViewType="2" then Response.Write(" selected=""selected""")%>><%= txtViewAll %></option>
	</select>
	<input type="hidden" name="StoreView" value="1" />
	</form>                         
    </td>
    <td width="50%" align="center">
    <form id="ReloadFrm" name="ReloadFrm" action="active_users.asp" method="post"><br />
    <select name="RefreshTime" size="1" onchange="autoReload();">
        <option value="0" <% if nRefreshTime = "0" then Response.Write(" selected=""selected""")%>><%= txtNoRld %></option>
        <option value="1" <% if nRefreshTime = "1" then Response.Write(" selected=""selected""")%>><%= txtRld1m %></option>
        <option value="5" <% if nRefreshTime = "5" then Response.Write(" selected=""selected""")%>><%= txtRld5m %></option>
        <option value="10" <% if nRefreshTime = "10" then Response.Write(" selected=""selected""")%>><%= txtRld10m %></option>
        <option value="15" <% if nRefreshTime = "15" then Response.Write(" selected=""selected""")%>><%= txtRld15m %></option>
        <option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>><%= txtRld30m %></option>
    </select>
    <input type="hidden" name="Cookie" value="1" />
    </form>
<script type="text/javascript">
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
  </script>
    </td>
  </tr>
</table><br />
<% If strdbntusername= "SkyDoggx" Then %>
<table width="400" border="0" cellspacing="5" cellpadding="2">
  <tr>
    <td height="50" class="tCellAlt0">&nbsp;tCellAlt0</td>
    <td height="50" class="tCellAlt1">&nbsp;tCellAlt1</td>
    <td height="50" class="tCellAlt2">&nbsp;tCellAlt2</td>
    <td height="50" class="tCellHover">&nbsp;tCellHover</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<% End If %>
<%
spThemeTitle = txtActvUsrs
spThemeBlock1_open(intSkin)
%>
<table class="grid" cellpadding="1" cellspacing="0" border="0" width="100%">
  <tr>
    <td class="tSubTitle" align="center" valign="top" nowrap="nowrap"><b><%= txtUsrName %></b></td>
<%	if hasAccess(1) then %>
    <td class="tSubTitle" align="center" valign="top" nowrap="nowrap"><b><%= txtIPaddy %></b></td>
<%	end if %>
	<td class="tSubTitle" align="center" valign="top" nowrap="nowrap"><b><%= txtLstAct %></b></td>
    <td class="tSubTitle" align="center" valign="top" nowrap="nowrap"><b><%= txtOnlTime %></b></td>
</tr>
<%
	set rs = Server.CreateObject("ADODB.Recordset")
	'
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.UserIP, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.LastChecked, " & strTablePrefix & "ONLINE.CheckedIn, " & strTablePrefix & "ONLINE.UserAgent "
	strSql = strSql & "FROM " & strMemberTablePrefix & "ONLINE "
	'Default value or first selected value is "Members Only" so we exclude
	'Guest enteries, if the selection is Guests Only then we exclude members
	'If ALL is selected then we show everything.
	if ViewType = "2" then
		'we do nothing if we want "everything"
	elseif ViewType = "1" then
		strSql = strSql & " where UserID='" & txtGuest & "'"
	else
		'The Default - Members Only View
		strSql = strSql & " where not UserID='" & txtGuest & "' "
	end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.CheckedIn DESC"

	rs.cachesize = 20
	'response.Write(strsql)
	rs.open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

	i = 0

	If rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write ""
	Else
	
		rs.movefirst
		num = 0
		rbt= 0
		rs.pagesize = mypagesize
		maxpages = cint(rs.pagecount)
		maxrecs = cint(rs.pagesize)
		rs.absolutepage = mypage
		howmanyrecs = 0
		rec = 1
		do until rs.EOF or (rec = mypagesize+1)
			if strI = 0 then
				CColor = "tCellAlt2"
			else
				CColor = "tCellAlt0"
			end if

  			strTheUserID = rs("UserID")
  			strTheUserID = OnlineSQLdecode(strTheUserID)

  			if Right(rs("UserID"), 5) <> txtGuest then
				strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_GLOW, " & strTablePrefix & "ONLINE.UserID "
				strSql = strSql & " FROM " & strTablePrefix & "MEMBERS, " & strTablePrefix & "ONLINE "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & rs("UserID") & "' "
	set rsMember =  my_Conn.Execute (strSql)
			end if
if Right(rs("UserID"), 5) = txtGuest then
num = num + 1
   '***********************>>  START ROBOT MOD  <<*********************************
   tmpRobot = lcase(left(rs("UserAgent"),3))
   select case tmpRobot
      case "sun"
           tmpRobot =txtGuest & " #"& num 
     case "sol"           
           tmpRobot =txtGuest & " #"& num
     case "nok"          
           tmpRobot =txtGuest & " #"& num
      case "fre"         
           tmpRobot =txtGuest & " #"& num 
     case "pal"          
           tmpRobot =txtGuest & " #"& num
      case "bsp"          
           tmpRobot =txtGuest & " #"& num
      case "ums"         
           tmpRobot =txtGuest & " #"& num
      case "os/"          
           tmpRobot =txtGuest & " #"& num
       case "dec"        
           tmpRobot =txtGuest & " #"& num
       case "ami"        
           tmpRobot =txtGuest & " #"& num
       case "lin"         
           tmpRobot =txtGuest & " #"& num
       case "win"         
           tmpRobot =txtGuest & " #"& num
       case "ava"         
           tmpRobot =txtGuest & " #"& num
       case "moz"        
	   	   if left(rs("UserIP"),6) = "72.30." then
             rbt = rbt +1
             tmpRobot = txtRobot & "&nbsp;#" & rbt & ": Inktomi"
		   elseif left(rs("UserIP"),7) = "68.142." then
		     tIP = split(rs("UserIP"),".")
			 if cLng(tIP(2)) >= 192 then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": Inktomi"
			 end if
		   elseif left(rs("UserIP"),5) = "74.6." then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": Inktomi"
		   elseif left(rs("UserIP"),6) = "66.94." then
		     tIP = split(rs("UserIP"),".")
			 if cLng(tIP(2)) >= 224 then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": YaHoo!"
			 end if
		   elseif left(rs("UserIP"),7) = "67.195." then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": YaHoo!"
		   elseif left(rs("UserIP"),7) = "66.249." then
		     tIP = split(rs("UserIP"),".")
			 if cLng(tIP(2)) >= 64 and cLng(tIP(2)) <= 95 then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": Google"
			 end if
		   elseif left(rs("UserIP"),6) = "72.14." then
		     tIP = split(rs("UserIP"),".")
			 if cLng(tIP(2)) >= 192 and cLng(tIP(2)) <= 255 then
               rbt = rbt +1
               tmpRobot = txtRobot & "&nbsp;#" & rbt & ": Google"
			 end if
		   else
             tmpRobot =txtGuest & " #"& num
		   end if
       case "net"        
           tmpRobot =txtGuest & " #"& num
       case "int"        
           tmpRobot =txtGuest & " #"& num
       case "ope"         
           tmpRobot =txtGuest & " #"& num
       case "mac"         
           tmpRobot =txtGuest & " #"& num
     case Else
           rbt = rbt +1
             pos=Instr(rs("UserAgent")," ")
             tmpRobot = txtRobot & "&nbsp;#" & rbt & ": " & rs("UserAgent")
    end select
end if

strOnlineDateCheckedIn = ChkDate2(rs("CheckedIn"))
strOnlineDateCheckedIn = strOnlineDateCheckedIn & ChkTime2(rs("CheckedIn"))
strOnlineLastDateChecked = ChkDate2(rs("LastChecked"))
strOnlineLastDateChecked = strOnlineLastDateChecked & ChkTime2(rs("LastChecked"))

strOnlineTotalTime = DateDiff("n",strOnlineDateCheckedIn,strOnlineLastDateChecked)

If strOnlineTotalTime > 60 then
' they must have been online for like an hour or so.
strOnlineHours = 0
do until strOnlineTotalTime < 60
strOnlineTotalTime = (strOnlineTotalTime - 60)
strOnlineHours = strOnlineHours + 1
loop
strOnlineTotalTime = strOnlineHours & " " & txtHours & " " & strOnlineTotalTime & " " & txtMinutes & ""
Else
strOnlineTotalTime = strOnlineTotalTime & " " & txtMinutes & ""
End If

%>
  <tr class="<% =CColor %>" onMouseOver="this.className='tCellHover';" onMouseOut="this.className='<% =CColor %>';">
<%  if Right(rs("UserID"), 5) = txtGuest then %>
    <td valign="middle" align="center" class="fNorm"><%= tmpRobot %></td>
 <% else %>
    <td valign="middle" align="center" class="fNorm">
<%
	'Response.Write("<a href=""cp_main.asp?cmd=8&member="">")
	Response.Write("<a href=""cp_main.asp?cmd=8&amp;member="&rsMember("MEMBER_ID")&""">")
	Response.Write(displayName(rs("UserID"),rsMember("M_GLOW")) & "</a></td>")
	'Response.Write(rs("UserID") & "</a></td>")
%>
 <% end if 
	if hasAccess(1) then %>
	<td valign="middle" align="center" class="fNorm"><a href="http://ws.arin.net/cgi-bin/whois.pl?queryinput=<%=rs("UserIP")%>" target="_blank"><% =rs("UserIP")%></a></td>
<%	end if
	if lcase(rs("UserID"))= strSiteOwner and not lcase(strDBNTUserName) = strSiteOwner then
	LastUserAction = txtInvis
	else
	LastUserAction = replace(replace(rs("M_BROWSE"),"&rsquo","'"),"&amp;#39","'")
	end if
%>
    <td valign="middle" align="center" class="fNorm"><%= LastUserAction %></td>
    <td align="center" valign="middle" width="100" class="fNorm" nowrap="nowrap"><%=strOnlineTotalTime%></td>
  </tr>
  
<%
		    rs.MoveNext
		    strI = strI + 1
		    if strI = 2 then
				strI = 0
			end if
		    rec = rec + 1
		loop
	end if %>
</table>
<% '
spThemeBlock1_close(intSkin)%>
<% if maxpages > 1 then %>
	<table border="0" width="100%">
  	<tr>
    	<td valign="top" align="center" class="fNorm"><b><%= txtPages %>: </b>
    	<% Call Paging() %></td>
  	</tr>
	</table>
<% end if %>
</td></tr></table>
<% rs.close
set rs = nothing %>
<!--#include file="inc_footer.asp" -->
<%
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = chkString(Request.QueryString("whichpage"),"numeric")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""1"" align=""top""><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then
				ref = "<td align=""right"">" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname
				ref = ref & "?whichpage=" & counter
				ref = ref & "&amp;pagesize=" & mypagesize
				if top = "1" then
					ref = ref & "'>"
					ref = ref & "<b>" & counter & "</b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=""right"">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></td>")
			end if
			if counter mod 15 = 0 and counter < maxpages then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub
%>