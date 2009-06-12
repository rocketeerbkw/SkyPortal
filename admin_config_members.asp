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
pgType = "memberConfig"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" and hasAccess(1) Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<%
if Request.Form("Method_Type") = "memberConfig" then 

			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRFULLNAME   = " & Request.Form("strFullName") & " "
			strSql = strSql & ",    C_STRPICTURE    = " & Request.Form("strPicture") & " "
			strSql = strSql & ",    C_STRSEX        = " & Request.Form("strSex") & " "
			strSql = strSql & ",    C_STRCITY       = " & Request.Form("strCity") & " "
			strSql = strSql & ",    C_STRSTATE      = " & Request.Form("strState") & " "
			strSql = strSql & ",    C_STRAGE        = " & Request.Form("strAge") & " "
			strSql = strSql & ",    C_STRCOUNTRY    = " & Request.Form("strCountry") & " "
			strSql = strSql & ",    C_STROCCUPATION = " & Request.Form("strOccupation") & " "
			strSql = strSql & ",    C_STRHOMEPAGE   = " & Request.Form("strHomepage") & " "
			strSql = strSql & ",    C_STRFAVLINKS   = " & Request.Form("strFavLinks") & " "
			strSql = strSql & ",    C_STRICQ        = " & Request.Form("strICQ") & " "
			strSql = strSql & ",    C_STRYAHOO      = " & Request.Form("strYAHOO") & " "
			strSql = strSql & ",    C_STRAIM        = " & Request.Form("strAIM") & " "
			strSql = strSql & ",    C_STRBIO        = " & Request.Form("strBio") & " "
			strSql = strSql & ",    C_STRHOBBIES	= " & Request.Form("strHobbies") & " "
			strSql = strSql & ",    C_STRLNEWS      = " & Request.Form("strLNews") & " "
			strSql = strSql & ",    C_STRQUOTE		= " & Request.Form("strQuote") & " "
			strSql = strSql & ",    C_STRMARSTATUS  = " & Request.Form("strMarStatus") & " "
			strSql = strSql & ",    C_STRRECENTTOPICS = " & Request.Form("strRecentTopics") & " "
			strSql = strSql & ",    C_STRMSN = " & Request.Form("strMSN") & " "
			strSql = strSql & ",     C_STRVAR1 = '" & ChkString(Request.Form("StrVar1"),"sqlstring") & "'"
			strSql = strSql & ",     C_STRVAR2 = '" & ChkString(Request.Form("StrVar2"),"sqlstring") & "'"
			strSql = strSql & ",     C_STRVAR3 = '" & ChkString(Request.Form("StrVar3"),"sqlstring") & "'"
			strSql = strSql & ",     C_STRVAR4 = '" & ChkString(Request.Form("StrVar4"),"sqlstring") & "'"
			strSql = strSql & ",    C_STRZIP      = " & ChkString(Request.Form("strZip"),"") & " "
			
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""

			Session.Contents("memberHome") = "<li><span class=""fSubTitle"">" & txtCMb01 & "</span></li>"
    closeAndGo("admin_config_members.asp")
end if

if Request.Form("Method_Type") = "memberRanks" then 
		Err_Msg = ""
		if Request.Form("strRankAdmin") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb08 & "</li>"
		end if
		if Request.Form("strRankMod") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb09 & "</li>"
		end if
		if Request.Form("strRankLevel0") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb10 & "</li>"
		end if
		if Request.Form("strRankLevel1") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb11 & "</li>"
		end if
		if Request.Form("strRankLevel2") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb12 & "</li>"
		end if
		if Request.Form("strRankLevel3") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb13 & "</li>"
		end if
		if Request.Form("strRankLevel4") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb14 & "</li>"
		end if
		if Request.Form("strRankLevel5") = "" then 
			Err_Msg = Err_Msg & "<li>" & txtEntrValFor & " " & txtCMb15 & "</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel2")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb17 & "</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel3")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb18 & "</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel3")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb19 & "</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb20 & "</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb21 & "</li>"
		end if
		if cint(Request.Form("intRankLevel3")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb22 & "</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb23 & "</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb24 & "</li>"
		end if
		if cint(Request.Form("intRankLevel3")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb25 & "</li>"
		end if
		if cint(Request.Form("intRankLevel4")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>" & txtCMb26 & "</li>"
		end if

		if Err_Msg = "" then

			'
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRSHOWRANK = " & Request.Form("strShowRank") & ""
			strSql = strSql & ",    C_STRRANKADMIN = '" & ChkString(Request.Form("strRankAdmin"),"name") & "'"
			strSql = strSql & ",    C_STRRANKMOD = '" & ChkString(Request.Form("strRankMod"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL0 = '" & ChkString(Request.Form("strRankLevel0"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL1 = '" & ChkString(Request.Form("strRankLevel1"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL2 = '" & ChkString(Request.Form("strRankLevel2"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL3 = '" & ChkString(Request.Form("strRankLevel3"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL4 = '" & ChkString(Request.Form("strRankLevel4"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL5 = '" & ChkString(Request.Form("strRankLevel5"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLORADMIN = '" & ChkString(Request.Form("strRankColorAdmin"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLORMOD = '" & ChkString(Request.Form("strRankColorMod"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR0 = '" & ChkString(Request.Form("strRankColor0"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR1 = '" & ChkString(Request.Form("strRankColor1"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR2 = '" & ChkString(Request.Form("strRankColor2"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR3 = '" & ChkString(Request.Form("strRankColor3"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR4 = '" & ChkString(Request.Form("strRankColor4"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR5 = '" & ChkString(Request.Form("strRankColor5"),"name") & "'"
			strSql = strSql & ",    C_INTRANKLEVEL0 = " & ChkString(Request.Form("intRankLevel0"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL1 = " & ChkString(Request.Form("intRankLevel1"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL2 = " & ChkString(Request.Form("intRankLevel2"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL3 = " & ChkString(Request.Form("intRankLevel3"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL4 = " & ChkString(Request.Form("intRankLevel4"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL5 = " & ChkString(Request.Form("intRankLevel5"),"number") & ""
			strSql = strSql & " WHERE CONFIG_ID = " & 1
		
			executeThis(strSql)
			Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
			Session.Contents("memberHome") = "<li><span class=""fSubTitle"">" & txtCMb27 & "</span></li>"
		else 
			Err_Msg1 = "<li><span class=""fSubTitle"">" & txtThereIsProb & "</span></li>"
			Session.Contents("memberHome") = Err_Msg1 & Err_Msg
		end if	
    closeAndGo("admin_config_members.asp?cmd=1")
end if

if request.querystring("mode")= "deleted" then
  cDays = cint(request.form("cDays"))
  cPosts = cint(request.form("cPosts"))
  cNum = cint(request.form("cNum"))
	strSql = "SELECT MEMBER_ID"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE M_POSTS <= " & cPosts & " AND M_LASTHEREDATE < '" & DateToStr(DateAdd("d", cDays*-1+1 , now())) & "' AND M_LEVEL = 1 AND M_STATUS = 1"

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.open  strSql, my_Conn, 3
  if rs.EOF or rs.BOF then
	Session.Contents("memberHome") = "<li><span class=""fSubTitle"">" & txtNoMemFnd & "</span></li>" 
  else
	do while not rs.EOF
	  Member_ID = rs("MEMBER_ID")
	  strSql = "SELECT COUNT(T_AUTHOR) AS POSTCOUNT "
	  strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	  strSql = strSql & " WHERE T_AUTHOR = " & Member_ID

	  set rs2 = my_Conn.Execute (strSql)
	  if not rs2.eof then
		intPostcount = rs2("POSTCOUNT")
	  else
		intPostcount = 0
	  end if
	  rs2.close

	  strSql = "SELECT COUNT(R_AUTHOR) AS REPLYCOUNT "
	  strSql = strSql & " FROM " & strTablePrefix & "REPLY "
	  strSql = strSql & " WHERE R_AUTHOR = " & Member_ID
	  set rs2 = my_Conn.Execute(strSql)
	  if not rs2.eof then
		intReplycount = rs2("REPLYCOUNT")
	  else
		intReplycount = 0
	  end if
	  rs2.close
	
	  if ((intReplycount + intPostCount) = 0) then
		strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Member_ID
		executeThis(strSql)
	  else
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 0
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_EMAIL = ' '"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 1
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_NAME = 'n/a'"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_COUNTRY = ' '"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_TITLE = 'deleted'"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE = ' '"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_AIM = ' '"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_YAHOO = ' '"
		strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_ICQ = ' '"
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Member_ID
		executeThis(strSql)
	  end if

	  rs.movenext
	loop
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT = " & strTablePrefix & "TOTALS.U_COUNT - " & cNum
	executeThis(strSql)
	Session.Contents("memberHome") = "<li><span class=""fSubTitle"">" & replace(txtCMb28,"[%count%]",cNum) & "</span></li>"

  end if
  set rs = nothing
end if
%>

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
  arg2 = txtCMb29 & "|admin_config_members.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
	<%
	if Session.Contents("memberHome") <> "" then
      call showMsgBlock(1,"<ul>" & Session.Contents("memberHome") & "</ul>")
	  Session.Contents("memberHome") = ""
	end if 
	   spThemeBlock1_open(intSkin)
		memberConfig()
		memberRanks()
		memberCleaning() %>
	<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp?target=admin_config_members.asp" %>
<% End IF

sub memberConfig() %>
	<div id="aa" style="display:<%= aa %>;">
<form action="admin_config_members.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="memberConfig">
<table border="0" cellspacing="0" cellpadding="0" width="400" align="center">
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1" width="100%">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><%= txtCMb30 %></b></td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right" width="50%"><b><%= txtFllName %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strFullName" value="1"<% if strFullName <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strFullName" value="0"<% if strFullName = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtMPic %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strPicture" value="1"<% if strPicture <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strPicture" value="0"<% if strPicture = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Picture')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtShoRecTopics %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strRecentTopics" value="1" <% if strRecentTopics <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strRecentTopics" value="0" <% if strRecentTopics = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RecentTopics')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb31 %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strSex" value="1"<% if strSex <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strSex" value="0"<% if strSex = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
    <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtAge %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strAge" value="1"<% if strAge <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strAge" value="0"<% if strAge = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCity %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strCity" value="1"<% if strCity <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strCity" value="0"<% if strCity = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtState %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strState" value="1"<% if strState <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strState" value="0"<% if strState = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtZipCd %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strZip" value="1"<% if strZip <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strZip" value="0"<% if strZip = "0" then Response.Write(" checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCntry %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strCountry" value="1" <% if (lcase(strCountry) <> "0") then Response.Write("checked")%>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strCountry" value="0" <% if (lcase(strCountry) = "0") then Response.Write("checked")%>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtMSN %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strMSN" value="1" <% if strMSN <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strMSN" value="0" <% if strMSN = "0" then Response.Write("checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtICQ %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strICQ" value="1" <% if strICQ <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strICQ" value="0" <% if strICQ = "0" then Response.Write("checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtYhoIM %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strYAHOO" value="1" <% if strYAHOO <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strYAHOO" value="0" <% if strYAHOO = "0" then Response.Write("checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtAIM %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strAIM" value="1" <% if strAIM <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strAIM" value="0" <% if strAIM = "0" then Response.Write("checked") %>>
            </td>
  </tr>
  
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtOcc %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strOccupation" value="1"<% if strOccupation <> "0" then Response.Write(" checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strOccupation" value="0"<% if strOccupation = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Occupation')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtHmPg %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strHomepage" value="1" <% if strHomepage <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strHomepage" value="0" <% if strHomepage = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#homepages')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtFavLinks %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strFavLinks" value="1" <% if strFavLinks <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strFavLinks" value="0" <% if strFavLinks = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#FavLinks')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtMarStat %>:</b>&nbsp;</td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strMarStatus" value="1" <% if strMarStatus <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strMarStatus" value="0" <% if strMarStatus = "0" then Response.Write("checked") %>>
            </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb32 %>:</b><br />
    <input type="text" class="textbox" maxlength="50" name="StrVar1" size="20" value="<% if StrVar1 <> "" then Response.Write(StrVar1) else Response.Write("Bio") %>">
    </td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strBio" value="1" <% if strBio <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strBio" value="0" <% if strBio = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Vars1to4')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
   <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb33 %>:</b><br />
    <input type="text" class="textbox" maxlength="50" name="StrVar2" size="20" value="<% if StrVar2 <> "" then Response.Write(StrVar2) else Response.Write("Hobbies") %>">
    </td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strHobbies" value="1" <% if strHobbies <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strHobbies" value="0" <% if strHobbies = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Vars1to4')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb34 %>:</b><br />
    <input type="text" class="textbox" maxlength="50" name="StrVar3" size="20" value="<% if StrVar3 <> "" then Response.Write(StrVar3) else Response.Write("Last News") %>">
    </td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strLNews" value="1" <% if strLNews <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strLNews" value="0" <% if strLNews = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Vars1to4')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb35 %>:</b><br />
    <input type="text" class="textbox" maxlength="50" name="StrVar4" size="20" value="<% if StrVar4 <> "" then Response.Write(StrVar4) else Response.Write("Quote") %>">
    </td>
    <td class="fNorm">
    <%= txtOn %>: <input type="radio" class="radio" name="strQuote" value="1" <% if strQuote <> "0" then Response.Write("checked") %>> 
    <%= txtOff %>: <input type="radio" class="radio" name="strQuote" value="0" <% if strQuote = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#Vars1to4')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  
  <tr valign="top">
    <td class="fNorm" colspan="2" align="center"><input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
  </tr>
</table>
    </td>
  </tr>
</table>
</form>
	</div>
<%
end sub

sub memberRanks() %>
	<div id="ab" style="display:<%= ab %>;">
<form action="admin_config_members.asp" method="post" id="formEle" name="Form12">
<input type="hidden" name="Method_Type" value="memberRanks">
<table border="0" width="550" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td class="tCellAlt2">
<table border="0" cellspacing="1" cellpadding="1" width="100%">
  <tr valign="top">
    <td class="tTitle" colspan="2"><b><%= txtCMb02 %></b></td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb03 %>:</b>&nbsp;</td>
    <td class="fNorm">
    <select name="strShowRank">
      <option value="0"<% if strShowRank = "0" then Response.Write(" selected")%>><%= txtNone %></option>
      <option value="1"<% if strShowRank = "1" then Response.Write(" selected")%>><%= txtCMb04 %></option>
      <option value="2"<% if strShowRank = "2" then Response.Write(" selected")%>><%= txtCMb05 %></option>
      <option value="3"<% if strShowRank = "3" then Response.Write(" selected")%>><%= txtCMb06 %></option>
    </select>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#ShowRank')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb08 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankAdmin" size="30" value="<% if strRankAdmin <> " " then Response.Write(strRankAdmin) end if %>">
    <%= icon(icnHelp,txtAdminst,"","","") %></td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColorAdmin value="gold"<% if strRankColorAdmin = "gold" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/gold.gif" border=0>
    <input type=radio name=strRankColorAdmin value="silver"<% if strRankColorAdmin = "silver" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/silver.gif" border=0>
    <input type=radio name=strRankColorAdmin value="bronze"<% if strRankColorAdmin = "bronze" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/bronze.gif" border=0>
    <input type=radio name=strRankColorAdmin value="orange"<% if strRankColorAdmin = "orange" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/orange.gif" border=0>
    <input type=radio name=strRankColorAdmin value="red"<% if strRankColorAdmin = "red" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/red.gif" border=0>
    <input type=radio name=strRankColorAdmin value="purple"<% if strRankColorAdmin = "purple" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/purple.gif" border=0>
    <input type=radio name=strRankColorAdmin value="blue"<% if strRankColorAdmin = "blue" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/blue.gif" border=0>
    <input type=radio name=strRankColorAdmin value="cyan"<% if strRankColorAdmin = "cyan" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/cyan.gif" border=0>
    <input type=radio name=strRankColorAdmin value="green"<% if strRankColorAdmin = "green" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/green.gif" border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb09 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankMod" size="30" value="<% if strRankMod <> " " then Response.Write(strRankMod) end if %>">
    <img src="<%= icnHelp %>" border="0" alt="(<%= txtModerator %>)" />
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColorMod value="gold"<% if strRankColorMod = "gold" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/gold.gif" border=0>
    <input type=radio name=strRankColorMod value="silver"<% if strRankColorMod = "silver" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/silver.gif" border=0>
    <input type=radio name=strRankColorMod value="bronze"<% if strRankColorMod = "bronze" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/bronze.gif" border=0>
    <input type=radio name=strRankColorMod value="orange"<% if strRankColorMod = "orange" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/orange.gif" border=0>
    <input type=radio name=strRankColorMod value="red"<% if strRankColorMod = "red" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/red.gif" border=0>
    <input type=radio name=strRankColorMod value="purple"<% if strRankColorMod = "purple" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/purple.gif" border=0>
    <input type=radio name=strRankColorMod value="blue"<% if strRankColorMod = "blue" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/blue.gif" border=0>
    <input type=radio name=strRankColorMod value="cyan"<% if strRankColorMod = "cyan" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/cyan.gif" border=0>
    <input type=radio name=strRankColorMod value="green"<% if strRankColorMod = "green" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/green.gif" border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb10 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel0" size="30" value="<% if strRankLevel0 <> " " then Response.Write(strRankLevel0) else Response.Write("Starting Member") end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel0" size="5" value="0">
    <%= icon(icnHelp,"("&txtCMb16&")","","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor0 value="gold"<% if strRankColor0 = "gold" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/gold.gif" border=0>
    <input type=radio name=strRankColor0 value="silver"<% if strRankColor0 = "silver" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/silver.gif" border=0>
    <input type=radio name=strRankColor0 value="bronze"<% if strRankColor0 = "bronze" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/bronze.gif" border=0>
    <input type=radio name=strRankColor0 value="orange"<% if strRankColor0 = "orange" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/orange.gif" border=0>
    <input type=radio name=strRankColor0 value="red"<% if strRankColor0 = "red" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/red.gif" border=0>
    <input type=radio name=strRankColor0 value="purple"<% if strRankColor0 = "purple" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/purple.gif" border=0>
    <input type=radio name=strRankColor0 value="blue"<% if strRankColor0 = "blue" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/blue.gif" border=0>
    <input type=radio name=strRankColor0 value="cyan"<% if strRankColor0 = "cyan" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/cyan.gif" border=0>
    <input type=radio name=strRankColor0 value="green"<% if strRankColor0 = "green" then Response.Write(" checked") %>><img src="Themes/<%= strTheme %>/Stars/green.gif" border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb11 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel1" size="30" value="<% if strRankLevel1 <> " " then Response.Write(strRankLevel1) end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel1" size="5" value="<% =intRankLevel1 %>">
    <%= icon(icnHelp,"("&txtCMb16&")","","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor1 value=gold<% if strRankColor1 = "gold" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/gold.gif border=0>
    <input type=radio name=strRankColor1 value=silver<% if strRankColor1 = "silver" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/silver.gif border=0>
    <input type=radio name=strRankColor1 value=bronze<% if strRankColor1 = "bronze" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/bronze.gif border=0>
    <input type=radio name=strRankColor1 value=orange<% if strRankColor1 = "orange" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/orange.gif border=0>
    <input type=radio name=strRankColor1 value=red<% if strRankColor1 = "red" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/red.gif border=0>
    <input type=radio name=strRankColor1 value=purple<% if strRankColor1 = "purple" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/purple.gif border=0>
    <input type=radio name=strRankColor1 value=blue<% if strRankColor1 = "blue" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/blue.gif border=0>
    <input type=radio name=strRankColor1 value=cyan<% if strRankColor1 = "cyan" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/cyan.gif border=0>
    <input type=radio name=strRankColor1 value=green<% if strRankColor1 = "green" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/green.gif border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
   </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb12 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel2" size="30" value="<% if strRankLevel2 <> " " then Response.Write(strRankLevel2) end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel2" size="5" value="<% =intRankLevel2 %>">
    <%= icon(icnHelp,"("&txtCMb16&")","","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor2 value=gold<% if strRankColor2 = "gold" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/gold.gif border=0>
    <input type=radio name=strRankColor2 value=silver<% if strRankColor2 = "silver" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/silver.gif border=0>
    <input type=radio name=strRankColor2 value=bronze<% if strRankColor2 = "bronze" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/bronze.gif border=0>
    <input type=radio name=strRankColor2 value=orange<% if strRankColor2 = "orange" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/orange.gif border=0>
    <input type=radio name=strRankColor2 value=red<% if strRankColor2 = "red" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/red.gif border=0>
    <input type=radio name=strRankColor2 value=purple<% if strRankColor2 = "purple" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/purple.gif border=0>
    <input type=radio name=strRankColor2 value=blue<% if strRankColor2 = "blue" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/blue.gif border=0>
    <input type=radio name=strRankColor2 value=cyan<% if strRankColor2 = "cyan" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/cyan.gif border=0>
    <input type=radio name=strRankColor2 value=green<% if strRankColor2 = "green" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/green.gif border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb13 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel3" size="30" value="<% if strRankLevel3 <> " " then Response.Write(strRankLevel3) end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel3" size="5" value="<% =intRankLevel3 %>">
    <%= icon(icnHelp,"("&txtCMb16&")","","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor3 value=gold<% if strRankColor3 = "gold" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/gold.gif border=0>
    <input type=radio name=strRankColor3 value=silver<% if strRankColor3 = "silver" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/silver.gif border=0>
    <input type=radio name=strRankColor3 value=bronze<% if strRankColor3 = "bronze" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/bronze.gif border=0>
    <input type=radio name=strRankColor3 value=orange<% if strRankColor3 = "orange" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/orange.gif border=0>
    <input type=radio name=strRankColor3 value=red<% if strRankColor3 = "red" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/red.gif border=0>
    <input type=radio name=strRankColor3 value=purple<% if strRankColor3 = "purple" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/purple.gif border=0>
    <input type=radio name=strRankColor3 value=blue<% if strRankColor3 = "blue" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/blue.gif border=0>
    <input type=radio name=strRankColor3 value=cyan<% if strRankColor3 = "cyan" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/cyan.gif border=0>
    <input type=radio name=strRankColor3 value=green<% if strRankColor3 = "green" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/green.gif border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb14 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel4" size="30" value="<% if strRankLevel4 <> " " then Response.Write(strRankLevel4) end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel4" size="5" value="<% =intRankLevel4 %>">
    <%= icon(icnHelp,"("&txtCMb16&")","","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor4 value=gold<% if strRankColor4 = "gold" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/gold.gif border=0>
    <input type=radio name=strRankColor4 value=silver<% if strRankColor4 = "silver" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/silver.gif border=0>
    <input type=radio name=strRankColor4 value=bronze<% if strRankColor4 = "bronze" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/bronze.gif border=0>
    <input type=radio name=strRankColor4 value=orange<% if strRankColor4 = "orange" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/orange.gif border=0>
    <input type=radio name=strRankColor4 value=red<% if strRankColor4 = "red" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/red.gif border=0>
    <input type=radio name=strRankColor4 value=purple<% if strRankColor4 = "purple" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/purple.gif border=0>
    <input type=radio name=strRankColor4 value=blue<% if strRankColor4 = "blue" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/blue.gif border=0>
    <input type=radio name=strRankColor4 value=cyan<% if strRankColor4 = "cyan" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/cyan.gif border=0>
    <input type=radio name=strRankColor4 value=green<% if strRankColor4 = "green" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/green.gif border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align="right"><b><%= txtCMb15 %>:</b>&nbsp;</td>
    <td class="fNorm"><input type="text" name="strRankLevel5" size="30" value="<% if strRankLevel5 <> " " then Response.Write(strRankLevel5) end if %>">
    <b><%= txtPosts %>:</b>&nbsp;<input type="text" name="intRankLevel5" size="5" value="<% =intRankLevel5 %>">
    <%= icon(icnHelp,txtCMb16,"","","") %>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" align=right><b><%= txtCMb07 %>:</b></td>
    <td class="fNorm">
    <input type=radio name=strRankColor5 value=gold<% if strRankColor5 = "gold" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/gold.gif border=0>
    <input type=radio name=strRankColor5 value=silver<% if strRankColor5 = "silver" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/silver.gif border=0>
    <input type=radio name=strRankColor5 value=bronze<% if strRankColor5 = "bronze" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/bronze.gif border=0>
    <input type=radio name=strRankColor5 value=orange<% if strRankColor5 = "orange" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/orange.gif border=0>
    <input type=radio name=strRankColor5 value=red<% if strRankColor5 = "red" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/red.gif border=0>
    <input type=radio name=strRankColor5 value=purple<% if strRankColor5 = "purple" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/purple.gif border=0>
    <input type=radio name=strRankColor5 value=blue<% if strRankColor5 = "blue" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/blue.gif border=0>
    <input type=radio name=strRankColor5 value=cyan<% if strRankColor5 = "cyan" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/cyan.gif border=0>
    <input type=radio name=strRankColor5 value=green<% if strRankColor5 = "green" then Response.Write(" checked") %>><img src=Themes/<%= strTheme %>/Stars/green.gif border=0>
    <a href="JavaScript:openWindow3('pop_help.asp?mode=2&place=1#RankColor')"><%= icon(icnHelp,txtHelp,"","","") %></a>
    </td>
  </tr>
  <tr valign="top">
    <td class="fNorm" colspan="2" align="center"><br /><input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" class="button">&nbsp;&nbsp;&nbsp;<input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"><br />&nbsp;</td>
  </tr>
</table>
    </td>
  </tr>
</table></form>
	</div>
<%
end sub

sub memberCleaning() %>
	<div id="ac" style="display:<%= ac %>;">
<table border="1" style="border-collapse: collapse" class="grid" width="75%" cellSpacing="0" cellPadding="0" align="center">
  <tr>
	<td class="tCellAlt2">
	<table align="center" width="100%" cellspacing="0" cellpadding="4" border="0">
<form method="post" action="admin_config_members.asp?cmd=2&mode=ready" id="formEle">
	  <tr>
    	<td class="tTitle"><b><%= txtCMb36 %></b></td>
	  </tr>
	  <tr>
    	<td class="fNorm"><%= txtCMb37 %>&nbsp;
              <select name="days">
                <option value="1000"<% if request.form("days") = "1000" then%> selected<%end if%>>1000</option>
                <option value="500"<% if request.form("days") = "500" then%> selected<%end if%>>500</option>
                <option value="365"<% if request.form("days") = "365" then%> selected<%end if%>>365</option>
                <option value="180"<% if request.form("days") = "180" then%> selected<%end if%>>180</option>
                <option value="120"<% if request.form("days") = "120" then%> selected<%end if%>>120</option>
                <option value="90"<% if request.form("days") = "90" or request.form("days") = "" then%> selected<%end if%>>90</option>
                <option value="60"<% if request.form("days") = "60" then%> selected<%end if%>>60</option>
                <option value="30"<% if request.form("days") = "30" then%> selected<%end if%>>30</option>
                <option value="10"<% if request.form("days") = "10" then%> selected<%end if%>>10</option>
                <option value="5"<% if request.form("days") = "5" then%> selected<%end if%>>5</option>
              </select>
    	&nbsp;<%= txtCMb38 %>&nbsp;

              <select name="posts">
                <option value="0"<% if request.form("posts") = "0" or request.form("posts") = "" then%> selected<%end if%>>0</option>
                <option value="5"<% if request.form("posts") = "5" then%> selected<%end if%>>less than 5</option>
                <option value="15"<% if request.form("posts") = "15" then%> selected<%end if%>>less than 15</option>
                <option value="30"<% if request.form("posts") = "30" then%> selected<%end if%>>less than 30</option>
                <option value="60"<% if request.form("posts") = "60" then%> selected<%end if%>>less than 60</option>
              </select>	
    	&nbsp;<%= txtCMb39 %></td>
    	  </tr>
<tr>
	<td class="fNorm"><input type="submit" value="<%= txtCMb40 %>" class="button"></td>
</tr></form>
	</table>
	</td>
  </tr>
<%if request.querystring("mode")= "ready" then
cDays = cint(request.form("days"))
cPosts = cint(request.form("posts"))

	strSql = "SELECT COUNT(MEMBER_ID) AS MEMCOUNT "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE M_POSTS <= " & cPosts & " AND M_LASTHEREDATE < '" & DateToStr(DateAdd("d", cDays*-1+1 , now())) & "' AND M_LEVEL = 1 AND M_STATUS = 1"

	set rs = my_Conn.Execute(strSql)
	
	memCount = rs("MEMCOUNT")
	
	rs.close
	set rs = nothing
%>
  <tr>
	<td class="tCellAlt2">
	<table align="center" width="100%" cellspacing="0" cellpadding="4" border="0">
<form method="post" action="admin_config_members.asp?cmd=2&mode=deleted">
  <tr>
<td class="fNorm"><%= replace(txtCMb42,"[%count%]",memCount) %></td>
  </tr>
  <tr>
<td class="fNorm">
<input type="hidden" name="cDays" value="<%=cDays%>">
<input type="hidden" name="cPosts" value="<%=cPosts%>">
<input type="hidden" name="cNum" value="<%=memCount%>">
<input type="submit" value="<%= txtCMb41 %>" class="button">&nbsp;<%= txtCMb43 %></td>
  </tr>
</form>
	</table>
	</td>
  </tr>
<%end if%>
</table>
	</div>
<%
end sub %>