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
curpagetype = "core"
%>
<!--#include file="inc_functions.asp" -->
<!--include file="custom_functions.asp" -->
<%
  dim sfTitle, sfSummary, sfURL, sfPublic, sfActive, sfFeatured
  dim sfWL, sfYH, sfBL, sfNG, sfGG, gRead, sfSCATid, sfID
  dim sfPModuleID, sfMCat, sfMSCat, sfPHDisplay, sfPHRead ,sfPHTZone
  dim sfPHImage, sfPFTable, sfPFid, sfPFTitle, sfPFAuthor
  dim sfPFAuthorInfo, sfPFSummary, sfPFPostDate, sfPFWhere, sfPFOrderBy
  
  dim iMode, iCmd, cid, app
  iMode = 0
  iCmd = 0
  cid = 0
  iShow = 5
  app = ""

if IsNumeric(Request("mode")) = True then
	iMode = cLng(Request("mode"))
else
  if Request("mode") <> "" then
	sMode = Request("mode")
  else
	closeAndGo("stop")
  end if
end if

if Request("cmd") <> "" and Request("cmd") <> " " then
	if IsNumeric(Request("cmd")) = True then
		iCmd = cLng(Request("cmd"))
	end if
end if

if Request("cid") <> "" and Request("cid") <> " " then
	if IsNumeric(Request("cid")) = True then
		cid = cLng(Request("cid"))
	end if
end if

if Request("sid") <> "" and Request("sid") <> " " then
	if IsNumeric(Request("sid")) = True then
		sid = cLng(Request("sid"))
	end if
end if

if Request("show") <> "" and Request("show") <> " " then
	if IsNumeric(Request("show")) = True then
		iShow = cLng(Request("show"))
	end if
end if

if Request("col") <> "" and Request("col") <> " " then
	if IsNumeric(Request("col")) = True then
		iCol = cLng(Request("col"))
	end if
end if
%>
<!--#include file="includes/inc_top_ajax.asp" -->
<%
  ':: set default module permissions
  'setAppPerms curpagetype,"iName"
  
  select case sMode
	  case "sajx_donate"
		Call sajx_sb_donate()
	  case "sajx_rateus" 
		Call sajx_sb_rateus()
	  case "sajx_affiliateBanners" 
		Call sajx_affiliateBanners()
	  case "sajx_login_box" 
		Call sajx_login_box()
	  case "sajx_theme_changer" 
		Call sajx_theme_changer()
	  case "sajx_welcome_fp" 
		Call sajx_welcome_fp()
	  case "sajx_announce_fp" 
		Call sajx_announce_fp()
	  case else
	    Response.Write "An error has occurred:" & sMode
  end select
  
'my_Conn.Close
'set my_Conn = nothing
closeObjects()

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

function sajx_sb_donate()
spThemeMM = "othrs"
spThemeTitle= "Support SkyPortal"
spThemeBlock1_open(intSkin)
'spThemeBlock2_open()%>
<table border="0" cellpadding="0" cellspacing="0"><tr>
  <td align="center">
  <p>Please help support the continued development of SkyPortal by making your donation today.</p>
  <p><a href="http://www.skyportal.net/site_donation.asp"><img src="http://www.skyportal.net/images/donation_sp.gif" border="0" alt="" title="Help support the SkyPortal Development" /></a></p>
  </td></tr></table>
<%
spThemeBlock1_close(intSkin)
'spThemeBlock2_close()
end function

function sajx_sb_rateus()
spThemeTitle = "Rate SkyPortal"
spThemeBlock1_open(intSkin) %>
    <table width="100%" border="0" cellspacing="0" cellpadding="2">
      <tr> 
        <td align="left" valign="middle">
		<p>If you use SkyPortal and think we are the best around, and you want everyone to know it, please vote for us below!</p>
        </td>
      </tr>
      <tr> 
        <td align="left" valign="middle"><hr />
		 <table width="100%" align="right" border="0" cellpadding="1" cellspacing="0">
		  <tr><td align="center">
		    <table width="100%" border="0" cellpadding="3" cellspacing="0">
			  <tr><td align="center">
			    <font style="font-size:10pt;font-family:Arial;"><b>Rated:</b> <a href="http://www.Aspin.com/func/review?id=6559210"><img src="http://ratings.Aspin.com/getstars?id=6559210" border="0" alt=""></a>  <font style="font-size:8pt;"><br />by <a href="http://www.Aspin.com" target="_blank">Aspin.com</a> users<br /></font></font> </td></tr>
			  <tr><td align="center" nowrap><form action="http://www.Aspin.com/func/review/write?id=6559210" method="post" target="_blank"> <font style="font-size:10pt;font-family:Arial;">What do you think?</font><br /> <select name="VoteStars"><option>5 Stars<option>4 Stars<option>3 Stars<option>2 Stars<option>1 Star</select>&nbsp;&nbsp;<input type=submit value="Vote"></form></td></tr>
			</table>
		  </td></tr>
		 </table>
        </td>
      </tr>
      <tr> 
        <td align="left" valign="middle"><hr />
		 <table width="100%" align="right" border="0" cellpadding="1" cellspacing="0">
		  <tr><td align="center">
		    <table width="100%" border="0" cellpadding="3" cellspacing="0">
			  <tr><td align="center" nowrap>
				<% sajx_codango_rating() %>
			  </td></tr>
			</table>
		  </td></tr>
		 </table> 
        </td>
      </tr>
      <tr> 
        
      <td align="center" valign="middle"> <a href="http://www.codango.com/asp/fnc/lab/review/1004/" target="_blank" title="Read the Codango Lab Review for SkyPortal RC5"><img src="http://www.codango.com/images/labReview/awards/CodangoSilverAwd86.png" alt="Codango Silver Award" title="Read the Codango Lab Review for SkyPortal RC5" border="0" /></a> 
      </td>
      </tr>
    </table>
<%
spThemeBlock1_close(intSkin)
end function

sub sajx_codango_rating()
  %>
  <div style="border:1px solid #aaaaaa;background-color:#ffffff;color:#555555;width:125px;text-align:center;font-family:Arial;font-size:12px;white-space:nowrap;">
<div style="margin-top:3px;border-bottom:4px solid #F8F8F8;"><a href="http://www.codango.com"><img src="http://images.codango.com/linkmedia/v1/RateBoxCdgLogo.gif" align="center" alt="Codango PHP, ASP .NET, JSP Scripts, Resources, Reviews" width="88" height="17" border="0"></a></div>
<div style="border-top:1px solid #D0D0D0;border-bottom:1px solid #aaaaaa;background:#ffffff url('http://images.codango.com/linkmedia/v1/RateBoxFade.gif') repeat-x;"><a href="http://www.codango.com/asp/fnc/review/?tree=aspin/webapps/instantw&id=6559210"><img src="http://images.codango.com/getstars/?id=6559210" align="right" Style="margin:6px 10px 0px 0px;" alt="Click to Read Reviews" width="62" height="12" border="0"></a>
<div style="font-size:11px;font-weight:bold;line-height:11px;text-align:center;margin:1px 0px 2px 0px;">User<br />Rated</div></div><div style="margin-top:4px;background-color:#F8F8F8">What do you think?</div>
<div style="border-bottom:5px solid #EBEBEB;background-color:#F3F3F3"><form style="margin:0px;" action="http://www.codango.com/asp/fnc/review/write/?id=6559210" method="post"><select style="width:70px;" name="VoteStars"><option>5 Stars<option>4 Stars<option>3 Stars<option>2 Stars<option>1 Star</select><input style="width:43px;" type="submit" value="Vote"></form></div></div>
  <%
end sub

sub sajx_affiliateBanners()
	  blkStart = timer
  showHowMany = 10 'ORDER BY ID DESC
  sSQL = "Select * FROM PORTAL_BANNERS WHERE B_LOCATION=2 AND B_ACTIVE=1"
  'executeThis(sSQL)
  'set rsAB = my_Conn.execute(sSQL)
  Set rsAB = oSpData.GetRecordset(sSql)
  if not rsAB.eof then
	spThemeMM = "aff_sm"
    spThemeTitle = txtAffiliates
    spThemeBlock1_open(intSkin)
     %><div>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	<% do until rsAB.eof
		strImage = rsAB("B_IMAGE")
		strHover = rsAB("B_ACRONYM")
		intID = rsAB("ID") %>
			<tr><td align="center" height="40"><a target="_blank" title="<%= strHover %>" href="banner_link.asp?id=<%= intID %>"><% If right(strImage,4) = ".swf" Then writeFlash2 strImage,intID,strHover Else response.write("<img alt=""" & strHover & """ name=""abImage"" border=""0"" src=""" & strImage & """ />") end if %></a></td></tr>
	<%     rsAB.movenext
	   loop %>
		</table></div>
<%
    spThemeBlock1_close(intSkin)
  end if
  set rsAB = nothing
	if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - blkStart),3)
	  response.Write(blkLoadTime)
	end if
end sub

sub sajx_login_box()
if not hasAccess(2) then
spThemeTitle= txtLogin
spThemeBlock1_open(intSkin) %>
<table border="0" cellpadding="0" cellspacing="0"><tr><td>
<form action="default.asp" method="post" id="logmex" name="logmex">
<table cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
<td align="center"><br />
  <%= txtUsrName %>:<br><input class="textbox" type="text" name="Name" size="15" maxlength="25" value="" style="margin:0px;" /><br /><br />
  <%= txtPass %>:<br><input class="textbox" type="password" name="Password" size="15" maxlength="25" style="margin:0px;" />
<% If SecImage >1 Then %>
<br /></td></tr><tr>
<td align="center" height="50">
<img src="includes/securelog/image.asp" alt="<%= txtSecImg %>" title="<%= txtSecImg %>" /><br />
</td></tr><tr>
<td align="center">
<input class="textbox" type="text" name="secCode" size="15" maxlength="8" value="<%= txtSecCode %>" onFocus="javascript:this.value='';" />
<%end if %><br />
</td></tr><tr>
<td align="center"><br />
<input type="checkbox" name="SavePassWord" value="true" checked="checked" />&nbsp;&nbsp;<%= txtSvPass %>
<br /><br />
<input type="submit" value="<%= txtLogin %>" id="logmein" name="logmein" class="btnLogin" /><input type="hidden" name="Method_Type" value="login" /><br />
<%if (lcase(strEmail) = "1") then %>
<br /><a href="password.asp"><%= txtForgotPass %>?<br /><span class="fSmall">Click Here</span></a><br />
<% end if %>
<%if strNewReg = 1 then %>
<br /><%= txtNotMember %>?<br /><a href="policy.asp"><span class="fAlert"><%= txtRegNow %>!</span></a><br />
<% End If %>
</td></tr>
</table></form></td></tr></table>
<%
spThemeBlock1_close(intSkin)
end if
end sub

function sajx_theme_changer()
  spThemeMM = "sknchgr"
  spThemeTitle = txtSknChgr
  spThemeBlock1_open(intSkin)
  %>
    <div class="spThemeChanger">
    <form name="themechanger" method="post" action="default.asp">
    <span class="fNorm">Select Skin:</span><br />
      <select name="thm" onChange="submit();">
        <% 
	ssSQL = "select C_STRAUTHOR, C_TEMPLATE, C_STRFOLDER, C_SKINLEVEL from portal_colors ORDER BY C_TEMPLATE"
	set rsThm = my_Conn.execute(ssSQL)
	if rsThm.eof then
		'strAuth = "anonymous"
	else
	  do until rsThm.eof
	    if hasAccess(rsThm("C_SKINLEVEL")) then
		  if rsThm("C_STRFOLDER") = strTheme then
	        Response.Write("<option value="""& rsThm("C_STRFOLDER") &""" selected=""selected"">"& rsThm("C_TEMPLATE") &"</option>")
		    strAuth = rsThm("C_STRAUTHOR")
		  else
	        Response.Write("<option value="""& rsThm("C_STRFOLDER") &""">"& rsThm("C_TEMPLATE") &"</option>")
		  end if
		end if
	    rsThm.movenext
	  loop
	end if
	set rsThm = nothing
	%>
      </select>
  </form>
    <span class="fSmall"><br /><%= txtAuthor %>:<b> <%= strAuth %> </b></span>
    </div>
  <% spThemeFooter = ""
  spThemeBlock1_close(intSkin)
end function 

function sajx_welcome_fp()
	if hasAccess(2) then
	  w_id = 1
	else
	  w_id = 2
	end if
	strSql = "SELECT * FROM " & strTablePrefix & "WELCOME WHERE W_ID = " & w_id & " AND W_ACTIVE=1"
	set rsWelcome =  my_Conn.Execute (strSql)
    if not rsWelcome.EOF then
	  W_ID = rsWelcome("W_ID")
	  W_TITLE = trim(replace(rsWelcome("W_TITLE"),"''","'"))
	  W_SUBJECT = trim(replace(rsWelcome("W_SUBJECT"),"''","'"))
	  W_SUBJECT	= replace(W_SUBJECT,"[%member%]",strdbntusername)
	  W_MESSAGE = trim(replace(rsWelcome("W_MESSAGE"),"''","'"))
	  W_MESSAGE	= replace(W_MESSAGE,"</p><p>","<br /><br />")
	  W_MESSAGE	= replace(W_MESSAGE,"<p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"</p>","")
	  W_MESSAGE	= replace(W_MESSAGE,"[%member%]",strdbntusername)
	  W_MESSAGE = FormatStr2(W_MESSAGE)

	  spThemeMM = "welcom"
	  'spThemeTitle = txtWelcomeTo & " " & strSiteTitle
	  spThemeTitle = W_TITLE
	  spThemeBlock1_open(intSkin) %>
		<p style="text-align:left;"><b><%= W_SUBJECT %></b><br><%= W_MESSAGE %></p>
	  <%
	  spThemeBlock1_close(intSkin)
	End if 
    set rsWelcome = nothing
end function

function sajx_announce_fp()
	intFirst = 1
	strSql = "SELECT * FROM " & strTablePrefix & "ANNOUNCEMENTS WHERE A_START_DATE <= '" & strCurDateString & "' and A_END_DATE >= '" & strCurDateString & "' ORDER BY A_START_DATE DESC;"
	set rsAnn =  my_Conn.Execute(strSql)
	if not rsAnn.EOF then 
	  spThemeTitle=txtAnnouncements
	  spThemeBlock1_open(intSkin)%>
	  <table id="annCat" border="0" cellspacing="0" cellpadding="2" width="100%"><%
  	  Do until rsAnn.EOF
		 A_ID = rsAnn("A_ID")
		 A_SUBJECT = trim(replace(rsAnn("A_SUBJECT"),"''","'"))
		 A_DATE = ChkDate(rsAnn("A_START_DATE"))
		 A_MESSAGE = trim(replace(rsAnn("A_MESSAGE"),"''","'"))
		 A_MESSAGE	= replace(A_MESSAGE,"</p><p>","<br /><br />")
		 A_MESSAGE	= replace(A_MESSAGE,"<p>","")
		 A_MESSAGE	= replace(A_MESSAGE,"</p>","")
		 
		 if intFirst = 1 then
			catHide = ""
			catImg = "min"
			catAlt = txtCollapse
		 else
			catHide = "none"
			catImg = "max"
			catAlt = txtExpand
		 end if %>			
	        <tr>
			<td width="80%" height="25" valign="top" class="tSubTitle"><% If hasAccess(1) Then %><a href="admin_announce.asp?cmd=1&a_id=<%= A_ID %>"><%= icon(icnEdit,txtEdit,"","","align=""right""") %></a><% End If %><img name="annCat<%=A_ID%>Img" id="annCat<%=A_ID%>Img" src="Themes/<%=strTheme%>/icon_<%=catImg%>.gif" onClick="javascript:mwpHS('annCat','<%=A_ID%>','tbody');" style="cursor:pointer;" title="<%=catAlt%>" alt="<%=catAlt%>" />&nbsp;<b><%=A_SUBJECT%></b>&nbsp;
			</td>
			<td width="20%" valign="middle" class="tSubTitle" align="right"><b><%= A_DATE %></b>
			</td>
			</tr>
			<tbody id="annCat<%=A_ID%>" style="display:<%=catHide%>;">
			<tr>
			<td width="100%" valign="top" colspan="2">
			<p align="justify"><%= A_MESSAGE %></p>
			</td>
			</tr>
			</tbody>
		<%
		intFirst = 0
		rsAnn.MoveNext
	  Loop 
	  set rsAnn = nothing %>
	</table>
 	<% spThemeBlock1_close(intSkin)
	End if
end function
%>