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

':: modify one of the 2 values below
':: sPage_id is the id of the record in the database
':: and is the preferred way to call the recordset.
sPage_INAME = "error"
sPage_id = 0

':::::::::::::::::::::::::::::::::::::::::::::::::

erType = chkString(request.QueryString("type"),"sqlstring")

	select case erType
	  case "referrer"
	    PageTitle = "xss"
  	    arg1 = "XSS"
	    tbTitle = "Possible XSS attempt"
	  case "luser"
	    PageTitle = txtWasErr
  	    arg1 = txtError
	    tbTitle = txtError
	  case "lsec"
	    PageTitle = txtWasErr
  	    arg1 = txtError
	    tbTitle = txtError
	  case "noperm"
	    PageTitle = txtWasErr
  	    arg1 = txtError
	    tbTitle = txtNoPerm
	  case "notmember"
	    PageTitle = txtMbrsOnly
  	    arg1 = txtLogin
	    tbTitle = txtPlzLogin
	  case "nopermtask"
	    PageTitle = txtNoAccess
  	    arg1 = txtNoAccess
	    tbTitle = txtNoAccess
	  case "lockdown"
	    strLockDown = 0
	    PageTitle = txtMbrsOnly
  	    arg1 = txtLogin
	    tbTitle = txtPlzLogin
	  case "forumdown"
  	    arg1 = "Forum down"
	    tbTitle = "Forums down"	    
	  case else
	    PageTitle = txtWasErr
  	    arg1 = txtError
	    tbTitle = txtError
	end select
%>
<!-- #include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<%
  select case Request.Form("Method_Type")
	case "login"
	  closeAndGo(strHomeUrl)
	case "logout"
	  closeAndGo(strHomeUrl)
  end select
'response.End()
  'get the default layout 
  cpSQL = "select * from PORTAL_PAGES where P_INAME = '"&sPage_INAME&"'"
  set rsCPs = my_Conn.execute(cpSQL)
  if not rsCPs.eof then
  	  left_Col = rsCPs("p_leftcol")
  	  maint_Col = rsCPs("p_maintop")
	  mainb_Col = rsCPs("p_mainbottom")
  	  right_Col = rsCPs("p_rightcol")
  else
    set rsCPs = nothing
    'closeAndGo("default.asp?no_error_page")
  end if
  set rsCPs = nothing
  
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if trim(left_Col) <> "" then
    cont = cont + 1
    response.Write("<td class=""leftPgCol"" valign=""top"" nowrap=""nowrap"">")
	intSkin = getSkin(intSubSkin,1)
	shoColumnBlocks(left_Col)
    response.Write("</td>")
  end if

    response.Write("<td class=""mainPgCol"" valign=""top"">")  
	intSkin = getSkin(intSubSkin,2)
    cont = cont + 1
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  if trim(maint_Col) <> "" then
	 shoColumnBlocks(maint_Col)
  end if

    spThemeTitle = tbTitle
    spThemeBlock1_open(intSkin)
	select case erType
	  case "referrer"
	    badreferrer("form")
	  case "luser"
	    badLogin("user")
	  case "lsec"
	    badLogin("sec")
	  case "noperm"
	    noAccessPage()
	  case "nopermtask"
	    noAccessTask()
	  case "notmember"
	    loginOrRegister()
	  case "lockdown"
	    lockDown()
	  case "forumdown"
	    forumDown()
	  case "maint"
	    siteMaintenance()
	  case else
	end select
	Response.Write "<p>&nbsp;</p>"
    spThemeBlock1_close(intSkin)

  if trim(mainb_Col) <> "" then
	 shoColumnBlocks(mainb_Col)
  end if
    response.Write("</td>")
  
  if trim(right_Col) <> "" then
    if cont = 2 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	shoColumnBlocks(right_Col)
    response.Write("</td>")
  end if
  response.Write("</tr></table>")
 %>
<!--#include file="inc_footer.asp" -->
<%
sub loginOrRegister3()
end sub

sub badreferrer(t)
  response.Write "<p>&nbsp;</p>"
  response.Write("<p class=""fTitle"">" & txtWasErr & "</p>")
  select case t
    case "form"
      response.Write("<p>Possible hacking attempt.<br><br>")
  	  response.Write("The form that you submitted did not originate from this website. Your IP and other information has been recorded and sent to the proper site administrators.</p>")
	case else
  end select
end sub

sub badLogin(t)
  response.Write "<p>&nbsp;</p>"
  response.Write("<p class=""fTitle"">" & txtWasErr & "</p>")
  select case t
    case "user"
      response.Write("<p>" & txtBadLogin1)
  	  response.Write("&nbsp;" & txtWereInc & "</p>")
	case "sec"
      response.Write(txtBadSecCode)
	case else
  end select
  response.Write("<p>" & txtPlsTryAgain)
  response.Write("&nbsp;" & txtOr & "&nbsp;")
  response.Write("<a href=""policy.asp""><u>")
  response.Write(lcase(txtRegister) & "</u></a>&nbsp;")
  response.Write(txtForAccnt & "</p>")
end sub

sub siteMaintenance()
  closeAndGo("maintenance.asp")
end sub

sub loginOrRegister()
  txtToPartic = "to access this area"
  lockDownLoginForm()
  response.Write("<p align=""center"">")
  response.Write("<a href=""JavaScript:history.go(-2);"">")
  response.Write(txtGoBack & "</a></p>")
end sub

sub forumDown()
  showForumDown()
end sub

sub lockDown()
  lockDownLoginForm()
end sub

sub noAccessTask()
  response.Write "<p>&nbsp;</p>"
  response.Write("<p class=""fTitle"">" & txtWasErr & "</p>")
  response.Write("<p><b>" & txtNoAccPerformTask & "</b></p>")
  response.Write("<p align=""center"">")
  response.Write("<a href=""JavaScript:history.go(-2);"">")
  response.Write("<b>" & txtGoBack & "</b></a></p>")
end sub

sub noAccessPage()
  response.Write "<p>&nbsp;</p>"
  response.Write("<p class=""fTitle"">" & txtWasErr & "</p>")
  response.Write("<p><b>" & txtNoPermViewPg & "</b></p>")
  response.Write("<p align=""center"">")
  response.Write("<a href=""JavaScript:history.go(-2);"">")
  response.Write("<b>" & txtGoBack & "</b></a></p>") %>
  <p> You are not logged in or you do not have permission to access this page.<br />
  This could be due to one of several reasons:</p>
<ol>
  <li>You are not logged in. Fill in the form at the bottom of this page and try again.</li>
  <li>You may not have sufficient privileges to access this page.<br />
    Are you trying to access administrative features or some other privileged system?</li>
</ol>
<p>The administrator may have required you to register before you can view this page.</p>
  <%
end sub
%>