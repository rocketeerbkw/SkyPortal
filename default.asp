<!--#include file="config.asp" --><%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> SkyPortal Software is
'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
'<>
'<> By using the SkyPortal software, you are agreeing to the
'<> terms of the SkyPortal End-User License Agreement.
'<>
'<> All copyright notices regarding SkyPortal must remain 
'<> intact in the scripts and in the outputted HTML.
'<> The "powered by" text/logo with a link back to 
'<> http://www.SkyPortal.net in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> SkyPortal Software Support can be obtained from support forums at:
'<> http://www.SkyPortal.net
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
dim CurPageType, CurPageInfoChk, spThemeTitle, spThemeMM
CurPageType = "home"
CurPageInfoChk = "1"
%>
<!--#include file="inc_functions.asp" -->
<%
function CurPageInfo ()
	PageName = txtHome 
	PageAction = txtBrows & "<br />" 
	PageLocation = "Default.asp" 
	CurPageInfo = PageAction & "<a href=""" & PageLocation & """>" & PageName & "</a>"
end function 
%>
<!--#include file="inc_top.asp" -->
<% 
function shoBlocks(arrCol)
    for fp = 0 to ubound(arrCol)
	  fTemp = split(arrCol(fp),":")
      if ubound(fTemp) = 2 then
	    fFunct = fTemp(1) & "(""" & fTemp(2) & """)"
	  else
	    fFunct = fTemp(1)
	  end if
  	  execute(fFunct)
    next
end function

l_sticky = ""
m_sticky = ""
r_sticky = ""

  'get the default layout including 'sticky' items
  fpSQL = "select * from PORTAL_FP_USERS where fp_uid = 0"
  set rsFPs = my_Conn.execute(fpSQL)
  if not rsFPs.eof then
    if rsFPs("fp_leftsticky") <> "" then
  	  left_Col = rsFPs("fp_leftsticky") & "," & rsFPs("fp_leftcol")
  	  l_sticky = rsFPs("fp_leftsticky")
	else
  	  left_Col = rsFPs("fp_leftcol")
	end if
    if rsFPs("fp_mainsticky") <> "" then
	  main_Col = rsFPs("fp_mainsticky") & "," & rsFPs("fp_maincol")
  	  m_sticky = rsFPs("fp_mainsticky")
	else
  	  main_Col = rsFPs("fp_maincol")
	end if
    if rsFPs("fp_rightsticky") <> "" then
	  right_Col = rsFPs("fp_rightsticky") & "," & rsFPs("fp_rightcol")
  	  r_sticky = rsFPs("fp_rightsticky")
	else
  	  right_Col = rsFPs("fp_rightcol")
	end if
  end if
  set rsFPs = nothing
  
if strdbntusername <> "" and intMyMax then
  fpSQL = "select * from PORTAL_FP_USERS where fp_uid = " & strUserMemberID
  set rsFP = my_Conn.execute(fpSQL)
  if not rsFP.eof then
    ':: use default layout
    if l_sticky <> "" then
    left_Col = l_sticky & "," & rsFP("fp_leftcol")
	else
	    left_Col = rsFP("fp_leftcol")
	end if
	if m_sticky <> "" then
	    main_Col = m_sticky & "," & rsFP("fp_maincol")
	else
	    main_Col = rsFP("fp_maincol")
	end if
	if r_sticky <> "" then
	    right_Col = r_sticky & "," & rsFP("fp_rightcol")
	else
	    right_Col = rsFP("fp_rightcol")
	end if
  end if
  set rsFP = nothing
end if

  cont = 0
  bLeft = false
  bMain = false
  bRight = false
  
  if right(left_Col,1) = "," then
    left_Col = left(left_Col,len(left_Col)-1)
  end if
  if right(main_Col,1) = "," then
    main_Col = left(main_Col,len(main_Col)-1)
  end if
  if right(right_Col,1) = "," then
    right_Col = left(right_Col,len(right_Col)-1)
  end if
  
  l_col = split(left_Col,",")
  m_col = split(main_Col,",")
  r_col = split(right_Col,",")
  
  if trim(left_Col) <> "" then
    bLeft = true
    cont = cont + 1
  end if
  if trim(main_Col) <> "" then
    bMain = true
    cont = cont + 1
  end if
  if trim(right_Col) <> "" then
    bRight = true
    cont = cont + 1
  end if
  
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if bLeft then
    response.Write("<td class=""leftPgCol"" valign=""top"">")
	intSkin = getSkin(intSubSkin,1)
    'response.Write("intSkin: " & intSkin)
	 shoBlocks(l_col)
    response.Write("</td>")
  end if
  
  if bMain then
    response.Write("<td class=""mainPgCol"" valign=""top"">")
	intSkin = getSkin(intSubSkin,2)
	arg2 = ""
    arg3 = ""
    arg4 = ""
    arg5 = ""
    arg6 = "Welcome to " & strSiteTitle
  
    shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
    'response.Write("intSkin: " & intSkin)
	 shoBlocks(m_col)
    response.Write("</td>")
  end if
  
  if bRight then
    if cont = 3 then
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
    'response.Write("intSkin: " & intSkin)
	shoBlocks(r_col)
    response.Write("</td>")
  end if
  response.Write("</tr></table>")
 %>
<!--#include file="inc_footer.asp" -->