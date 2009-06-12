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
dim CurPageType, CurPageInfoChk, spThemeTitle, spThemeMM
CurPageType = "core"
sPage_INAME = "privacy"
pgname = "ERROR!"
sPage_id = 0
CurPageInfoChk = "1"
%>
<!--#include file="inc_functions.asp" -->
<%

  'get the default layout 
  if sPage_id = 0 then
    cpSQL = "select * from PORTAL_PAGES where P_INAME = '" & sPage_INAME & "'"
  else
    cpSQL = "select * from PORTAL_PAGES where P_ID = " & sPage_id & ""
  end if
  set rsCPs = my_Conn.execute(cpSQL)
  if not rsCPs.eof then
	  pgtitle = rsCPs("P_TITLE")
	  pgname = rsCPs("P_NAME")
	  pgbody = rsCPs("P_CONTENT")
  	  left_Col = rsCPs("P_LEFTCOL")
  	  maint_Col = rsCPs("P_MAINTOP")
	  mainb_Col = rsCPs("P_MAINBOTTOM")
  	  right_Col = rsCPs("P_RIGHTCOL")
	  
	  m_title = rsCPs("P_META_TITLE")
	  addToMeta "NAME","Description",rsCPs("P_META_DESC")
	  addToMeta "NAME","Keywords",rsCPs("P_META_KEY")
	  addToMeta "HTTP-EQUIV","Expires",rsCPs("P_META_EXPIRES")
	  addToMeta "NAME","Rating",rsCPs("P_META_RATING")
	  addToMeta "NAME","Distribution",rsCPs("P_META_DIST")
	  addToMeta "NAME","Robots",rsCPs("P_META_ROBOTS")
		'm_description = rsCPs("P_META_DESC")
		'm_keywords = rsCPs("P_META_KEY")
		'm_expires = rsCPs("P_META_EXPIRES")
		'm_rating = rsCPs("P_META_RATING")
		'm_distribution = rsCPs("P_META_DIST")
		'm_robots = rsCPs("P_META_ROBOTS")
  end if
  set rsCPs = nothing

PageTitle = m_title

function CurPageInfo () 
	PageName = pgname 
	PageAction = txtBrows & "<br />" 
	PageLocation = request.ServerVariables("URL")
	CurPageInfo = PageAction & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function 
%>
<!--#include file="inc_top.asp" -->
<% 
  cont = 0
  bLeft = false
  bMaint = false
  bMainb = false
  bRight = false
  
  if trim(left_Col) <> "" then
    if right(left_Col,1) = "," then
      left_Col = left(left_Col,len(left_Col)-1)
    end if
    if instr(left_Col,",") > 0 then
      l_col = split(left_Col,",")
	else
	  dim l_col(0)
      l_col(0) = left_Col
	end if
    bLeft = true
    cont = cont + 1
  end if
  if trim(maint_Col) <> "" then
    if right(maint_Col,1) = "," then
      maint_Col = left(maint_Col,len(maint_Col)-1)
    end if
    if instr(maint_Col,",") > 0 then
      mt_col = split(maint_Col,",")
	else
	  dim mt_col(0)
      mt_col(0) = maint_Col
	end if
    bMaint = true
    cont = cont + 1
  end if
  if trim(mainb_Col) <> "" then
    if right(mainb_Col,1) = "," then
      mainb_Col = left(mainb_Col,len(mainb_Col)-1)
    end if
    if instr(mainb_Col,",") > 0 then
      mb_col = split(mainb_Col,",")
	else
	  dim mb_col(0)
      mb_col(0) = mainb_Col
	end if
    bMainb = true
    cont = cont + 1
  end if
  if trim(right_Col) <> "" then
    if right(right_Col,1) = "," then
      right_Col = left(right_Col,len(right_Col)-1)
    end if
    if instr(right_Col,",") > 0 then
      r_col = split(right_Col,",")
	else
	  dim r_col(0)
      r_col(0) = right_Col
	end if
    bRight = true
    cont = cont + 1
  end if
  
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if bLeft then
    response.Write("<td class=""leftPgCol"" valign=""top"" nowrap=""nowrap"">")
	intSkin = getSkin(intSubSkin,1)
	 shoBlocks(l_col)
    response.Write("</td>")
  end if

    response.Write("<td class=""mainPgCol"" valign=""top"">")  
	intSkin = getSkin(intSubSkin,2)
  if bMaint then
	 shoBlocks(mt_col)
  end if

    spThemeTitle = pgtitle
    spThemeBlock1_open(intSkin) %>
         <table width="100%"><tr><td>
             <% 
			 include.writeSource(pgbody)
			 'response.write pgbody %>
         </td></tr></table>
    <%
    spThemeBlock1_close(intSkin)

  if bMainb then
	 shoBlocks(mb_col)
  end if
    response.Write("</td>")
  
  if bRight then
    if cont = 3 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	shoBlocks(r_col)
    response.Write("</td>")
  end if
  response.Write("</tr></table>")

 %>
<!--#include file="inc_footer.asp" -->