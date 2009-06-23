<!-- #INCLUDE FILE="config.asp" --><%
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
CurPageType = "article"
	dim strPoster
	dim strAuthor
	dim strAuthorEmail
	dim intVotes, intRating
	dim isOwner
	
  cmd = 0
  intMMID = 0
  lastid = 0
  articleid = 0

if Request.QueryString("item") <> "" or  Request.QueryString("item") <> " " then
	if IsNumeric(trim(Request.QueryString("item"))) then
		articleid = cLng(Request.QueryString("item"))
	else
		closeAndGo("article.asp")
	end if
end if
intItemID = articleid
if Request.QueryString("mm") <> "" then
	if IsNumeric(Request.QueryString("mm")) then
		intMMID = cLng(Request.QueryString("mm"))
	else
		closeAndGo("article.asp")
	end if
end if
if Request.QueryString("cmd") <> "" then
	intCmd = chkString(Request.QueryString("cmd"),"sqlstring")
end if
if Request.QueryString("mode") <> "" then
	sMode = chkString(Request.QueryString("mode"),"sqlstring")
end if
%>
<!-- #INCLUDE FILE="lang/en/article_lang.asp" -->
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #INCLUDE FILE="Modules/articles/article_functions.asp" -->
<% 
	
  strSQL = mod_singleItemSql(item_tbl)
  strSQL = strSQL & "WHERE (((ARTICLE.ARTICLE_ID)=" & articleid & ") AND ((ARTICLE.ACTIVE)=1));"
  'strSQL = "SELECT ARTICLE_ID, TITLE, POST_DATE, CONTENT, POSTER, KEYWORD, HIT, CATEGORY, PARENT_ID, AUTHOR, AUTHOR_EMAIL, SHOW, HIT from ARTICLE where show = 1 and ARTICLE_ID = " & articleid & ""

  set rs = my_Conn.Execute(strSQL)
  if not rs.eof then
	':: Populate variables
	catid = rs(sMCPre & "CAT_ID")
	catname = rs("CAT_NAME")
	scatid = rs("SUBCAT_ID")
	scatname = rs("SUBCAT_NAME")
	title = rs("TITLE")
    strPoster = rs("POSTER")
    strPosterEmail = rs("POSTER_EMAIL")
    strPostDate = ChkDate2(rs("POST_DATE"))
    strUpdated = rs("UPDATED")
    strPosterURL = rs("TDATA1")
    strSource = rs("TDATA2")
    strSourceURL = rs("TDATA3")
    strSourceContact = rs("TDATA4")
    intHit = rs("Hit")
    intVotes = rs("VOTES")
    intRating = rs("RATING")
	hp = rs("FEATURED")
    if strUpdated <> "0" then
	  strUpdated = ChkDate2(rs("UPDATED"))
    end if
	
	':: search engine stuff
	PageTitle = chkString(rs("TITLE"),"display")
	addToMeta "NAME","Description",replace(replace(rs("SUMMARY"),"""",""),"<br />","")
	addToMeta "NAME","Keywords",replace(rs("KEYWORD"),"""","")
  end if
  cont = 0
  bLeft = false
  bMaint = false
  bMainb = false
  bRight = false
  bShoRight = true
  
cpSQL = "select * from " & strTablePrefix & "PAGES where p_iname = 'article_read'"
set rsCPs = my_Conn.execute(cpSQL)
if not rsCPs.eof then
  pgtitle = rsCPs("p_title")
  pgname = rsCPs("p_name")
	  
  'm_title = rsCPs("P_META_TITLE")
  'addToMeta "NAME","Description",rsCPs("P_META_DESC")
  'addToMeta "NAME","Keywords",rsCPs("P_META_KEY")
  addToMeta "HTTP-EQUIV","Expires",rsCPs("P_META_EXPIRES")
  addToMeta "NAME","Rating",rsCPs("P_META_RATING")
  addToMeta "NAME","Distribution",rsCPs("P_META_DIST")
  addToMeta "NAME","Robots",rsCPs("P_META_ROBOTS")
	  
  left_Col = rsCPs("p_leftcol")
  maint_Col = rsCPs("p_maintop")
  mainb_Col = rsCPs("p_mainbottom")
  right_Col = rsCPs("p_rightcol")
  
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
end if
set rsCPs = nothing
  
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = "An Article"
	PageAction = "Reading<br />"
	PageLocation = "article_read.asp?item=" & articleid
	CurPageInfo = PageAction&"<a href="&PageLocation&">"&PageName&"</a>"
end function
%>
<!-- #INCLUDE FILE="inc_top.asp" --> 
<% 
  setAppPerms "article","iName"

  if rs.eof then
	  ':: article not found
	  intCmd = 90
  else
    setPermVars rs,2
    if not bSCatRead then
	  closeAndGo("article.asp")
    end if
	
	  ':: check and increment item hit count
	  if readCookie("articleid") <> "" then
	    lastid=cLng(readCookie("articleid"))
	  end if
	  if lastid <> "" and not isnumeric(lastid) then
	    closeAndGo("article.asp")
	  end if
	  if lastid <> articleid then
	    executeThis("UPDATE ARTICLE SET HIT = HIT + 1 Where ARTICLE_ID =" & articleid)
		setCookie "articleid",articleid,7
	  end If
	  
	  ':: get item comment count
	  sSQL = "SELECT count(*) as Comments FROM " & strTablePrefix & "M_RATING"
	  sSQL = sSQL & " WHERE COMMENTS NOT LIKE ' '"
	  sSQL = sSQL & " AND ITEM_ID = " & articleid & ""
	  sSQL = sSQL & " AND APP_ID=" & intAppID & ""
	  set rsA = my_Conn.execute(sSQL)
	  if not rsA.eof then
		Comments = rsA("Comments")
	  else
	    Comments = "0"
	  end if
	  set rsA = nothing
	  
	  if hp = 0 then
	    hp = 1
	    sTxt = "Make this a 'Featured Article'"
	    sImg = icnFeature
	  else
	    hp = 2
	    sTxt = "Remove from 'Featured Articles'"
	    sImg = icnUnfeature
	  end if
  end if
 %>
  <script type="text/javascript">
  function jsDelArt(nam,s){
	var stM
	stM = "This will delete the Article:\n\n";
	stM += ""+nam+"\n";
	stM += "\nRemember, this cannot be undone!\n";
	var del=confirm(stM);
	if (del==true){
	  window.location="<%= sScript %>?cmd=24&item="+s;
	}else{
	  return;
	}
  }
  </script><% 
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if bLeft then
    response.Write("<td class=""leftPgCol"" valign=""top"" nowrap=""nowrap"">")
	intSkin = getSkin(intSubSkin,1)
	  cStart = timer
	 shoBlocks(l_col)
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
  end if

  response.Write("<td class=""mainPgCol"" valign=""top"">")
  cStart = timer
  intSkin = getSkin(intSubSkin,2)
  
  arg1 = "Articles|article.asp"
  if intCmd = 23 then
    arg2 = txtEdit & "|"
  elseif intCmd = 90 then
    ':: nothing
  else
    arg2 = catname & "|article.asp?title=" & server.URLEncode(catname) & "&amp;cmd=1&amp;cid=" & catid
    arg3 = scatname & "|article.asp?title=" & server.URLEncode(scatname) & "&amp;cmd=2&amp;cid=" & catid & "&amp;sid=" & scatid
  end if
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if bMaint then
	 shoBlocks(mt_col)
  end if

':: start main content
  select case intCmd
    case 23 ':: edit item
      spThemeTitle = "Edit Article"
	  spThemeBlock1_open(intSkin)
	  if sMode = 322 then
	    processArticleEditForm()
	  else
	    editArticle()
	  end if
	case 24 '::delete item
	  Call deleteArticle(articleID)
	  set rs = nothing
	  closeAndGo("article.asp")
	case 25 '::delete comment
	  Call mod_deleteComment(item_tbl,item_fld)
	  set rs = nothing
	  closeAndGo("article_read.asp?item=" & articleID)
	case 90 ':: article not found
      spThemeTitle = "&nbsp;"
	  spThemeBlock1_open(intSkin)
	  articleNotFound()
	case else
	  mainArticle()
  end select
  spThemeBlock1_close(intSkin)
':: end main content

  if bMainb then
	 shoBlocks(mb_col)
  end if
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
	
  if bRight and bShoRight then
    if cont = 3 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	  cStart = timer
	shoBlocks(r_col)
	  if shoBlkTimer then
	  blkLoadTime = formatnumber((timer - cStart),3)
	  response.Write(blkLoadTime)
	  end if
    response.Write("</td>")
  end if
  response.Write("</tr></table>")
  app_Footer()
  %>
<!-- #INCLUDE FILE="inc_footer.asp" -->
<%
set rs = nothing 

sub editArticle()
  Response.Write("Edit Article")
end sub

sub processArticleEditForm()
  Response.Write("Process Edit Article Form")
end sub

sub articleNotFound()
  sMsg = "&nbsp;</p><p>Article not found</p><p>&nbsp;"
  showMsgBlock 0,sMsg
end sub

function GetRating(ArticleID)
	strSQL = "SELECT VOTES, RATING FROM ARTICLE WHERE ARTICLE_ID = " & ArticleID
	set rsArticleRating = server.CreateObject("adodb.recordset")
	rsArticleRating.Open strSQL, my_Conn
		
	dim intVotes
	dim intRating
	intVotes = rsArticleRating("VOTES")
	intRating = rsArticleRating("RATING")
	rsArticleRating.Close
	set rsArticleRating = nothing
												
	if intVotes > 0 then
	intRating = Round(intRating/intVotes)%>
	Rating:<% =intRating%>&nbsp; Votes:<% =intVotes%> (Rating Scale: <span class="fAlert">1</span> = worst, <span class="fAlert">10</span> = best)
<%else%>
	Rating:<% =intRating%>&nbsp; Votes:0 (Rating Scale: <span class="fAlert">1</span> = worst, <span class="fAlert">10</span> = best)
<%end if
end function%>