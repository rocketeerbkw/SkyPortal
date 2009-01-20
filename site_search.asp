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
Server.ScriptTimeout = 200
Response.Buffer=true
PageTitle = txtSitSrch
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = txtSitSrch
	PageLocation = "site_search.asp"
	CurPageInfo = "<a href=""" & PageLocation & """>" & PageName & "</a>"

end function
%>
<!-- #include file="inc_functions.asp" -->
<!-- #include file="inc_top.asp" -->
<%
'search = ChkString(Request("search"), "SQLString")
search = chkString(Request("search"),"sqlstring")
qsearch=replace(search," ","+")
'show = chkString(Request("num"),"sqlstring")
show = 10
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr>
<td width="190" class="leftPgCol" valign="top">
<% 
  intSkin = getSkin(intSubSkin,1)
  menu_fp() 
%></td>
<td width="100%" class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtSitSrch & "|site_search.asp"
  arg2 = txtSrchRslts & " : " & search & "|javascript:;"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

spThemeBlock1_open(intSkin)
Response.Write("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
response.write "<tr><td align=""left""><br />"
 
Dim iPageSize       
Dim iPageCount      
Dim iPageCurrent    
Dim strOrderBy      
Dim strSQL          
Dim objPagingConn   
Dim objPagingRS     
Dim iRecordsShown   
Dim I

'############################## Forum Search ####################
If chkApp("forums","USERS") Then
  strSQL = "select FORUM_ID from Portal_Topics where T_Subject like '%" & search & "%' or T_Message like '%" & search & "%' and T_Status=1" 

Set objPagingRS = Server.CreateObject("ADODB.Recordset")
objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
  reccount = 0
  if not objPagingRS.eof then
    do until objPagingRS.eof
      if chkForumAccess(strUserMemberID,objPagingRS("FORUM_ID")) then
	    reccount = reccount + 1
	  end if
	  objPagingRS.movenext
	loop
  end if

    objPagingRS.Close
    Set objPagingRS = Nothing

    strSQL = "select FORUM_ID from Portal_Reply where R_Message like'%" & search & "%' order by Reply_ID" 

    Set objPagingRS = Server.CreateObject("ADODB.Recordset")
	objPagingRS.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  if not objPagingRS.eof then
    do until objPagingRS.eof
      if chkForumAccess(strUserMemberID,objPagingRS("FORUM_ID")) then
	    reccount = reccount + 1
	  end if
	  objPagingRS.movenext
	loop
  end if
  objPagingRS.Close
  Set objPagingRS = Nothing
    'reccount = reccount + objPagingRS.recordcount
	%>
<center><span class="fTitle"><b><%= txtForums %> - <%= txtSrchRslts %> : "<%=search%>" <%= txtFound %>&nbsp;<%=reccount%>&nbsp;<%= txtSitems %></b></span></center>
<% If reccount > 0 Then %>	
<center><a href="forum_search.asp?mode=DoIt&search=<%=qsearch%>&searchdate=0&Searchmember=0&SearchMessage=0&andor=phrase&forum=0">
<%= txtVSrchRslts %></a></center>
<br />
<% end if

response.Write("<hr />")
response.Flush()
end if
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::

moduleSiteSearch()

response.write "</td></tr>"
Response.Write("</table>")
spThemeBlock1_close(intSkin)
%>
<br />
</td>
</tr>
</table>
<!-- #include file="inc_footer.asp" -->
<% 
sub moduleSiteSearch()
	sSql = "SELECT * FROM " & strTablePrefix & "MODS WHERE M_CODE='siteSrch'"
	set rsP = my_Conn.execute(sSql)
	if not rsP.eof then
	  do until rsP.eof
	    execute("Call " & rsP("M_VALUE"))
	    rsP.movenext
	  loop
	end if
	set rsP = nothing
end sub
%>