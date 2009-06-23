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

CurPageType = "forums"
CurPageInfoChk = "1"
%><!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" --><%
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = txtNewsArch
	PageAction = txtViewing & "<br />" 
	PageLocation = "fnews.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<!-- #INCLUDE file="includes/inc_ADOVBS.asp" -->
<table cellpadding="0" cellspacing="0" width="100%">
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
  arg1 = txtNewsArch & "|fnews.asp"
  arg2 = ""
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

Set NobjDict =   CreateObject("Scripting.Dictionary")
iPageSize = 15
If Request.QueryString("page") = "" Then
	iPageCurrent = 1
Else
	iPageCurrent = CInt(Request.QueryString("page"))
End If

    strSQL      =   "SELECT m_code, m_value FROM " & strTablePrefix & "mods WHERE m_name = 'news';"
    set NobjRec  =   my_conn.Execute(strSQL)

    while not NobjRec.EOF    
        NobjDict.Add NobjRec.Fields.Item("m_code").Value, NobjRec.Fields.Item("m_value").Value
        NobjRec.moveNext
    wend     
    
	NslLength    	=   cint(NobjDict.Item("slLength"))
	NslSort      	=   cint(NobjDict.Item("slSort"))
	NslEncode		=	cint(NobjDict.Item("slEncode"))
	NstrIMGInPosts	=	cint(NobjDict.Item("slImages"))
	NstrIcons		=	0

    set NobjDict 	=   nothing

    strSQL      	=  "SELECT " & strTablePrefix & "TOPICS.TOPIC_ID, " & _
	strTablePrefix & "TOPICS.T_SUBJECT, " & _
	strTablePrefix & "TOPICS.T_AUTHOR, " & _
	strTablePrefix & "MEMBERS.M_NAME, " & _
	strTablePrefix & "TOPICS.T_REPLIES, " & _
	strTablePrefix & "TOPICS.T_DATE, " & _
	strTablePrefix & "TOPICS.T_MESSAGE " & _
	"FROM " & strTablePrefix & "TOPICS, " & _
	strTablePrefix & "FORUM, " & _
	strMemberTablePrefix & "MEMBERS " & _
	"WHERE " & strTablePrefix & "FORUM.F_PRIVATEFORUMS = 0 AND " & _
	strTablePrefix & "TOPICS.FORUM_ID = " & strTablePrefix & "FORUM.FORUM_ID AND " & _
	strTablePrefix & "TOPICS.T_NEWS = 1 AND " & _	    				
	strTablePrefix & "TOPICS.T_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
    Select Case NslSort
    Case "2"    '   last post
	    strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"
	Case "3"    '   hot topics
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_REPLIES DESC;"
	Case Else   '   last created
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.TOPIC_ID DESC;"
	End Select
	
Set NobjRec = Server.CreateObject("ADODB.Recordset")
NobjRec.PageSize = iPageSize
NobjRec.CacheSize = iPageSize
NobjRec.Open strSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

reccount = NobjRec.recordcount
iPageCount = NobjRec.PageCount

If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
<%
spThemeTitle= txtNews & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""forum_search.asp?mode=news"">" & txtSrchNews & "</a>"
spThemeCellCustomCode = "colspan=""3"""
spThemeBlock1_open(intSkin)
Response.Write("<table cellpadding=""0"" cellspacing=""0"" width=""100%"">")
If iPageCount = 0 Then
%>
<tr><td class="tCellAlt0"><b><%= txtNoNewsFnd %>.</b></td></tr>
<%
Else
	NobjRec.AbsolutePage = iPageCurrent

 iRecordsShown = 0
    Do While iRecordsShown < iPageSize And Not NobjRec.EOF And NOT NobjRec.EOF
      NT_Subject     =   NobjRec("T_SUBJECT")
      NT_Author      =   NobjRec("T_AUTHOR")
      NM_NAME        =   NobjRec("M_NAME")
      NT_Message     =   NobjRec("T_MESSAGE")
      NT_REPLIES     =   NobjRec("T_REPLIES")
      NT_DATE        =   NobjRec("T_DATE")
      NTOPIC_ID      =   NobjRec("TOPIC_ID")
      

      If Len(NT_Message) > CInt(NslLength) Then  
      	NT_Message = Left(NT_Message, NslLength) & "..."
      Else
      	NT_Message = NT_Message
      End If

	  if NslEncode = 1 then
	    NT_Message = formatStr(NT_MESSAGE)
	  else
	  	NT_Message = HTMLencode(NT_MESSAGE)
	  end if
%> 
<tr>
<td width="65%" class="tCellAlt0"><a href="link.asp?TOPIC_ID=<%= NTOPIC_ID %>"><b><%= NT_SUBJECT%></b></a></td>
<td width="15%" class="tCellAlt0"><span class="fSmall"><%= txtPstdBy %>&nbsp;<b><%= NM_NAME %></b></span></td>
<td width="20%" class="tCellAlt0" align="right"><b><%= ChkDate(NT_DATE) %></b><br /><%= ChkTime(NT_DATE) %></td>
</tr>
<tr>
<td width="100%" colspan="3" class="tCellAlt1"><%= NT_MESSAGE %><br /><br /></td>
</tr>

<tr>
<td width="100%" colspan="3"  class="tCellAlt2">
<table border="0" width="100%" nowrap align="left"><tr>
<td><a href="link.asp?TOPIC_ID=<%= NTOPIC_ID %>"><%= txtReadAll %></a></td>
<td><a href="JavaScript:openWindow5('forum_pop.asp?mode=5&amp;cid=<%= NTOPIC_ID %>')"><img border="0" src="images/icons/print.gif" width="16" height="17" title="<%= txtPrint %>" alt="<%= txtPrint %>"></a></td>
<%if (lcase(strEmail) = "1") then %><td><a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<% =NTOPIC_ID %>')"><img border="0" src="images/icons/icon_email.gif" height="15" width="15" title="<%= txtEmail %>" alt="<%= txtEmail %>"></a></td><%end if%>
<td><a href="link.asp?TOPIC_ID=<%=NTOPIC_ID %>&view=lasttopic"><%= txtLstComment %>&nbsp;(<%= NT_REPLIES %>)</a></td>
</tr>
</table>

</td>
</tr>
<%
 iRecordsShown = iRecordsShown + 1
 NobjRec.MoveNext
 Loop
%> 
	</td>
<%
	NobjRec.close
	Set NobjRec  =  nothing
	iRecordsShown = 0
End If%>
</tr><tr>
<td class="tCellAlt1" colspan="3" align="center">
<%if iPageCurrent <> 1 Then%>
	<a HREF="fnews.asp?page=<%= iPageCurrent - 1 %>"><%= txtPrevious %>&nbsp;<% =iPageSize%>&nbsp;<%= txtTopics %></a>&nbsp;|&nbsp;
<%End If%>
 <%= txtPage %>&nbsp;<b><%= iPageCurrent %></b>&nbsp;<%= txtof %>&nbsp;<b><%= iPageCount %></b>
<%If iPageCurrent < iPageCount Then%>
	&nbsp;|&nbsp;<a HREF="fnews.asp?page=<%= iPageCurrent + 1 %>"><%= txtNext %>&nbsp;<% =iPageSize%>&nbsp;<%= txtTopics %></a>
<%End If%>
</td></tr></table>
<%
spThemeBlock1_close(intSkin)%>
</td>
  </tr>
</table>
    </td>
</tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
