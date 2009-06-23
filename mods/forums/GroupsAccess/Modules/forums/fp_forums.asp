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
%><!-- include file="lang/en/forum_core.asp" --><%
NstrIMGInPosts = 0

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Forum NEWS large
' :::::::::::::::::::::::::::::::::::::::::::::::
Function cleanfrontnews(fstring)
' New function by Hawk92 - source code box mod - 11-2004 version 1.5
' Provides processing of messages containing code to display in last topics block
ptr=InStr(1,fString,"[@@]",1)
ptr2=InStr(ptr+4,fString,"[@@]",1)+4
ptr3=InStr(ptr2,fString,"[/@@]",1)
ptr4=InStr(ptr3+5,fString,"[/@@]",1)+5
plen=len(fstring)-ptr4+1
If ptr>0 then
fString=Mid(fString,1,ptr-1)& " [...Code Snippet ...]"& Mid(fString,ptr4,plen)
End if
cleanfrontnews=fString
End Function

function cntReportedPosts()
  cntReportedPosts = "&nbsp;(" & getCount("R_STATUS",strTablePrefix & "REPORTED_POST","R_STATUS=0") & ")"
end function

'cntActiveTopics = cntActiveTopics2()
'cntActiveTopics = ""
function cntActiveTopics()
  cntActiveTopics = "&nbsp;"
end function

function cntActiveTopics2()
  ':::::::::::::::: get Active topics count :::::::::::::::::::::::
  tmpAT = 0
  if not IsNull(Session(strUniqueID & "last_here_date")) then 
	'get counts from all forums
	strSql = "SELECT " & strTablePrefix & "TOPICS.T_LAST_POST, " & strTablePrefix & "TOPICS.FORUM_ID"
 	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
 	strSql = strSql & "WHERE (((" & strTablePrefix & "TOPICS.T_LAST_POST)>'"& Session(strUniqueID & "last_here_date") & "'))"
 	set rs = my_Conn.Execute(strSql)
 	if not rs.EOF then
	  do until rs.EOF
		if chkForumAccess(strUserMemberID,rs("FORUM_ID")) then
 		tmpAT = tmpAT + 1
		end if
		rs.movenext
	  loop
 	else
 		tmpAT = 0
 	end if
  end if
  set rs = nothing
  cntActiveTopics2 = "&nbsp;(" & tmpAT & ")"
end function


function f_news_fp()
  if chkApp("forums","USERS") then
	Set NobjRec  =   Server.CreateObject("ADODB.RecordSet")
    Set NobjDict =   CreateObject("Scripting.Dictionary")   

    strSQL = "SELECT m_code, m_value FROM " & strTablePrefix & "mods WHERE m_name = 'news';"
    set NobjRec  =   my_conn.Execute(strSQL)

    while not NobjRec.EOF    
        NobjDict.Add NobjRec.Fields.Item("m_code").Value, NobjRec.Fields.Item("m_value").Value
        NobjRec.moveNext
    wend     
    
    NslPosts = cint(NobjDict.Item("slPosts"))
	NstrColumns	= cint(NobjDict.Item("slColumns"))
    NslLength = cint(NobjDict.Item("slLength"))
    NslSort = cint(NobjDict.Item("slSort"))
	NslEncode = cint(NobjDict.Item("slEncode"))
	NstrIMGInPosts = cint(NobjDict.Item("slImages"))
	NstrDefimg = chkString(NobjDict.Item("slDefimg"),"displayimage")
	NstrIcons = 0
    set NobjDict = nothing

    strSQL = "SELECT TOP " & NslPosts & " " & strTablePrefix & "TOPICS.TOPIC_ID, " & _
    	strTablePrefix & "TOPICS.CAT_ID, " & _
    	strTablePrefix & "TOPICS.FORUM_ID, " & _
    	strTablePrefix & "TOPICS.T_SUBJECT, " & _
	    strTablePrefix & "TOPICS.T_AUTHOR, " & _
    	strTablePrefix & "MEMBERS.M_NAME, " & _
		strTablePrefix & "TOPICS.T_REPLIES, " & _
		strTablePrefix & "TOPICS.T_DATE, " & _
		strTablePrefix & "TOPICS.T_MESSAGE " & _
		"FROM " & strTablePrefix & "TOPICS, " & _
		strTablePrefix & "FORUM, " & _
		strMemberTablePrefix & "MEMBERS " & _
		"WHERE " & strTablePrefix & "FORUM.F_PRIVATEFORUMS >= 0 AND " & _
		strTablePrefix & "TOPICS.FORUM_ID = " & strTablePrefix & "FORUM.FORUM_ID AND " & _
        strTablePrefix & "TOPICS.T_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
	strSQL = strSQL & " AND " & strTablePrefix & "TOPICS.T_NEWS = 1 "

	' ***** We will use this values later on
	ShowEntireTopic = False ' i think it's obvious
	' ***************************************

    Select Case NslSort
    Case "2"    '   last post
	    strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

	Case "3"    '   hot topics
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_REPLIES DESC;"
					
	Case Else   '   last created
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.TOPIC_ID DESC;"

	End Select
    
    set NobjRec = my_Conn.Execute(strSql)

spThemeTitle="News"
'spThemeTitle=spThemeTitle&" [" & intSkin & "]"
spThemeBlock1_open(intSkin)

%><table cellspacing="3" cellpadding="0" border="0" bordercolor="#FF0000" width="100%">
<tr><%

	MyRecordTotal = 0	' the total number of records
	MyRecordcount = 1	' the record i am working with
	MyStartingRows = 0  ' How many single row news it will be displayed before the multiple column layout
	MyColumns = NstrColumns		' How many columns per row
	MyColumnsCount = 0
	
	if MyColumns = 2 then
	  colSpan = " colspan=""3"""
	else
	  colSpan = ""
	end if
%> <td width="<%=100/MyColumns%>%" valign="top">
 <% 
	while not NobjRec.EOF
		MyRecordTotal = MyRecordTotal + 1
		NobjRec.MoveNext
	wend
	if not (NobjRec.EOF and NobjRec.BOF) then NobjRec.Movefirst

	While NOT NobjRec.EOF
      NT_Subject     =   ChkString(NobjRec("T_SUBJECT"),"display")
      NT_Author      =   NobjRec("T_AUTHOR")
      NM_NAME        =   NobjRec("M_NAME")
      NT_Message     =   NobjRec("T_MESSAGE")
      NT_REPLIES     =   NobjRec("T_REPLIES")
      NT_DATE        =   NobjRec("T_DATE")
      NTOPIC_ID      =   NobjRec("TOPIC_ID")
      NFORUM_ID      =   NobjRec("FORUM_ID")
      NCATEG_ID      =   NobjRec("CAT_ID")
      
	  ' we will only parse the string if we don't want to show the entire topic
      If Len(NT_Message) > CInt(NslLength) Then  
        NT_Message=cleanfrontnews(NT_Message)
		NT_Message = replace(NT_MESSAGE, strHomeUrl, "", 1, -1, 1)
		'NT_Message = trimNews(NT_Message,NslLength) & "..."
		NT_Message = left(NT_Message,NslLength) & " ..."
      Else
      	'NT_Message = trimNews(NT_Message,len(NT_Message))
      End If
      if NslEncode = 1 then
	    if strAllowHtml <> 1 then
	      NT_Message = formatStr(NT_MESSAGE)
		else
		  'NT_Message = replace(NT_MESSAGE, strHomeUrl, "", 1, -1, 1)
		end if
	  else
	  	NT_Message = HTMLencode(NT_MESSAGE)
	  end if

      IF Left( NT_Message, 1 ) = " " Then NT_Message = Right( NT_Message, Len( NT_Message ) - 1 ) 
	  %>
	
	<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%">
	<tr><td class="tSubTitle">
	<div><a href="link.asp?TOPIC_ID=<%= NTOPIC_ID %>"><b><%= NT_SUBJECT%></b></a></div>
	<div style="text-align:left;"><span class="fSmall"><b><%= ChkDate(NT_DATE) %></b></span></div></td></tr>
	<tr><td width="100%" valign="top"><div align="justify"><%= SetImgLink( NT_MESSAGE, NTOPIC_ID, NT_SUBJECT ) %></div></td></tr>
	<tr><td width="100%" class="tCellAlt2">
	
	<table border="0" cellspacing="1" cellpadding="0" align="left" width="100%"><tr>
	<td>
	
	<table width="100%" border="0" cellpadding="4" cellspacing="0"><tr><td width="50%" align="left"><a href="link.asp?TOPIC_ID=<%= NTOPIC_ID %>">&nbsp;&nbsp;Read News</a></td><td width="50%" align="right" nowrap><a href="link.asp?view=lasttopic&amp;TOPIC_ID=<%=NTOPIC_ID %>">Last Reply (<%= NT_REPLIES %>)</a></td></tr></table>
	
	</td>
	<td align="right"><a href="JavaScript:openWindow5('forum_pop.asp?mode=5&amp;cid=<%= NTOPIC_ID %>')"><img border="0" src="images/icons/print.gif" width="16" height="17" alt="Print News" title="Print News" /></a></td>
	<td align="center" width="25"><%if (lcase(strEmail) = "1") then %><a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<% =NTOPIC_ID %>')"><img border="0" src="images/icons/icon_email.gif" height="15" width="15" title="Send this news to a friend" alt="Send this news to a friend" /></a><%end if%>
	</td></tr></table>
	
	</td></tr></table>
<%
	NobjRec.MoveNext()

	MyColumnsCount = MyColumnsCount + 1
	If MyColumnsCount = MyColumns and not NobjRec.eof then
		MyColumnsCount = 0
		Response.write "</td></tr><tr><td align=""center"" style=""height:1px;""" & colSpan & "><img src=""themes/" & strTheme & "/line.gif"" height=""1"" width=""98%"" alt="""" /></td></tr>"
		'Response.write "<tr><td colspan=""" & MyColumns + 10 & """ align=""center"" style=""background:url(themes/" & strTheme & "/dot_light.gif);""><img src=""images/spacer.gif"" height=""1"" alt="""" /></td></tr>" 
		if MyRecordCount < MyRecordTotal then
		  response.Write("<tr><td width=""" & 100/MyColumns & "%"" valign=""top"">")
		else
		  'response.Write("<tr><td" & colSpan & " valign=""top"">")
		end if
	elseif MyColumnsCount = MyColumns and NobjRec.eof then
	  	Response.write "</td></tr>"
	elseif MyColumnsCount < MyColumns and NobjRec.eof then
			Response.write "</td><td style=""background:url(themes/" & strTheme & "/line.gif);""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td><td width=""" & 100/MyColumns & "%"" valign=""top"">"
	  	Response.write "</td></tr>"
	
	elseif MyColumnsCount < MyColumns and not NobjRec.eof then
		if MyRecordCount <> MyRecordTotal then
			Response.write "</td><td style=""background:url(themes/" & strTheme & "/line.gif);""><img src=""images/spacer.gif"" width=""1"" alt="""" /></td><td width=""" & 100/MyColumns & "%"" valign=""top"">"
		else
			Response.write "</td></tr>"
		end if
	End IF
	if NobjRec.eof then
	  'Response.write "</td></tr>"
	end if
	MyRecordCount = MyRecordCount + 1
Wend
%> 
<% If NOT ShowEntireTopic Then %>
<tr><td width="100%"<%=colSpan%> align="center"><span class="fSmall"><a href="fnews.asp">Archived News</a> - <a href="forum_search.asp?mode=news">Search News</a></span></td></tr>
<% End IF %>
</table>
<%
spThemeBlock1_close(intSkin)
NobjRec.close
Set NobjRec  =  nothing
end if
end function

' News parser function: check if news contains any picture in the begining 
'                       and set that image as LINK to the complete post, 
'                       with the ALT tag matching the post subject
Function SetImgLink (Message, TOPIC_ID, SUBJECT) ' News Parser Function
	Message = replace(Message,"</p><p>","<br /><br />")
	Message = replace(Message,"<p>","")
	Message = replace(Message,"<br />","<br />")
	Message = replace(Message,"<hr />","<hr />")
	TempMessage = Message
'		 If DefaultImage <> "" Then
		 If NstrIMGInPosts = 1 Then 
          TempMessage = "<img align=""left"" src=""themes/" & strTheme & "/news_fp.gif""  border=""0"" hspace=""4"" vspace=""0"" title=""" & SUBJECT & """ alt=""" & SUBJECT & """ />" & Message
         end if
	SetImgLink = TempMessage	
End Function

function trimnews(messg,maxchars)
 strtag1 = "<"
 strtag2 = ">"
 tempmessage = messg
 nmsglen = len(tempmessage)
 ntrimlen = maxchars ' default
 
 if nmsglen > maxchars then ' too long so we have to trim
  ntag1pos = instr(1, tempmessage, strtag1, 1) ' strtag1 character
  ntag2pos = 0 ' strtag2 character - start at 0
  if ntag1pos > 0 then ' there's a  "<" characters, so find where to cut.
   for i = 1 to maxchars  ' capture the last "<" and ">" positions within the maxchars
    if mid(tempmessage,i,1) = strtag1 then
     ntag1pos=i
    end if
    if mid(tempmessage,i,1) = strtag2 then
     ntag2pos=i
    end if
   next
   if (ntag2pos = 0 ) or (ntag1pos > ntag2pos) then ' within maxchars, there's a "<" but not a ">"
    nnewendpos = instr(ntag1pos, tempmessage, strtag2,1)
    if nnewendpos = 0 then  ' there's a < in the message with no ">" - probably not a tag.
     ntrimlen = maxchars
    else
     ntrimlen = nnewendpos  ' end right after the trailing tag
    end if
   end if
  end if
 end if
 
 tempmessage = left(tempmessage, ntrimlen)
 trimnews = tempmessage
 
end function

function trimNews2(messg,maxChars)
	TempMessage = ""
	postmsg = ""
	wh = 1
	ncnt = 0
  if InStr(1,messg , "<img", 1 ) > 0 then
	do until ncnt >= maxChars
		if InStr( wh , messg , "<img", 1 ) > 0 then
			i = InStr( wh , messg , "<img", 1 ) 
			TempMessage = Left( messg, i -1 )
			ncnt = ncnt + (i - wh)
			For j = i to Len( messg )
				If Mid( messg , j , 1 ) <> ">" then
					TempMessage = TempMessage & Mid( messg , j , 1 )
				else
					wh = j + 1
					TempMessage = TempMessage & Mid( messg , j , 1 )
					exit for
				end if
			Next
  		else
			TempMessage = TempMessage & Mid( messg, wh, maxChars - ncnt )
			ncnt = len(TempMessage)
		end if
	loop
  else
	TempMessage = Left( Message, maxChars )
	ncnt = len(TempMessage)
  end if
  trimNews = TempMessage
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Forum recent topics small
' :::::::::::::::::::::::::::::::::::::::::::::::
Function cleanlasttopic(fstring)
' New function by Hawk92 - source code box mod - 11-2004 version 1.5
' Provides processing of messages containing code to display in last topics block

ptr=InStr(1,fString,"[@@]",1)
If ptr>0 then
fString=Mid(fString,1,ptr-1)& " ...Code Snippet ..."
End if
cleanlasttopic=fString
End Function
' end of new Function
incForumFp= true

function f_topics_sm()
  if chkApp("forums","USERS") then
  if not strForumStatus = "down" then

    Set objRec  =   Server.CreateObject("ADODB.RecordSet")
    Set objDict =   CreateObject("Scripting.Dictionary")   


    strSQL      =   "SELECT m_code, m_value FROM " & strTablePrefix & "mods WHERE m_name = 'slash';"
    set objRec  =   my_conn.Execute(strSQL)

    while not objRec.EOF    
        objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
        objRec.moveNext
    wend     
    
    slPosts     	=   cint(objDict.Item("slPosts"))
    slLength    	=   cint(objDict.Item("slLength"))
    slSort      	=   cint(objDict.Item("slSort"))
	slEncode		=	cint(objDict.Item("slEncode"))
	strIMGInPosts	=	cint(objDict.Item("slImages"))
	strIcons		=	0

    set objDict 	=   nothing

    strSQL      	=   "SELECT TOP " & slPosts & " " & strTablePrefix & "TOPICS.TOPIC_ID, " & _
				strTablePrefix & "TOPICS.T_SUBJECT, " & _
		                strTablePrefix & "TOPICS.T_AUTHOR, " & _
				strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, " & _
	    	                strTablePrefix & "MEMBERS.M_NAME, " & _
				strTablePrefix & "TOPICS.T_REPLIES, " & _
				strTablePrefix & "TOPICS.T_DATE, " & _
				strTablePrefix & "TOPICS.T_LAST_POST, " & _
		   		strTablePrefix & "TOPICS.T_MESSAGE, " & _
				strTablePrefix & "TOPICS.T_POLL "	& _					    		
			    "FROM " & strTablePrefix & "TOPICS, " & _
				strTablePrefix & "FORUM, " & _
				strMemberTablePrefix & "MEMBERS " & _
			    	"WHERE " & strTablePrefix & "FORUM.F_PRIVATEFORUMS = 0 AND " & _
	    			strTablePrefix & "TOPICS.FORUM_ID = " & strTablePrefix & "FORUM.FORUM_ID AND NOT " & _
				strTablePrefix & "TOPICS.T_STATUS = 0 AND " & _
				strTablePrefix & "TOPICS.T_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "

    Select Case slSort
    Case "2"    '   last post
	    strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"
		DTString = "Last [slPosts] topics"
	Case "3"    '   hot topics
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.T_REPLIES DESC;"
		DTString = "Top [slPosts] hottest"						
	Case Else   '   last created
		strSQL = strSQL & "ORDER BY " & strTablePrefix & "TOPICS.TOPIC_ID DESC;"
		DTString = "Last [slPosts] topics"
	End Select
    
    DTString = Replace(DTString,"[slPosts]", slPosts)
    DTString = Replace(DTString,"[ForumName]", strSiteTitle)

    set objRec = my_Conn.Execute(strSql)
	
spThemeMM = "lstTopic_sm"
spThemeTitle= DTString
spThemeBlock1_open(intSkin)
response.Write("<table width=""100%"" cellpadding=""0"" cellspacing=""0"">")

    While NOT objRec.EOF
      T_Subject     =   ChkString(objRec("T_SUBJECT"),"display")
      T_Author      =   objRec("T_AUTHOR")
      T_LastAuthor  =   objRec("T_LAST_POST_AUTHOR")
      M_NAME        =   objRec("M_NAME")
      T_Message     =   objRec("T_MESSAGE")
      T_REPLIES     =   objRec("T_REPLIES")
      T_DATE        =   objRec("T_DATE")
      TOPIC_ID      =   objRec("TOPIC_ID")
      T_LAST_POST   =   objRec("T_LAST_POST")

	  'if slEncode = 1 then
	  	'T_Message = SlashCode(T_MESSAGE)
	  'else	  		  
	  	T_Message = HTMLencode(T_MESSAGE)
	  'end if
	  T_Message = cleanlasttopic(T_Message) 
	  T_Message   =   replace(T_Message, "/"," ", 1, -1, 1)
	  T_Message   =   replace(T_Message, "_"," ", 1, -1, 1)
	  T_Message   =   replace(T_Message, "*"," ", 1, -1, 1)
	  T_Message   =   replace(T_Message, "^"," ", 1, -1, 1)
	  T_Message   =   replace(T_Message, "-"," ", 1, -1, 1)
	  T_Message   =   replace(T_Message, "#"," ", 1, -1, 1)
  	  T_Message   =   replace(T_Message, "%"," ", 1, -1, 1)
  	  T_Message   = chkBadWords(T_Message)

	if instr(T_SUBJECT, " ") = 0 and Len(T_SUBJECT) > 18 then
		T_SUBJECT = Left(T_SUBJECT, 15) & "..."
	end if
      If Len(T_Message) > CInt(slLength) Then  
      	T_Message = Left(T_Message, slLength) & "..."
      Else
      	T_Message = T_Message
      End If	  
%> 
<tr><td><a href="link.asp?TOPIC_ID=<%= TOPIC_ID %>"><b><%= T_SUBJECT%></b></a><% if objRec("T_POLL") <> 0 then %>&nbsp;<img src="images/icons/icon_topic_poll.gif" alt="" /><% end if %></td></tr>
<tr><td>Posted by <b><%= M_NAME %></b><%if T_REPLIES <> 0 then%><br />&nbsp;Last replied by <b><%=getMemberName(T_LastAuthor)%></b><%end if%><%if T_REPLIES <> 0 then%><br />&nbsp;on <b><%= ChkDate(T_LAST_POST) %></b> @ <%= ChkTime(T_LAST_POST) %><%else%><br />&nbsp;on <b><%= ChkDate(T_DATE) %></b> @ <%= ChkTime(T_DATE) %><%end if%></td></tr>
<%If not CInt(slLength) = 0 then%>
<tr><td><%= T_MESSAGE %></td></tr>
<%end if%>
<tr><td><%if not T_REPLIES = 0 then%>[<a href="link.asp?view=lasttopic&amp;TOPIC_ID=<%= TOPIC_ID %>">Last Reply (<%= T_REPLIES %>)</a>]<%end if%><br /><br /></td></tr>
<%
      objRec.MoveNext()
    Wend
%> 
<%
response.Write("</table>")
spThemeBlock1_close(intSkin)
objRec.close
Set objRec  =  nothing
end if
end if
end function

' :::::::::::::::::::::::::::::::::::::::::::::::
' :::		Forum FEATURED POLL
' :::::::::::::::::::::::::::::::::::::::::::::::
function f_polls_fp()
  if chkApp("forums","USERS") then

if strFeaturedPoll <> 0 then
Dim PollError
PollError = False

	strSql = "SELECT  " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "FORUM.F_SUBJECT "
	strSql = strSql & " FROM  " & strTablePrefix & "TOPICS , " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_POLL = " & strFeaturedPoll
	strSql = strSql & " AND   " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID"

	set rsPoll = my_Conn.Execute (strSql)
	if (rsPoll.EOF or rsPoll.BOF) then
		PollError = True
	else
		strRqTopicID = rsPoll("TOPIC_ID")
		strRqForumID = rsPoll("FORUM_ID")
		strRqCatID = rsPoll("CAT_ID")
		strRqTopic_Title = replace(rsPoll("T_SUBJECT"), "#", "")
		strRqForum_Title = replace(rsPoll("F_SUBJECT"), "#", "")
	end if
	
	rsPoll.Close
	set rsPoll = nothing

	strSql = "SELECT POLL_TYPE, POLL_ID, POLL_ALLOW, POLL_QUESTION," 
        strSql = strSql & " ANSWER1, ANSWER2, ANSWER3, ANSWER4, ANSWER5, ANSWER6, ANSWER7, ANSWER8, ANSWER9, ANSWER10, ANSWER11, ANSWER12,"
        strSql = strSql & " RESULT1, RESULT2, RESULT3, RESULT4, RESULT5, RESULT6, RESULT7, RESULT8, RESULT9, RESULT10, RESULT11, RESULT12,"
        strSql = strSql & " POST_DATE, END_DATE, POLL_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "POLLS "
	strSql = strSql & " WHERE POLL_ID = " & strFeaturedPoll

	set rs = my_Conn.Execute (strSql)

	if not(rs.eof or rs.bof) then

		if rs("POLL_TYPE") = "0" then
		strPollType = 0
		else
		strPollType = 1
		end if
		strPoll_ID = rs("POLL_ID")
		strPollAllow = rs("POLL_ALLOW")
		strPollQuestion = rs("POLL_QUESTION")
		
		strPollAns1 = rs("ANSWER1")
		strPollAns2 = rs("ANSWER2")
		strPollAns3 = rs("ANSWER3")
		strPollAns4 = rs("ANSWER4")
		strPollAns5 = rs("ANSWER5")
		strPollAns6 = rs("ANSWER6")
		strPollAns7 = rs("ANSWER7")
		strPollAns8 = rs("ANSWER8")
		strPollAns9 = rs("ANSWER9")
		strPollAns10 = rs("ANSWER10")
		strPollAns11 = rs("ANSWER11")
		strPollAns12 = rs("ANSWER12")
		
	if rs("RESULT1") <> "" then
		strPollRes1 = cInt(rs("RESULT1"))
	else
		strPollRes1 = 0
        end if
	if rs("RESULT2") <> "" then
		strPollRes2 = cInt(rs("RESULT2"))
	else
		strPollRes2 = 0
        end if
	if rs("RESULT3") <> "" then
		strPollRes3 = cInt(rs("RESULT3"))
	else  
		strPollRes3 = 0
        end if		
	if rs("RESULT4") <> "" then
		strPollRes4 = cInt(rs("RESULT4"))
	else
		strPollRes4 = 0
        end if
	if rs("RESULT5") <> "" then
		strPollRes5 = cInt(rs("RESULT5"))
	else  
		strPollRes5 = 0
        end if		
	if rs("RESULT6") <> "" then
		strPollRes6 = cInt(rs("RESULT6"))
	else  
		strPollRes6 = 0
        end if		
	if rs("RESULT7") <> "" then
		strPollRes7 = cInt(rs("RESULT7"))
	else  
		strPollRes7 = 0
        end if		
	if rs("RESULT8") <> "" then
		strPollRes8 = cInt(rs("RESULT8"))
	else  
		strPollRes8 = 0
        end if		
	if rs("RESULT9") <> "" then
		strPollRes9 = cInt(rs("RESULT9"))
	else		
		strPollRes9 = 0
        end if
	if rs("RESULT10") <> "" then
		strPollRes10 = cInt(rs("RESULT10"))
	else  
		strPollRes10 = 0
        end if		
	if rs("RESULT11") <> "" then
		strPollRes11 = cInt(rs("RESULT11"))
	else  
		strPollRes11 = 0
        end if		
	if rs("RESULT12") <> "" then
		strPollRes12 = cInt(rs("RESULT12"))
	else  
		strPollRes12 = 0
        end if		
		
		strPostDate = rs("POST_DATE")
		strEndDate = rs("END_DATE")
		strPollAuthor = rs("POLL_AUTHOR")
	end if

	rs.Close
	set rs = nothing

strResultTotal = cInt(strPollRes1 + strPollRes2 + strPollRes3 + strPollRes4 + strPollRes5 + strPollRes6 + strPollRes7 + strPollRes8 + strPollRes9 + strPollRes10 + strPollRes11 + strPollRes12)

if not strResultTotal = 0 then
barPercent1 = round((strPollRes1/strResultTotal)*100,0)
barPercent2 = round((strPollRes2/strResultTotal)*100,0)
barPercent3 = round((strPollRes3/strResultTotal)*100,0)
barPercent4 = round((strPollRes4/strResultTotal)*100,0)
barPercent5 = round((strPollRes5/strResultTotal)*100,0)
barPercent6 = round((strPollRes6/strResultTotal)*100,0)
barPercent7 = round((strPollRes7/strResultTotal)*100,0)
barPercent8 = round((strPollRes8/strResultTotal)*100,0)
barPercent9 = round((strPollRes9/strResultTotal)*100,0)
barPercent10 = round((strPollRes10/strResultTotal)*100,0)
barPercent11 = round((strPollRes11/strResultTotal)*100,0)
barPercent12 = round((strPollRes12/strResultTotal)*100,0)

else
barPercent1 = 0
barPercent2 = 0
barPercent3 = 0
barPercent4 = 0
barPercent5 = 0
barPercent6 = 0
barPercent7 = 0
barPercent8 = 0
barPercent9 = 0
barPercent10 = 0
barPercent11 = 0
barPercent12 = 0

end if

if strPostDate <> strEndDate then
	pollExpireT = 1
	if strEndDate >= strCurDateString then
		pollExpireT = 0
	end if
else
	pollExpireT = 0
end if

if trim(strDBNTUserName) = "" then
	tmpUserId = 0
	tmpUserId2 = -1
else
	tmpUserId = strUserMemberID
	tmpUserId2 = strUserMemberID
end if

	strSql = "SELECT POLL_ID"
	strSql = strSql & " FROM " & strTablePrefix & "POLL_ANS "
	strSql = strSql & " WHERE POLL_ID = " & strFeaturedPoll & " AND MEMBER_ID = " & tmpUserId2

	set rs = my_Conn.Execute (strSql)

	if not(rs.eof or rs.bof) then
        	alreadyVoted = 1
 	else
        	alreadyVoted = 0
 	end if
	rs.Close
	set rs = nothing

if not trim(Request.Cookies(strUniqueID & "poll")(""&strFeaturedPoll&"")) = "" then
	cookied = 1
else
	cookied = 0
end if

if (pollExpireT = 0 and strPollAllow = 1 and cookied = 0) or (pollExpireT = 0 and strPollAllow = 0 and alreadyVoted = 0) then%>
<form action="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&amp;FORUM_ID=<%=strRqForumID%>&amp;CAT_ID=<%=strRqCatID%>&amp;Topic_Title=<% =strRqTopic_Title %>&amp;Forum_Title=<% =strRqForum_Title %>&amp;POLL_ID2=<%=strFeaturedPoll%>&amp;pollMode=vote" method="post">
<%
if not PollError then
if strPollType = 0 then
  iPollType = "radio"
else
  iPollType = "checkbox"
end if
spThemeMM = "feat_poll" & strFeaturedPoll
spThemeTitle= "Featured Poll"
spThemeBlock1_open(intSkin)
%>
<table cellpadding="0" cellspacing="0" width="100%">
<tr>
      <td>
	&nbsp;<b><% =strPollQuestion %></b><hr />
      </td>
  </tr>
<tr>
      <td>
      <input name="voteAns" value="1" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns1 %></b></span>
      </td>
  </tr>
<%if trim(strPollAns2) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="2" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns2 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns3) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="3" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns3 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns4) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="4" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns4 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns5) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="5" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns5 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns6) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="6" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns6 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns7) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="7" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns7 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns8) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="8" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns8 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns9) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="9" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns9 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns10) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="10" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns10 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns11) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="11" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns11 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns12) <> "" then%>
  <tr>
      <td>
      <input name="voteAns" value="12" type="<%= iPollType %>" />
	&nbsp;<span class="fSmall"><b><% =strPollAns12 %></b></span>
      </td>
  </tr>
<%end if%>
  <tr>
      <td align="center"><br />
<input src="images/vote.gif" type="image" />
<a href="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&amp;FORUM_ID=<%=strRqForumID%>&amp;CAT_ID=<%=strRqCatID%>&amp;Topic_Title=<%= strRqTopic_Title %>&amp;Forum_Title=<%= strRqForum_Title %>&amp;pollMode=result"><img src="images/voteresults.gif" title="View Results" alt="View Results" border="0" /></a>
      </td>
  </tr>
</table> 
<%
spThemeBlock1_close(intSkin)
end if %>
</form> <%

else
if not PollError then
spThemeMM = "feat_poll" & strFeaturedPoll
spThemeTitle= "Poll Results:"
spThemeBlock1_open(intSkin)
%>
  <table cellpadding="0" cellspacing="0" width="100%"><tr>
      <td colspan="2"><%if pollExpireT = 1 then%><span class="fAlert">&nbsp;<b>Poll has expired</b></span><br /><%end if%>
	  <b>Question: </b><br /><% =strPollQuestion %>
      <hr /></td>
  </tr>
  <tr>
      <td>&nbsp;</td>
      <td nowrap="nowrap">
	  <span class="fSmall"><b><% =strPollAns1 %>:&nbsp;<% =strPollRes1 %></b></span>
	  <% If trim(barPercent1) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent1/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent1%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%if trim(strPollAns2) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns2 %>:&nbsp;<% =strPollRes2 %></b></span><% If trim(barPercent2) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent2/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent2%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns3) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns3 %>: &nbsp;<% =strPollRes3 %></b></span><% If trim(barPercent3) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent3/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent3%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns4) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns4 %>: &nbsp;<% =strPollRes4 %></b></span><% If trim(barPercent4) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent4/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent4%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns5) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns5 %>: &nbsp;<% =strPollRes5 %></b></span><% If trim(barPercent5) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent5/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent5%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns6) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns6 %>: &nbsp;<% =strPollRes6 %></b></span><% If trim(barPercent6) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent6/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent6%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns7) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns7 %>: &nbsp;<% =strPollRes7 %></b></span><% If trim(barPercent7) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent7/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent7%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if
if trim(strPollAns8) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns8 %>: &nbsp;<% =strPollRes8 %></b></span><% If trim(barPercent8) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent8/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent8%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if
if trim(strPollAns9) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns9 %>: &nbsp;<% =strPollRes9 %></b></span><% If trim(barPercent9) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent9/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent9%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if
if trim(strPollAns10) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns10 %>: &nbsp;<% =strPollRes10 %></b></span><% If trim(barPercent10) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent10/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent10%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if
if trim(strPollAns11) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns11 %>: &nbsp;<% =strPollRes11 %></b></span><% If trim(barPercent11) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent11/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent11%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if
if trim(strPollAns12) <> "" then%>
  <tr>
      <td></td>
      <td>
	  <span class="fSmall"><b><% =strPollAns12 %>: &nbsp;<% =strPollRes12 %></b></span><% If trim(barPercent12) > 0 Then %><br /><span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent12/2.5%>%" height="8" alt="" />&nbsp;(<% =barPercent12%>%)<% End If %><br />&nbsp;</span>
      </td>
  </tr>
<%end if%>
  <tr>
      <td colspan="2" align="center"><hr /><span class="fSmall"><b>Total:&nbsp;<% =strResultTotal %><%if strResultTotal = "1" then response.write "&nbsp;vote" else response.write "&nbsp;votes"%></b></span></td>
  </tr>
  <tr>
      <td colspan="2" align="center"><br />
<a href="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&amp;FORUM_ID=<%=strRqForumID%>&amp;CAT_ID=<%=strRqCatID%>&amp;Topic_Title=<% =strRqTopic_Title %>&amp;Forum_Title=<% =strRqForum_Title %>&amp;pollMode=result"><img src="images/voteresults.gif" alt="View Results" border="0" /></a>
      </td>
  </tr></table>
<%
spThemeBlock1_close(intSkin)
end if

end if
end if
end if
end function

function chkForumAccess(fMemID,fForum)
	if hasAccess(1) then 
		chkForumAccess = true
		exit function
	end if
	'
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PASSWORD_NEW"
	strSql = strSql & " FROM " & strTablePrefix & "FORUM"
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.Forum_ID = " & fForum
	set rsStatus = my_conn.Execute (strSql)
	dim Users
	dim MatchFound
	If cint(rsStatus("F_PRIVATEFORUMS")) <> 0 then
			Select case cint(rsStatus("F_PRIVATEFORUMS"))
				case 0
					chkForumAccess = true
				case 1, 6 '## Allowed Users
					UserNum = fMemID
					if isAllowedMember(fForum,UserNum) = 1 then
					  chkForumAccess = true
					else
					  chkForumAccess = false
					end if
				case 2 '## password
					select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							chkForumAccess = true
						case else
							if trim(chkString(Request("pass"), "urlpath")) = "" then
								chkForumAccess = false
							else
								if trim(chkString(Request("pass"), "urlpath")) <> rsStatus("F_PASSWORD_NEW") then
									chkForumAccess = false
								else
									Response.Cookies(strUniqueID & "User").Path = strCookieURL
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
									chkForumAccess = true
								end if
							end if
					end select
				case 3    '## Either Password or Allowed
					UserNum = fMemID					
					if isAllowedMember(fForum,UserNum) = 1 then
						chkForumAccess = true
					else
						chkForumAccess = false
					end if
					if not(chkForumAccess) then 
					  if fMemID = strUserMemberID then
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if trim(chkString(Request("pass"), "urlpath")) = "" then
									chkForumAccess = false
								else
									if trim(chkString(Request("pass"), "urlpath")) <> rsStatus("F_PASSWORD_NEW") then
										chkForumAccess = false
									else
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					  end if
					end if
				
				case 7    '## members or password
					if fMemID < 1 then
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if trim(chkString(Request("pass"), "urlpath")) = "" then
									chkForumAccess = false
								else
									if trim(chkString(Request("pass"), "urlpath")) <> rsStatus("F_PASSWORD_NEW") then
										chkForumAccess = false
									else
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					end if						
					
				case 4, 5 '## members only
					if fMemID < 1 then
						chkForumAccess = false
					else
						chkForumAccess = true
					end if

				case 8, 9   
					test="test db"
					chkForumAccess = FALSE
					if strAuthType="db" then
						chkForumAccess = true
						exit function
					end if              
					NTGroupSTR = Split(strNTGroupsSTR, ", ")
					for j = 0 to ubound(NTGroupSTR)
						NTGroupDBSTR = Split(rsStatus("F_PASSWORD_NEW"), ", ")
						for i = 0 to ubound(NTGroupDBSTR)
							if NTGroupDBSTR(i) = NTGroupSTR(j) then
								chkForumAccess = True    
								exit function
							end if
						next
					next

				case 13, 14 ' New group stuff
					chkForumAccess = false
					if instr(rsStatus("F_PASSWORD_NEW"), ",") then
					    strPass = Split(rsStatus("F_PASSWORD_NEW"), ",")
					    for i=0 to ubound(strPass)
					         if hasAccess(cLng(Trim(strPass(i)))) then
					              chkForumAccess = true
					         end if
					    next
					elseif hasAccess(rsStatus("F_PASSWORD_NEW")) then
					    chkForumAccess = true
					end if

				case else    
					chkForumAccess = true
			end select
	else
		chkForumAccess = true
	end if
	set rsStatus = nothing
end function

function isAllowedMember(fForum_ID,fMemberID)
		isAllowedMember = 0
		on error resume next
		strSql = "SELECT MEMBER_ID, FORUM_ID FROM " & strMemberTablePrefix & "ALLOWED_MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "ALLOWED_MEMBERS.FORUM_ID = " & fForum_ID
		strSql = strSql & " AND " & strMemberTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID = " & fMemberID

		set rsAllowedMember = my_Conn.execute (strSql)
		if (rsAllowedMember.EOF or rsAllowedMember.BOF) then
			isAllowedMember = 0
			set rsAllowedMember = nothing
			exit function
		else
			isAllowedMember = 1
			set rsAllowedMember = nothing
		end if
		on error goto 0
end function

function chkDisplayForum(fForum_ID)
  if hasAccess(1) then
	chkDisplayForum= true
  else
	' - load the user list       
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS,  " & strTablePrefix & "FORUM.F_PASSWORD_NEW  "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID

	set rsAccess = my_Conn.Execute(strSql)
	select case rsAccess("F_PRIVATEFORUMS")

	case 5
		UserNum = strUserMemberID
		if not hasAccess(2) then
			chkDisplayForum= false
		else
			chkDisplayForum= true
		end if
	case 6
		UserNum = strUserMemberID
		if not hasAccess(2) then
			chkDisplayForum= false
		else
		  MatchFound = isAllowedMember(fForum_ID,UserNum)
		  if MatchFound = 1 then
			chkDisplayForum= true
		  Else
			chkDisplayForum= false
		  end if 
		end if
 	case 8
		chkDisplayForum= false
			if strAuthType="nt" THEN
				NTGroupSTR = Split(strNTGroupsSTR, ", ")
				for j = 0 to ubound(NTGroupSTR)
					NTGroupDBSTR = Split(rsAccess("F_PASSWORD_NEW"), ", ")
					for i = 0 to ubound(NTGroupDBSTR)
						if NTGroupDBSTR(i) = NTGroupSTR(j) then
							chkDisplayForum= true
						end if
					next
				next
			End if
			
	case 14 ' New group stuff
		strGroupPassed = false
		if instr(rsAccess("F_PASSWORD_NEW"), ",") then
		    strPass = Split(rsAccess("F_PASSWORD_NEW"), ",")
		    for i=0 to ubound(strPass)
		         if hasAccess(cLng(Trim(strPass(i)))) then
		              strGroupPassed = true
		         end if
		    next
		elseif hasAccess(rsAccess("F_PASSWORD_NEW")) then
		    strGroupPassed = true
		end if
		if strGroupPassed then
			chkDisplayForum = true
		end if

	case else 
	  chkDisplayForum= true
	end select 
	set rsAccess = nothing
  end if
end function

':: forum admin menu
sub forumConfigMenu(typ)
	if typ = 1 then
	  cls = "block"
	  icn = "min1"
	  alt = "Collapse"
	else
	  cls = "none"
	  icn = "max1"
	  alt = "Expand"
	end if %>
    <div class="tCellAlt1" onMouseOver="this.className='tCellHover';" onMouseOut="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block4<%= typ %>','2');"><span style="margin: 2px;"><img name="block4<%= typ %>Img" id="block4<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b>Forums</b></div>
      <div class="menu" id="block4<%= typ %>" style="display: <%= cls %>; text-align:left;">
	<% if PgType = "adminForums" then %>
		<%if intIsSuperAdmin then%>
		<a onclick="show('aa');hide('ab');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Forum Features<br /></a>
		<a onclick="show('ab');hide('aa');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Moderator Setup<br /></a>
		<a onclick="show('ac');hide('ab');hide('aa');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Merge Forums<br /></a>
		<a href="admin_count.asp"><%= icn_bar %>Update Counts<br /></a>
		<a onclick="show('ad');hide('ab');hide('aa');hide('ac');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Forum Archiving<br /></a>
		<%end if%>
		<a onclick="show('ae');hide('ab');hide('ac');hide('ad');hide('aa');hide('af');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Forum Order<br /></a>
		<a onclick="show('af');hide('ab');hide('ac');hide('ad');hide('ae');hide('aa');hide('ag');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Forum Status<br /></a>
		<a onclick="show('ag');hide('ab');hide('ac');hide('ad');hide('ae');hide('af');hide('aa');hide('ah');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Last Topics<br /></a>
		<a onclick="show('ah');hide('ab');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('aa');hide('ai');hide('zz');" href="javascript:;"><%= icn_bar %>Front Page News<br /></a>
		<a onclick="show('ai');hide('ab');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('aa');hide('zz');" href="javascript:;"><%= icn_bar %>Polls<br /></a>
	<% else %>		
		<%if intIsSuperAdmin then%>
		<a href="admin_forums.asp"><%= icn_bar %>Forum Features<br /></a>
		<a href="admin_forums.asp?cmd=1"><%= icn_bar %>Moderator Setup<br /></a>
		<a href="admin_forums.asp?cmd=2"><%= icn_bar %>Merge Forums<br /></a>
		<a href="admin_count.asp"><%= icn_bar %>Update Counts<br /></a>
		<a href="admin_forums.asp?cmd=3"><%= icn_bar %>Forum Archiving<br /></a>
		<%end if%>
		<a href="admin_forums.asp?cmd=4"><%= icn_bar %>Forum Order<br /></a>
		<a href="admin_forums.asp?cmd=5"><%= icn_bar %>Forum Status<br /></a>
		<a href="admin_forums.asp?cmd=6"><%= icn_bar %>Last Topics<br /></a>
		<a href="admin_forums.asp?cmd=7"><%= icn_bar %>Front Page News<br /></a>
		<a href="admin_forums.asp?cmd=8"><%= icn_bar %>Polls<br /></a>
	<% end if %>
		   </div>
<%
end sub

sub forum_ads() %>
<div class="tCellAlt1" style="padding:4px;">
<script type="text/javascript">
google_ad_client = "pub-0322107212113565";
google_ad_width = 728;
google_ad_height = 90;
google_ad_format = "728x90_as";
google_ad_channel ="3136682314";
google_color_border = "336699";
google_color_bg = "FFFFFF";
google_color_link = "0000FF";
google_color_url = "008000";
google_color_text = "000000";
//</script>
<script type="text/javascript"
  src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
</script></div>
<%
end sub
%>