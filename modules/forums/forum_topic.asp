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
%>
<!--#INCLUDE FILE="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="modules/forums/forum_functions.asp" -->
<%
if strQuickReply = 1 then
  hasEditor = true  
  strEditorType = "advanced"
  'strEditorType = "default"
 'strEditorType = "simple"
  strEditorElements = "Message"
end if

bOnlineUsers = true
CurPageInfoChk = "1"
function CurPageInfo ()
if Request.QueryString("SearchStrings") = "" then
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = chkstring(Request.QueryString("Topic_Title"), "sqlstring")
	PageAction = "Viewing Message<br />" 
	PageLocation = "forum_topic.asp?" & strOnlineQueryString & ""
	CurPageInfo = PageAction & " " & "<a href=""" & PageLocation & """>" & PageName & "</a>"
else
	PageName = chkstring(Request.QueryString("Topic_Title"), "sqlstring")
	PageAction = "Viewing Highlighted Message:<br />" 

	CurPageInfo = PageAction & PageName
end if
end function 

if Request.QueryString("TOPIC_ID") <> "" or Request.QueryString("TOPIC_ID") <> " " then
	if IsNumeric(Request.QueryString("TOPIC_ID")) = True then
		strRqTopicID = cLng(Request.QueryString("TOPIC_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if
if Request.QueryString("FORUM_ID") <> "" or Request.QueryString("FORUM_ID") <> " " then
	if IsNumeric(Request.QueryString("FORUM_ID")) = True then
		strRqForumID = cLng(Request.QueryString("FORUM_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if
if Request.QueryString("CAT_ID") <> "" or Request.QueryString("CAT_ID") <> " " then
	if IsNumeric(Request.QueryString("CAT_ID")) = True then
		strRqCatID = cLng(Request.QueryString("CAT_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if
if Request.QueryString("REPLY_ID") <> "" or Request.QueryString("REPLY_ID") <> " " then
	if IsNumeric(Request.QueryString("REPLY_ID")) = True then
		strRqReplyID = cLng(Request.QueryString("REPLY_ID"))
	else
		closeAndGo("fhome.asp")
	end if
end if

dim t_count
	t_count = ""
dim strPreHighlight
	strPreHighlight = "<span style='background-color: #FFFF00'>"
dim strPostHighlight
	strPostHighlight = "</span>"
dim strNoHighlightMsg
	strNoHighlightMsg = "<br /><br /><span class=""fAlert"">[Note: Some of your search words were contained within the "
	strNoHighlightMsg = strNoHighlightMsg & "message html tags and cannot be highlighted.]</span><br />"

dim strUniqueHolderPrefix
	strUniqueHolderPrefix = "$&@$"
dim strUniqueHolderSuffix
	strUniqueHolderSuffix = "#)(#"
dim nUniqueHolderIndex
	nUniqueHolderIndex=-1
dim arrHTMLReplacements()
	
dim strSearchStrings
	strSearchStrings = chkString(Request.QueryString("SearchStrings"),"sqlstring") ' convert to space delimited for keyword processing
dim strURLSearchStrings
	strURLSearchStrings = strSearchStrings
	strURLSearchStrings = ChkString(strURLSearchStrings,"urlpath")   ' convert for use in url when building jump links
dim arrSearchStrings
	arrSearchStrings = split(strSearchStrings," ")
dim nSearchStringsTopIndex 
	nSearchStringsTopIndex = ubound(arrSearchStrings)
dim strSearchType 
	strSearchType = trim(chkString(Request.QueryString("SearchType"),"sqlstring"))
dim bKeywordsPresent
	if strSearchType <> "" then

		bKeywordsPresent=true
	else
		bKeywordsPresent=false
	end if

' START Fixed Variables: Do not alter. See equivalent variables in forum_search.asp for comment
dim PrevSearchItem
PrevSearchItem = "<img src=""images/icons/icon_prev_find.gif"" border=""0"" title=""Go to previous search item"" alt=""Go to previous search item"">"
dim NextSearchItem
NextSearchItem = "<img src=""images/icons/icon_next_find.gif"" border=""0"" title=""Go to next search item"" alt=""Go to next search item"">"
dim sWhichPage
sWhichPage = "&whichpage="  ' used to tell the forum_topic.asp page which page to show
dim arrReplyIDs
arrReplyIDs = split(chkString(Request.QueryString("ReplyIDs"),"sqlstring"),",") ' array of all 'search hit' ReplyIDs for this topic
dim nRepliesArrayUpperBound
nRepliesArrayUpperBound = ubound(arrReplyIDs) ' index of upper bound of ReplyID array
dim arrAllReplyIDs
arrAllReplyIDs = split(chkString(Request.QueryString("AllReplyIDs"),"sqlstring"),",") ' array of ALL ReplyIDs for this topic
dim nAllRepliesArrayUpperBound
nAllRepliesArrayUpperBound = ubound(arrAllReplyIDs) ' index of upper bound of AllReplyIDs array
' END Fixed variables
 %>
<!--#INCLUDE FILE="inc_top.asp" -->
<%
	' get module id
	sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'forums'"
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  intAppID = rsA("APP_ID")
	end if
Member_ID = strUserMemberID
%>
<SCRIPT language="JavaScript" type="text/JavaScript">

function swapIMAV(id) {
	var o1=document.getElementById("imPanel"+id);
	var o2=document.getElementById("avPanel"+id);
	if(o2.style.display!="block") {
		o1.style.display="none";
		o2.style.display="block";
	} else {
		o1.style.display="block";
		o2.style.display="none";
	}
}

</script>
<% 
if strPrivateForums = 1 then
	if Request("Method_Type") = "" then
		chkUser4()
	end if
end if

if (hasAccess(1)) or (chkForumModerator(strRqForumID, STRdbntUserName)= "1") or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if

'Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID

set rsCStatus = my_Conn.Execute(StrSql)
if rsCStatus.EOF then
	set rsCStatus = nothing
	closeAndGo("fhome.asp")
else
	staCatStatus = rsCStatus("CAT_STATUS")
end if
set rsCStatus = nothing

'Find out if the Forum is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS, " & strTablePrefix & "FORUM.F_SUBJECT " 
strSql = strSql & " FROM " & strTablePrefix & "FORUM "
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumID

set rsFStatus = my_Conn.Execute(StrSql)
if rsFStatus.EOF then
	closeAndGo("fhome.asp")
else
	staFStatus = rsFStatus("F_STATUS")
	staFTitle = chkString(rsFStatus("F_SUBJECT"), "urlpath")
end if
set rsFStatus = nothing

' Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "TOPICS.T_STATUS, " & strTablePrefix & "TOPICS.T_POLL, " & strTablePrefix & "TOPICS.T_SUBJECT " 
strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & strRqTopicID

set rsTStatus = my_Conn.Execute(StrSql)
if rsTStatus.EOF or rsTStatus.BOF then
	closeAndGo("fhome.asp")
else
	staTStatus = rsTStatus("T_STATUS")
	staTTitle = rsTStatus("T_SUBJECT")
	staTPoll = rsTStatus("T_POLL")
end if
set rsTStatus = nothing

dim mypage : mypage = request("whichpage")
if ((Trim(mypage) = "") Or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)


' Paging Variables
dim scriptname, intPagingLinks, strQS
scriptname = request.servervariables("script_name")
intPagingLinks = 6 ' ## Number of links per page...


strQS = "&TOPIC_ID=" & chkString(Request("TOPIC_ID"), "sqlstring") &_
	"&FORUM_ID=" & chkString(Request("FORUM_ID"), "sqlstring") &_
	"&CAT_ID=" & chkString(Request("CAT_ID"), "sqlstring") &_
	"&Forum_Title=" & ChkString(Request("FORUM_Title"),"urlpath") &_
	"&Topic_Title=" & ChkString(Request("Topic_Title"),"urlpath")
 %>
<script type="text/javascript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<%
	'
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strTablePrefix & "REPLY.REPLY_ID, " & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " & strTablePrefix & "REPLY.R_MESSAGE, " & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & strRqTopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"

	if strDBType = "mysql" then 'MySql specific code

		' Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
		strSql2 = strSql2 & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
		strSql2 = strSql2 & " AND   TOPIC_ID = " & strRqTopicID & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 
	
		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize= strPageSize

		rs.open  strSql,  my_Conn, 3
		if not(rs.EOF) then
			rs.movefirst
		end if
		
	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")

		rs.cachesize = strPageSize
		rs.open  strSql,  my_Conn, 3

		If not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize = strPageSize
			rs.absolutepage = mypage '**
			maxpages = cint(rs.pagecount)
		end if
	end if
	i = 0 

':::::::::::: PAGE BREADCRUMB ::::::::::::::::::::::::::::::::::::::::::::::::
%><table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td class="mainPgCol">
	<% intSkin = getSkin(intSubSkin,2) %>
<div id="formEles" class="breadcrumb">
<table border="0" width="100%" cellpadding="0" cellspacing="0">
  <tr>
	<td width="50%" align="left" class="fNorm" nowrap>
	<img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="fhome.asp">All Forums</a><br />
	<img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="FORUM.asp?FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(Request.QueryString("FORUM_Title"),"display") %></a><br />
<%	if staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0 then %>
	<img src="images/icons/icon_blank.gif" height=15 width=15 border="0"><img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("Topic_Title"),"display") %>
<%	else %>
	<img src="images/icons/icon_blank.gif" height=15 width=15 border="0"><img src="images/icons/icon_bar.gif" height=15 width=15 border="0"><img src="images/icons/icon_folder_closed_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("Topic_Title"),"display") %>
<%	end if %>
    </td>
    <td align="center" width="50%" class="fNorm"><% call PostingOptions() %></td>
  </tr>
</table></div>
<% if maxpages > 1 then %>
  <div style="text-align:center;padding:4px;">
      <% Call Paging() %>
  </div>
<% end if %>
<%
':::::::::::: END PAGE BREADCRUMB ::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::: START POLL DISPLAY :::::::::::::::::::::::::::::::::::::::
poll_ID = staTPoll
if poll_ID <> "0" then
pollMode = trim(chkString(Request.Querystring("pollMode"),"SQLString"))

if pollMode = "result" or pollMode = "" or pollMode = "vote" then
	strSql = "SELECT POLL_TYPE, POLL_ID, POLL_ALLOW, POLL_QUESTION," 
        strSql = strSql & " ANSWER1, ANSWER2, ANSWER3, ANSWER4, ANSWER5, ANSWER6, ANSWER7, ANSWER8, ANSWER9, ANSWER10, ANSWER11, ANSWER12,"
        strSql = strSql & " RESULT1, RESULT2, RESULT3, RESULT4, RESULT5, RESULT6, RESULT7, RESULT8, RESULT9, RESULT10, RESULT11, RESULT12,"
        strSql = strSql & " POST_DATE, END_DATE, POLL_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "POLLS "
	strSql = strSql & " WHERE POLL_ID = " & poll_ID

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
	tmpUserId = getMemberID(strDBNTUserName)
	tmpUserId2 = getMemberID(strDBNTUserName)
end if

	strSql = "SELECT POLL_ID"
	strSql = strSql & " FROM " & strTablePrefix & "POLL_ANS "
	strSql = strSql & " WHERE POLL_ID = " & poll_ID & " AND MEMBER_ID = " & tmpUserId2

	set rs = my_Conn.Execute (strSql)

	if not(rs.eof or rs.bof) then
       	alreadyVoted = 1
 	else
        alreadyVoted = 0
 	end if
	set rs = nothing

if not trim(Request.Cookies(strUniqueID & "poll")(""&POLL_ID&"")) = "" then
	cookied = 1
else
	cookied = 0
end if

if pollMode = "vote" then
	if strPollAllow = 0 and tmpUserId = 0 then%>
		<p align=center><span class="fTitle">There Was A Problem</span></p>
		<span class="fSubTitle">You must a registered member to vote in this poll</span>
		<p align=center><a href="JavaScript:history.go(-1)">Go Back</a></p>
<%elseif trim(request.form("voteAns")) = "" then%>
		<p align=center><span class="fTitle">There Was A Problem</span></p>

		<span class="fSubTitle">You must select at least one option.</span>
		<p align=center><a href="JavaScript:history.go(-1)">Go Back</a></p>
<%elseif pollExpireT = "1" then%>
		<p align="center"><span class="fTitle">There Was A Problem</span></p>
		<span class="fSubTitle">The poll has expired.</span> 
		<p align="center"><a href="JavaScript:history.go(-1)">Go Back</a></p>
<%elseif alreadyVoted = "1" and strPollAllow = 0 then%>
		<p align="center"><span class="fTitle">There Was A Problem</span></p>
		<span class="fSubTitle">You have already voted.</span>
			
		<p align=center><a href="JavaScript:history.go(-1)">Go Back</a></p>
<%else
		strSql = "INSERT INTO " & strTablePrefix & "POLL_ANS (POLL_ID"
		strSql = strSql & ", ANS_VALUE"
		strSql = strSql & ", MEMBER_ID"
		strSql = strSql & ", ANS_DATE"
		strSql = strSql & ", IP"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Poll_ID
		strSql = strSql & ", '" & chkstring(request.form("voteAns"), "sqlstring")& "'"
		strSql = strSql & ", " & tmpUserId
 		strSql = strSql & ", '" & strCurDateString & "'"
		strSql = strSql & ", '" & Request.ServerVariables("REMOTE_ADDR") & "'"
		strSql = strSql & " )"

		my_Conn.Execute (strSql)
		
		strSql = "UPDATE " & strTablePrefix & "POLLS "
		strSql = strSql & " SET POLL_AUTHOR = " & strPollAuthor
If InStr(request.form("voteAns")& ",", "1,") Then
		strSql = strSql & ", RESULT1 = " & strPollRes1 + 1
end if
If InStr(request.form("voteAns")& ",", "2,") Then
		strSql = strSql & ", RESULT2 = " & strPollRes2 + 1
end if
If InStr(request.form("voteAns")& ",", "3,") Then
		strSql = strSql & ", RESULT3 = " & strPollRes3 + 1
end if
If InStr(request.form("voteAns")& ",", "4,") Then
		strSql = strSql & ", RESULT4 = " & strPollRes4 + 1
end if
If InStr(request.form("voteAns")& ",", "5,") Then
		strSql = strSql & ", RESULT5 = " & strPollRes5 + 1
end if
If InStr(request.form("voteAns")& ",", "6,") Then
		strSql = strSql & ", RESULT6 = " & strPollRes6 + 1
end if
If InStr(request.form("voteAns")& ",", "7,") Then
		strSql = strSql & ", RESULT7 = " & strPollRes7 + 1
end if
If InStr(request.form("voteAns")& ",", "8,") Then
		strSql = strSql & ", RESULT8         = " & strPollRes8 + 1
end if
If InStr(request.form("voteAns")& ",", "9,") Then
		strSql = strSql & ", RESULT9 = " & strPollRes9 + 1
end if                                   
If InStr(request.form("voteAns")& ",", "10,") Then
		strSql = strSql & ", RESULT10 = " & strPollRes10 + 1
end if
If InStr(request.form("voteAns")& ",", "11,") Then
		strSql = strSql & ", RESULT11 = " & strPollRes11 + 1
end if
If InStr(request.form("voteAns")& ",", "12,") Then
		strSql = strSql & ", RESULT12 = " & strPollRes12 + 1
end if
		strSql = strSql & " WHERE POLL_ID = " & POLL_ID

		my_Conn.Execute (strSql)

if not POLL_ID = "" then
Response.Cookies(strUniqueID & "poll")(""&POLL_ID&"") = (""&POLL_ID&"") 
Response.Cookies(strUniqueID & "poll").Expires = dateadd("d",30,strCurDateAdjust)
else
POLL_ID2 = Request.QueryString("POLL_ID2")
Response.Cookies(strUniqueID & "poll")(""&POLL_ID2&"") = (""&POLL_ID2&"") 
Response.Cookies(strUniqueID & "poll").Expires = dateadd("d",30,strCurDateAdjust)
end if
%>
		<p align=center><span class="fTitle">Thank you for voting. You will now be taken back to the poll.</span></p><br /><br />
		<p align=center><a href="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&FORUM_ID=<%=strRqForumID%>&CAT_ID=<%=strRqCatID%>&Topic_Title=<% =ChkString(left(Request.QueryString("TOPIC_TITLE"), 50),"urlpath") %>&Forum_Title=<% =ChkString(left(Request.QueryString("FORUM_TITLE"), 50),"urlpath") %>&pollMode=result">Back to Topic</a></p>
		<meta http-equiv="Refresh" content="2; URL=forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&FORUM_ID=<%=strRqForumID%>&CAT_ID=<%=strRqCatID%>&Topic_Title=<% =ChkString(left(Request.QueryString("TOPIC_TITLE"), 50),"urlpath") %>&Forum_Title=<% =ChkString(left(Request.QueryString("FORUM_TITLE"), 50),"urlpath") %>&pollMode=result">
<%end if%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
	Response.End
end if

if (pollMode = "" and pollExpireT = 0 and strPollAllow = 1 and cookied = 0) or (pollMode = "" and pollExpireT = 0 and strPollAllow = 0 and alreadyVoted = 0) then%>
<form action="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&FORUM_ID=<%=strRqForumID%>&CAT_ID=<%=strRqCatID%>&Topic_Title=<% =ChkString(left(Request.QueryString("TOPIC_TITLE"), 50),"urlpath") %>&Forum_Title=<% =ChkString(left(Request.QueryString("FORUM_TITLE"), 50),"urlpath") %>&pollMode=vote" method=post>
<br />
<%
spThemeTableCustomCode = " border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" align=""center"" width=""95%"""
spThemeSmallBlock_open()%>
  <tr>
    <td class="tSubTitle"><b>Poll: </b><% =strPollQuestion %><%if hasAccess(1) then%>&nbsp;<a href="forum_post_info.asp?TOPIC_IDP=<%=strRqTopicID%>&POLL_ID=<%=POLL_ID%>&Method_Type=editPoll">[Edit Poll]</a><%end if%></td>
  </tr>

<tr><td>
<table align=center border=0 width="100%" cellpadding="2" cellspacing="0">
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="1" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns1 %></b></span>
      </td>
  </tr>
<%if trim(strPollAns2) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="2" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns2 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns3) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="3" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns3 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns4) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="4" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns4 %></b></span>
      </td>
  </tr> 
<%end if
if trim(strPollAns5) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="5" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns5 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns6) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="6" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns6 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns7) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="7" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns7 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns8) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="8" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns8 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns9) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="9" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns9 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns10) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="10" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns10 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns11) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="11" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns11 %></b></span>
      </td>
  </tr>
<%end if
if trim(strPollAns12) <> "" then%>
  <tr>
      <td class="tCellAlt1">
      <input name="voteAns" value="12" <%if strPollType = 0 then%>type="radio"<%else%>type="checkbox"<%end if%>>
	&nbsp;<span class="fSmall"><b><% =strPollAns12 %></b></span>
      </td>
  </tr>
<%end if%>
  <tr>
      <td class="tCellAlt1">
<INPUT src="images/vote.gif" type="image" border="0">
<a href="forum_topic.asp?TOPIC_ID=<%=strRqTopicID%>&FORUM_ID=<%=strRqForumID%>&CAT_ID=<%=strRqCatID%>&Topic_Title=<% =ChkString(left(Request.QueryString("TOPIC_TITLE"), 50),"urlpath") %>&Forum_Title=<% =ChkString(left(Request.QueryString("FORUM_TITLE"), 50),"urlpath") %>&pollMode=result"><img src="images/voteresults.gif" alt="View Results" title="View Results" border="0"></a>
      </td>
  </tr>
</table></td></tr>
</form>  
<%spThemeSmallBlock_close()%>

<%else%>
<br />
<%
spThemeTableCustomCode = " border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" align=""center"" width=""95%"""
spThemeSmallBlock_open()%>
<tr><td>
<table align=center border=0 width="100%" cellpadding="2" cellspacing="0">
  <tr>
    <td class="tSubTitle" colspan="2"><b>Poll Results:</b><%if hasAccess(1) then%>&nbsp;<a href="forum_post_info.asp?TOPIC_IDP=<%=strRqTopicID%>&POLL_ID=<%=POLL_ID%>&Method_Type=editPoll">[Edit Poll]</a><%end if%></td>
  </tr>
  <tr>
      <td class="tCellAlt1" colspan="2">
	  <b>Question: </b><% =strPollQuestion %>&nbsp;(Total: <%=strResultTotal%>) 100%
<%if pollExpireT = 1 then%>&nbsp;<span class="fSubTitle"><b>Poll has expired</b></span><%end if%>
      </td>
  </tr>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns1 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent1/2.5%>%" height="8">&nbsp;<% =strPollRes1 %>&nbsp;(<% =barPercent1%>%)</span>
      </td>
  </tr>
<%if trim(strPollAns2) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns2 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent2/2.5%>%" height="8">&nbsp;<% =strPollRes2 %>&nbsp;(<% =barPercent2%>%)</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns3) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns3 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent3/2.5%>%" height="8">&nbsp;<% =strPollRes3 %>&nbsp;(<% =barPercent3%>%)</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns4) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns4 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent4/2.5%>%" height="8">&nbsp;<% =strPollRes4 %>&nbsp;(<% =barPercent4%>%)</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns5) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns5 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent5/2.5%>%" height="8">&nbsp;<% =strPollRes5 %>&nbsp;(<% =barPercent5%>%)</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns6) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns6 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent6/2.5%>%" height="8">&nbsp;<% =strPollRes6 %>&nbsp;(<% =barPercent6%>%)</span>
      </td>
  </tr> 
<%end if
if trim(strPollAns7) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns7 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent7/2.5%>%" height="8">&nbsp;<% =strPollRes7 %>&nbsp;(<% =barPercent7%>%)</span>
      </td>
  </tr>
<%end if
if trim(strPollAns8) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns8 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent8/2.5%>%" height="8">&nbsp;<% =strPollRes8 %>&nbsp;(<% =barPercent8%>%)</span>
      </td>
  </tr>
<%end if
if trim(strPollAns9) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns9 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent9/2.5%>%" height="8">&nbsp;<% =strPollRes9 %>&nbsp;(<% =barPercent9%>%)</span>
      </td>
  </tr>
<%end if
if trim(strPollAns10) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns10 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent10/2.5%>%" height="8">&nbsp;<% =strPollRes10 %>&nbsp;(<% =barPercent10%>%)</span>
      </td>
  </tr>
<%end if
if trim(strPollAns11) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns11 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent11/2.5%>%" height="8">&nbsp;<% =strPollRes11 %>&nbsp;(<% =barPercent11%>%)</span>
      </td>
  </tr>
<%end if
if trim(strPollAns12) <> "" then%>
  <tr>
      <td class="tCellAlt1">
	  <span class="fSmall"><b><% =strPollAns12 %>: </b></span>
      </td>
      <td class="tCellAlt1" width="90%">
	  <span class="fSmall"><img src="images/icons/bar.gif" width="<% =barPercent12/2.5%>%" height="8">&nbsp;<% =strPollRes12 %>&nbsp;(<% =barPercent12%>%)</span>
      </td>
  </tr>
<%end if%>
  
</table></td></tr>
<%spThemeSmallBlock_close()
end if
end if%>
<br /><br />
<%
end if
'::::::::::::::::::::: end polls ::::::::::::::::::::::::::::::::::::::::::::::

spThemeBlock1_open(intSkin)
%><table cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td>
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" class="tSubTitle" width="120" nowrap><b>Author</b></td>
        <td align="center" class="tSubTitle" width="100%"><b><% Call Topic_nav() %></b></td>
<%	if (AdminAllowed = 1) then %>
        <td align=right class="tSubTitle" colspan=2 nowrap><% call AdminOptions() %></td>
<%	else %>
        <td align=right class="tSubTitle" nowrap>&nbsp;</td>
<%	end if %>
      </tr>
<% 
	lastTopicVisited = 0
	if Request.Cookies(strCookieURL & strUniqueID & "topic_id") <> "" then
	  lastTopicVisited = Request.Cookies(strCookieURL & strUniqueID & "topic_id")
	end if
	if lastTopicVisited <> "" and not isnumeric(lastTopicVisited) then
		closeAndGo("default.asp")
	end if

	if lastTopicVisited <> strRqTopicID then
		upCount strRqTopicID,""
	    Response.Cookies(strCookieURL & strUniqueID & "topic_id") = strRqTopicID
		Response.Cookies(strCookieURL & strUniqueID & "topic_id").Expires = dateadd("d",7,strCurDateAdjust)
	end If
	GetFirst()
	if mypage = 1 then 
		'Call GetFirst()
	else
	  'strSql = "UPDATE " & strTablePrefix & "TOPICS "
	  'strSql = strSql & " SET T_VIEW_COUNT = T_VIEW_COUNT + 1"
	  'strSql = strSql & " WHERE TOPIC_ID = " & topicID & ""
	  'my_conn.Execute (strSql)
		'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
	end if

	' Get all topics from DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LASTNAME, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_CITY, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_STATE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_DATE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_MSN, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_ICQ, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_Homepage, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_GLOW, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " 	
	strSql = strSql & strTablePrefix & "REPLY.REPLY_ID, " 
	strSql = strSql & strTablePrefix & "REPLY.R_AUTHOR, " 
	strSql = strSql & strTablePrefix & "REPLY.TOPIC_ID, " 
	strSql = strSql & strTablePrefix & "REPLY.R_MESSAGE, " 
	strSql = strSql & strTablePrefix & "REPLY.R_DATE, "
	strSql = strSql & strTablePrefix & "REPLY.R_MSGICON, "
	strSql = strSql & strTablePrefix & "REPLY.R_SIG, "
	strSql = strSql & strTablePrefix & "COUNTRIES.CO_FLAG"
	strSql = strSql & " FROM (" & strMemberTablePrefix & "MEMBERS "
    strSql = strSql & "LEFT JOIN "& strTablePrefix &"COUNTRIES ON "
    strSql = strSql & strMemberTablePrefix & "MEMBERS.M_COUNTRY = "&strTablePrefix&"COUNTRIES.CO_NAME)"	
	strSql = strSql & " INNER JOIN "&strTablePrefix&"REPLY ON "&strMemberTablePrefix&"MEMBERS.MEMBER_ID = "&strTablePrefix&"REPLY.R_AUTHOR"
	strSql = strSql & " WHERE " & strTablePrefix &"REPLY.TOPIC_ID = " & strRqTopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"
	if strDBType = "mysql" then 'MySql specific code
		if mypage > 1 Then 
			intOffSet = CInt((mypage - 1) * strPageSize) - 1
			strSql = strSql & " LIMIT " & intOffSet & ", " & CInt(strPageSize) & " "
		end if

		' Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strTablePrefix & "REPLY "
		strSql2 = strSql2 & " WHERE  TOPIC_ID = " & strRqTopicID & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 
		rsCount.close
		set rsCount = nothing

		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize = strPageSize

		rs.open  strSql,  my_Conn, 3

	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.cachesize = 20
		rs.open  strSql,  my_Conn, 3
	
		if not(rs.EOF or rs.BOF) then  '## Replies found in DB
			rs.movefirst
			rs.pagesize = strPageSize
			maxpages = cint(rs.pagecount)
			rs.absolutepage = mypage
		end if
	end if		
	if rs.EOF or rs.BOF then  '## No replies found in DB
		Response.Write ""
	else %><!--
      <tr>
<%		  if (AdminAllowed = 1) then %>
        <td height="25" valign="middle" align="center" class="tSubTitle" colspan="4" nowrap><% 'forum_ads() %></td>
<%		  else %>
        <td height="25" valign="middle" align="center" class="tSubTitle" colspan="3" nowrap><% 'forum_ads() %></td>
<%		  end if %>
      </tr> -->
<%		if maxpages > 1 then %>
      <tr>
<%		  if (AdminAllowed = 1) then %>
        <td height="25" valign="middle" align="center" class="tSubTitle" colspan="4" nowrap>Replies</td>
<%		  else %>
        <td height="25" valign="middle" align="center" class="tSubTitle" colspan="3" nowrap>Replies</td>
<%		  end if %>
      </tr>
      <tr>
        <td align="center" class="tSubTitle" width="120" nowrap><b>Author</b></td>
        <td align="center" class="tSubTitle" width="100%"><b><% Call Paging() %></b></td>
<%		  if (AdminAllowed = 1) then %>
        <td align=right class="tSubTitle" colspan=2 nowrap>&nbsp;</td>
<%	 	  else %>
        <td align=right class="tSubTitle" nowrap>&nbsp;</td>
<%		  end if %>
      </tr>
<%		
		else %>
      <tr>
        <td align="center" class="tSubTitle" width="120" nowrap><b>Author</b></td>
        <td height="25" valign="middle" align="center" class="tSubTitle" width="100%"><b>Replies</b></td>
<%		  if (AdminAllowed = 1) then %>
        <td align=right class="tSubTitle" colspan=2 nowrap>&nbsp;</td>
<%	 	  else %>
        <td align=right class="tSubTitle" nowrap>&nbsp;</td>
<%		  end if %>
      </tr>
<%		end if
':::: end reply-bar ::::: %>
<%		'rs.movefirst			
		intI = 0 
		howmanyrecs = 0
		rec = 1
	
		do until rs.EOF or (mypage = 1 and rec > strPageSize) or (mypage > 1 and rec > strPageSize) '**		
			if intI = 0 then 
				CColor = "tCellAlt2"
				MColor = "tCellAlt1"
			else
				CColor = "tCellAlt1"
				MColor = "tCellAlt2"
			end if
%>
      <tr>
        <td class="<% =MColor %>" valign="top">
		<a name="pid<% =rs("REPLY_ID") %>"></a>
		<span class="fNorm"><% getAuthorInfo(rs("REPLY_ID")) %></span>
	</td>
  
        <td class="<% =CColor %>" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top"><img src='images/icons/icon_mi_<% =rs("R_MSGICON") %>.gif' border="0" hspace="3"><span class="fSmall">Posted&nbsp;-&nbsp;<% =ChkDate(rs("R_DATE")) %>&nbsp;:&nbsp;<% =ChkTime(rs("R_DATE")) %></span>
<%			if (staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0 and hasAccess(2)) then %>
		&nbsp;<a href="javascript:;"><img src="images/icons/icon_exclaim.gif" height=15 width=15 alt="Report this post to a Moderator" title="Report this post to a Moderator" border="0" align="middle" hspace="2" onClick="openWindow3('forum_report_post.asp?rid=<% =rs("REPLY_ID") %>&page=<%= mypage %>')"></a>
		<% End If %>
<%			if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
<%				if (staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0) or (AdminAllowed = 1) then %>
        &nbsp;<a href="forum_post.asp?method=Edit&REPLY_ID=<% =rs("REPLY_ID") %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&auth=<% =ChkString(rs("R_AUTHOR"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_edit_topic.gif" height=15 width=15 alt="Edit Message" title="Edit Message" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<%			if (staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0 and hasAccess(2)) or (AdminAllowed = 1) then %>
        &nbsp;<a href="forum_post.asp?method=ReplyQuote&REPLY_ID=<% =rs("REPLY_ID") %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>&M=<% =Request.QueryString("M") %>"><img src="images/icons/icon_reply_topic.gif" height=15 width=15 alt="Reply with Quote" title="Reply with Quote" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%			if (strIPLogging = "1") then %>
<%				if (AdminAllowed = 1) or (strNoCookies = "1") then %>
        &nbsp;<a href="JavaScript:;"><img src="images/icons/icon_ip.gif" onClick="openWindow('forum_pop.asp?mode=12&sid=<% =rs("REPLY_ID") %>&cmd=<% =strRqForumID %>')" height=15 width=15 alt="View user's IP address" title="View user's IP address" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<%			if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
<%				if (staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0) or (AdminAllowed = 1) then %>
        &nbsp;<a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Reply&REPLY_ID=<% =Rs("REPLY_ID") %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>')"><img src="images/icons/icon_delete_reply.gif" height=15 width=15 title="Delete Reply" alt="Delete Reply" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>

        <hr noshade>
        <%
        Rmessage = rs("R_MESSAGE")
		if strAllowHtml = 1 then
          'signature = ReplaceUrls(GetSig(getMemberName(rs("R_AUTHOR"))))
          signature = GetSig(getMemberName(rs("R_AUTHOR")))		
		else
          signature = GetSig(getMemberName(rs("R_AUTHOR")))
		end if
	if rs("R_SIG") = 1 and signature <> ""  then
		if strAllowHtml = 1 then
		Rmessage = Rmessage & "<br /><br /><br />" & replace(ChkString(signature,"signature"),"''","'")
		else
		Rmessage = Rmessage & vbCrLf & vbCrLf & replace(ChkString(signature,"signature"),"''","'")
		end if
	end if
		'Rmessage = replace(formatstr(ReplaceUrls(Rmessage)),"''","'")
		Rmessage = replace(formatstr(Rmessage),"''","'")
        %>        
        <span class="fNorm"><% =Rmessage %></span><a href="#top"><img src="themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right" title="Go to Top of Page" alt="Go to Top of Page"></a></td>
      </tr>
<%			
		    rs.MoveNext
		    intI  = intI + 1
		    if intI = 2 then 
				intI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>
    </table></td>
  </tr>
<%if maxpages > 1 or (AdminAllowed = 1) then%>
  <tr>
    <td colspan="2">
    <table border="0" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" class="fNorm">
<% if maxpages > 1 then %>
        <b>Topic is <% =maxpages %> Pages Long:</b><br />
        <% Call Paging() %>
<% else %>
	&nbsp;
<% end if %>
        </td>
        <td align="right" width="100" class="fNorm">
<% if (AdminAllowed = 1) then %>
<%	call AdminOptions() %>
<% else %>
        &nbsp;
<% end if %>
        </td> 
      </tr>
    </table>
	</td>
  </tr>
<%end if%>
</table>
<%
spThemeBlock1_close(intSkin)%>
</div>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top" class="fNorm" width="60%"><% Call PostingOptions() %></td>
    <td align="center" valign="top" width="50%">
	<!--#INCLUDE file="modules/forums/inc_jump_to.asp" --></td>
  </tr>
</table>
<br />
<%if hasAccess(2) and (staCatStatus = 1) and (staFStatus = 1) and (staTStatus = 1) and not strQuickReply = "0" then
strCkPassWord = Request.Cookies(strUniqueID & "User")("Pword") %>
<center><div style="width: 700px;">	<%
spThemeTitle = "<center>Quick Reply Box</center>"
spThemeBlock1_open(intSkin)
%>
<form action="forum_post_info.asp" method="post" name="PostTopic" id="PostTopic">
<input name="Method_Type" type="hidden" value="Reply">
<input name="Type" type="hidden" value="<% =Request.QueryString("type") %>">
<input name="REPLY_ID" type="hidden" value="<% =strRqReplyID %>">
<input name="TOPIC_ID" type="hidden" value="<% =strRqTopicID %>">
<input name="FORUM_ID" type="hidden" value="<% =strRqForumID %>"> 
<input name="CAT_ID" type="hidden" value="<% =strRqCatID %>">
<input name="FORUM_Title" type="hidden" value="<% =ChkString(Request.QueryString("FORUM_Title"), "hidden") %>">
<input name="Topic_Title" type="hidden" value="<% =ChkString(Request.QueryString("TOPIC_Title"), "edit") %>">
<input name="Refer" type="hidden" value="<% =strHomeURL %>link.asp?TOPIC_ID=<% =strRqTopicID %>&view=lasttopic">
<input name="cookies" type="hidden" value="yes">
<input name="UserName" type="hidden" Value="<% =strDBNTUserName%>">
<input name="Password" type="hidden" value="<% =strCkPassWord%>">
<!--input name="Sig" type="hidden" value="yes"-->
<table>
	<% 	
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "",""
  else
  	displayPLAINeditor 2,""
  end if %>
<tr>
	<td align="center" colspan="2">
<input name="Sig" type="hidden" value="yes">
<input name="Submit" type="submit" value="Post New Reply" accesskey="s" title="Shortcut Key: Alt+S" class="button">&nbsp;<% If strAllowHtml <> 1 Then %><input name="Preview" type="button" class="Button" value=" Preview " onclick="OpenPreview()"><% End If %></td>
</tr></table></form>
<%
spThemeBlock1_close(intSkin)%></div></center>
<%end if
set rs = nothing

'end if
%><br /></td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
sub GetFirst()
	' Get Original Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMSTATUS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_CITY, " & strMemberTablePrefix & "MEMBERS.M_STATE, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_DATE, " & strMemberTablePrefix & "MEMBERS.M_GLOW, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AVATAR_URL, " & strMemberTablePrefix & "MEMBERS.M_MSN, " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL, "	
	strSql = strSql & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_SIG, " 
	strSql = strSql & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_MSGICON, " & strTablePrefix & "TOPICS.T_MESSAGE, " & strTablePrefix & "TOPICS.T_VIEW_COUNT, "
    strSql = strSql & strTablePrefix & "COUNTRIES.CO_FLAG "
    strSql = strSql & " FROM (" & strMemberTablePrefix & "MEMBERS "
    strSql = strSql & "LEFT JOIN "& strTablePrefix &"COUNTRIES ON "
    strSql = strSql & strMemberTablePrefix & "MEMBERS.M_COUNTRY = "&strTablePrefix&"COUNTRIES.CO_NAME)"	
	strSql = strSql & " INNER JOIN "&strTablePrefix&"TOPICS ON "&strMemberTablePrefix&"MEMBERS.MEMBER_ID = "&strTablePrefix&"TOPICS.T_AUTHOR"
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " &  strRqTopicID 
	set rs = my_Conn.Execute (strSql)

	if rs.EOF or rs.BOF then  '## No categories found In DB
		Response.Write "  <tr>" & vbCrLf
		Response.Write "    <td colspan=5 class=""fSubTitle"">No Topics Found</td>" & vbCrLf
		Response.Write "  </tr>" & vbCrLf
	else
	  strIMmsg = "View " & ChkString(rs("M_NAME"),"display") & "'s profile"
 %>
      <tr>
        <td class="tCellAlt0" valign="top" nowrap="nowrap">
		<span class="fNorm"><% getAuthorInfo("t" & rs("TOPIC_ID")) %></span>
	</td>
           
        <td class="tCellAlt0" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top">
		<table width="100%" height="10" border="0" cellspacing="0" cellpadding="5">
  <tr>
    <td valign="top"> <img src='images/icons/icon_mi_<% =rs("T_MSGICON") %>.gif' border="0" hspace="3"><span class="fSmall">Posted&nbsp;-&nbsp;
      <% =ChkDate(rs("T_DATE")) %>
      &nbsp;:&nbsp;
      <% =ChkTime(rs("T_DATE")) %>
       </span>
      <%			if (staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0 and hasAccess(2)) then %>
      &nbsp;<a href="javascript:;"><img src="images/icons/icon_exclaim.gif" height=16 width=16 title="Report this post to a Moderator" alt="Report this post to a Moderator" border="0" align="middle" hspace="4" onClick="openWindow3('forum_report_post.asp?tid=<% =rs("TOPIC_ID") %>&page=<%= mypage %>')"></a> 
      <% 			End If %>
      <%		if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
      <%			if ((staCatStatus <> 0) and (staFStatus <> 0) and (staTStatus <> 0)) or (hasAccess(1) or mlev = 3) then %>
      &nbsp;<a href="forum_post.asp?method=EditTopic&REPLY_ID=<% =rs("TOPIC_ID") %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&auth=<% =rs("T_AUTHOR") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_edit_topic.gif" height=15 width=15 alt="Edit Message" title="Edit Message" border="0" align="absmiddle" hspace="6"></a> 
      <%			end if %>
      <%		end if %>
      <%			if staCatStatus <> 0 and staFStatus <> 0 and staTStatus <> 0 and hasAccess(2) then %>
      &nbsp;<a href="forum_post.asp?method=TopicQuote&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" height=15 width=15 title="Reply with Quote" alt="Reply with Quote" border="0" align="absmiddle" hspace="6"></a> 
      <%			end if %>
      <%		if (strIPLogging = "1") then %>
      <%			if (AdminAllowed = 1) or (strNoCookies = "1") then %>
      &nbsp;<a href="JavaScript:;"><img src="images/icons/icon_ip.gif" onClick="openWindow('forum_pop.asp?mode=12&cid=<% =rs("TOPIC_ID") %>&cmd=<% =strRqForumID %>')" height=15 width=15 alt="View user's IP address" title="View user's IP address" border="0" align="absmiddle" hspace="6"></a> 
      <%			end if %>
      <%		end if %>
      <hr noshade size="1">
    </td>
  </tr>
  <tr>
    <td valign="top" class="fNorm">
        <%
        Tmessage = rs("T_MESSAGE")
		if strAllowHtml = 1 then
          signature = ReplaceUrls(GetSig(getMemberName(rs("T_AUTHOR"))))		
		else
          signature = GetSig(getMemberName(rs("T_AUTHOR")))
		end if
	if rs("T_SIG") = 1 and signature <> ""  then
		if strAllowHtml = 1 then
		Tmessage = Tmessage & "<br /><br /><br />" & replace(ChkString(signature,"signature"),"''","'")
		else
		Tmessage = Tmessage & vbCrLf & vbCrLf & replace(ChkString(signature,"signature"),"''","'")
		end if
	end if
		Tmessage = replace(formatstr(Tmessage),"''","'")
        %><% =Tmessage %>
      <% 'ProcessMsg2(rs("T_MESSAGE"),0, rs("T_SIG"), rs("T_AUTHOR")) %>
	</td>
  </tr>
</table> 
		</td>
      </tr>
<%	end if
	
	't_count = rs("T_VIEW_COUNT") + 1
	'   strRqTopicID
	set rs = nothing

End Sub

function upCount(topicID,cnt)
	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & "SET " & strTablePrefix & "TOPICS.T_VIEW_COUNT = " & strTablePrefix & "TOPICS.T_VIEW_COUNT+1 "
	strSql = strSql & "WHERE (" & strTablePrefix & "TOPICS.TOPIC_ID = " & topicID & ");"
	executeThis(strSql)
end function

sub PostingOptions() 
	if (hasAccess(2)) then
		if ((staCatStatus = 1) and (staFStatus = 1)) then %>
    <a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="Post New Topic" alt="Post New Topic" height=15 width=15 border=0></a>&nbsp;<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">New Topic</a>
<%  else
			if (AdminAllowed = 1) then %>
    <a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_locked.gif" height=15 width=15 title="Forum Locked" alt="Forum Locked" border=0></a>&nbsp;<a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">New Topic</a>
<%			else %>
    <img src="images/icons/icon_folder_locked.gif" height=15 width=15 border=0>&nbsp;Forum Locked
<%			end if 
   end if 
	if (staCatStatus = 1) and (staFStatus = 1) and (staTStatus = 1) then %>
    <a href="forum_post.asp?method=Reply&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" height=15 width=15 alt="Reply to Topic" title="Reply to Topic" border=0></a>&nbsp;<a href="forum_post.asp?method=Reply&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>">Reply to Topic</a>
<%	Else 
			if (AdminAllowed = 1)  then %>
    <a href="forum_post.asp?method=Reply&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_closed_topic.gif" height=15 width=15 title="Topic Locked" alt="Topic Locked" border=0></a>&nbsp;<a href="forum_post.asp?method=Reply&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>">Reply to Topic</a>
<%			Else %>
    <img src="images/icons/icon_closed_topic.gif" height=15 width=15 border=0>&nbsp;Topic Locked
<%			end if
	end if 
	if (lcase(strEmail) = 1) then 
		if (hasAccess(2)) or (not hasAccess(2) and  strLogonForMail <> "1") then %>
				<br />	
				<a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<%= strRqTopicID  %>')"><img border="0" src="images/icons/icon_send_topic.gif" height=15 width=15 title="Send Topic to a Friend" alt="Send Topic to a Friend"></a>&nbsp;<a href="JavaScript:openWindow('forum_pop.asp?mode=4&amp;cid=<%= strRqTopicID  %>')">Email Topic</a>
<%		end if
	end if %>
	<a href="JavaScript:openWindow5('forum_pop.asp?mode=5&amp;cid=<% =strRqTopicID %>')"><img border="0" src="images/icons/print.gif" title="Printer Friendly Version" width="16" height="17"></a>&nbsp;<a href="JavaScript:openWindow5('forum_pop.asp?mode=5&amp;cid=<% =strRqTopicID%>')">Print Topic</a><br />
	<%
  If hasAccess(2) and intBookmarks = 1 Then 
	bookmark_id = chkIsBookmarked(intAppID,"0","0",strRqTopicID,strUserMemberID)
	  if bookmark_id <> 0 then
		response.Write("<a href=""javascript:;"" onclick=""JavaScript:openWindow('forum_pop.asp?mode=8&amp;cid=" & bookmark_id & "');""><img src=""themes/" &  strTheme & "/icons/unbookmark.gif"" title=""Remove Bookmark for this Web Link"" alt='remove bookmark' border='0' style=""display:inline;"" hspace=""4"">&nbsp;Remove Bookmark</a>&nbsp;")
	  else
		response.Write("<a href=""javascript:;"" onclick=""JavaScript:openWindow('forum_pop.asp?mode=6&amp;cmd=3&amp;cid=" & strRqTopicID & "');""><img src=""themes/" &  strTheme & "/icons/bookmark.gif"" title=""Bookmark this Web Link"" alt=""bookmark"" border=""0"" style=""display:inline;"" hspace=""4"">&nbsp;Bookmark Topic</a>&nbsp;")
	  end if
  end if %>
		<%
			if intSubscriptions = 1 and hasAccess(2) then 
			  subscription_id = chkIsSubscribed(intAppID,"0","0",strRqTopicID,strUserMemberID)
			  if subscription_id <> 0 then
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=9&amp;cid=" & subscription_id &"');""><img src=""themes/" &  strTheme & "/icons/unsubscribe.gif"" title=""UnSubscribe from this Topic"" alt='unsubscribe' border='0'>&nbsp;UnSubscribe</a>" 
			  else
				Response.Write " <a href=""javascript:;"" onclick=""javascript:openWindow3('forum_pop.asp?mode=7&amp;cmd=3&amp;cid="& strRqTopicID &"');""><img src=""themes/" &  strTheme & "/icons/subscribe.gif"" title=""Subscribe to this Topic"" alt='subscribe' border='0'>&nbsp;Subscribe To Topic</a>" 
			  end if
			end if %>			
<%	end if %>
    
<% 
end sub 

sub AdminOptions() 
 %>
<%	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then
		if (staCatStatus = 0) then 
			if (hasAccess(1)) then %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Category&CAT_ID=<% =strRqCatID %>')"><img border="0" src="images/icons/icon_folder_unlocked.gif" title="Un-Lock Category" alt="Un-Lock Category" height=15 width=15></a>
<%			else %>
    <img border="0" src="images/icons/icon_folder_unlocked.gif" title="Cateogry Locked" alt="Cateogry Locked" height=15 width=15>
<%			end if
		else 
			if (staFStatus = 0) then %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Forum&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')"><img border="0" src="images/icons/icon_folder_unlocked.gif" title="Un-Lock Forum" alt="Un-Lock Forum" height=15 width=15></a>
<%			else
				if (staTStatus <> 0) then %>
    <a href="JavaScript:openWindow('forum_pop_lock.asp?mode=Topic&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="images/icons/icon_folder_locked.gif" title="Lock Topic" alt="Lock Topic" height=15 width=15></a>
<%				else %>
    <a href="JavaScript:openWindow('forum_pop_open.asp?mode=Topic&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="images/icons/icon_folder_unlocked.gif" title="Un-Lock Topic" alt="Un-Lock Topic" height=15 width=15></a>
<%				end if
			end if
		end if %>
<%		if ((staCatStatus <> 0) and (staFStatus <> 0) and (staTStatus <> 0)) or (AdminAllowed = 1) then %>
    <a href="forum_post.asp?method=EditTopic&REPLY_ID=<% =strRqTopicID %>&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_folder_pencil.gif" alt="Edit Topic" title="Edit Topic" border="0" hspace="0"></a>
<%		end if %>
    <a href="JavaScript:openWindow('forum_pop_delete.asp?mode=Topic&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="images/icons/icon_folder_delete.gif" title="Delete Topic" alt="Delete Topic" height=15 width=15></a>
    <a href="forum_post.asp?method=Topic&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="images/icons/icon_folder_new_topic.gif" title="New Topic" alt="New Topic" height=15 width=15 border=0></a>
    <a href="forum_post.asp?method=Reply&TOPIC_ID=<% =strRqTopicID %>&FORUM_ID=<% =strRqForumID %>&CAT_ID=<% =strRqCatID %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="images/icons/icon_reply_topic.gif" alt="Reply to Topic" title="Reply to Topic" height=15 width=15 border=0></a>
<%	end if %>
    

<% 
end sub 


sub Paging()

	if (IsNumeric(intPagingLinks) = 0) AND (Trim(intPagingLinks) = "") then intPagingLinks = 10
	if (maxpages > 1) and (Trim(strQS) <> "") then

		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0"" valign=""top"" align=""center"">" & vbCrLf &_
			"<tr align=""center"">" & vbCrLf)

		if maxpages > 10 then
			Response.Write("<td>")
			Response.Write("<form method=""post"" name=""pagelist"" id=""pagelist"" action=""" & scriptname & "?n=0"& strQS & """>")
			Response.Write("<table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""right""><tr><td><b>Go to Page</b>:&#160;</td><td>")
			Response.Write("<select name=""whichpage"" onchange=""jumpToPage(this)"" style=""font-size:10px;"">" & vbCrLf)
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """>&#160;-" & vbCrLf)
			pgeselect = ""
			if pgenumber = mypage then pgeselect = " selected"
			Response.Write("<option value=""" & scriptname & "?whichpage=1" & strQS & """" & pgeselect & ">1" & vbCrLf)
			for counter = 1 to (maxpages/5)
				pgenumber = (counter*5)
				pgeselect = ""
				if pgenumber = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & pgenumber & strQS & """" & pgeselect & ">" & pgenumber & vbCrLf)
			next
			if (maxpages mod 5) > 0 then
				pgeselect = ""
				if maxpages = mypage then pgeselect = " selected"
				Response.Write("<option value=""" & scriptname & "?whichpage=" & maxpages & strQS & """" & pgeselect & ">" & maxpages & vbCrLf)
			end if
			Response.Write("</select>")
			Response.Write("</td></tr></table>" & vbCrLf)
			Response.Write("</form>")
			Response.Write("</td><td nowrap>&#160;&#160;</td>")
		end if
		dim pgelow, pgehigh, pgediv

		if maxpages > intPagingLinks then
			pgediv = Int(Abs(intPagingLinks/2))
			pgelow = mypage - pgediv
			pgehigh = mypage + (intPagingLinks - (pgediv + 1))
			if pgelow < 1 then
				pgelow = 1
				pgehigh = pgelow + (intPagingLinks - 1)
			end if
			if pgehigh > maxpages then
				pgehigh = maxpages
				pgelow = pgehigh - (intPagingLinks - 1)
			end if
		else
			pgelow = 1
			pgehigh = maxpages
		end if

		Response.Write("<td nowrap>")
		if pgelow > 1 then
			response.write("<a href=""" & scriptname & "?whichpage=1" & strQS & """>First Page</a>&#160;&#160;")
		end if
		' Previous Page Link
		if mypage = 1 then
			response.write("Previous Page")
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage - 1) & strQS & """>Previous Page</a>")
		end if
		response.write("&#160;&#160;&#160;")
		
		for counter = pgelow to pgehigh
			if counter <> mypage then
				response.write("&#160;<a href=""" & scriptname & "?whichpage=" & counter & strQS & """>" & counter & "</a>")
			else
				response.write("&#160;<span class=""fAlert"">[" & counter & "]</span>")
			end if
			if counter < pgehigh then response.write("&#160;")
		next
		'Response.Write("</td><td nowrap>&#160;")
		
		response.write("&#160;&#160;&#160;")
		
		' Next Page Link
		if mypage = maxpages then
			response.write("Next Page")
		else
			response.write("<a href=""" & scriptname & "?whichpage=" & (mypage + 1) & strQS & """>Next Page</a>")
		end if
		if pgehigh < maxpages then
			response.write("&#160;&#160;<a href=""" & scriptname & "?whichpage=" & maxpages & strQS & """>Last Page</a>")
		end if
		Response.Write("</td></tr></table>")


	else
		response.write("<div style=""font-size:6px;"">&#160;</div>")
	end if

end sub

Sub Topic_nav()    
    set rsLastPost = Server.CreateObject("ADODB.Recordset")
    
    strSql = "SELECT T_LAST_POST FROM " & strTablePrefix & "TOPICS " 
    strSql = strSql & "WHERE TOPIC_ID = " & strRqTopicID
                
    set rsLastPost = my_Conn.Execute (StrSql)
    
    T_LAST_POST = rsLastPost("T_LAST_POST")
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
    strSql = strSql & "WHERE T_LAST_POST > '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & strRqForumID
    strSql = strSql & " ORDER BY T_LAST_POST;"
                
    set rsPrevTopic = my_conn.Execute (strSQL)
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
    strSql = strSql & "WHERE T_LAST_POST < '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & strRqForumID
    strSql = strSql & " ORDER BY T_LAST_POST DESC;"
                
    set rsNextTopic = my_conn.Execute (strSQL)
    
    if rsPrevTopic.EOF then
        prevTopic = "<img src=""images/icons/icon_blank.gif"" height=15 width=15 title=""Previous Topic"" alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        prevTopic = "<a href=forum_topic.asp?cat_id=" & strRqCatID & _
                    "&FORUM_ID=" & strRqForumID & _
                    "&TOPIC_ID=" & rsPrevTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsPrevTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""images/icons/icon_topic_prev.gif"" title=""Previous Topic"" alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
                    
    if rsNextTopic.EOF then
        NextTopic = "<img src=""images/icons/icon_blank.gif"" height=15 width=15 title=""Previous Topic"" alt=""Previous Topic"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        NextTopic = "<a href=forum_topic.asp?cat_id=" & strRqCatID & _
                    "&FORUM_ID=" & strRqForumID & _
                    "&TOPIC_ID=" & rsNextTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsNextTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""images/icons/icon_topic_next.gif"" title=""Next Topic"" alt=""Next Topic"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
    
    Response.Write (prevTopic & "<b><b>&nbsp;Topic&nbsp;</b>" & nextTopic)
    
 '   rsLastPost.close
 '   rsPrevTopic.close
 '   rsNextTopic.close
    set rsLastPost = nothing
    set rsPrevTopic = nothing
    set rsNextTopic = nothing
    
end sub

function ProcessMsg2(strMsg, nReplyID, sig, aut_id)
	signature = GetSig(getMemberName(aut_id))
	if sig = 1 and signature <> ""  then
		strMsg = strMsg & vbCrLf & vbCrLf & ChkString(signature, "signature" )
	end if
	ProcessMsg2 = replace(ProcessMsg(strMsg, nReplyID),"''","'")	
end function

sub getAuthorInfo(uniqueID) %>
		<center><% If trim(strdbntusername) <> "" Then %><a href="javascript:;" onClick="swapIMAV('<% =uniqueID %>')" title="Click for Member Contact info"><img src="images/icons/icon_group.gif" height=15 width=15 title="Click for Member Contact info" alt="Click for Member Contact info" border="0"></a><br /><% End If %>
        <% 
		strIMmsg = "View " & ChkString(rs("M_NAME"),"display") & "'s profile" %>
		<a href="cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>" title="<%= strIMmsg %>">	
	  <b><%= displayName(ChkString(rs("M_NAME"),"display"),rs("M_GLOW")) %></b></a><br />
<% If trim(strdbntusername) <> "" Then %>
	  <%= chkIsOnline(rs("M_NAME"),2) %>
<DIV ID="imPanel<% =uniqueID %>" STYLE="display:none;">
  <table width="100" align="center">
	<tr>
		<td width="100%" valign="top" align="center" colspan="2" nowrap>
		<span class="fSmall">
		<% If trim(rs("M_FIRSTNAME")) <> "" or trim(rs("M_LASTNAME")) <> "" Then 
				response.Write(rs("M_FIRSTNAME") & " " & rs("M_LASTNAME") & "<BR>")
			  End If 
				'response.Write("status: " & chkIsOnline(arrCurOnline(onl,0),1) & "<BR>")
				response.Write("joined: " & split(strtodate(rs("M_DATE"))," ")(0) & "<BR>")
			  If trim(rs("M_CITY")) <> "" Then 
				response.Write("city: " & rs("M_CITY") & "<BR>")
			  end if
			  If trim(rs("M_STATE")) <> "" Then 
				response.Write("state: " & rs("M_STATE") & "<BR>")
			  end If
			  If Trim(rs("M_COUNTRY")) <> "" Then 
			       Response.Write(rs("M_COUNTRY") & "<BR>")
			  end If
		%>
		</span>
		</td>
	</tr>
	<tr>
		<td width="50%" align="right" nowrap>
<%			hasIM = "" %>
		<a href="cp_main.asp?cmd=8&member=<% =rs("MEMBER_ID") %>"> <small>Bio&nbsp;</small><img src="images/icons/icon_profile.gif" height=15 width=15 title="View Profile" alt="View Profile" border="0" align="absmiddle"></a>&nbsp;</td><td width="50%" align="left" nowrap><% if chkApp("PM","USERS") and rs("M_PMSTATUS") = 1 and rs("M_PMRECEIVE") = 1 then %>
		&nbsp;<a href="Javascript:;" onclick="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getmemberid(rs("M_NAME")) %>');"><img src="images/icons/pm.gif" height=17 width=11 title="Send a private message to <% =rs("M_NAME")%>" alt="Send a private message to <% =rs("M_NAME")%>" border="0" align="absmiddle"><small>&nbsp;PM</small></a><% else %>&nbsp;<% end if %></td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			hasIM = "1" %>
<%	if (lcase(strEmail) = "1" and rs("M_HIDE_EMAIL") = 0) then 
			if (hasAccess(2)) or (not hasAccess(2) and  strLogonForMail <> "1") then  %>
				<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><small>Email&nbsp;</small><img src="images/icons/icon_email.gif" height=15 width=15 title="Email Poster" alt="Email Poster" border="0" align="absmiddle"></a>&nbsp;
<%			hasIM = "1" %>
<%		end if
		else %>
			&nbsp;<img src="images/spacer.gif" height=15 width=15 title="No Email Available" alt="No Email Available" border="0" align="absmiddle">&nbsp;
<%			hasIM = "1" %>
<%	end if %>  
		</td><td width="50%" nowrap align="left">
<%			if strHomepage = "1" then %>
<%				if rs("M_Homepage") <> " " then %>
        &nbsp;<a href="<% =ChkString(rs("M_Homepage"),"displayimage") %>" target="_blank"><img src="images/icons/icon_homepage.gif" height=15 width=15 alt="Visit <% = ChkString(rs("M_NAME"),"display") %>'s Homepage" title="Visit <% = ChkString(rs("M_NAME"),"display") %>'s Homepage" border="0" align="absmiddle"><small>&nbsp;Web</small></a>
<%			hasIM = "1" %>
<%				end if %>
<%			end if %></td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			if (strMSN = "1") then %>
<%				if Trim(rs("M_MSN")) <> "" then %>
        <a href="JavaScript:;" onClick="openWindow('pop_portal.asp?cmd=7&mode=3&msn=<% =ChkString(replace(rs("M_MSN"),"@","[no-spam]@"), "displayimage") %>&M_NAME=<% =ChkString(rs("M_NAME"), "JSurlpath") %>')"><small>msn&nbsp;</small><img src="images/icons/icon_msn.gif" title="MSN id" alt="MSN id" border="0" align="absmiddle"></a>&nbsp;
<%			hasIM = "1" %>
<%				end if %>
<%			end if %>
		</td><td width="50%" align="left" nowrap>
<%			if (strAIM = "1") then %>
<%				if Trim(rs("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=2&AIM=<% =ChkString(rs("M_AIM"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="images/icons/icon_aim.gif" height=15 width=15 title="AIM id" alt="AIM id" border="0" align="absmiddle"><small>&nbsp;aim</small></a>
<%			hasIM = "1" %>
<%				end if %>
<%			end if %></td></tr>
	<tr><td width="50%" align="right" nowrap>
<%			if strICQ = "1" then %>
<%			  if Trim(rs("M_ICQ")) <> "" then %>
        <a href="JavaScript:openWindow('pop_portal.asp?cmd=7&mode=1&ICQ=<% =ChkString(rs("M_ICQ"), "displayimage") %>&M_NAME=<% =ChkString(rs("M_NAME"),"JSurlpath") %>')"><small>icq&nbsp;</small><img src="http://web.icq.com/whitepages/online?icq=<% = ChkString(rs("M_ICQ"),"display")  %>&img=5" alt="ICQ number" title="ICQ number" border="0" align="absmiddle"></a>&nbsp;
<%			hasIM = "1" %>
<%			  end if %>
<%			end if %>
		</td><td width="50%" align="left" nowrap>
<%			if strYAHOO = "1" then %>
<%			  if Trim(rs("M_YAHOO")) <> "" then 
					if instr(rs("M_YAHOO"),"@") then
					Yhoo = ChkString(replace(rs("M_YAHOO"),"@","[no-spam]@"), "display") 
					else
					Yhoo = ChkString(rs("M_YAHOO"), "display")
					end if %>
        &nbsp;<a href="http://edit.yahoo.com/config/send_webmesg?.target=<% =ChkString(rs("M_YAHOO"), "JSurlpath") %>&.src=pg" target="_blank"><img src="images/icons/icon_yahoo.gif" height=15 width=15 alt="Send <% =ChkString(rs("M_NAME"),"display")  %> a Yahoo! Message" title="Send <% =ChkString(rs("M_NAME"),"display")  %> a Yahoo! Message" border="0" align="absmiddle"><small>ahoo</small></a>
<%			hasIM = "1" %>
<%			  end if %>
<%			end if %>
		</td>
	</tr>
</table>
</DIV>
<% End If %>

<DIV ID="avPanel<% =uniqueID %>" STYLE="display:block;">
  <table width="100%" align="center">
	<tr>
		<td width="50%" valign="top" align="center">
<%		if strShowRank = 1 or strShowRank = 3 then
        Response.Write "<span class=""fSmall"">" & ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") & "</span><br />" & vbcrlf
		end if
		if strShowRank = 2 or strShowRank = 3 then
        Response.Write getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) & vbcrlf
		end if 
		  dnrLvl = getDonor_Level(rs("MEMBER_ID"))
		  if dnrLvl <> "" then
		  response.Write("<br />" & dnrLvl)
		  end if
		%><br />
        <% if len(Trim(rs("M_AVATAR_URL"))) > 10 then %>
		<%	' Get Avatar Settings from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR2.A_WSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_HSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"

		set rsav = my_Conn.Execute (strSql) 
		
		strAvatarURL = chkString(rs("M_AVATAR_URL"), "")
		strAvatarURL = Replace(strAvatarURL, "javascript", "")
		strAvatarURL = Replace(strAvatarURL, "alert", "")
		strAvatarURL = Replace(strAvatarURL, "<", "")
		strAvatarURL = Replace(strAvatarURL, ">", "")
		%>
        <img src="<% =strAvatarURL %>" width=<% =rsav("A_WSIZE") %> height=<% =rsav("A_HSIZE") %> border=<% =rsav("A_BORDER") %> hspace="0" vspace="5" >
		<% 
        rsav.close 
        set rsav = nothing 
        %><% end if %>
                <br /><span class="fSmall"><% =rs("M_POSTS") %> Posts</span>
	  <%If Trim(rs("CO_FLAG")) <> "" Then%>
			    <br /><img src="<%=rs("CO_FLAG")%>" title="<% =rs("M_COUNTRY") %>" />
			   <% End If%>
	  </td></tr></table>
</DIV>
<%
end sub
 %>