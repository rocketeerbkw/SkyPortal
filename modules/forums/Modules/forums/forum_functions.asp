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

function ChkQuoteOk(fString)
	ChkQuoteOk = not(InStr(1, fString, "'", 0) > 0)
end function

Function RemoveHTML(strText)
	' Removed the following from TAGLIST: B P I U since we don't allow single letter searches so these tags
	' shouldn't interfere and leaving them in keeps more of the original message formatting.
    Const TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;BASE;BASEFONT;BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;IFRAME;IMG;INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;PARAM;PLAINTEXT;PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;UL;VAR;WBR;XMP;"
    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"

    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
	' find tag beginning
    nPos1 = InStr(strText, "<")
	
	' while there is a tag to process...
    Do While nPos1 > 0
	
		' look for tag end
        nPos2 = InStr(nPos1 + 1, strText, ">")
		
		' if ending tag found...
        If nPos2 > 0 Then
		
			' get the tag name (minus the "<" and ">"
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
			
			' Replace CR/LF with spaces
	    	strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

			' look for trailing space
            nPos3 = InStr(strTagName, " ")
			
			' trailing space found?
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)  ' yes, clip trailing space
            End If
            
			' Does tag begin with "/"? (is it a closing tag)
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
           Else
		   		' no closing tag so this is a block search?
                bSearchForBlock = True
            End If
            
			' Is the the tag in the tag list?
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
			
				' yes, set remove flag
                bRemove = True
				
				' If searching for block removal...
                If bSearchForBlock Then
				
					' is tag in block list?
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
					
						' yes, set positions
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
                bRemove = False
            End If
            
			' Here it is. Is the item to be removed?
            If bRemove Then
			
				' copy over the first part (non-tag part) to the return string
                strResult = strResult & Left(strText, nPos1 - 1)
				
				
				' increment our array holder. (first time in it is -1 so incrementing it to zero)
				nUniqueHolderIndex = nUniqueHolderIndex + 1
				
				' allocate another position in the array for the new element
				ReDim Preserve arrHTMLReplacements(nUniqueHolderIndex+1)
				
				' store the original text
				arrHTMLReplacements(nUniqueHolderIndex) = Mid(strText,nPos1,nPos2-nPos1+1) 
				
				' build the temporary holding string
				sTmp = strUniqueHolderPrefix & CStr(nUniqueHolderIndex) & strUniqueHolderSuffix
				
				' Copy it to the return string
				strResult = strResult & sTmp
				
				' Skip over the tag part and append the remaining part to the return string
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    strResult = strResult & strText
    
    RemoveHTML = strResult
End Function

function ProcessMsg(strMsg, nReplyID)
	' all msgs get formatted for output
	ProcessMsg = FormatStr(strMsg)
	
	' if there are no keywords to process, then exit immediately
	if bKeywordsPresent=false then
		exit function
	end if
	
	' There are keywords so see if this msg is one that should be 
	' processed for color coding
	bFound = false
	nPos=0
		
	for i = 0 to nRepliesArrayUpperBound
		if nReplyID = CInt(arrReplyIDs(i)) then
			bFound = true
			nPos = i
			exit for
		end if
	next
		
	' was the passed message in the list of hit IDs?
	if bFound = false then
		exit function
	end if

	' this is a reply requiring a color code so get to it...
	
	' First remove all the html tags that will get in the way when we color code
	ProcessMsg = RemoveHTML(ProcessMsg)

	' indicates whether we shoudl show a special note regarding color coding
	' if some of the search words don't appear in the final output
	bShowHighlightNote=false
	nColorHits=0 ' used for individual word color code counting
	
	' Now find and highlight the keywords
	if strSearchType = "phrase" then
		' color code the phrase
		ProcessMsg = Highlight(ProcessMsg, strSearchStrings, strPreHighlight, strPostHighlight)
		
		' no color coding done (tags containing keyword(s) removed before they could be color coded)
		if bHighlightMade=false then bShowHighlightNote=true
	else
		for each word in arrSearchStrings
			' color code the word
			ProcessMsg = Highlight(ProcessMsg, word, strPreHighlight, strPostHighlight)
			
			if bHighlightMade=true then nColorHits = nColorHits+1
		next
	end if
	
	' check non-phrase processing
	if strSearchType <> "phrase" then
		if strSearchType = "and" then ' requires all words to be hit
			if nSearchStringsTopIndex <> (nColorHits-1) then bShowHighlightNote=true
		else
			if nColorHits=0 then bShowHighlightNote=true ' requires any words, so if NO words hit, show msg
		end if
	end if
	

	if nUniqueHolderIndex <> -1 then
		for x = 0 to (nUniqueHolderIndex+1)
			strTmp = strUniqueHolderPrefix & CStr(x) & strUniqueHolderSuffix
			ProcessMsg = Replace(ProcessMsg, strTmp, arrHTMLReplacements(x))
		next  
	end if
	
	' If not all hits were highlighted, show message
	if bShowHighlightNote=true then
		ProcessMsg = ProcessMsg + strNoHighlightMsg
	end if
	
end function

function GetWhichPage(nReplyID)

	if nReplyID < 0 then
		GetWhichPage = sWhichPage & CStr(1)
		Response.Write("ID less than zero")
		exit function
	end if

	' init
	nIndex = -1
	
	' Get the index of the nReplyID 
	' examine each array occurrence to try and match the passed replyID
	for x = 0 to nAllRepliesArrayUpperBound
		if nReplyID = CInt(arrAllReplyIDs(x)) then
			nIndex = x
			exit for
		end if
	next

	' Make nIndex 1-based for mathpurposes
	nIndex = nIndex + 1

	if nIndex <= CInt(strPageSize) then
		GetWhichPage = sWhichPage & "1"
	else
		' how appropriate...MOD
		nMod = CInt(nIndex) MOD CInt(strPageSize)
		
		nPage = nIndex \ CInt(strPageSize)
		if nMod <> 0 then
			nPage = nPage + 1
		end if
	
		GetWhichPage = sWhichPage & CStr(nPage)
	end if
		
end function

'##############################################
'##              Private Forums              ##
'##############################################

sub chkUser4()
	if hasAccess(1) then 
		exit sub
	end if
	if len(Request.QueryString("FORUM_ID")) < 1 then
	  closeAndGo("fhome.asp")
	end if
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.Forum_ID = " & chkstring(Request.QueryString("FORUM_ID"), "numeric")
	set rsStatus = my_conn.Execute (strSql)
	dim Users
	If cint(rsStatus("F_PRIVATEFORUMS")) <> 0 then
			Select case cint(rsStatus("F_PRIVATEFORUMS"))
				case 0
					'## Do Nothing
				case 1, 6 '## Allowed Users
					UserNum = strUserMemberID
					MatchFound = isAllowedMember(chkstring(Request.QueryString("FORUM_ID"), "numeric"), cint(UserNum))
					if MatchFound then
						exit sub
					else
						doNotAllowed
						closeAndGo("stop")
					end if
				case 2 '## password
					select case chkString(Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")), "")
						case rsStatus("F_PASSWORD_NEW")
							'## OK
						case else
							if trim(chkString(Request("pass"), "urlpath")) = "" then
								doPasswordForm
								closeAndGo("stop")
							else
								if chkString(Request("pass"), "urlpath") <> rsStatus("F_PASSWORD_NEW") then
									Response.Write "Invalid password! <a href='" & chkString(Request.ServerVariables("HTTP_REFERER"), "refer") & "'>Back</a>"
									closeAndGo("stop")
								else
									Response.Cookies(strUniqueID & "User").Path = strCookieURL
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
								end if
							end if
					end select
				case 3    '## Either Password or Allowed
					UserNum = strUserMemberID
					MatchFound = isAllowedMember(chkString(Request.QueryString("FORUM_ID"),"numeric"), cint(UserNum))
					if MatchFound then
						exit sub
					else
					select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							'## OK
						case else
							if trim(chkString(Request("pass"), "urlpath")) = "" then
								doLoginForm
								closeAndGo("stop")
							else
								if chkString(Request("pass"), "urlpath") <> rsStatus("F_PASSWORD_NEW") then
									Response.Write "Invalid password! <a href='" & chkString(Request.ServerVariables("HTTP_REFERER"), "refer") & "'>Back</a>"
									closeAndGo("stop")
								else
									Response.Cookies(strUniqueID & "User").Path = strCookieURL
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
								end if
							end if
					end select
					end if
				'## code added 07/13/2000
				case 7    '## members or password
					if (strDBNTUserName = "") then
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								'## OK
							case else
								if trim(chkString(Request("pass"), "urlpath")) = "" then
									doLoginForm
									closeAndGo("stop")
								else
									if trim(chkString(Request("pass"), "urlpath")) <> rsStatus("F_PASSWORD_NEW") then
										Response.Write "Invalid password! <a href='" & chkString(Request.ServerVariables("HTTP_REFERER"), "refer") & "'>Back</a>"
										closeAndGo("stop")
									else
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
									end if
								end if
						end select
					end if
				'## end code added 07/13/2000
				case 4, 5 '## members only
					if strDBNTUserName = "" then
						doNotLoggedInForm
						closeAndGo("stop")
					end if
				case 8, 9
					NTGroupSTR = Split(strNTGroupsSTR, ", ")
					NTGroupDBSTR = Split(rsStatus("F_PASSWORD_NEW"), ", ")
						For i = 0 to ubound(NTGroupDBSTR)
							for j = 0 to ubound(NTGroupSTR)
								if NTGroupDBSTR(i) = NTGroupSTR(j) then
									exit SUB
								end if
							next
						next
					doNotAllowed
					closeAndGo("stop")
				case else
					Response.Write "<br />ERROR: Invalid forum type: " & rsStatus("F_PRIVATEFORUMS")
					closeAndGo("stop")
			end select
	end if
	set rsStatus = nothing
end sub

'##############################################
'##           Multi-Moderators               ##
'##############################################

function chkForumModerator(fForum_ID, fMember_Name)
	strSql = "SELECT * FROM " & strTablePrefix & "MODERATOR"
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID
	strSql = strSql & " AND MEMBER_ID = " & getMemberID(fMember_Name)
	set rsChk = my_Conn.Execute (strSql)
	if rsChk.eof then
		chkForumModerator = "0"
	else
		chkForumModerator = "1"
	end if 
	set rsChk = nothing
end function

function listForumModerators(fForum_ID)
	tmpModList = ""
	strSql = "SELECT * FROM " & strTablePrefix & "MODERATOR"
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID
	set rsChk = my_Conn.Execute (strSql)
	if not rsChk.EOF then
	  tmpModList = getMemberName(rsChk("MEMBER_ID"))
	  rsChk.MoveNext
	  do until rsChk.EOF
		tmpModList = tmpModList & ", " & getMemberName(rsChk("MEMBER_ID"))
		rsChk.MoveNext
	  loop
	end if
	set rsChk = nothing
	listForumModerators = tmpModList
end function

function ChkIsNew(fDateTime)
	if strHotTopic = "1" then
		if fDateTime > Session(strUniqueID & "last_here_date") then
			if rs("T_REPLIES") >= intHotTopicNum then
			        ChkIsNew =  "<img src=images/icons/icon_folder_new_hot.gif height=15 width=15 alt=HotTopic border=0 />"
			else
			        ChkIsNew =  "<img src=images/icons/icon_folder_new.gif height=15 width=15 alt=NewTopic border=0 />"
			end if
		else
			if rs("T_REPLIES") >= intHotTopicNum then
			        ChkIsNew =  "<img src=images/icons/icon_folder_hot.gif height=15 width=15 alt=HotTopic border=0 />"
			else
			        ChkIsNew = "<img src=images/icons/icon_folder.gif height=15 width=15 border=0 />" 
			end if
		end if
	else
		if fDateTime > Session(strUniqueID & "last_here_date") then
			ChkIsNew =  "<img src=images/icons/icon_folder_new.gif height=15 width=15 alt=NewTopic border=0 />" 
		else
			ChkIsNew = "<img src=images/icons/icon_folder.gif height=15 width=15 border=0 />" 
		end if
	end if
end function

function ChkUser(fName, fPassword)

	'
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & fPassword &"'"
	End IF
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		ChkUser = 0
	else
		if cstr(rsCheck("MEMBER_ID")) = Request.Form("Author") then
			ChkUser = 1 '## Author
		else
			Select case cint(rsCheck("M_LEVEL"))
				case 1
					ChkUser = 2 '## Normal User
				case 2
					ChkUser = 3 '## Moderator
				case 3
					ChkUser = 4 '## Admin
				case else
					ChkUser = cint(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if

	rsCheck.close
	set rsCheck = nothing

end function

function ChkUser3(fName, fPassword, fReply)
	'
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	StrSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then	
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & fPassword &"' "
	End If
	strSql = strSql & " AND   " & strTablePrefix & "REPLY.REPLY_ID = " & fReply
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		tmpChk = 0 '##  Invalid Password
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(rsCheck("R_AUTHOR")) then 
			tmpChk = 1 '## Author
		else
			Select case cint(rsCheck("M_LEVEL"))
				case 1
					tmpChk = 2 '## Normal User
				case 2
					tmpChk = 3 '## Moderator
				case 3
					tmpChk = 4 '## Admin
				case else
					tmpChk = cint(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing
	ChkUser3 = tmpChk
end function

'##############################################
'##                Do Counts                 ##
'##############################################

sub DoULastPost(sUser_Name)
	'Updates the M_LASTPOSTDATE in the MEMBERS table
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LASTPOSTDATE = '" & strCurDateString & "' "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & sUser_Name & "'"
	executeThis(strSql)
end sub

sub deleteCount(sUser_ID)
	' - Update Total Post for user
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & strMemberTablePrefix & "MEMBERS.M_POSTS - 1 "
	strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_REP = " & strMemberTablePrefix & "MEMBERS.M_REP - 1 "
	strSql = strSql & " WHERE MEMBER_ID = " & sUser_ID
	my_Conn.Execute (strSql)
end sub

sub DoRepAdd(sUser_Name)
	' - Update Total Reputation for user ADD
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_REP = " & strMemberTablePrefix & "MEMBERS.M_REP + 1 "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & sUser_Name & "'"
	my_Conn.Execute (strSql)
end sub

sub DoPCount()
	' - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT + 1"
	my_Conn.Execute (strSql)
end sub

sub DoTCount()
	' - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET " & strTablePrefix & "TOTALS.T_COUNT = " & strTablePrefix & "TOTALS.T_COUNT + 1"
    my_Conn.Execute (strSql)
end sub

sub DoUCount(sUser_Name)
	' - Update Total Post for user
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & strMemberTablePrefix & "MEMBERS.M_POSTS + 1 "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & sUser_Name & "'"
	my_Conn.Execute (strSql)
end sub

sub doNotAllowed()
%>
<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
<p align="center"><span class="fTitle">
<%= txtNoForumAcc %>.
</span></p>
<p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBackData %></a></p>
<p align="center"><a href="default.asp"><%= txtReturnHome %></a></p>
<!--INCLUDE FILE="inc_footer.asp"-->
<%
end sub

sub doPasswordForm()
%>
<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
<p align="center">
<span class="fTitle"><%= txtEnterForumPass %>.</span>
<form action="<% =Request.ServerVariables("SCRIPT_NAME") %>" id=form62 name=form62>
<%
	for each q in Request.QueryString
		Response.Write "<input TYPE=hidden name=""" & chkstring(q, "hidden") & """ value=""" & chkstring(Request(q), "hidden") & """>"
	next
%>
<input class="textbox" name=pass type=password size="20" />
<input class="button" type=submit value=Enter id=submit61 name=submit61 />
</form>
</p>
<p align="center"><a href="JavaScript:history.go(-1)"><%= txtGoBackData %></a></p>
<p align="center"><a href="default.asp"><%= txtReturnHome %></a></p>
<!--INCLUDE FILE="inc_footer.asp"-->
<%
end sub
%>