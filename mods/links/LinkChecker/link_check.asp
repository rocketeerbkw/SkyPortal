<%
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|              Coded by Brandon Williams.             |'|'
'|'|           Copyright 2007 Brandon Williams.          |'|'
'|'|            Distributed under the MIT Open           |'|'
'|'|         Source License included with software.      |'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'

'Some parts Copyright (C) 2005-2006 Dogg Software All Rights Reserved

DIM blFlagBadLinks, blSendReport, strWhereToSend, strSendTo, strSendFrom

'Set this to true if you want to set a "Bad Link" flag for bad links (becomes an Admin Pending Task)
blFlagBadLinks = true

'Do you want to send a bad link report?
blSendReport = false
	'Set this to EMAIL or PM
	strWhereToSend = "PM"
	'Set this to Email Address (if sending to Email) or Member Name (if sending to PM)
	strSendTo = "Admin"
	'Set this to Email Address (if sending to Email) or Member Name (if sending to PM)
	strSendFrom = "Admin"
	
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''DO NOT EDIT BELOW THIS LINE UNLESS YOU KNOW WHAT YOU ARE DOING'''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'This function will take either an array or comma delimited
'string of URLs, and return a multidimensional array where
'(x,0) = URL (x,1) = Status
'The time it took to run the check is returned in (0,x) where
'x = the last array item
function batchLinkCheck(arrLinks)
	Dim oXmlHttp, oReturnStatus, arrReturn, linkStartTime, linkTime
	
	linkStartTime = timer
	
	if NOT isArray(arrLinks) then
		arrLinks = split(arrLinks, ",")
	end if

	Set oXmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
	on error resume next 
	Redim arrReturn(1,UBOUND(arrLinks) + 1)
		for i=0 to (UBOUND(arrLinks))
			oXmlHttp.OPEN "HEAD", TRIM(arrLinks(i)), False
			oXmlHttp.SEND
			
			if err.number <> 0 then
				arrReturn(0,i) = TRIM(arrLinks(i))
				arrReturn(1,i) = 000
			else			
				oReturnStatus = oXmlHttp.STATUS
			
				arrReturn(0,i) = TRIM(arrLinks(i))
				arrReturn(1,i) = oReturnStatus
			end if
			
		next
	Set oXmlHttp = Nothing
	
	linkTime = formatnumber((timer - linkStartTime),3)
	arrReturn(0,Ubound(arrReturn,2)) = linkTime
	arrReturn(1,ubound(arrReturn,2)) = 999
	
	batchLinkCheck = arrReturn
	
end function

'This function is not my original work. Credits for the original author
'are not available because I don't know who that is.  You can email me
'if you are the author at rocketeerbkw@yahoo.com
function arraySort( arToSort, sortBy, compareDates )
	Dim c, d, e, smallestValue, smallestIndex, tempValue
	
	For c = 0 To uBound( arToSort, 2 ) - 1
		smallestValue = arToSort( sortBy, c )
		smallestIndex = c
		
		For d = c + 1 To uBound( arToSort, 2 )
			if not compareDates then
				if strComp( arToSort( sortBy, d ), smallestValue ) < 0 Then
					smallestValue = arToSort( sortBy, d )
					smallestIndex = d
				end if
			else
				if not isDate( smallestValue ) then
					arraySort = arraySort( arToSort, sortBy, false)
					exit function
				else
					if dateDiff( "d", arToSort( sortBy, d ), smallestValue ) > 0 Then
						smallestValue = arToSort( sortBy, d )
						smallestIndex = d
					end if
				end if
			end if
		Next
		
		if smallestIndex <> c Then 'swap
			For e = 0 To uBound( arToSort, 1 )
				tempValue = arToSort( e, smallestIndex )
				arToSort( e, smallestIndex ) = arToSort( e, c )
				arToSort( e, c ) = tempValue
			Next
		End if
	Next
end function

'This function will disable links in a skyportal database
'array should be in the same input as outputed by batchLinkCheck()
function disableLinks(arrLinks)
	Dim dlSQL
	if NOT isArray(arrLinks) then
		exit function
	end if
	
	dlSQL = "UPDATE LINKS SET BADLINK = 1 WHERE "
	
	for i=0 to (ubound(arrLinks, 2) - 1)
		'check for bad links using status code
		if (arrLinks(1,i) = 404) or (arrLinks(1,i) = 410) or (arrLinks(1,i) = 400) or (arrLinks(1,i) = 503) then
			'build sql
			if i=(ubound(arrLinks, 2) - 1) then
				dlSQL = dlSQL & "URL = '" & arrLinks(0,i) & "'"
			else
				dlSQL = dlSQL & "URL = '" & arrLinks(0,i) & "' OR "
			end if
		end if
	next
	
	'execute sql
	executeThis(dlSQL)
	
end function

function status2Readable(statuscode)
	Select Case statuscode
		case 000
			status2Readable = "<b>Unable To Connect To Server!</b>"
		case 200
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-200"" target=""_blank"">OK</a>"
		case 301 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-301"" target=""_blank"">Moved Permanently</a>"
		case 307 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-307"" target=""_blank"">Temporary Redirect</a>"
		case 400 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-400"" target=""_blank"">Bad Request</a>"
		case 401 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-401"" target=""_blank"">Unauthorized</a>"
		case 403 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-403"" target=""_blank"">Forbidden</a>"
		case 404 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-404"" target=""_blank"">Not Found</a>"
		case 410 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-410"" target=""_blank"">Gone</a>"
		case 500 
			status2Readable = "<a href=""http://www.seoconsultants.com/tools/headers.asp#code-500"" target=""_blank"">Internal Server Error</a>"
		case 503
			status2Readable = "<a href="""" target=""_blank"">Temporarily Unavailable</a>"
		case else
			status2Readable = statuscode
	End Select
end function

'This function will check ALL links from the database and Email or PM
'the results to you
function checkLinkAll()
	DIM arrLinks2Check, arrLinksChecked, arrLinksDisplay
	'Get all the links from the database
	strSql = "SELECT URL, CATEGORY, PARENT_ID FROM LINKS"
	Set RS=Server.CreateObject("ADODB.Recordset")
	RS.Open strSql, my_Conn, 3
	
	if NOT (rs.BOF or rs.EOF) then
		rs.movefirst
		i=0
		ReDIM arrLinks2Check(rs.recordcount - 1)
		Do WHILE NOT rs.EOF			
			arrLinks2Check(i) = rs("url")
			i = i + 1
			rs.MoveNext
		Loop
	Else
		response.write("no links found")
	End if

	arrLinksChecked = batchLinkCheck(arrLinks2Check)
	
	arraySort arrLinksChecked, 1, false
	
	if blSendReport then
	
		strMessage = "<p>This is a Bad Link Report sent to you by the <a href=""http://skyportal.net/forum_topic.asp?TOPIC_ID=4003&FORUM_ID=69&CAT_ID=7&Topic_Title=Automatic+Link+Checking&Forum_Title=Mods+in+Development"" target=""_blank"" >Automatic Link Checker</a> mod by <a href=""http://skyportal.net/cp_main.asp?cmd=8&member=47"" target=""_blank"">Battousai</a>.  This report checked every link in your database and displays only those which don't work below.  You may click on the status for more information.</p><p>Generally a status in the 400 range is safe to delete.  A status in the 300 or 500 range should be investigated further before being purged.</p>"
		strMessage = strMessage & "<p><table border=""1"" cellpadding=""2"" cellspacing=""0"" rules=""rows"" frame=""box"">"
		strMessage = strMessage & "<tr style=""font-weight: bold;""><td>Status</td><td>Category</td><td>SubCategory</td><td>URL</td><td>Status Code</td></tr>"
		for i=0 to (ubound(arrLinksChecked, 2) - 1)
			if arrLinksChecked(1,i) <> 200 then
				strMessage = strMessage & "<tr>"
				strMessage = strMessage & "<td>" & status2Readable(arrLinksChecked(1,i)) & "</td>"
				rs.MoveFirst
				Do While NOT rs.EOF
					if rs("URL") = arrLinksChecked(0,i) then
						subSQL = "SELECT CAT_ID, SUBCAT_NAME FROM LINKS_SUBCATEGORIES WHERE SUBCAT_ID=" & rs("CATEGORY")
						set rsSub = my_Conn.execute(subSQL)
						  catId = rsSub("CAT_ID")
						  subcatName = rsSub("SUBCAT_NAME")
						set rsSub = nothing
						
						catSQL = "SELECT CAT_NAME FROM LINKS_CATEGORIES WHERE CAT_ID = " & catId
						set rsCat = my_Conn.execute(catSQL)
						  catName = rsCat("CAT_NAME")
						set rsCat = nothing
	
						strMessage = strMessage & "<td><a href=""" & strHomeURL & "links.asp?cmd=1&cid=" & catId & """ target=""_blank"">" & catName & "</a></td><td><a href=""" & strHomeURL & "links.asp?cmd=2&cid=" & catId & "&sid=" & rs("CATEGORY") & """ target=""_blank"">" & subcatName & "</a></td>"
					end if
					rs.MoveNext
				Loop
				strMessage = strMessage & "<td><a href=""" & arrLinksChecked(0,i) & """ target=""_blank"">" & arrLinksChecked(0,i) & "</a></td>"
				strMessage = strMessage & "<td align=""right"">" & arrLinksChecked(1,i) & "</td"
				strMessage = strMessage & "</tr>"
			end if
		next
		strMessage = strMessage & "<tr><td style=""font-weight: bold;"">Links Checked:</td><td>" & ubound(arrLinksChecked, 2) - 1 & "</td><td style=""font-weight: bold;"">Time It Took:</td><td>" & arrLinksChecked(0,ubound(arrLinksChecked, 2)) & "</td></tr>"
		strMessage = strMessage & "</tr></table></p>"
		
		'Response.Write(chkString(strMessage, "message"))
		
		if strWhereToSend = "PM" then
			sendPMtoMember getMemberID(strSendTo),getMemberID(strSendFrom),"Bad Link Report " & formatDateTime(now(), 0),chkString(strMessage, "message"),0,"Err Msg"
			
		elseif strWhereToSend = "EMAIL" then
			sendOutEmail strSendTo, "Bad Link Report " & formatDateTime(now(), 0), strMessage, 2, 0
		end if
		
	end if 'send report
	
	if blFlagBadLinks then
		disableLinks(arrLinksChecked)
	end if
	
end function
%>