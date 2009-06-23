<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Classic ASP Form Creator v1.0
' Copyright David Angell, http://www.angells.com/FormCreator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/**
' * SkyPortal Forms Module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://www.opensource.org/licenses/mit-license.php MIT License
' */

Function GetID(FieldID)
    tmpGetID = trim(request.querystring(FieldID)&" ")
    if len(tmpGetID) > 0 then if not isnumeric(tmpGetID) then response.redirect "default.asp"
    GetID = tmpGetID
end function

'Replaced by SkyPortals builtin chkString()
'Function ChkString (fString, action)
'    dim tmpString
'	tmpString = trim(fString)
'    If tmpString = "" Then tmpString = " "
'    If isnull(tmpString) Then tmpString = " "
'    if action = "in" then
'    	tmpString = replace(tmpString, "+", "&#043;")
'    	tmpString = replace(tmpString, """", "&quot;")
'        tmpString = replace(tmpString, ">", "&gt;")
'        tmpString = replace(tmpString, "<", "&lt;")
'        tmpString = Replace(tmpString, "'", "&#039;")
'    else
'    	tmpString = replace(tmpString, "&#043;", "+")
'    	tmpString = replace(tmpString, "&quot;", """")
'        tmpString = replace(tmpString, "&gt;", ">")
'        tmpString = replace(tmpString, "&lt;", "<")
'        tmpString = Replace(tmpString, "&#039;", "'")
'    end if
'    ChkString = trim(tmpString)
'End Function

Function IsValidEmail(strEmail)
  Dim bIsValid
  bIsValid = True
  If Len(strEmail) < 6 Then
    bIsValid = False
  Else
    If Instr(1, strEmail, " ") <> 0 Then
      bIsValid = False
    Else
      If InStr(1, strEmail, "@", 1) < 2 Then
        bIsValid = False
      Else
        If InStrRev(strEmail, ".") < InStr(1, strEmail, "@", 1) + 2 Then
          bIsValid = False
        else
          if InStrRev(strEmail, ".") = len(strEmail) then
            bIsValid = False
          end if
        End If
      End If
    End If
  End If
  IsValidEmail = bIsValid
End Function

Function FormatPhoneNumber(strInput)
	Dim strTemp
	Dim strCurrentChar
	Dim I
	For I = 1 To Len(strInput)
		strCurrentChar = Mid(strInput, I, 1)
		If Asc("0") <= Asc(strCurrentChar) And Asc(strCurrentChar) <= Asc("9") Then
			strTemp = strTemp & strCurrentChar
		End If
	Next

	strInput = strTemp
	strTemp = ""

	If Len(strInput) = 11 And Left(strInput, 1) = "1" Then
		strInput = Right(strInput, 10)
	End If

	If Len(strInput) <> 10 Then
		strTemp = False
    else
    	strTemp = "("
    	strTemp = strTemp & Left(strInput, 3)
    	strTemp = strTemp & ") "
    	strTemp = strTemp & Mid(strInput, 4, 3)
    	strTemp = strTemp & "-"
    	strTemp = strTemp & Right(strInput, 4)
	End If

	FormatPhoneNumber = strTemp
End Function

Function FormatZipCode(strInput)
	Dim strTemp
	Dim strCurrentChar
	Dim I
	For I = 1 To Len(strInput)
		strCurrentChar = Mid(strInput, I, 1)
		If Asc("0") <= Asc(strCurrentChar) And Asc(strCurrentChar) <= Asc("9") Then
			strTemp = strTemp & strCurrentChar
		End If
	Next

	strInput = strTemp
	strTemp = ""

	If Len(strInput) <> 5 And Len(strInput) <> 9 Then
		strTemp = False
    else
        if len(strInput) = 9 then
        	strTemp = Left(strInput, 5) & "-" & Right(strInput, 4)
        else
        	strTemp = strInput
        end if
	End If

	FormatZipCode = strTemp
End Function

Function requiredfield (enteredfield, description)
dim temp
    if len(trim(enteredfield&" ")) < 1 then
        temp = "<li>" & description & " is a required field.</li>"
    else
        temp = ""
    end if
    requiredfield = msg & temp
end function

Function DanDate (strDate, strFormat)
	'call using:  strFormattedDate = DanDate(dtmDate, strFormat)
	'dtmDate should be a Date variable
	'strFormat should be a template for the output date.
    '%m Month as a decimal no. 02
    '%b Abbreviated month name Feb
    '%B Full month name February
    '%d Day of the month 23
    '%j Day of the year 54
    '%y Year without century 98
    '%Y Year with century 1998
    '%w Weekday as integer 5 (0 is Sunday)
    '%a Abbreviated day name Fri
    '%A Weekday Name Friday
    '%I Hour in 12 hour format 12
    '%H Hour in 24 hour format 24
    '%M Minute as an integer 01
    '%S Second as an integer 55
    '%P AM/PM Indicator PM
    '%% Actual Percent sign %%
    'The resulting string will be the same as the format string,
    'but with the following key characters replaced with the relevant date/time part
    'Example DanDate(dtmDate, "%b %d, %Y"), where dtmDate = 02/25/2005 = "Feb 25, 2005"

	Dim intPosItem
	Dim intHourPart
	Dim strHourPart
	Dim strMinutePart
	Dim strSecondPart
	Dim strAMPM


	If not IsDate(strDate) Then
		DanDate = strDate
		Exit Function
	End If

	intPosItem = Instr(strFormat, "%m")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("m",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%m")
	Loop

	intPosItem = Instr(strFormat, "%b")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%b")
	Loop

	intPosItem = Instr(strFormat, "%B")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        MonthName(DatePart("m",strDate),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%B")
	Loop

	intPosItem = Instr(strFormat, "%d")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("d",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%d")
	Loop

	intPosItem = Instr(strFormat, "%j")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("y",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%j")
	Loop

	intPosItem = Instr(strFormat, "%y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        Right(DatePart("yyyy",strDate),2) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%y")
	Loop

	intPosItem = Instr(strFormat, "%Y")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("yyyy",strDate) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%Y")
	Loop

	intPosItem = Instr(strFormat, "%w")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        DatePart("w",strDate,1) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%w")
	Loop

	intPosItem = Instr(strFormat, "%a")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),True) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%a")
	Loop

	intPosItem = Instr(strFormat, "%A")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & _
                        WeekDayName(DatePart("w",strDate,1),False) & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%A")
	Loop

	intPosItem = Instr(strFormat, "%I")
	Do While intPosItem > 0
		intHourPart = DatePart("h",strDate) mod 12
		if intHourPart = 0 then intHourPart = 12
                strFormat = Left(strFormat, intPosItem-1) & _
                        intHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%I")
	Loop

	intPosItem = Instr(strFormat, "%H")
	Do While intPosItem > 0
		strHourPart = DatePart("h",strDate)
		if strHourPart < 10 Then strHourPart = "0" & strHourPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strHourPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%H")
	Loop

	intPosItem = Instr(strFormat, "%M")
	Do While intPosItem > 0
		strMinutePart = DatePart("n",strDate)
		if strMinutePart < 10 then strMinutePart = "0" & strMinutePart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strMinutePart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%M")
	Loop

	intPosItem = Instr(strFormat, "%S")
	Do While intPosItem > 0
		strSecondPart = DatePart("s",strDate)
		if strSecondPart < 10 then strSecondPart = "0" & strSecondPart
                strFormat = Left(strFormat, intPosItem-1) & _
                        strSecondPart & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%S")
	Loop

	intPosItem = Instr(strFormat, "%P")
	Do While intPosItem > 0
		if DatePart("h",strDate) >= 12 then
			strAMPM = "PM"
		Else
			strAMPM = "AM"
		End If
                strFormat = Left(strFormat, intPosItem-1) & _
                        strAMPM & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%P")
	Loop

	intPosItem = Instr(strFormat, "%%")
	Do While intPosItem > 0
                strFormat = Left(strFormat, intPosItem-1) & "%" & _
                        Right(strFormat, Len(strFormat) - (intPosItem + 1))
		intPosItem = Instr(strFormat, "%%")
	Loop

	DanDate = strFormat

End Function
%>
