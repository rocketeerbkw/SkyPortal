<%
function DoRosterDropDown(fSql,fDisplayField,fValueField,fSelectValue,fName,fExtra,fFirstOption)
	Dim strOutput
	set rsdrop = my_Conn.execute(fSql)
	strOutput = strOutput & "<select name=""" & fName & """ " & fExtra & ">" & vbCrLf
	if rsdrop.EOF or rsdrop.BOF then
		strOutput = strOutput & "<option value=""0"">None</option>"  & vbCrLf
	else
		if trim(fFirstOption) <> "" then
			strOutput = strOutput & "<option value=""0"">" & fFirstOption & "</option>" & vbCrLf
		end if
		do until rsdrop.EOF
			if lcase(rsdrop(fValueField)) = lcase(fSelectValue) then
			  strOutput = strOutput & "<option value=""" & rsdrop(fValueField) & """ selected=""selected"">" & rsdrop(fDisplayField) & "</option>" & vbCrLf
			else
			  strOutput = strOutput & "<option value=""" & rsdrop(fValueField) & """>" & rsdrop(fDisplayField) & "</option>" & vbCrLf
			end if
			rsdrop.MoveNext
		loop
	end if
	strOutput = strOutput & "</select>" & vbCrLf
	rsdrop.Close
	set rsdrop = nothing
	
	DoRosterDropDown = strOutput
end function

function DoRosterDropDownSm(fTableName,fDisplayField,fValueField,fSelectValue,fName,fExtra,fFirstOption,fWhere,fOrderBy)
	Dim strOutput
    strSql = "SELECT " & fDisplayField & ", " & fValueField 
	strSql = strSql & " FROM " & fTableName
	if trim(fWhere) <>  "" then
		strSQL = strSQL & " WHERE " & fWhere
	end if
	if trim(fOrderBy) <> "" then
		strSQL = strSQL & " ORDER BY " & fOrderBy
	end if

	set rsdrop = my_Conn.execute(strSQL)
	strOutput = strOutput & "<select name=""" & fName & """ " & fExtra & ">" & vbCrLf
	if rsdrop.EOF or rsdrop.BOF then
		strOutput = strOutput & "<option value=""0"">None</option>"  & vbCrLf
	else
		if trim(fFirstOption) <> "" then
			strOutput = strOutput & "<option value=""0"">" & fFirstOption & "</option>" & vbCrLf
		end if
		do until rsdrop.EOF
			if lcase(rsdrop(fValueField)) = lcase(fSelectValue) then
			  strOutput = strOutput & "<option value=""" & rsdrop(fValueField) & """ selected=""selected"">" & rsdrop(fDisplayField) & "</option>" & vbCrLf
			else
			  strOutput = strOutput & "<option value=""" & rsdrop(fValueField) & """>" & rsdrop(fDisplayField) & "</option>" & vbCrLf
			end if
			rsdrop.MoveNext
		loop
	end if
	strOutput = strOutput & "</select>" & vbCrLf
	rsdrop.Close
	set rsdrop = nothing
	
	DoRosterDropDownSm = strOutput
end function

Function plain2HTMLtxt(str)
	Dim output
    
    if isBarren(str) then
        plain2HTMLtxt = str
        exit function
    end if
	
	output = displayEmail(str)
	output = ChkUrls(output,"http://", 1)
	output = ChkUrls(output,"https://", 2)
	output = ChkUrls(output,"file:///", 3)
	output = ChkUrls(output,"www.", 4)
	output = ChkUrls(output,"mailto:",5)
		
	plain2HTMLtxt = output

End Function

Function rw(rwme)
	response.write rwme & "<br />"
End Function

Function gw(gwme)
	response.write gwme
	response.end
End Function

Function jsa(jsame)
	response.write "<script type=""text/javascript"">"
	response.write "	alert('" & Replace(jsame, "'", "\'") & "');"
	response.write "</script>"
End Function

Function isBarren(str2Chk)
  Dim isIt
  
  isIt = False
  
  If len(Trim(str2Chk&" ")) = 0 Then
    isIt = True
  Elseif str2Chk = "" Then
    isIt = True
  Elseif isNull(str2Chk) Then
    isIt = True
  Elseif isEmpty(str2Chk) Then
    isIt = True
  End If
  
  isBarren = isIt
End Function

function iif(iif1, iif2, iif3)
	if iif1 then
		iif = iif2
	else
		iif = iif3
	end if
end function

sub showMsg(msgType, msgTxt)
    %>
    <style type="text/css">
        .info, .success, .warning, .error, .validation {
            font-family:Arial, Helvetica, sans-serif; 
            font-size:13px;
            text-align: left;
            border: 1px solid;
            margin: 10px 0px;
            padding: 10px 10px 10px 50px;
            background-repeat: no-repeat;
            background-position: 10px center;
            /* If you want the images, remove the following lines */
        }
        .info {
            color: #00529B;
            background-color: #BDE5F8;
            background-image: url('images/icons/info.png');
        }
        .success {
            color: #4F8A10;
            background-color: #DFF2BF;
            background-image:url('images/icons/success.png');
        }
        .warning {
            color: #9F6000;
            background-color: #FEEFB3;
            background-image: url('images/icons/warning.png');
        }
        .error {
            color: #D8000C;
            background-color: #FFBABA;
            background-image: url('images/icons/error.png');
        }
        .validation {
            color: #D63301;
            background-color: #FFCCBA;
            background-image: url('images/icons/validation.png');
        }

    </style>
    <%
    select case msgType
        case "note"
            response.write "<div class=""info"">"
            response.write msgTxt
            response.write "</div>"
        
        case "warn"
            response.write "<div class=""warning"">"
            response.write msgTxt
            response.write "</div>"
        
        case "err"
            response.write "<div class=""error"">"
            response.write msgTxt
            response.write "</div>"
            
        case "validation"
            response.write "<div class=""validation"">"
            response.write msgTxt
            response.write "</div>"
        
        case "success"
            response.write "<div class=""success"">"
            response.write msgTxt
            response.write "</div>"
        
        case else
            response.write msgTxt
    end select
end sub

sub rosterDropDownAddPlayers2Team(iTeam)
    strSql = "SELECT D.[STARTAGE], D.[ENDAGE] FROM " & STRTABLEPREFIX & "TEAM T LEFT OUTER JOIN " & STRTABLEPREFIX & "DIVISION D ON T.[DIVISION_ID] = D.[ID] WHERE T.[ID] = " & iTeam

    
end sub

'Function IsValidEmail - copyright asp101 (http://www.asp101.com/samples/email.asp)
'Function FormatPhoneNumber - copyright asp101 (http://www.asp101.com/samples/phone_format.asp)

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
		strTemp = strInput
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

Function rosterGetNextId()
    Dim sSql,tmpID, rs
    sSql = ""
    tmpID = 0
    set rs = nothing

    sSql = "SELECT [M_VALUE] FROM " & STRTABLEPREFIX & "MODS WHERE [M_NAME] = 'roster' AND [M_CODE] = 'pCount'"
    set rs = my_conn.execute(sSql)
    
    if rs.EOF or rs.BOF then
        gw("Fatal Error")
    else
        tmpID = rs.Fields("M_VALUE")
        set rs = nothing
    end if
    
    if not isNumeric(tmpID) then
        gw("Fatal Error")
    end if
    
    tmpID = cInt(tmpID) + 1
    
    sSql = "UPDATE " & STRTABLEPREFIX & "MODS SET [M_VALUE] = '" & tmpID & "' WHERE [M_NAME] = 'roster' AND [M_CODE] = 'pCount'"
    executeThis(sSql)
    
    rosterGetNextId = tmpID

End Function
%>