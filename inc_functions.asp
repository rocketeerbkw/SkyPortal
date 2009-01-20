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

%><!--#include file="includes/inc_encryption.asp" -->
<!-- #include file="files/config/core_icons.asp" --><%
intSkin = 1
dim sAppRead, sAppWrite, sAppFull, bAppRead, bAppWrite, bAppFull
dim sCatRead, sCatWrite, sCatFull, bCatRead, bCatWrite, bCatFull
dim sSCatRead, sSCatWrite, sSCatFull, bSCatRead, bSCatWrite, bSCatFull
dim cust_meta, cust_js, cust_xml, cust_css, intSecCode
dim intAppID, intAppActive
dim iDATA1, tDATA1, iDATA2, tDATA2, iDATA3, tDATA3, iDATA4, tDATA4
dim iDATA5, tDATA5, iDATA6, tDATA6, iDATA7, tDATA7, iDATA8, tDATA8
dim iDATA9, tDATA9, iDATA10, tDATA10

function randomNum(i)
  if isnull(i) or i = "" then i = 99999
  if not isnumeric(i) or i < 2 then i = 99999
  i = clng(i)
  Randomize()
  randomNum=Int(Rnd()*i)
end function

function doublenum(fNum)
	if len(fNum) > 1 then 
		doublenum = fNum 
	else 
		doublenum = "0" & fNum
	end if
end function

function widenum(fNum)
	if fNum > 9 then 
		widenum = "" 
	else 
		widenum = "&nbsp;"
	end if
end function

function Chked(fYN)
   if fYN = "yes" or fYN = "1" or fYN = 1 then '**
      Chked = " checked=""checked"""
   else 
      Chked = ""
   end if    
end function

function chkRadioB(actualValue, thisValue, boltf)
if isNumeric(actualValue) then actualValue = cLng(actualValue)
if actualValue = thisValue EQV boltf then
chkRadioB = " checked=""checked"""
else 
chkRadioB = ""
end if
end function

function chkSelect(ByVal actualValue, thisValue)
if isNumeric(actualValue) then actualValue = cLng(actualValue)
if actualValue = thisValue then
chkSelect = " selected"
else 
chkSelect = ""
end if
end Function

function CheckSelected(ByVal chkval1, chkval2)
   if IsNumeric(chkval1) Then chkval1 = CInt(chkval1)
   if (chkval1 = chkval2) then
      CheckSelected = " selected"
   else
      CheckSelected = ""
   end if
end function

function chkRadio(ByVal actualValue, thisValue)
if isNumeric(actualValue) then actualValue = cLng(actualValue)
if actualValue = thisValue then
chkRadio = " checked=""checked"""
else 
chkRadio = ""
end if
end Function

function chkCheckbox(actualValue, thisValue, boltf)
if isNumeric(actualValue) then actualValue = cLng(actualValue)
if actualValue = thisValue EQV boltf then
chkCheckbox = " checked"
else 
chkCheckbox = ""
end if
end Function 

function chkExist(actualValue)
if trim(actualValue) <> "" then
chkExist = actualValue
else 
chkExist = ""
end if
end function


function chkExistElse(actualValue, elseValue)
if trim(actualValue) <> "" then
chkExistElse = actualValue
else 
chkExistElse = elseValue
end if
end Function

function DoDropDown(fTableName,fDisplayField,fValueField,fSelectValue,fName,fFirstOption,fWhere,fOrderBy)
	strSql = "SELECT " & fDisplayField & ", " & fValueField 
	strSql = strSql & " FROM " & fTableName
	if trim(fWhere) <>  "" then
		strSQL = strSQL & " WHERE " & fWhere
	end if
	if trim(fOrderBy) <> "" then
		strSQL = strSQL & " ORDER BY " & fOrderBy
	end if
	set rsdrop = my_Conn.execute(strSql)
	Response.Write "<select name=""" & fName & """>" & vbCrLf
	if rsdrop.EOF or rsdrop.BOF then
		Response.Write "<option>" & txtNotFound & "</option>"  & vbCrLf
	else
		if trim(fFirstOption) <> "" then
			Response.Write "<option value=""0"">&nbsp;"
			Response.Write fFirstOption & "&nbsp;</option>" & vbCrLf
		end if
		do until rsdrop.EOF
			if lcase(rsdrop(fValueField)) = lcase(fSelectValue) then
			  Response.Write "<option value=""" & rsdrop(fValueField) & """ selected=""selected"">&nbsp;"
			  Response.Write rsdrop(fDisplayField) & "&nbsp;</option>" & vbCrLf
			else
			  Response.Write "<option value=""" & rsdrop(fValueField) & """>&nbsp;"
			  Response.Write rsdrop(fDisplayField) & "&nbsp;</option>" & vbCrLf
			end if
			rsdrop.MoveNext
		loop
	end if
	Response.Write "</select>" & vbCrLf
	rsdrop.Close
	set rsdrop = nothing	
end function

function DoSubmitDropDown(fTableName,fDisplayField,fValueField,fSelectValue,fName,fFirstOption,fWhere,fOrderby)
	strSql = "SELECT " & fDisplayField & ", " & fValueField 
	strSql = strSql & " FROM " & fTableName
	if trim(fWhere) <>  "" then
		strSQL = strSQL & " WHERE " & fWhere
	end if
	if trim(fOrderBy) <> "" then
		strSQL = strSQL & " ORDER BY " & fOrderBy
	end if
	set rsdrop = my_Conn.execute(strSql)
		
	Response.Write "<select name=""" & fName & """ onchange=""submit()"">" & vbCrLf
	if rsdrop.EOF or rsdrop.BOF then
		Response.Write "<Option>" & txtNotFound & "</option>"  & vbCrLf
	else
		if trim(fFirstOption) <> "" then
			Response.Write "<option value=""" & rsdrop(fValueField) & """>&nbsp;"
			Response.Write fFirstOption & "&nbsp;</option>" & vbCrLf
		end if
		do until rsdrop.EOF
			if lcase(rsdrop(fValueField)) = lcase(fSelectValue) then
			  Response.Write "<option value=""" & rsdrop(fValueField) & """ selected=""selected"">&nbsp;"
			  Response.Write rsdrop(fDisplayField) & "&nbsp;</option>" & vbCrLf
			else
				Response.Write "<option value=""" & rsdrop(fValueField) & """>&nbsp;"
				Response.Write rsdrop(fDisplayField) & "&nbsp;</option>" & vbCrLf
			end if
			rsdrop.MoveNext
		loop
	end if
	Response.Write "</select>" & vbCrLf
	rsdrop.Close
	set rsdrop = nothing	
end function

function ChkBadWords(fString)
	bwords = split(strBadWords, "|")
	for i = 0 to ubound(bwords)
		fString = Replace(fString, bwords(i), string(len(bwords(i)),"*"), 1,-1,1) 
	next
	ChkBadWords = fString
end function

sub showMsgBlock(t,sMsg)
  if t = 1 then
	spThemeBlock1_open(intSkin)
  end if%>
	<table class="tPlain" cellpadding="10">
		<tr align="center"><td>
			<p><b><%= sMsg %></b></p>
		</td></tr>
	</table>
	<%
  if t = 1 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub displayABanner(b_type,d_type,rInt,tBlkT,tBlk)
  'b_type = banner type: numeric: large, affiliate
  'd_type = display type: 1 - random or 2 - rotating
  'rInt = integer : Interval in seconds for banners to rotate
  'tBlk = boolian : dispkay a themeblock with the banner
  'tBlkT = string : Themeblock Title
  
  activeBanners = false
  tJS = ""
  
  iRandNum = randomNum(99999)
	' Select random banner to display on page first
	strSQL = "SELECT * FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.B_ACTIVE=1 and " & strTablePrefix & "BANNERS.B_LOCATION=" & b_type
	'set rsB = server.CreateObject("adodb.recordset")
	'rsB.CursorType = 3
	'rsB.CursorLocation = 3
	'rsB.LockType = 3
	'rsB.Open strSQL, my_Conn
	
  Set rsB = oSpData.GetRecordset(strSQL)
     if not rsB.eof then
		numLinks = ubound(rsB.getrows(),2) + 1
		rsB.MoveFirst
		if numLinks > 1 then
			Randomize
			RndNumber = Int(Rnd * numLinks) 
			rsB.move RndNumber
		end if
		bannerTxt = chkstring(rsB("B_ACRONYM"),"display") & "|" & numLinks & "|" & RndNumber
		bannerID = rsB("ID")
		bannerImg = chkstring(rsB("B_IMAGE"),"displayimage")
		bannerName = chkstring(rsB("B_NAME"),"display")
		activeBanners = true
	end if
	if d_type = 1 and activeBanners then 
	  'random banner - add 1 to impression count
		sSql = "UPDATE " & strTablePrefix & "BANNERS SET " & strTablePrefix & "BANNERS.B_IMPRESSIONS = " & strTablePrefix & "BANNERS.B_IMPRESSIONS + 1  WHERE " & strTablePrefix & "BANNERS.ID=" &  bannerID
		on error resume next
		my_Conn.execute(sSql)
		on error goto 0
	end if
	
'response.Write(d_type & " : " & b_type)
'random rotating banners
  if d_type = 2 and activeBanners = true then  
	strSQL = "SELECT B_IMAGE, B_ACRONYM, ID FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.B_ACTIVE=1 and " & strTablePrefix & "BANNERS.B_LOCATION=" & b_type & " and " & strTablePrefix & "BANNERS.B_IMAGE NOT LIKE '%.swf'"
	'set rsLinks = my_Conn.execute(strSQL)
    Set rsLinks = oSpData.GetRecordset(strSQL)
	if not rsLinks.eof then
	  tJS = vbcrlf & "<script type=""text/javascript"">" & vbcrlf
	  tJS = tJS & "<!-- Begin" & vbcrlf
	  ' // delay between rotating images (in seconds)
	  tJS = tJS & "var interval" & iRandNum & " = " & rInt & ";" & vbcrlf
	  ' // random display? 0 = no, 1 = yes
	  tJS = tJS & "var random_display" & iRandNum & " = 1;" & vbcrlf
	  tJS = tJS & "interval" & iRandNum & " *= 1000;" & vbcrlf
	  tJS = tJS & "var image_index" & iRandNum & " = 0;" & vbcrlf
	  tJS = tJS & "image_list" & iRandNum & " = new Array();" & vbcrlf
	  tJS = tJS & "link_list" & iRandNum & " = new Array();" & vbcrlf
	  tJS = tJS & "text_list" & iRandNum & " = new Array();" & vbcrlf
	  tJS = tJS & "var url" & iRandNum & " = ""banner_link.asp?id=2"";" & vbcrlf
	  Do until rsLinks.eof
	    tJS = tJS & "image_list" & iRandNum & "[image_index" & iRandNum & "++] = new imageItem("""&rsLinks("B_IMAGE")&""");" & vbcrlf
	    tJS = tJS & "text_list" & iRandNum & "[image_index" & iRandNum & "] = """&rsLinks("B_ACRONYM")&""";" & vbcrlf
	    tJS = tJS & "link_list" & iRandNum & "[image_index" & iRandNum & "] = "&rsLinks("ID")&";" & vbcrlf
		rsLinks.movenext
	  loop
	  tJS = tJS & "var number_of_image" & iRandNum & " = image_list" & iRandNum & ".length;" & vbcrlf
	  
	  tJS = tJS & "function getNextImage" & iRandNum & "() {" & vbcrlf
	  tJS = tJS & "if (random_display" & iRandNum & "==1) {" & vbcrlf
	  tJS = tJS & "image_index" & iRandNum & " = generate(0, number_of_image" & iRandNum & "-1);" & vbcrlf
	  tJS = tJS & "}else{" & vbcrlf
	  tJS = tJS & "image_index" & iRandNum & " = (image_index" & iRandNum & "+1) % number_of_image" & iRandNum & ";"
	  tJS = tJS & vbcrlf & "}" & vbcrlf
	  tJS = tJS & "var new_image = get_ImageItemLocation(image_list" & iRandNum & "[image_index" & iRandNum & "]);" & vbcrlf
	  tJS = tJS & "return(new_image);" & vbcrlf
	  tJS = tJS & "}" & vbcrlf
	  
	  tJS = tJS & "function rotateImage" & iRandNum & "(place) {" & vbcrlf
	  tJS = tJS & "var new_image = getNextImage" & iRandNum & "();" & vbcrlf
	  tJS = tJS & "document[place].src = new_image;" & vbcrlf
	  tJS = tJS & "url" & iRandNum & " = ""banner_link.asp?id=""+link_list" & iRandNum & "[image_index" & iRandNum & "+1];+""""" & vbcrlf
	  tJS = tJS & "document[place].alt = """"+text_list" & iRandNum & "[image_index" & iRandNum & "+1]+"""";" & vbcrlf
	  tJS = tJS & "document[place].title = """"+text_list" & iRandNum & "[image_index" & iRandNum & "+1]+"""";" & vbcrlf
	  tJS = tJS & "var recur_call" & iRandNum & " = ""rotateImage" & iRandNum & "('""+place+""')"";" & vbcrlf
	  tJS = tJS & "setTimeout(recur_call" & iRandNum & ", interval" & iRandNum & ");" & vbcrlf
	  tJS = tJS & "}" & vbcrlf
	  
	  tJS = tJS & "window.onLoad=rotateImage" & iRandNum & "('bImage" & iRandNum & "')" & vbcrlf
	  tJS = tJS & "//  End -->" & vbcrlf
	  tJS = tJS & "</script>" & vbcrlf
	end If
  end if
  
  If activeBanners Then
    ':: display the banner
    if tBlk then
	  spThemeTitle = tBlkT
      spThemeBlock1_open(intSkin)
    end if
	Response.Write("<div>")
	If d_type = 2 then
	  Response.Write("<a href=""javascript:void();"" onClick=""window.open(url" & iRandNum & ",'BannerWin');"" name=""banner"">")
	  Response.Write("<img alt=""" & bannertxt & """ title=""" & bannertxt & """ name=""bImage" & iRandNum & """ border=""0"" src=""" & bannerImg & """ />")
	  Response.Write("</a>")
	ElseIf d_type = 1 Then
	  Response.Write("<a target=""_blank"" title=""" & bannerTxt & """ href=""banner_link.asp?id=" & bannerID & """>")
	  If right(bannerImg,4) = ".swf" Then
	    writeFlash(bannerImg)
	  Else 
	    response.write("<img alt=""" & bannertxt & """ name=""bImage" & iRandNum & """ border=""0"" src=""" & bannerImg & """ />")
	  end if
	  Response.Write("</a>")
	End If
	Response.Write("</div>")
    if tBlk then
      spThemeBlock1_close(intSkin)
    end if
  response.Write(tJS)
  end if
end sub

'######### IPGate MOD #############################

dbHitCnt = 0

Function OnlineSQLencode(byVal strPass)
 If not isNull(strPass) and strPass <> "" Then
 	strPass = Replace(strPass, "'", "")
 	strPass = Replace(strPass, "|", "")
 	strPass = Replace(strPass, "(", "")
 	strPass = Replace(strPass, ")", "")
 	strPass = Replace(strPass, ";", "")
 	OnlineSQLencode = strPass
 End If
End Function

Function OnlineSQLdecode(byVal strPass)
 If not isNull(strPass) and strPass <> "" Then
 	strPass = Replace(strPass, "'%'", "%")
 	strPass = Replace(strPass, "''", "'")
 	strPass = Replace(strPass, "'|'", "|")
 	OnlineSQLdecode = strPass
 End If
End Function

Function SetConfigValue(bUpdate, fVariable, fValue)

' bUpdate = 1 : if it exists then overwrite with new values
' bUpdate = 0 : if it exists then leave unchanged

Dim strSql

strSql = "SELECT C_" & fVariable & " FROM " & strTablePrefix & "CONFIG "

Set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSql, my_Conn
dbHitCnt = dbHitCnt + 1

if bUpdate <> 0 then 
SetConfigValue = "updated"
my_conn.execute ("UPDATE " & strTablePrefix & "CONFIG SET C_" & fVariable & " = '" & fValue & "'"),,adCmdText + adExecuteNoRecords
else ' not changed
SetConfigValue = "unchanged"
end if

rs.close
set rs = nothing
end function 

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::: :::
'::: This script performs 'RC4' Stream Encryption :::
'::: (Based on what is widely thought to be RSA's RC4 :::
'::: algorithm. It produces output streams that are identical :::
'::: to the commercial products) :::
'::: :::
'::: This script is Copyright © 1999 by Mike Shaffer :::
'::: ALL RIGHTS RESERVED WORLDWIDE :::
'::: :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Dim sbox(255)
Dim enckey(255)


Sub RC4Initialize(strPwd)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::: This routine called by EnDeCrypt function. Initializes the :::
'::: sbox and the key array) :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

dim tempSwap
dim a
dim b

intLength = len(strPwd)
For a = 0 To 255
enckey(a) = asc(mid(strpwd, (a mod intLength)+1, 1))
sbox(a) = a
next

b = 0
For a = 0 To 255
b = (b + sbox(a) + enckey(a)) Mod 256
tempSwap = sbox(a)
sbox(a) = sbox(b)
sbox(b) = tempSwap
Next

End Sub

Function EnDeCrypt(plaintxt, psw)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::: This routine does all the work. Call it both to ENcrypt :::
'::: and to DEcrypt your data. :::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

dim temp
dim a
dim i
dim j
dim k
dim cipherby
dim cipher

i = 0
j = 0

RC4Initialize psw

For a = 1 To Len(plaintxt)
i = (i + 1) Mod 256
j = (j + sbox(i)) Mod 256
temp = sbox(i)
sbox(i) = sbox(j)
sbox(j) = temp

k = sbox((sbox(i) + sbox(j)) Mod 256)

cipherby = Asc(Mid(plaintxt, a, 1)) Xor k
cipher = cipher & Chr(cipherby)
Next

EnDeCrypt = cipher

End Function
'############### end IPGate mod ##################################3

Function ChkActUsrUrl(strTString)
strTString   =   replace(strTString, "<","", 1, -1, 1)
strTString   =   replace(strTString, ">","", 1, -1, 1)
strTString   =   replace(strTString, """","", 1, -1, 1)
strTString   =   replace(strTString, "'","", 1, -1, 1)
strTString   =   replace(strTString, ";","", 1, -1, 1)
ChkActUsrUrl = strTString

end function

Function ReplaceImageTags(fString)
Dim oTag, cTag
Dim roTag, rcTag
Dim oTagPos, cTagPos
Dim nTagPos
Dim counter1, counter2
Dim strUrlText
Dim Tagcount
Dim strTempString, strResultString
TagCount = 6
Dim ImgTags(6,2,2)
Dim strArray, strArray2

ImgTags(1,1,1) = "[img]"
ImgTags(1,2,1) = "[/img]"
ImgTags(1,1,2) = "<img hspace=7 vspace=5 src="""
ImgTags(1,2,2) = """ border=0 />"

ImgTags(2,1,1) = "[image]"
ImgTags(2,2,1) = "[/image]"
ImgTags(2,1,2) = ImgTags(1,1,2)
ImgTags(2,2,2) = ImgTags(1,2,2)

ImgTags(3,1,1) = "[img=right]"
ImgTags(3,2,1) = "[/img=right]"
ImgTags(3,1,2) = "<img align=right hspace=7 vspace=5 src="""
ImgTags(3,2,2) = """ id=right border=0 />"

ImgTags(4,1,1) = "[image=right]"
ImgTags(4,2,1) = "[/image=right]"
ImgTags(4,1,2) = ImgTags(3,1,2)
ImgTags(4,2,2) = ImgTags(3,2,2)

ImgTags(5,1,1) = "[img=left]"
ImgTags(5,2,1) = "[/img=left]"
ImgTags(5,1,2) = "<img align=left hspace=7 vspace=5 src="""
ImgTags(5,2,2) = """ id=left border=0 />"

ImgTags(6,1,1) = "[image=left]"
ImgTags(6,2,1) = "[/image=left]"
ImgTags(6,1,2) = ImgTags(5,1,2)
ImgTags(6,2,2) = ImgTags(5,2,2)

strResultString = ""
strTempString = fString

for counter1 = 1 to TagCount

oTag = ImgTags(counter1,1,1)
roTag = ImgTags(counter1,1,2)
cTag = ImgTags(counter1,2,1)
rcTag = ImgTags(counter1,2,2)
oTagPos = InStr(1, strTempString, oTag, 1)
cTagPos = InStr(1, strTempString, cTag, 1)

if (oTagpos > 0) and (cTagPos > 0) then

strArray = Split(strTempString, oTag, -1)
for counter2 = 0 to Ubound(strArray)
if (Instr(1, strArray(counter2), cTag) > 0) then
strArray2 = split(strArray(counter2), cTag, -1)
strUrlText = strArray2(0)
strUrlText = replace(strUrlText, """", "") ' ## filter out "
strUrlText = replace(strUrlText, "<", "") ' ## filter out <
strUrlText = replace(strUrlText, ">", "") ' ## filter out >
strUrlText = replace(strUrlText, "+", "") ' ## filter out +
strUrlText = replace(strUrlText, "(", "") ' ## filter out (
strUrlText = replace(strUrlText, ")", "") ' ## filter out )
strUrlText = replace(strUrlText, ";", "") ' ## filter out ;
strUrlText = replace(strUrlText, "'", "") ' ## filter out '
strUrlText = replace(strUrlText, "=", "") ' ## filter out =
strUrlText = replace(strUrlText, "&", "") ' ## filter out &
strUrlText = replace(strUrlText, "#", "") ' ## filter out #
strUrlText = replace(strUrlText, vbTab, " ", 1, -1, 1) ' ## filter out Tabs
strUrlText = replace(strUrlText, "view-source", " ", 1, -1, 1) ' ## filter out view-source
strUrlText = replace(strUrlText, "javascript", " ", 1, -1, 1) ' ## filter out javascript
strUrlText = replace(strUrlText, "jscript", " ", 1, -1, 1) ' ## filter out jscript
strUrlText = replace(strUrlText, "vbscript", " ", 1, -1, 1) ' ## filter out vbscript
strUrlText = replace(strUrlText, "mailto", " ", 1, -1, 1) ' ## filter out mailto

strResultString = strResultString & roTag & strUrlText & rcTag & strArray2(1)
else
strResultString = strResultString & strArray(counter2)
end if 
next 

strTempString = strResultString
strResultString = ""
end if
next

ReplaceImageTags = strTempString

end function

function ChkUrls(fString, fTestTag, fType)

Dim strArray
Dim Counter
Dim strTempString

strTempString = fString
if Instr(1, fString, fTestTag) > 0 then
	strArray = Split(fString, fTestTag, -1)
	strTempString = strArray(0)
	for counter = 1 to UBound(strArray)
		if ((strArray(counter-1) = "" or len(strArray(counter-1)) < 5) and strArray(counter)<> "") then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		elseif ((UCase(right(strArray(counter-1),6)) <> "HREF=""") and (UCase(right(strArray(counter-1),5)) <> "[URL]") and (UCase(right(strArray(counter-1),6)) <> "[URL=""") and (UCase(right(strArray(counter-1),7)) <> "FILE:///") and (UCase(right(strArray(counter-1),7)) <> "HTTP://") and (UCase(right(strArray(counter-1),8)) <> "HTTPS://") and (UCase(right(strArray(counter-1),5)) <> "SRC=""") and (UCase(right(strArray(counter-1),5)) <> "SRC=""") and strArray(counter)<> "") then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		else
			strTempString = strTempString & fTestTag & strArray(counter)
		end if
	next
end if

ChkUrls = strTempString

end function


function ChkMail(fString, fTestTag, fType)

Dim strArray
Dim Counter
Dim strTempString

strTempString = fString

if Instr(1, fString, fTestTag) > 0 then
	strArray = Split(fString, fTestTag, -1)
	strTempString = ""
'	strTempString = strArray(0)
	for counter = 0 to UBound(strArray)
		if (Instr(strArray(counter), "@") > 0) and not(Instr(strArray(counter), "mailto:") > 0) and not(Instr(UCase(strArray(counter)), "[URL") > 0) then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		else
			strTempString = strTempString & fTestTag & strArray(counter)
		end if
	next
end if

ChkMail = strTempString

end function


'#################################################################################
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## by Hawk92 - 11-2004
'## Original Formatstr replaced by 3 functions Formatstr,Formatstr2,Formatstr3
'## ##############################################################################
function FormatStr(fString)
    strMatch=InStr(1,fString,"[@@X]",1)
    If strMatch >0 Then
    	arrStr = split(fString,"[@@]")
		tmpStr = ""
		for xu = 0 to ubound(arrStr)
			if inStr(1,arrStr(xu),"[/@@]",1) = 0 then
				tmpStr = tmpStr & FormatStr2(arrStr(xu))
			else
				arrTmp = split(arrStr(xu),"[/@@]")
				tmpStr = tmpStr & arrTmp(0)
				for xy = 1 to ubound(arrTmp)
					tmpStr = tmpStr & FormatStr2(arrTmp(xy))
				next
				set arrTmp = nothing
			end if
		next
		set arrStr = nothing
		tmpStr=Replace(tmpStr,"[APOS]","'",1,-1,1)
    	FormatStr=tmpStr   
    Else
	  FormatStr = FormatStr2(fString)
    End If	
end Function

function FormatStr_worksB4NewCodebox(fString)
    strMatch=InStr(1,fString,"[@@]",1)
    If strMatch >0 Then
    	arrStr = split(fString,"[@@]")
		tmpStr = ""
		for xu = 0 to ubound(arrStr)
			if inStr(1,arrStr(xu),"[/@@]",1) = 0 then
				tmpStr = tmpStr & FormatStr2(arrStr(xu))
			else
				arrTmp = split(arrStr(xu),"[/@@]")
				tmpStr = tmpStr & arrTmp(0)
				for xy = 1 to ubound(arrTmp)
					tmpStr = tmpStr & FormatStr2(arrTmp(xy))
				next
				set arrTmp = nothing
			end if
		next
		set arrStr = nothing
		tmpStr=Replace(tmpStr,"[APOS]","'",1,-1,1)
    	FormatStr=tmpStr   
    Else
	  FormatStr = FormatStr2(fString)
    End If	
end Function

Function FormatStr3(fString,ptr)
' New function by Hawk92 - source code box mod 1.5
' This function processes messages with code for display

' This grabs any content before the first code marker
strtemp= Mid(fString,1,ptr-1)
strFinal=FormatStr2(strtemp)
      sptr=1
      eptr=1
      cntr=1
' This is a loop to parse the string and determine the start and end of the code segment
      Do While sptr < Len(fString)And cntr=<2
        eptr=InStr(sptr,fString,"[/@@]",1)+4
        If eptr>0 Then
           cntr=cntr+1
           If cntr=<2 Then
            sptr=eptr
           End if
        End if
      Loop
' This removes the [@@] markers and sets [APOS] back to single quotes   
  strMid=Mid(fString,ptr,eptr-ptr+1)
  strMid=Replace(strMid,"[@@]","",1,-1,1)
  strMid=Replace(strMid,"[/@@]","",1,-1,1)
  strMid=Replace(strMid,"[APOS]","'",1,-1,1)
' Grab any string after the code segment
  strEnd=Mid(fString,eptr+1,Len(fString)-(eptr))
  FormatStr3=strFinal&strMid&strEnd    
End Function

function FormatStr2(fString)
' Renamed function by Hawk92 - source code box mod 1.5 
' Original code from original FormatStr function 1.5 beta3 
	'if strAllowHtml <> 1 then
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	'fString = Replace(fString, CHR(10), "<br />")
	'end if
	'fString = Replace(fString, vbcrlf, CHR(10))
	'fString = Replace(fString, vbcr, CHR(10))
	'fString = Replace(fString, vblf, CHR(10))
	if strBadWordFilter = 1 then
		fString = ChkBadWords(fString)
	end if
	if strAllowHtml <> 1 then
	fString = ChkUrls(fString,"http://", 1)
	fString = ChkUrls(fString,"https://", 2)
	fString = ChkUrls(fString,"file:///", 3)
	fString = ChkUrls(fString,"www.", 4)
	fString = ChkUrls(fString,"mailto:",5)
	fString = ChkMail(fString," ",5)
	'fString = edit_hrefs(fString, 5)
	end if
	fString = ReplaceUrls(fString)
  If InStr(lcase(fString),"[code]")>0 Then
	fString = doMsgCode(fString)
  end if
	
	FormatStr2 = fString
end Function

Function doMsgCode(fStr)
 If InStr(lcase(fStr),"[/code]")>0 Then
	' This is the end string for the codebox
	endstr="</textarea></td></tr></table></form>"
	
	ptr=1
	eptr=1
	max=Len(fStr)
	cntr=0
	strFinal = ""
	strTemp = ""
	' This is the loop to process each part of the message String
	Do While (ptr < max)
  	  forceclose=False
	  strCode=""
 	  If InStr(ptr,fStr,"[code]",1)>0 Then
  		sptr=InStr(ptr,fStr,"[code]",1)
  		eptr=InStr(sptr,fStr,"[/code]",1)+7
  		If eptr=<7 Then
		  forceclose=True
  		  eptr=max
  		End If
  		strTemp=Mid(fStr,ptr,(sptr-ptr))
  		'strTemp = HTMLEncode(strTemp)
  		'strTemp=saveForumCode(strTemp,"message")
  		strCode=Mid(fStr,sptr,eptr-sptr)
  		If forceclose=True Then
  		  strCode=strCode&"[/code]"
  		End If
		if strAllowHTML = 0 then
  		  strCode = server.HTMLEncode(strCode)
		else
		  strCode = Replace(strCode,"<br />", vbcrlf)
		  strCode = Replace(strCode,"<br />", vbcrlf)
  		  'strCode = server.HTMLEncode(strCode)
		end if
  
		' Replace the [code] markers with the html for the codebox
		cdStrt = codeStart()
  		strCode=Replace(strCode,"[code]",cdStrt,1)
  		strCode=Replace(strCode,"[/code]",endstr,1) 
  		strFinal=strFinal & strTemp& strCode
  		ptr=eptr
 	  Else
		' If there are no more code markers set prt to end of string  
  		ptr=max
 	  End if
	Loop
	' This picks up any part of the string after the Code
  	strFinal=strFinal& Mid(fStr,eptr,max-eptr+1)
  	'strFinal = doCode(strFinal, "[quote]", "[/quote]", "<BLOCKQUOTE id=quote><font id=quote>" & txtQuote & ":<hr width=99% height=1 noshade id=quote>", "<hr width=99% height=1 noshade id=quote></font id=quote></BLOCKQUOTE id=quote>")		
	doMsgCode=strFinal
   if strAllowForumCode = 12 and strAllowHTML = 20 then
     strFinal = fStr
	 stSCode = ""
	 if InStr(lcase(strFinal),"[code]") > 0 then
       testStr = InStr(lcase(strFinal),"[code]")
	 else
	   testStr = 0
	 end if
     do until testStr = 0
	    stSCode = codeStart()
  		strFinal=Replace(lcase(strFinal),"[code]",stSCode,1)
		if InStr(lcase(strFinal),"[code]") > 0 then
          testStr = InStr(lcase(strFinal),"[code]")
		else
		  testStr = 0
		end if
	 loop
  	 strFinal=Replace(lcase(strFinal),"[/code]",endstr,1,1,1) 
	 doMsgCode=strFinal
	 'doMsgCode=codeStart() 
   end if
  else
    doMsgCode=fStr
  end if
End function

function codeStart()
  		Randomize()
  		ccntr=Int(Rnd()*9000)
		' This is the start string for the codebox
		ststr="<form method=""post"" name=""selectcode" & ccntr & """ action="""">"
		ststr=ststr&"<table style=""border: 1 solid #000000"" cellpadding=""0"" cellspacing=""0""><tr><td class=""spThemeCodeboxHead"" width=""100%"" valign=""middle"">"
		ststr=ststr&"&nbsp;&nbsp;Source Code:"
		ststr=ststr&"<img border=""0"" src=""images/clear.gif"" width=""144"" height=""4"" />"
		ststr=ststr&"<img border=""0"" src=""images/click_select.gif"" align=""middle"" onclick=""docodebox("&ccntr&");"" />"
		ststr=ststr&"<img border=""0"" src=""images/clear.gif"" width=""146"" height=""3"" />"
		ststr=ststr&"<img border=""0"" src=""images/icon_minus2.gif"" title=""" & txtCollapse & """  align=""top"" onclick=""contract("&ccntr&");"" />&nbsp;"
		ststr=ststr&"<img border=""0"" src=""images/icon_plus2.gif"" title=""" & txtExpand & """ align=""top"" onclick=""expand("&ccntr&");"" />&nbsp;&nbsp;&nbsp;&nbsp;"
		ststr=ststr&"<img border=""0"" src=""images/icon_help.gif"" title=""" & txtHelp & """ align=""middle"" onclick=""dohelp();"" /></td></tr>"
		ststr=ststr&"<tr><td></td></tr>"
		ststr=ststr&"<tr><td  valign=""middle""><textarea rows=""3"" READONLY=""Yes"" id=""thecode"&ccntr&""" name=""thecode"&ccntr&""" cols=""70"" style=""color: #008000; font-family: courier; font-size: 10pt; text-align: left; vertical-align: top; background-image: none; background-repeat: no-repeat; border: none"" wrap=""virtual"">"
		codeStart=ststr
end function
'#################################################################################
'## End of Replacement functions for FormatStr
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## ##############################################################################

function doCode(fString, fOTag, fCTag, fROTag, fRCTag)
  if fOTag = "[quote=mebaby]" then
'	fOTagPos = Instr(1, fString, fOTag, 1)
'	fCTagPos = Instr(1, fString, fCTag, 1)
'	while (fCTagPos > 0 and fOTagPos > 0)
'		quoted = mid(fString,fOTagPos+7,instr(fOTagPos+7,fString,"]",1))
'		newTag = fOTag&quoted&"]"
'		replace fString,newTag,"[quote]",fOTagPos,1
'		fOnTag = "[quote]"
'		replace(fString,"@@who@@",quoted,fOTagPos,1)
'		fString = replace(fString, fOnTag, fROTag&quoted&" said:<br /><br />", fOTagPos, 1, 1)
'		fString = replace(fString, fCTag, fRCTag, fCTagPos, 1, 1)
'		fOTagPos = Instr(1, fString, fOTag, 1)
'		fCTagPos = Instr(1, fString, fCTag, 1)
'	wend
'	doCode = fString
  else
	fOTagPos = Instr(1, fString, fOTag, 1)
	fCTagPos = Instr(1, fString, fCTag, 1)
	while (fCTagPos > 0 and fOTagPos > 0)
		fString = replace(fString, fOTag, fROTag, 1, 1, 1)
		fString = replace(fString, fCTag, fRCTag, 1, 1, 1)
		fOTagPos = Instr(1, fString, fOTag, 1)
		fCTagPos = Instr(1, fString, fCTag, 1)
	wend
	doCode = fString
  end if
end function

'#################################################################################
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## by Hawk92 - 11-2004
'## Original CleanCode replaced by 3 functions CleanCode,CleanCode2,CleanCode3
'## ##############################################################################
Function CleanCode(fString)
' Modified by Hawk92 - source code box mod - 11-2004
' Provides Branching control if [@@] markers (which indicate code) are in String
' Cleancode2 is original MWP Cleancode Logic
' Cleancode3 is new logic to process messages containing code
if fString = "" then 
		fString = " "
else 
    strMatch2=InStr(1,fString,"[@@]",1)
    If strMatch2 >0 Then
	  fString=CleanCode3(fstring)
	Else
	  fstring=CleanCode2(fString)
    End If
End if
CleanCode=fString
End Function

function CleanCode2(fString)
' New function by Hawk92 - source code box mod - 11-2004 (mostly original MWP1.5 Cleancode()
		if strAllowForumCode = "1" then
			fString = replace(fString, "<marquee>", "[marquee]", 1, -1, 1)
			fString = replace(fString, "</marquee>", "[/marquee]", 1, -1, 1)

			fString = replace(fString, "<sup>", "[sup]", 1, -1, 1)
			fString = replace(fString, "</sup>", "[/sup]", 1, -1, 1)

			fString = replace(fString, "<sub>", "[sub]", 1, -1, 1)
			fString = replace(fString, "</sub>", "[/sub]", 1, -1, 1)

			fString = replace(fString, "<tt>", "[tt]", 1, -1, 1)
			fString = replace(fString, "</tt>", "[/tt]", 1, -1, 1)

			fString = replace(fString, "<span style='background-color: #FFFF00'>", "[hl]", 1, -1, 1)
			fString = replace(fString, "<b></b></span>", "[/hl]", 1, -1, 1)

			fString = replace(fString, "<pre>", "[pre]", 1, -1, 1)
			fString = replace(fString, "</pre>", "[/pre]", 1, -1, 1)

			fString = replace(fString, "<hr />","[hr]", 1, -1, 1)
			fString = replace(fString, "<hr>","[hr]", 1, -1, 1)
			
			fString = replace(fString, "<b>","[b]", 1, -1, 1)
			fString = replace(fString, "</b>","[/b]", 1, -1, 1)
			fString = replace(fString, "<strong>","[b]", 1, -1, 1)
			fString = replace(fString, "</strong>","[/b]", 1, -1, 1)
		    fString = replace(fString, "<s>", "[s]", 1, -1, 1)
		    fString = replace(fString, "</s>", "[/s]", 1, -1, 1)
			fString = replace(fString, "<u>","[u]", 1, -1, 1)
			fString = replace(fString, "</u>","[/u]", 1, -1, 1)
			fString = replace(fString, "<i>","[i]", 1, -1, 1)
			fString = replace(fString, "</i>","[/i]", 1, -1, 1)
			fString = replace(fString, "<font face='Andale Mono'>", "[font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "</font id='Andale Mono'>", "[/font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial'>", "[font=Arial]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial'>", "[/font=Arial]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial Black'>", "[font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial Black'>", "[/font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "<font face='Book Antiqua'>", "[font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "</font id='Book Antiqua'>", "[/font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "<font face='Century Gothic'>", "[font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "</font id='Century Gothic'>", "[/font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "<font face='Comic Sans MS'>", "[font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Comic Sans MS'>", "[/font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Courier New'>", "[font=Courier New]", 1, -1, 1)
			fString = replace(fString, "</font id='Courier New'>", "[/font=Courier New]", 1, -1, 1)
			fString = replace(fString, "<font face='Georgia'>", "[font=Georgia]", 1, -1, 1)
			fString = replace(fString, "</font id='Georgia'>", "[/font=Georgia]", 1, -1, 1)
			fString = replace(fString, "<font face='Impact'>", "[font=Impact]", 1, -1, 1)
			fString = replace(fString, "</font id='Impact'>", "[/font=Impact]", 1, -1, 1)
			fString = replace(fString, "<font face='Tahoma'>", "[font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "</font id='Tahoma'>", "[/font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "<font face='Times New Roman'>", "[font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "</font id='Times New Roman'>", "[/font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "<font face='Trebuchet MS'>", "[font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Trebuchet MS'>", "[/font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Script MT Bold'>", "[font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "</font id='Script MT Bold'>", "[/font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "<font face='Stencil'>", "[font=Stencil]", 1, -1, 1)
			fString = replace(fString, "</font id='Stencil'>", "[/font=Stencil]", 1, -1, 1)
			fString = replace(fString, "<font face='Verdana'>", "[font=Verdana]", 1, -1, 1)
			fString = replace(fString, "</font id='Verdana'>", "[/font=Verdana]", 1, -1, 1)
			fString = replace(fString, "<font face='Lucida Console'>", "[font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "</font id='Lucida Console'>", "[/font=Lucida Console]", 1, -1, 1)
		    
		      fString = replace(fString, "<font color=red>", "[red]", 1, -1, 1)
		      fString = replace(fString, "</font id=red>", "[/red]", 1, -1, 1)
		      fString = replace(fString, "<font color=green>", "[green]", 1, -1, 1)
		      fString = replace(fString, "</font id=green>", "[/green]", 1, -1, 1)
		      fString = replace(fString, "<font color=blue>", "[blue]", 1, -1, 1)
		      fString = replace(fString, "</font id=blue>", "[/blue]", 1, -1, 1)
		      fString = replace(fString, "<font color=white>", "[white]", 1, -1, 1)
		      fString = replace(fString, "</font id=white>", "[/white]", 1, -1, 1)
		      fString = replace(fString, "<font color=purple>", "[purple]", 1, -1, 1)
		      fString = replace(fString, "</font id=purple>", "[/purple]", 1, -1, 1)
	  	      fString = replace(fString, "<font color=yellow>", "[yellow]", 1, -1, 1)
	  	      fString = replace(fString, "</font id=yellow>", "[/yellow]", 1, -1, 1)
		      fString = replace(fString, "<font color=violet>", "[violet]", 1, -1, 1)
		      fString = replace(fString, "</font id=violet>", "[/violet]", 1, -1, 1)
		      fString = replace(fString, "<font color=brown>", "[brown]", 1, -1, 1)
		      fString = replace(fString, "</font id=brown>", "[/brown]", 1, -1, 1)
		      fString = replace(fString, "<font color=black>", "[black]", 1, -1, 1)
		      fString = replace(fString, "</font id=black>", "[/black]", 1, -1, 1)
		      fString = replace(fString, "<font color=pink>", "[pink]", 1, -1, 1)
		      fString = replace(fString, "</font id=pink>", "[/pink]", 1, -1, 1)
		      fString = replace(fString, "<font color=orange>", "[orange]", 1, -1, 1)
		      fString = replace(fString, "</font id=orange>", "[/orange]", 1, -1, 1)
		      fString = replace(fString, "<font color=gold>", "[gold]", 1, -1, 1)
		      fString = replace(fString, "</font id=gold>", "[/gold]", 1, -1, 1)

		      fString = replace(fString, "<font color=beige>", "[beige]", 1, -1, 1)
		      fString = replace(fString, "</font id=beige>", "[/beige]", 1, -1, 1)
		      fString = replace(fString, "<font color=teal>", "[teal]", 1, -1, 1)
		      fString = replace(fString, "</font id=teal>", "[/teal]", 1, -1, 1)
		      fString = replace(fString, "<font color=navy>", "[navy]", 1, -1, 1)
		      fString = replace(fString, "</font id=navy>", "[/navy]", 1, -1, 1)
		      fString = replace(fString, "<font color=maroon>", "[maroon]", 1, -1, 1)
		      fString = replace(fString, "</font id=maroon>", "[/maroon]", 1, -1, 1)
		      fString = replace(fString, "<font color=limegreen>", "[limegreen]", 1, -1, 1)
		      fString = replace(fString, "</font id=limegreen>", "[/limegreen]", 1, -1, 1)

			fString = replace(fString, "<h1>", "[h1]", 1, -1, 1)
			fString = replace(fString, "</h1>", "[/h1]", 1, -1, 1)
			fString = replace(fString, "<h2>", "[h2]", 1, -1, 1)
			fString = replace(fString, "</h2>", "[/h2]", 1, -1, 1)
			fString = replace(fString, "<h3>", "[h3]", 1, -1, 1)
			fString = replace(fString, "</h3>", "[/h3]", 1, -1, 1)
			fString = replace(fString, "<h4>", "[h4]", 1, -1, 1)
			fString = replace(fString, "</h4>", "[/h4]", 1, -1, 1)
			fString = replace(fString, "<h5>", "[h5]", 1, -1, 1)
			fString = replace(fString, "</h5>", "[/h5]", 1, -1, 1)
			fString = replace(fString, "<h6>", "[h6]", 1, -1, 1)
			fString = replace(fString, "</h6>", "[/h6]", 1, -1, 1)
			fString = replace(fString, "<font size=1>", "[size=1]", 1, -1, 1)
			fString = replace(fString, "</font id=size1>", "[/size=1]", 1, -1, 1)
			fString = replace(fString, "<font size=2>", "[size=2]", 1, -1, 1)
			fString = replace(fString, "</font id=size2>", "[/size=2]", 1, -1, 1)
			fString = replace(fString, "<font size=3>", "[size=3]", 1, -1, 1)
			fString = replace(fString, "</font id=size3>", "[/size=3]", 1, -1, 1)
			fString = replace(fString, "<font size=4>", "[size=4]", 1, -1, 1)
			fString = replace(fString, "</font id=size4>", "[/size=4]", 1, -1, 1)
			fString = replace(fString, "<font size=5>", "[size=5]", 1, -1, 1)
			fString = replace(fString, "</font id=size5>", "[/size=5]", 1, -1, 1)
			fString = replace(fString, "<font size=6>", "[size=6]", 1, -1, 1)
			fString = replace(fString, "</font id=size6>", "[/size=6]", 1, -1, 1)
			fString = replace(fString, "<br />","[br]" & vbcrlf, 1, -1, 1)
			fString = replace(fString,"<br />",vbcrlf)
			fString = replace(fString,"</p><p>",vbcrlf & vbcrlf)
			fString = replace(fString,"<p>","")
			fString = replace(fString,"</p>","")
		    fString = replace(fString, "<div align=left>", "[left]", 1, -1, 1)
		    fString = replace(fString, "</div id=left>", "[/left]", 1, -1, 1)
			fString = replace(fString, "<center>","[center]", 1, -1, 1)
			fString = replace(fString, "</center>","[/center]", 1, -1, 1)
		    fString = replace(fString, "<div align=right>", "[right]", 1, -1, 1)
		    fString = replace(fString, "</div id=right>", "[/right]", 1, -1, 1)
			fString = replace(fString, "<ul>","[list]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "</ul>","[/list]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "<ol>","[list=1]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "</ol>","[/list=1]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "<ol type=1>","[list=1]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "</ol id=1>","[/list=1]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "<ol type=a>","[list=a]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "</ol id=a>","[/list=a]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "<li>","[*]", 1, -1, 1)
			fString = replace(fString, "</li>","[/*]" & vbcrlf, 1, -1, 1)
			fString = replace(fString, "<BLOCKQUOTE id=quote><font id=quote>quote: <hr width=99% height=1 noshade id=quote>","[quote]", 1, -1, 1)
			fString = replace(fString, "<hr width=99% height=1 noshade id=quote></font id=quote></BLOCKQUOTE id=quote><font id=quote>","[/quote]", 1, -1, 1)
			fString = replace(fString, "<pre id=code><font face=courier id=code>","[code]", 1, -1, 1)
			fString = replace(fString, "</font id=code></pre id=code>","[/code]", 1, -1, 1)
				fString = replace(fString, "<img hspace=7 vspace=5 src=""","[img]", 1, -1, 1)
				fString = replace(fString, "<img align=right hspace=7 vspace=5 src=""","[img=right]", 1, -1, 1)
				fString = replace(fString, "<img align=left hspace=7 vspace=5 src=""","[img=left]", 1, -1, 1)
				fString = replace(fString, """ border=0 />","[/img]", 1, -1, 1)
				fString = replace(fString, """ id=right border=0 />","[/img=right]", 1, -1, 1)
				fString = replace(fString, """ id=left border=0 />","[/img=left]", 1, -1, 1)
		end if
		if strIcons = "1" then
			fString= replace(fString, "<img src=images/Smilies/angry.gif border=0 align=middle />", "[:(!]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/blackeye.gif border=0 align=middle />", "[B)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/dead.gif border=0 align=middle />", "[xx(]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/dead.gif border=0 align=middle />", "[XX(]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/shock.gif border=0 align=middle />", "[:O]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/shock.gif border=0 align=middle />", "[:o]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/shock.gif border=0 align=middle />", "[:0]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/blush.gif border=0 align=middle />", "[:I]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/sad.gif border=0 align=middle />", "[:(]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/shy.gif border=0 align=middle />", "[8)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/smile.gif border=0 align=middle />", "[:)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/evil.gif border=0 align=middle />", "[}:)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/big.gif border=0 align=middle />", "[:D]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/cool.gif border=0 align=middle />", "[8D]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/sleepy.gif border=0 align=middle />", "[|)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/clown.gif border=0 align=middle />", "[:o)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/clown.gif border=0 align=middle />", "[:O)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/clown.gif border=0 align=middle />", "[:0)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/tongue.gif border=0 align=middle />", "[:P]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/tongue.gif border=0 align=middle />", "[:p]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/wink.gif border=0 align=middle />", "[;)]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/8ball.gif border=0 align=middle />", "[8]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/question.gif border=0 align=middle />", "[?]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/approve.gif border=0 align=middle />", "[^]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/dissapprove.gif border=0 align=middle />", "[V]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/dissapprove.gif border=0 align=middle />", "[v]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/kisses.gif border=0 align=middle />", "[:X]", 1, -1, 1)
			fString= replace(fString, "<img src=images/Smilies/kisses.gif border=0 align=middle />", "[:x]", 1, -1, 1)
		end if
	fString = Replace(fString, "'", "'")
	CleanCode2 = fString
end Function


Function CleanCode3(fString)
' New function by Hawk92 - source code box mod - 11-2004
' Provides processing of messages containing code to display for editing
' replaces new codebox html back to [CODE] [/CODE] markers
' replaces [APOS] back to single quotes
' all parts of the strings outside the markers are processed through CleanCode2

' This replaces the starting html for the sourcecode box with the [CODE] marker
strPattern="(\[@@\])(.*\[@@\])"
Set regEx = New RegExp
regEx.Pattern = strPattern
regEx.IgnoreCase = True
regEx.Global = True
fString = regEx.replace(fString,"[CODE]")
' This part replaces the ending html for the soucecodebox with [/CODE] markers
regEx.Pattern ="(\[\/@@\])(.*\[\/@@\])"
fString = regEx.replace(fString,"[/CODE]")
' This turns [APOS] back into a single quote
fString=Replace(fString,"[APOS]","'",1,-1,1)
ptr=1
max=Len(fString)
strFinal = ""
' This is a loop to parse the string and clean all the non code portions
Do While (ptr < max)
  If InStr(ptr,fString,"[CODE]",1)>0 then
  sptr=InStr(ptr,fString,"[CODE]",1)
  eptr=InStr(ptr,fString,"[/CODE]",1)+7
  strTemp=Mid(fString,ptr,(sptr-ptr) )
  strTemp=CleanCode2(strTemp)
  strFinal=strFinal & strTemp& Mid(fstring,sptr,eptr-sptr)
  ptr=eptr
  Else
  ptr=max
  End if
Loop
  strFinal=strFinal& CleanCode2(Mid(fString,eptr,max-eptr+1))
  CleanCode3=strFinal
End Function
'#################################################################################
'## End of Replacement functions for CleanCode
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## ##############################################################################

function Smile(fString)
	fString = replace(fString, "[:(!]", "<img src=images/Smilies/angry.gif border=0 align=middle />")
	fString = replace(fString, "[B)]", "<img src=images/Smilies/blackeye.gif border=0 align=middle />")
	fString = replace(fString, "[xx(]", "<img src=images/Smilies/dead.gif border=0 align=middle />")
	fString = replace(fString, "[XX(]", "<img src=images/Smilies/dead.gif border=0 align=middle />")
	fString = replace(fString, "[:I]", "<img src=images/Smilies/blush.gif border=0 align=middle />")
	fString = replace(fString, "[:(]", "<img src=images/Smilies/sad.gif border=0 align=middle />")
	fString = replace(fString, "[:o]", "<img src=images/Smilies/shock.gif border=0 align=middle />")
	fString = replace(fString, "[:O]", "<img src=images/Smilies/shock.gif border=0 align=middle />")
	fString = replace(fString, "[:0]", "<img src=images/Smilies/shock.gif border=0 align=middle />")
	fString = replace(fString, "[|)]", "<img src=images/Smilies/sleepy.gif border=0 align=middle />")
	fString = replace(fString, "[:)]", "<img src=images/Smilies/smile.gif border=0 align=middle />")
	fString = replace(fString, "[:D]", "<img src=images/Smilies/big.gif border=0 align=middle />")
	fString = replace(fString, "[}:)]", "<img src=images/Smilies/evil.gif border=0 align=middle />")
	fString = replace(fString, "[:o)]", "<img src=images/Smilies/clown.gif border=0 align=middle />")
	fString = replace(fString, "[:O)]", "<img src=images/Smilies/clown.gif border=0 align=middle />")
	fString = replace(fString, "[:0)]", "<img src=images/Smilies/clown.gif border=0 align=middle />")
	fString = replace(fString, "[8)]", "<img src=images/Smilies/shy.gif border=0 align=middle />")
	fString = replace(fString, "[8D]", "<img src=images/Smilies/cool.gif border=0 align=middle />")
	fString = replace(fString, "[:P]", "<img src=images/Smilies/tongue.gif border=0 align=middle />")
	fString = replace(fString, "[:p]", "<img src=images/Smilies/tongue.gif border=0 align=middle />")
	fString = replace(fString, "[;)]", "<img src=images/Smilies/wink.gif border=0 align=middle />")
	fString = replace(fString, "[8]", "<img src=images/Smilies/8ball.gif border=0 align=middle />")
	fString = replace(fString, "[?]", "<img src=images/Smilies/question.gif border=0 align=middle />")
	fString = replace(fString, "[^]", "<img src=images/Smilies/approve.gif border=0 align=middle />")
	fString = replace(fString, "[V]", "<img src=images/Smilies/dissapprove.gif border=0 align=middle />")
	fString = replace(fString, "[v]", "<img src=images/Smilies/dissapprove.gif border=0 align=middle />")
	fString = replace(fString, "[:X]", "<img src=images/Smilies/kisses.gif border=0 align=middle />")
	fString = replace(fString, "[:x]", "<img src=images/Smilies/kisses.gif border=0 align=middle />")
	Smile = fString
end function

function HTMLEncode(fString)
        if trim(fString) = "" or isNull(fString) then
		HTMLEncode = " "
	else
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")

		HTMLEncode = fString
	end if
	
end function

function HTMLDecode(fString)
        if trim(fString) = "" or isNull(fString) then
		HTMLDecode = " "
	else
		fString = replace(fString, "&gt;", ">")
		fString = replace(fString, "&lt;", "<")

		HTMLDecode = fString
	end if
end function


'#################################################################################
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## by Hawk92 - 11-2004
'## Original ChkString replaced by 3 functions ChkString,DoMsgCode,SaveForumCode
'## ##############################################################################
function ChkString(fString,fField_Type) 
'## Types - name, password, title, message, url, urlpath, email, number, list
'## this function cleans data for saving to db, displaying etc
'## Modified by Hawk92 - source code box mod - 11-2004
'## This function now provides branching control to new DoMsgcode function if its a message
'## with code and to SaveForumcode if not a message with code
	fString = trim(fString)
	fField_Type = lcase(fField_Type)
	if fString = "" then
		fString = " "
		exit function
	else
		' ChkBadWords(fString)
	end if
	If fField_Type = "clean" then
			fString = Replace(fString, """", "", 1, -1, 1)
			fString = Replace(fString, "<", "", 1, -1, 1)
			fString = Replace(fString, ">", "", 1, -1, 1)
			fString = Replace(fString, "%", "", 1, -1, 1)
			fString = Replace(fString, "'", "", 1, -1, 1)
			fString = Replace(fString, "\", "", 1, -1, 1) 
			fString = Replace(fString, "|", "", 1, -1, 1) 
			fString = Replace(fString, "--", "", 1, -1, 1) 
			fString = Replace(fString, ";", "", 1, -1, 1)
			fString = Replace(fString, "#", "", 1, -1, 1)
			fString = Replace(fString, "*", "", 1, -1, 1)
			'fString = HTMLEncode(fString)
			ChkString = fString
			Exit function
	end If
	If fField_Type = "cleanurl" then
			fString = Replace(fString, " ", "_", 1, -1, 1)
			fString = Replace(fString, """", "", 1, -1, 1)
			fString = Replace(fString, "<", "", 1, -1, 1)
			fString = Replace(fString, ">", "", 1, -1, 1)
			fString = Replace(fString, "'", "", 1, -1, 1)
			fString = Replace(fString, "\", "/", 1, -1, 1) 
			fString = Replace(fString, "|", "", 1, -1, 1) 
			fString = Replace(fString, "--", "", 1, -1, 1) 
			fString = Replace(fString, ";", "", 1, -1, 1)
			fString = Replace(fString, "#", "", 1, -1, 1)
			fString = Replace(fString, "*", "", 1, -1, 1)
			fString = Replace(fString, "(", "", 1, -1, 1)
			fString = Replace(fString, ")", "", 1, -1, 1)
			fString = Replace(fString, "[", "", 1, -1, 1)
			fString = Replace(fString, "]", "", 1, -1, 1)
			fString = Replace(fString, "{", "", 1, -1, 1)
			fString = Replace(fString, "}", "", 1, -1, 1)
			fString = Replace(fString, "^", "", 1, -1, 1)
			fString = Replace(fString, "!", "", 1, -1, 1)
			fString = Replace(fString, "$", "", 1, -1, 1)
			'fString = Replace(fString, ":", "", 1, -1, 1)
			'fString = Replace(fString, ".", "", 1, -1, 1)
			fString = Replace(fString, ",", "", 1, -1, 1)
			fString = Replace(fString, "`", "", 1, -1, 1)
			'fString = HTMLEncode(fString)
			ChkString = fString
			Exit function
	end If
	if fField_Type = "refer" then
			fString = Replace(fString, "&#", "#")
			fString = Replace(fString, """", "&quot;")
			fString = HTMLEncode(fString)
			ChkString = fString
			Exit function
	end if
	if fField_Type = "decode" then
			fString = HTMLDecode(fString)
			ChkString = fString
			exit function
	end if
	if fField_Type = "urlpath" then
			fString = Server.URLEncode(fString)
			ChkString = trim(fString)
			exit function
	end if
	If fField_Type = "sqlstring" or fField_Type = "display" then
			'fString = Replace(fString, "&", "&amp;", 1, -1, 1) 
			fString = Replace(fString, ";", "&#59;", 1, -1, 1) 
			'fString = Replace(fString, "'", "''", 1, -1, 1)
			'fString = Replace(fString, "'", "&#039;", 1, -1, 1)
			fString = Replace(fString, "<", "&lt;", 1, -1, 1) 
			fString = Replace(fString, ">", "&gt;", 1, -1, 1) 
			'fString = Replace(fString, "[", "&#091;", 1, -1, 1) 
			'fString = Replace(fString, "]", "&#093;", 1, -1, 1) 
			fString = Replace(fString, """", "&quot;", 1, -1, 1) 
			'fString = Replace(fString, "=", "", 1, -1, 1) 
			fString = Replace(fString, "'", "&#39;", 1, -1, 1) 
			'fString = Replace(fString, "%", "", 1, -1, 1) 
			'fString = Replace(fString, "*", "", 1, -1, 1)
			fString = Replace(fString, "\", "", 1, -1, 1) 
			fString = Replace(fString, "|", "", 1, -1, 1) 
			fString = Replace(fString, "--", "", 1, -1, 1) 
			'fString = Replace(fString, "#", "", 1, -1, 1) 
			'fString = Replace(fString, "select", "sel&#101;ct", 1, -1, 1) 
			'fString = Replace(fString, "join", "jo&#105;n", 1, -1, 1) 
			'fString = Replace(fString, "union", "un&#105;on", 1, -1, 1) 
			'fString = Replace(fString, "where", "wh&#101;re", 1, -1, 1) 
			'fString = Replace(fString, "exec", "ex&#101;c", 1, -1, 1) 
			'fString = Replace(fString, "insert", "ins&#101;rt", 1, -1, 1) 
			'fString = Replace(fString, "delete", "del&#101;te", 1, -1, 1) 
			'fString = Replace(fString, "update", "up&#100;ate", 1, -1, 1) 
			'fString = Replace(fString, "like", "lik&#101;", 1, -1, 1) 
			'fString = Replace(fString, "drop", "dro&#112;", 1, -1, 1) 
			'fString = Replace(fString, "create", "cr&#101;ate", 1, -1, 1) 
			'fString = Replace(fString, "modify", "mod&#105;fy", 1, -1, 1) 
			'fString = Replace(fString, "rename", "ren&#097;me", 1, -1, 1) 
			'fString = Replace(fString, "alter", "alt&#101;r", 1, -1, 1) 
			'fString = Replace(fString, "cast", "ca&#115;t", 1, -1, 1) 
			'fString = Replace(fString, "char", "ch&#97;r", 1, -1, 1)
			ChkString = trim(fString)
		exit Function
	end if
	if fField_Type = "jsurlpath" then
		fString = Replace(fString, "'", "\'")
		fString = Server.URLEncode(fString)
		ChkString = fString
		exit function
	end if
	if fField_Type = "edit" then		
		if strAllowHTML <> "1" then			
			fString = HTMLEncode(fString)		
		end if		
		fString = Replace(fString, """", "&quot;", 1, -1, 1)
		ChkString = fString		
		exit function	
	end if
	if fField_Type = "display2" then
		if strBadWordFilter = "1" then
		  fString = chkBadWords(fString)
		end if
		fString = replace(fString,"+","&#043;", 1, -1, 1)
		fString = replace(fString, """", "&quot;", 1, -1, 1)
		fString = replace(fString, "''", "'", 1, -1, 1)
        ChkString = fString
		exit function
	elseif fField_Type = "message" then
		if strAllowHTML = 0 and strAllowForumCode = 1 Then
			 'If InStr(1,fString,"[CODE]",1)>0 Then
	         '   fString=doMsgCode(fstring)
	         'Else
				fString = HTMLEncode(fString)
				fString=saveForumCode(fString,fField_Type)
			 'End if
			 fstring = replace(fstring,vbcrlf,"<br />")
		else
			'fString = HTMLEncode(fString)
			 fstring = replace(fstring,vbcrlf,"<br />")
			 fstring = replace(fstring,vblf,"<br />")
			 fstring = replace(fstring,vbcr,"<br />")
			 fstring = replace(fstring,"</p><p>","<br /><br />")
			 fstring = replace(fstring,"<p>","")
			 fstring = replace(fstring,"</p>","")
			fString=replace(fString,"'","&#39;")
			fString=replace(fString,"%&#62;","%&gt;")
			fString=replace(fString,"&#60;%","&lt;%")
			fString=replace(fString,"object","obj&#101;ct")
			fString=replace(fString,"embed","emb&#101;d")
			fString=replace(fString,"iframe","ifr&#097;me")
			fString=replace(fString,"script","scr&#105;pt")
			fString=replace(fString,"javascript","jav&#097;script")
			fString=replace(fString,"http-equiv","http_&#101;quiv")
			fString=replace(fString,"alert(","al&#101;rt(")
        	'ChkString = fString
        	ChkString = chkHtmlCode(fString)
			exit function
		end if
				
				'fString = replace(fString, """", "&quot;")
	elseif fField_Type = "hidden" then
		fString = HTMLEncode(fString)
	elseif fField_Type = "numeric" then
		if not isNumeric(fString) then
          ChkString = 0
		else
          ChkString = fString
		end if
		exit function
	end if
	if fField_Type = "displayimage" then
		fString = Replace(fString, " ", "")
		fString = Replace(fString, """", "")
		fString = Replace(fString, "'", "")
		fString = Replace(fString, "<", "")
		fString = Replace(fString, ">", "")
		chkString = fString
	exit function
	end If
	if fField_Type = "preview" then
		if strAllowHTML <> "1" then
			fString = HTMLEncode(fString)
		end if
	end If

  if fField_Type <> "message"  then
	fString=saveForumCode(fString,fField_Type)
  End if
ChkString = fString
end Function

Function doMsgCode2(fString)
' New function by Hawk92 - source code box mod - 11-2004
' This function processes messages that have code in them
' This will parse the message and append html to generate the codebox
' Non code portions of the string are passed through the normal SaveForumCode Function
'
' This is the end string for the codebox
endstr="[/@@]</textarea></td></tr></table></form>[/@@]"

ptr=1
max=Len(fString)
cntr=0
strFinal=""
' This is the loop to process each part of the message String
Do While (ptr < max)
  forceclose=False
  If InStr(ptr,fString,"[CODE]",1)>0 Then
  sptr=InStr(ptr,fString,"[CODE]",1)
  eptr=InStr(ptr,fString,"[/CODE]",1)+7
  If eptr=<7 Then
  forceclose=True
  eptr=max
  End If
  Randomize()
  cntr=Int(Rnd()*9000)
  strTemp=Mid(fString,ptr,(sptr-ptr) )
  strTemp = HTMLEncode(strTemp)
  strTemp=saveForumCode(strTemp,"message")
  strCode=Mid(fstring,sptr,eptr-sptr)
  If forceclose=True Then
  strCode=strCode&"[/CODE]"
  End If
  strCode = HTMLEncode(strCode)
  
' This is the start string for the codebox
startstr="[@@]<form name=""selectcode"&cntr&""" method=""post"" action="" "">" &_
"<table style=""border: 1 solid #000000"" cellpadding=""0"" cellspacing=""0""><tr><td class=""spThemeCodeboxHead"" width=""100%"" valign=""middle"">" &_
"&nbsp;&nbsp;Source Code:" &_
"<img border=""0"" src=""images/clear.gif"" width=""144"" height=""4"" /> <img border=""0"" src=""images/click_select.gif"" onclick=""selectCodeBoxCode("&cntr&");"" title=""" & txtSelCode & """  align=""middle"" />" &_
"<img border=""0"" src=""images/clear.gif"" width=""146"" height=""13"" />" &_
"<img border=""0"" src=""images/icon_minus2.gif"" title=""" & txtCollapse & """  align=""top"" onclick=""contract("&cntr&");"" />&nbsp;" &_
"<img border=""0"" src=""images/icon_plus2.gif"" title=""" & txtExpand & """ align=""top"" onclick=""expand("&cntr&");"" />&nbsp;&nbsp;&nbsp;&nbsp;" &_
"<img border=""0"" src=""images/icon_help.gif"" title=""" & txtHelp & """ align=""middle"" onclick=""dohelp();"" /></td></tr>" &_  
"<tr><td></td></tr>" &_
"<tr><td  valign=""middle""><textarea rows=""3"" READONLY=""Yes""  id=""thecode"&cntr&""" name=""thecode"&cntr&""" cols=""70"" style=""color: #008000; font-family: courier; font-size: 10pt; text-align: left; vertical-align: top; background-image: none; background-repeat: no-repeat; border: none"">[@@]"
' Replace single quotes with [APOS] ton enable sql database saving
  strCode=Replace(strCode,"'","[APOS]",1,-1,1)
' Replace the [code] markers with the html for the codebox
  strCode=Replace(strCode,"[CODE]",startstr,1,1,1)
  strCode=Replace(strCode,"[/CODE]",endstr,1,1,1) 
  strFinal=strFinal & strTemp& strCode
  ptr=eptr
  Else
' If there are no more code markers set prt to end of string  
  ptr=max
  End if
Loop
' This picks up any part of the string after the Code
  strFinal=strFinal& saveForumCode(Mid(fString,eptr,max-eptr+1),"message") 
  strFinal = doCode(strFinal, "[quote]", "[/quote]", "<BLOCKQUOTE id=quote><font id=quote>" & txtQuote & ":<hr width=99% height=1 noshade id=quote>", "<hr width=99% height=1 noshade id=quote></font id=quote></BLOCKQUOTE id=quote>")		
doMsgCode2=strFinal
End function
sUrlK = ""
ssUrl = request.ServerVariables("SERVER_NAME")
aURL = split(ssUrl,".")
if ubound(aURL) = 1 then
  sx = 0
elseif ubound(aURL) = 2 then
  if len(aURL(1)) <= 3 then
    sx = 0
  else
    sx = 1
  end if
elseif ubound(aURL) = 3 then
  sx = 1
end if
for a = sx to ubound(aURL)
  if sUrlK <> "" then
    sUrlK = sUrlK & "."
  end if
  sUrlK = sUrlK & aURL(a)
next
ssUrl = sUrlK
s = pEncrypt(lcase(ssUrl))
sUrlK = ""
Function saveForumCode(fString,fField_Type)
' New function by Hawk92 - source code box mod - 11-2004 (mostly part of original 1.5 chkstring()
	if strAllowForumCode = "1" and fField_Type <> "signature" then
		fString = doCode(fString, "[marquee]", "[/marquee]", "<marquee>", "</marquee>")
		fString = doCode(fString, "[sup]", "[/sup]", "<sup>", "</sup>")
		fString = doCode(fString, "[sub]", "[/sub]", "<sub>", "</sub>")
		fString = doCode(fString, "[tt]", "[/tt]", "<tt>", "</tt>")
		fString = doCode(fString, "[hl]", "[/hl]", "<span style='background-color: #FFFF00'>", "<b></b></span>")
		fString = doCode(fString, "[pre]", "[/pre]", "<pre>", "</pre>")
		fString = replace(fString, "[hr]", "<hr />", 1, -1, 1)
		fString = doCode(fString, "[b]", "[/b]", "<b>", "</b>")
		fString = doCode(fString, "[s]", "[/s]", "<s>", "</s>")
		fString = doCode(fString, "[strike]", "[/strike]", "<s>", "</s>")
		fString = doCode(fString, "[u]", "[/u]", "<u>", "</u>")
		fString = doCode(fString, "[i]", "[/i]", "<i>", "</i>")
		if fField_Type <> "title" then
			fString = doCode(fString, "[font=Andale Mono]", "[/font=Andale Mono]", "<font face='Andale Mono'>", "</font id='Andale Mono'>")
			fString = doCode(fString, "[font=Arial]", "[/font=Arial]", "<font face='Arial'>", "</font id='Arial'>")
			fString = doCode(fString, "[font=Arial Black]", "[/font=Arial Black]", "<font face='Arial Black'>", "</font id='Arial Black'>")
			fString = doCode(fString, "[font=Book Antiqua]", "[/font=Book Antiqua]", "<font face='Book Antiqua'>", "</font id='Book Antiqua'>")
			fString = doCode(fString, "[font=Century Gothic]", "[/font=Century Gothic]", "<font face='Century Gothic'>", "</font id='Century Gothic'>")
			fString = doCode(fString, "[font=Courier New]", "[/font=Courier New]", "<font face='Courier New'>", "</font id='Courier New'>")
			fString = doCode(fString, "[font=Comic Sans MS]", "[/font=Comic Sans MS]", "<font face='Comic Sans MS'>", "</font id='Comic Sans MS'>")
			fString = doCode(fString, "[font=Georgia]", "[/font=Georgia]", "<font face='Georgia'>", "</font id='Georgia'>")
			fString = doCode(fString, "[font=Impact]", "[/font=Impact]", "<font face='Impact'>", "</font id='Impact'>")
			fString = doCode(fString, "[font=Tahoma]", "[/font=Tahoma]", "<font face='Tahoma'>", "</font id='Tahoma'>")
			fString = doCode(fString, "[font=Times New Roman]", "[/font=Times New Roman]", "<font face='Times New Roman'>", "</font id='Times New Roman'>")
			fString = doCode(fString, "[font=Trebuchet MS]", "[/font=Trebuchet MS]", "<font face='Trebuchet MS'>", "</font id='Trebuchet MS'>")
			fString = doCode(fString, "[font=Script MT Bold]", "[/font=Script MT Bold]", "<font face='Script MT Bold'>", "</font id='Script MT Bold'>")
			fString = doCode(fString, "[font=Stencil]", "[/font=Stencil]", "<font face='Stencil'>", "</font id='Stencil'>")
			fString = doCode(fString, "[font=Verdana]", "[/font=Verdana]", "<font face='Verdana'>", "</font id='Verdana'>")
			fString = doCode(fString, "[font=Lucida Console]", "[/font=Lucida Console]", "<font face='Lucida Console'>", "</font id='Lucida Console'>")

			fString = doCode(fString, "[red]", "[/red]", "<font color=red>", "</font id=red>")
			fString = doCode(fString, "[green]", "[/green]", "<font color=green>", "</font id=green>")
			fString = doCode(fString, "[blue]", "[/blue]", "<font color=blue>", "</font id=blue>")
			fString = doCode(fString, "[white]", "[/white]", "<font color=white>", "</font id=white>")
			fString = doCode(fString, "[purple]", "[/purple]", "<font color=purple>", "</font id=purple>")
			fString = doCode(fString, "[yellow]", "[/yellow]", "<font color=yellow>", "</font id=yellow>")
			fString = doCode(fString, "[violet]", "[/violet]", "<font color=violet>", "</font id=violet>")
			fString = doCode(fString, "[brown]", "[/brown]", "<font color=brown>", "</font id=brown>")
			fString = doCode(fString, "[black]", "[/black]", "<font color=black>", "</font id=black>")
			fString = doCode(fString, "[pink]", "[/pink]", "<font color=pink>", "</font id=pink>")
			fString = doCode(fString, "[orange]", "[/orange]", "<font color=orange>", "</font id=orange>")
			fString = doCode(fString, "[gold]", "[/gold]", "<font color=gold>", "</font id=gold>")

			fString = doCode(fString, "[beige]", "[/beige]", "<font color=beige>", "</font id=beige>")
			fString = doCode(fString, "[teal]", "[/teal]", "<font color=teal>", "</font id=teal>")
			fString = doCode(fString, "[navy]", "[/navy]", "<font color=navy>", "</font id=navy>")
			fString = doCode(fString, "[maroon]", "[/maroon]", "<font color=maroon>", "</font id=maroon>")
			fString = doCode(fString, "[limegreen]", "[/limegreen]", "<font color=limegreen>", "</font id=limegreen>")

			fString = doCode(fString, "[h1]", "[/h1]", "<h1>", "</h1>")
			fString = doCode(fString, "[h2]", "[/h2]", "<h2>", "</h2>")
			fString = doCode(fString, "[h3]", "[/h3]", "<h3>", "</h3>")
			fString = doCode(fString, "[h4]", "[/h4]", "<h4>", "</h4>")
			fString = doCode(fString, "[h5]", "[/h5]", "<h5>", "</h5>")
			fString = doCode(fString, "[h6]", "[/h6]", "<h6>", "</h6>")
			fString = doCode(fString, "[size=1]", "[/size=1]", "<font size=1>", "</font id=size1>")
			fString = doCode(fString, "[size=2]", "[/size=2]", "<font size=2>", "</font id=size2>")
			fString = doCode(fString, "[size=3]", "[/size=3]", "<font size=3>", "</font id=size3>")
			fString = doCode(fString, "[size=4]", "[/size=4]", "<font size=4>", "</font id=size4>")
			fString = doCode(fString, "[size=5]", "[/size=5]", "<font size=5>", "</font id=size5>")
			fString = doCode(fString, "[size=6]", "[/size=6]", "<font size=6>", "</font id=size6>")
			fString = doCode(fString, "[list]", "[/list]", "<ul>", "</ul>")
			fString = doCode(fString, "[list=1]", "[/list=1]", "<ol type=1>", "</ol id=1>")
			fString = doCode(fString, "[list=a]", "[/list=a]", "<ol type=a>", "</ol id=a>")
			fString = doCode(fString, "[*]", "[/*]", "<li>", "</li>")
			fString = doCode(fString, "[left]", "[/left]", "<div align=left>", "</div id=left>")
			fString = doCode(fString, "[center]", "[/center]", "<center>", "</center>")
			fString = doCode(fString, "[centre]", "[/centre]", "<center>", "</center>")
			fString = doCode(fString, "[right]", "[/right]", "<div align=right>", "</div id=right>")
			fString = doCode(fString, "[code]", "[/code]", "<pre id=code><font face=courier id=code>", "</font id=code></pre id=code>")
'			fString = doCode(fString, "[quote=", "[/quote]", "<BLOCKQUOTE id=quote><font size=" & strFooterFontSize & " id=quote><hr height=1 noshade id=quote>", "<hr height=1 noshade id=quote></BLOCKQUOTE id=quote></font id=quote><font size=" & strDefaultFontSize & " id=quote>")
			fString = doCode(fString, "[quote]", "[/quote]", "<BLOCKQUOTE id=quote><font id=quote>" & txtQuote & ": <hr width=99% height=1 noshade id=quote>", "<hr width=99% height=1 noshade id=quote></font id=quote></BLOCKQUOTE id=quote><font id=quote>")
			fString = replace(fString, "[br]", "<br />", 1, -1, 1)
				fString = ReplaceImageTags(fString)
		end if
	end if
	if strIcons = "1" and _
	fField_Type <> "title" and _
	fField_Type <> "hidden" then
		fString= smile(fString)
	end if
	if fField_Type = "preview" then
		if strAllowHTML <> "1" then
			fString = HTMLEncode(fString)
		end if
	end if
	if fField_Type <> "hidden" and _
	fField_Type <> "preview" then
		fString = Replace(fString, "'", "''")
		'fString = HTMLEncode(fString)
	end If
saveForumCode=fString
End Function
'#################################################################################
'## End of Replacement functions for ChkString
'## Functions Replaced for Sourcecode Box MOD ver 1.5
'## ##############################################################################

function chkHtmlCode(hstring) 
lookFor = "<br/><br/><small><b>" & txtQuote & ":</b></small><br/><span class=quote>"

replaceWith = "<span class=quote>"
tmpCode = replace(hString, lookFor, replaceWith)
tmpCode = replace(tmpCode, replaceWith, lookFor)
chkHtmlCode = tmpCode
end function

function ReadLastHereDate(UName)
	dim tLastHereDate
	dim rsLHD

  if UName <> "" then
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & UName & "'"
	set rsLHD = my_conn.Execute(strSql)

	if rsLHD.EOF then
	  tLastHereDate = DateToStr(DateAdd("d",-10,strCurDateAdjust))
	else
	  tLastHereDate = rsLHD("M_LASTHEREDATE")
	  if tLastHereDate = "" or IsNull(tLastHereDate) then
		tLastHereDate = DateToStr(DateAdd("d",-10,strCurDateAdjust))
	  end if	
	end if
	set rsLHD = nothing
	
	' - Do DB Update
	strCurDateString = dateToStr(now())
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & "SET M_LASTHEREDATE = '" & strCurDateString & "', M_LAST_IP = '" & Request.ServerVariables("REMOTE_HOST") & "'"
	strSql = strSql & " WHERE M_NAME = '" & UName & "'"
	executeThis(strSql)
	Call setCookie("last_here_date",strCurDateString,360)
  else
    tLastHereDate = readCookie("last_here_date")
	if len(tLastHereDate & "x") = 1 then
	  tLastHereDate = DateToStr(DateAdd("d",-10,strCurDateAdjust))
	  Call setCookie("last_here_date",strCurDateString,360)
	end if
  end if	
	ReadLastHereDate = tLastHereDate
end function

function chkIsSuperAdmin(typ,chk)
 tmpName = ""
 tmpChk = 0
 if typ = 1 then 'id is passed
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	StrSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & chk
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	set rsChk = my_Conn.Execute(strSql)
	if not rsChk.eof then
	  tmpName = lcase(rsChk("M_NAME"))
	end if
	set rsChk = nothing
 elseif typ = 2 then ' member name is passed
   tmpName = lcase(chk)
 end if
 'we have the name, now do the work
  if trim(tmpName) <> "" then
    if isArray(tempArr) then
      for isw = 0 to ubound(tempArr)
	    if tempArr(isw) = tmpName then
	      tmpChk = 1
	    end if
      next
	else
	  if tempArr = tmpName then
	    tmpChk = 1
	  end if
	end if
  end if
  chkIsSuperAdmin = tmpChk
end function

function chkIsAdmin(uID)
	sSql = "SELECT FROM " & strTablePrefix & "GROUP_MEMBERS"
	sSql = sSql & " WHERE G_GROUP_ID = 1 AND G_MEMBER_ID = " & uID
	on error resume next
	set rsChk = my_Conn.Execute(sSql)
	if rsChk.EOF then
	  'not in db
	  tmpResult = 0
	else
	  tmpResult = 1
	end if
	set rsChk = nothing
	on error goto 0
	chkIsAdmin = tmpResult
end function

function chkIsMbr(fName, fPassword)
    t = 0
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	StrSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & fName & "'"
	if strAuthType="db" then	
	  strSql = strSql &" AND "& strMemberTablePrefix &"MEMBERS.M_PASSWORD='"& fPassword &"'"
	End If
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	on error resume next
	set rsCheck = my_Conn.Execute(strSql)
	if (rsCheck.BOF and rsCheck.EOF) then
	  t = 0 ':: Invalid Username or Password
	else
	  t = 1	
	end if	
	set rsCheck = nothing
	on error goto 0
	chkIsMbr = t
end function

function GetSig(fUser_Name)
	strSql = "SELECT M_SIG "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & fUser_Name & "'"

	set rsSig = my_Conn.Execute (strSql)

	if rsSig.EOF or rsSig.BOF then
		'## Do Nothing
		GetSig = ""
	else
	  if len(rsSig("M_SIG") & "x") > 1 then
		GetSig = replace(rsSig("M_SIG"),"<a","<a rel=""nofollow""")
	  end if 
	end if

	rsSig.close
	set rsSig = nothing

end function

function getMember_Level(fM_TITLE, fM_LEVEL, fM_POSTS)
	dim Member_Level
	Member_Level = ""
	if Trim(fM_TITLE) <> "" then
		Member_Level = fM_TITLE
	else
		select case fM_LEVEL
			case "1"  
				if (fM_POSTS < intRankLevel1) then Member_Level = Member_Level & strRankLevel0
				if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Member_Level = Member_Level & strRankLevel1
				if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Member_Level = Member_Level & strRankLevel2
				if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Member_Level = Member_Level & strRankLevel3
				if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Member_Level = Member_Level & strRankLevel4
				if (fM_POSTS >= intRankLevel5) then Member_Level = Member_Level & strRankLevel5
			case "2"
				Member_Level = Member_Level & strRankMod
			case "3"
				Member_Level = Member_Level & strRankAdmin
			case else  
				Member_Level = Member_Level & "Error" 
		end select
	end if
    getMember_Level = Member_Level
end function

function getMemberName(fUser_ID)
	'
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE MEMBER_ID = " & fUser_ID
	set rsGetMemberID = my_Conn.Execute(strSql)
	if rsGetMemberID.EOF or rsGetMemberID.BOF then
		getMemberName = ""
	else
		getMemberName = rsGetMemberID("M_NAME")
	end if
	set rsGetMemberID = nothing
end function

function getMemberID(fUser_Name)
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & fUser_Name & "'"

	Set rsGetMemID = my_Conn.Execute(strSql)
  	If Not ( rsGetMemID.BOF and rsGetMemID.EOF ) Then
      tMemberID = rsGetMemID("MEMBER_ID")
	else
	  tMemberID = 0
    End If
	set rsGetMemID = nothing
	getMemberID = tMemberID
end function

'::::::::::::::::::::::::::::::::::::::::::::::
'::        Cookie and Session routines       ::
'::::::::::::::::::::::::::::::::::::::::::::::

sub DoCookies(fSave)
	Response.Cookies(strUniqueID & "User").Path = strCookieURL
	Response.Cookies(strUniqueID & "User")("Name") = strDBNTFUserName
	Response.Cookies(strUniqueID & "User")("PWord") = pEncrypt(pEnPrefix & chkString(Request.Form("Password"),""))
	Response.Cookies(strUniqueID & "User")("Cookies") = cInt(Request.Form("Cookies"))
	if fSave = "true" then
	  Response.Cookies(strUniqueID & "User").Expires = dateAdd("d", 30, now())
	end if
	Session.Contents(strUniqueID & "last_here_date") = ReadLastHereDate(strDBNTFUserName)
	
	Response.Cookies(strUniqueID & "hide").Path = strCookieURL
	Response.Cookies(strUniqueID & "hide")("Name") = strDBNTFUserName
	Response.Cookies(strUniqueID & "hide").Expires = dateAdd("d", 30, now())
	'delete from online table
	executeThis("DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP = '" & Request.ServerVariables("REMOTE_ADDR") & "'")
	strDBNTUserName = strDBNTFUserName
end sub

sub ClearCookies()
	executeThis("DELETE FROM " & strTablePrefix & "ONLINE WHERE UserID = '" & chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring") & "'")
	executeThis("DELETE FROM " & strTablePrefix & "ONLINE WHERE UserIP = '" & Request.ServerVariables("REMOTE_ADDR") & "'")
	deleteCookie("User")
	deleteCookie("hide")
	deleteSession("last_here_date")
end sub

sub deleteCookie(key)
	Response.Cookies(strUniqueID & key).Path = strCookieURL
	Response.Cookies(strUniqueID & key) = ""
	Response.Cookies(strUniqueID & key).Expires = dateadd("d", -2, now())
end sub

function setMultiCookie(key1,key2,val,ex)
	Response.Cookies(strUniqueID & key).Path = strCookieURL
	Response.Cookies(strUniqueID & key) = val
	if len(ex & "x") = 1 then
	else
	  Response.Cookies(strUniqueID&key).Expires = dateadd("d",ex,now())
	end if
	setCookie = true
end function

function readMultiCookie(key1,key2)
  dim t
  t = ""
    t=chkString(Request.Cookies(strUniqueID & key1)(key2),"sqlstring")
	readMultiCookie = t
end function

function setCookie(key,val,expire)
	Response.Cookies(strUniqueID&key).Path = strCookieURL
	Response.Cookies(strUniqueID&key) = val
	Response.Cookies(strUniqueID&key).Expires = dateadd("d",expire,now())
	setCookie = true
end function

function readCookie(key)
    readCookie=chkString(Request.Cookies(strUniqueID & key),"sqlstring")
end function
':: End cookie routines

':: Session Routines
sub deleteSession(key)
	Session.Contents.Remove(strUniqueID & key)
end sub

function setSession(key,val)
	Session.Contents(strUniqueID & key) = "x"
	Session.Contents(strUniqueID & key) = val
end function

function readSession(key)
  dim t
  t = ""
  If Session.Contents(strUniqueID & key) <> "" Then
    t = Session.Contents(strUniqueID & key)
	Call setSession(key,t)
  end if
  readSession = t
end function
':: End Session Routines

sub doLoginForm()
%>
<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
<p align="center"><span class="fTitle">
<%= txtNoForumAcc %>.
</span></p>
<p align="center"><%= txtGotSpecPerm %>:
<form method="post" action="<% =Request.ServerVariables("SCRIPT_NAME") %>" id="form62" name="form62">
<%
	for each q in Request.QueryString
		Response.Write "<input type=""hidden"" name=""" & chkstring(q, "hidden") & """ value=""" & chkstring(Request(q), "hidden") & """ />"
	next
%>
<input class="textbox" name="pass" type="password" size="20" />
<input class="button" type="submit" value="<%= txtSubmit %>" id="submit62" name="submit62" />
</form></p>
<p align="center"><a href="JavaScript:history.go(-1);"><%= txtGoBackData %></a></p>
<p align="center"><a href="default.asp"><%= txtReturnHome %></a></p>
<!--INCLUDE FILE="inc_footer.asp"-->
<%
end sub

sub doNotLoggedInForm()
%>
<p align="center"><span class="fTitle"><%= txtThereIsProb %></span></p>
<p align="center"><span class="fTitle">
<%= txtNotLogIn %>.
</span></p>
<p align="center"><a href="JavaScript:history.go(-1);"><%= txtGoBackData %></a></p>
<p align="center"><a href="default.asp"><%= txtReturnHome %></a></p>
<!--INCLUDE FILE="inc_footer.asp"-->
<%
	'Response.End
end sub

sub doNotLoggedInGame()
  doNotLoggedInForm()
end sub

sub chkSiteLockDown()
  if strLockDown = 1 Then
    if strAuthType = "db" and not hasAccess(2) Then
      if (curPageType = "register" and strNewReg = 1) or (strEmail = 1 and curpagetype = "password") then
	    ' do nothing
      else
        closeAndGo("error.asp?type=lockdown")
	  end if
    end If
  end if
end sub

sub lockDownLoginForm()
  Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""500"">")
  Response.Write("<tr><td>")
  msgStr = "<span class=""fSubTitle""><b>" & txtMustBMember1 & "<br />" & txtToPartic & ".</b></span><br /><br />"
  Response.Write("<form action=""default.asp"" method=""post"">")
  Response.Write("<input type=""hidden"" name=""Method_Type"" value=""login"" />")
  Call showPasswordBlock2(0,"",msgStr,1,1,1)
  Response.Write("</form>")
  Response.Write("</td></tr></table>")
end sub

'showPasswordBlock2(block_type,block_title_text,"Message",bool_show_save-password,show_reg-now,bool_secimage)
sub showPasswordBlock2(p_type,tb_title,m_msg,sav_pass,reg_now,secimg_val)
if p_type = 1 then
spThemeTitle= tb_title
spThemeBlock1_open(intSkin)
end if %>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr><td width="100%" colspan="2">&nbsp;</td></tr>
<% If m_msg <> "" Then %>
<tr><td width="100%" align="center" colspan="2"><%= m_msg %></td></tr>
<% end if %>
<%
 if strAuthType="db" then %>
      <tr>
        <td width="40%" align="right" nowrap><b><%= txtUsrName %>:&nbsp;</b></td>
        <td><input type="text" name="Name" value="<%= chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size="20"></td>
      </tr>
      <tr>
        <td align="right" nowrap><b><%= txtPass %>:&nbsp;</b></td>
        <td><input type="Password" name="Password" size="20"></td>
      </tr>
<% else %>
	<tr>
	  <td width="40%" align="right" nowrap><b>NT Account:</b></td>
	  <td>&nbsp;<%=Session(strUniqueID & "userID")%></td>
	</tr>
<% end if %>
<% If secimg_val = 1 and SecImage > 1 Then %>
	<tr>
	<td align="right" nowrap><b><%= txtSecImg %>:&nbsp;</b></td>
	<td><img src="includes/securelog/image.asp" alt="<%= txtSecImg %>" title="<%= txtSecImg %>" /></td>
	</tr>
	<tr>
	<td width="40%" align="right" nowrap><b><%=txtSecCode%>:&nbsp;</b></td>
	<td><input class="textbox" type="text" name="secCode" size="15" maxlength="8" value="<%= txtSecCode %>" onfocus="javascript:this.value='';" /></td>
	</tr>
<% end if %>
<% If sav_pass = 1 Then %>
<tr><td align="center" colspan="2">
<input type="checkbox" name="SavePassWord" value="true" checked="checked" />&nbsp;&nbsp;
<span class="fSmall"><%= txtSvPass %></span></td></tr>
<% end if %>
<tr><td height="34" align="center" valign="middle" colspan="2">
	  <input class="btnLogin" type="submit" value="<%= txtSubmit %>" id="submitw1" name="submitw1" />
</td></tr>
<% If sav_pass = 1 or reg_now = 1 Then %>
<!-- <tr><td align="right" colspan="2">&nbsp;</td></tr> -->
<tr><td align="center" colspan="2">
		<%if strEmail = 1 and sav_pass = 1 then %>
            <br />
            <a href="password.asp"><span class="fSmall"><%= txtForgotPass %>?</span></a>
        <% end if %>
		<%if strNewReg = 1 and reg_now = 1 then %>
			<br /><br />
            <span class="fSmall"><%= txtNotMember %>?<br /><a href="policy.asp"><span class="fAlert"><%= txtRegNow %>!</span></a></span>
        <% End If %>
          </td>
        </tr>
<% End If %>
		<tr><td colspan="2">&nbsp;</td></tr>
		</table>
<%
  if p_type = 1 then
	spThemeBlock1_close(intSkin)
  end if
end sub

'showPasswordBlock(block_type,block_title_text,"Message",bool_show_save-password,show_reg-now)
sub showPasswordBlock(p_type,tb_title,m_msg,sav_pass,reg_now)
if p_type = 1 then
spThemeTitle= tb_title
spThemeBlock1_open(intSkin)
end if %>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr><td width="100%" align="right" colspan="2">&nbsp;</td></tr>
<% If m_msg <> "" Then %>
<tr><td width="100%" align="center" colspan="2"><%= m_msg %></td></tr>
<% end if %>
<tr>
<%
 if strAuthType="db" then %>
      <tr>
        <td class="tCellAlt0" width="40%" align="right" nowrap><b><%= txtUsrName %>:&nbsp;</b></td>
        <td class="tCellAlt0"><input type="text" name="User" value="<% =chkString(Request.Cookies(strUniqueID & "User")("Name"),"sqlstring")%>" size="20"></td>
      </tr>
      <tr>
        <td class="tCellAlt0" align="right" nowrap><b><%= txtPass %>:&nbsp;</b></td>
        <td class="tCellAlt0"><input type="Password" name="Pass" size="20"></td>
      </tr>
<% else %>
	<tr>
	  <td class="tCellAlt0" width="40%" align="right" nowrap><b>NT Account:</b></td>
	  <td class="tCellAlt0">&nbsp;<%=Session(strUniqueID & "userID")%></td>
	</tr>
<% end if %>
<tr><td height="34" align="center" valign="middle" colspan="2">
	  <input class="btnLogin" type="submit" value="<%= txtLogin %>" id="submitw1" name="submitw1" />
</td></tr>
<% If sav_pass = 1 or reg_now = 1 Then %>
<tr><td align="right" colspan="2"></td></tr>
<tr><td align="center" colspan="2">
<input type="checkbox" name="SavePassWord" value="true" checked="checked" />
<span class="fSmall"><%= txtSvPass %></span>
		<%if strEmail = 1 and sav_pass = 1 then %>
            <br />
            <a href="password.asp"><span class="fSmall"><%= txtForgotPass %>?</span></a>
        <% end if %>
		<%if strNewReg = 1 and reg_now = 1 then %>
			<br /><br />
            <span class="fSmall"><%= txtNotMember %>?<br /><a href="policy.asp"><%= txtRegNow %>!</a></span>
        <% End If %>
          </td>
        </tr>
<% End If %>
		<tr><td align="right" colspan="2"></td></tr>
		</table>
<%
  if p_type = 1 then
	spThemeBlock1_close(intSkin)
  end if
end sub

sub showLoginBlock(strLogMsg)
spThemeTitle= txtLogin
spThemeBlock1_open(intSkin)
If strSiteLockdown <> 5 Then %>
<form action="default.asp" method="post" id="formw1" name="formw1">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<input type="hidden" name="Method_Type" value="login" />
<tr><td width="100%" align="right" colspan="2">&nbsp;</td></tr>
<% If strLogMsg <> "" Then %>
<tr><td width="100%" align="center" colspan="2"><%= strLogMsg %></td></tr>
<% end if %>
<tr>
          <td width="50%" align="right" valign="middle"><b><%= txtUsrName %>:&nbsp; 
            </b></td>
<td width="50%" align="left" valign="middle"><input class="textbox" type="text" name="Name" size="10" /></td></tr>
<tr>
          <td align="right" valign="middle"><b><%= txtPass %>:&nbsp; 
            </b></td>
<td align="left" valign="middle"><input class="textbox" type="password" name="Password" size="10" /></td></tr>
<tr><td height="34" align="center" valign="middle" colspan="2">
	  <input class="btnLogin" type="submit" value="<%= txtLogin %>" id="submitw1" name="submitw1" />
</td></tr>
<tr><td align="right" colspan="2"></td></tr>
<tr><td align="center" colspan="2">
<input type="checkbox" name="SavePassWord" value="true" checked="checked" />
<span class="fSmall"><%= txtSvPass %></span>
		<%if (strEmail = 1) then %>
            <br />
            <a href="password.asp"><span class="fSmall"><%= txtForgotPass %>?</span></a>
        <% end if %>
		<%if strNewReg = 1 then %>
			<br /><br />
            <span class="fSmall"><%= txtNotMember %>?<br /><a href="policy.asp"><%= txtRegNow %>!</a></span>
        <% End If %>
          </td>
        </tr></table></form>
<% Else %>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr><td width="100%" align="center" colspan="2">
	  <b>
	  <%= txtSiteDown %><br />
	  <%= txtChkBack %>.
	  </b>
	  </td></tr>
	</table>
<% End If 
	spThemeBlock1_close(intSkin)
end sub

Function ReplaceUrls(fString)
	Dim oTag, c1Tag, c2Tag
	Dim roTag, rc1Tag, rc2Tag
	Dim oTagPos, c1TagPos, c2TagPos
	Dim nTagPos
	Dim counter2
	Dim strArray, strArray2, strArray3

    oTag   = "[url="""
    oTag2  = "[url]"
    roTag  = "<a href="""
    c1Tag  = """]"
    c1Tag2 = "[/url]"
    rc1Tag = """ target=""_New"">"
    c2Tag  = "[/url]"
    rc2Tag = "</a>"
    oTagPos = InStr(1, fString, oTag, 1)
'	if oTagpos > 0 then
      c1TagPos = InStr(1, fString, c1Tag, 1)
'	else
'	  c1TagPos = 0
'	end if
   
strTempString = ""
if (oTagpos > 0) and (c1TagPos > 0) then
	strArray = Split(fString, oTag, -1)

	for counter2 = 0 to UBound(strArray)
		if (InStr(1, strArray(counter2), c2Tag, 1) > 0) and (InStr(1, strArray(counter2), c1Tag, 1) > 0) then
			strArray2 = Split(strArray(counter2), c1Tag, -1)
			if Instr(1, strArray2(1), c2Tag) and not ((Instr(1, UCase(strArray2(1)), "[URL]") >0) and not (Instr(1, UCase(strArray2(1)), "[/URL]") >0)) then
'			if Instr(1, strArray2(1), c2Tag) then  
				strFirstPart = Left(strArray2(1), Instr(1, strArray2(1),c2Tag)-1)
				strSecondPart = Right(strArray2(1), (Len(strArray2(1)) - Instr(1, strArray2(1), c2Tag) - len(c2Tag)+1))
				if strFirstPart <> "" then
					if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
						strTempString = strTempString & roTag & "mailto:" & replace(strArray2(0),"""","") & rc1Tag & strFirstPart & rc2Tag & strSecondPart
					else
						strTempString = strTempString & roTag & replace(strArray2(0),"""","") & rc1Tag & strFirstPart & rc2Tag & strSecondPart
					end if
				else
					if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
						strTempString = strTempString & roTag & "mailto:" & replace(strArray2(0),"""","") & rc1Tag & replace(strArray2(0),"""","") & rc2Tag & strSecondPart
					else
						strTempString = strTempString & roTag & replace(strArray2(0),"""","") & rc1Tag & replace(strArray2(0),"""","") & rc2Tag & strSecondPart
					end if
				end if
				if ubound(strArray2) >= 2 then
					for cnt = 2 to ubound(strArray2)
						strTempString = strTempString & """]" & strArray2(cnt)
					next
				end if
			else
				strTempString = strTempString & roTag & replace(strArray2(0),"""","") & rc1Tag & replace(strArray2(0),"""","") & rc2Tag & strArray2(1)
				if ubound(strArray2) >= 2 then
					for cnt = 2 to ubound(strArray2)
						strTempString = strTempString & """]" & strArray2(cnt)
					next
				end if
			end if
		elseif (InStr(1, strArray(counter2), c1Tag, 1) > 0) then
			if counter2 = 0 then
				strTempString = strTempString & strArray(counter2)
			else
				strArray2 = Split(strArray(counter2), c1Tag, -1)
				strTempString = strTempString & roTag & replace(strArray2(0),"""","") & rc1Tag & replace(strArray2(0),"""","") & rc2Tag & strArray2(1)
				if ubound(strArray2) >= 2 then
					for cnt = 2 to ubound(strArray2)
						strTempString = strTempString & """]" & strArray2(cnt)
					next
				end if
			end if
		else
			strTempString = strTempString & strArray(counter2)
		end if
	next

else
	strTempString = fString
end if

oTagPos2 = InStr(1, strTempString, oTag2, 1)
'	if oTagpos2 > 0 then
      c1TagPos2 = InStr(1, strTempString, c1Tag2, 1)
'	else
'	  c1TagPos2 = 0
'	end if

if (oTagpos2 > 0) and (c1TagPos2 > 0) then
 	strTempString2 = ""
 	strArray = Split(strTempString, oTag2, -1)
 	for counter3 = 0 to Ubound(strArray)
 		if (Instr(1, strArray(counter3), c1Tag2) > 0) then
 			strArray2 = split(strArray(counter3), c1Tag2, -1)
			if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
	 			strTempString2 = strTempString2 & roTag & "mailto:" & replace(strArray2(0),"""","") & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
			else
	 			strTempString2 = strTempString2 & roTag & replace(strArray2(0),"""","") & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
			end if
 		else
 			strTempString2 = strTempString2 & strArray(counter3)
 		end if	
 	next  
 	strTempString = strTempString2
end if
	ReplaceUrls = strTempString
end function

Function displayName(nam,glo) 'name,glow colors
  shoGlo = ""
  if varBrowser = "ie" then
	if trim(glo) <> "" then
	  gloStr = split(glo,":")
	  gloColor = gloStr(0)
	  txColor = gloStr(1)
	  if len(gloColor) < 6 then
		gloColor = gloColor
	  else
	    if left(gloColor,1) <> "#" then
		  gloColor = "#" & gloColor
		end if
	  end if
	  if len(txColor) < 6 then
		txColor = txColor
	  else
		if left(txColor,1) <> "#" then
		  txColor = "#" & txColor
		end if
	  end if
	  shoGlo = "<font style=""filter:glow(color:" & gloColor & ",strength:4); width:100%; cursor:pointer;"" color=""" & txColor & """>" & nam & "</font>"
	else
	  shoGlo = nam
    end if
  Else
    shoGlo = nam
  End If
  displayName = shoGlo
End Function

Function showGlow(strng) 'mname,lev
	gloStr = split(strng,":")
	gloColor = gloStr(0)
	txColor = gloStr(1)
	showGlow = "style=""filter:glow(color:" & gloColor & ",strength:4); width:100%; cursor:pointer;"" color=""" & txColor & """"
End Function

Function pmCheck() ' Start check for new PM
	if strDBNTUserName = "" Then
		pmimage = ""
		pmCount = 0
	else 
	  'if chkApp("pm","USERS") then
		if strDBType = "access" then
			strSqL = "SELECT COUNT(M_TO) AS [pmcount] " 
		else
			strSqL = "SELECT COUNT(M_TO) AS pmcount" 
		end if
		strSql = strSql & " FROM " & strTablePrefix & "PM"
		strSql = strSql & " WHERE M_TO = " & strUserMemberID
		strSql = strSql & " AND M_READ = 0" 

		Set rsPMN = my_Conn.Execute(strSql)
		'if not rsPMN.EOF then
		'else
		'end if
		  pmCnt = clng(rsPMN("pmcount"))
		'rsPMN.Close
		set rsPMN = nothing
		if pmCnt <> 0 then
			'pmimage = "Hello"
			If strPMtype = 0 or strPMtype = 2 then
			pmimage = "<a href=""pm.asp""><img src=""images/icons/icon_pm2.gif"" alt=""" & txtNewPM & """ border=""0"" /></a>"
			end if
			pmCount = pmCnt
		end if
	  'end if
	end if
End Function ' End check for new PM

function newPM()
  np = ""
  if pmCount <> 0 then
    np = "&nbsp;(" & pmCount & ")"
  end if
  newPM = np
end function

'  This gets counts. Default is zero
'  Use: getCount("url","LINKS","SHOW=0")
'  Produces: Select count(url) FROM LINKS WHERE SHOW=0
'  Returns: integer
Function getCount(cntField,cntTable,cntParam) 
	rcont = 0  ' This sets the default count
	If cntField <> "" and cntTable <> "" and cntParam <> "" then
		strSQL = "SELECT count(" & cntField & ") FROM " & cntTable & " WHERE " & cntParam
		on error resume next
		Set RScont = my_Conn.Execute(strSQL)
		if not RScount.eof then
			rcont = RScont(0)
		end if
		set RScont = nothing
		on error goto 0
	end if
	getCount = rcont
End Function

' Get the pending tasks count
Function getPendingTaskCount()
	PTcnt = 0
	getPendingTaskCount = PTcnt
End Function

Function MYURLEncode(strInput)
Dim I
Dim strTemp
Dim strChar
Dim strOut
Dim intAsc

strTemp = Trim(strInput)
For I = 1 To Len(strTemp)
    strChar = Mid(strTemp, I, 1)
    intAsc = Asc(strChar)
    If intAsc = 10 or intAsc = 13 Then
        strOut = strOut & "\n"
    else
        strOut = strOut & strChar
    End If
Next

MYURLEncode = strOut

End Function

function getReported()
If hasAccess(1) Then
	If getCount("R_STATUS",strTablePrefix & "REPORTED_POST","R_STATUS=0") <> 0 Then
		getReported = "<a href=""forum_report_post_moderate.asp""><img src=""" & strHomeUrl & "images/icons/icon_reported.gif"" width=""16"" height=""16"" alt=""" & txtRptdPst & """ title=""" & txtViewRptPost & "."" border=""0"" style""display:inline; position:absolute;"" /></a>"
	else
		getReported = ""
	End If
else
	getReported = ""
End If 
end function

Function GetAge(dtBd)
	Dim dtToday
	Dim iAge
	dtToday = date()
	iAge = Year(dtToday) - Year(dtBd)
	If (Month(dtToday) * 100 + Day(dtToday)) < (Month(dtBd) * 100 + Day(dtBd)) then iAge = iAge -1
	GetAge = iAge
End Function

'##############################################
'##        Security Image check function     ##
'##############################################
Function DoSecImage(fSecCode)
Dim strEncCode
strEncCode = pEncrypt(ucase(fSecCode))
If strEncCode = Session("secCode") Then
  DoSecImage = true
Else
  DoSecImage = false
End If
End Function

sub shoSecurityImg() %>
  <p><%= txtCopySecImg %>
  <br /><img src="includes/securelog/image.asp" vspace="4" alt="<%= txtSecImg %>" title="<%= txtSecImg %>" /><br />
  <input class="textbox" type="text" id="secCode" name="secCode" size="15" maxlength="8" /></p>
<%
end sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::         SP v1.0 additions                     :::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
function ipgate_js() %>
<script type="text/javascript">
<!-- Begin
var SpecialWord = "sp20login",
    SpecialUrl = "admin_login.asp?target=admin_ipgate.asp",
    SpecialLetter = 0;
function getKeyP(keyStroke) {
var isNetscape=(document.layers);
// Cross-browser key capture routine couresty
// of Randy Bennett (rbennett@thezone.net)
var eventChooser = (isNetscape) ? keyStroke.which : event.keyCode;
var which = String.fromCharCode(eventChooser).toLowerCase();
if (which == SpecialWord.charAt(SpecialLetter)) {
SpecialLetter++;
if (SpecialLetter == SpecialWord.length) window.location = SpecialUrl;
}
else SpecialLetter = 0;
}
document.onkeypress = getKeyP;
//  End -->
</script>
<%
end function

function ipgate_noaccess()
  ipgate_js()
  spThemeTitle=txtNoAccess
  spThemeBlock1_open(intSkin)
  %><p>&nbsp;</p>
  <table align="center" border="0" cellpadding="0" cellspacing="0" width="500">
    <tr>
        <td align="center" valign="top" class="fTitle"><%= txtIPGateNoAccess %>!<br /><br />  <%= txtIPGtxt1 %> <a href="mailto:<% =strSender %>"><%= txtAdminst %></a> <%= txtIPGtxt2 %>.</td>
    </tr>
  </table><p>&nbsp;</p>
  <%
  spThemeBlock1_close(intSkin)
end function

function ipgate_lockdown()
  ipgate_js()
  spThemeTitle=txtNoAccess
  spThemeBlock1_open(intSkin)
  %><p>&nbsp;</p>
  <table align="center" border="0" cellpadding="0" cellspacing="0" width="500">
    <tr><td align="center" valign="top" class="fTitle"><%= txtIPGlockdown %>.</td></tr>
  </table><p>&nbsp;</p>
  <%
  spThemeBlock1_close(intSkin)
end function

function ipgate_banned()
  ipgate_js()
  spThemeTitle=txtNoAccess
  spThemeBlock1_open(intSkin)
  %><p>&nbsp;</p>
  <table align="center" border="0" cellpadding="0" cellspacing="0" width="500">
    <tr>
        <td align="center" valign="top" class="fTitle"><%= txtIPGbanned %>!<br /><br />  <%= txtIPGtxt1 %> <a href="mailto:<% =strSender %>"><%= txtAdminst %></a> <%= txtIPGtxt2 %>.</td>
    </tr>
  </table><p>&nbsp;</p>
  <%
  spThemeBlock1_close(intSkin)
end function

function showForumDown()
	strSql = "SELECT C_DOWNMSG"
	strSql = strSql & " FROM " & strTablePrefix & "CONFIG "
	set rs = my_Conn.Execute (strSql)
response.write "<br />"
spThemeBlock1_open(intSkin)
%><table class="tPlain" width="95%">
<tr>
<td class="tCellAlt1" align="center"><span class="fTitle"><% =rs("C_DOWNMSG")%></span></td>
</tr></table><%
spThemeBlock1_close(intSkin)%>
<p align="center">
<a href="default.asp"><%= txtBack %></a></p>
<%
end function

function chkIsBookmarked(s_app,s_cat,s_sub,s_item,s_id)
  intTemp = 0
  cSQL = "SELECT * FROM " & strTablePrefix & "BOOKMARKS WHERE "
  cSQL = cSQL & "APP_ID=" & s_app & " AND M_ID=" & s_id
  'if s_cat <> 0 then
    cSQL = cSQL & " AND CAT_ID=" & s_cat
  'end if
  'if s_sub <> 0 then
    cSQL = cSQL & " AND SUBCAT_ID=" & s_sub
  'end if
  'if s_item <> 0 then
    cSQL = cSQL & " AND ITEM_ID=" & s_item
  'end if
  set rsC = my_Conn.execute(cSQL)
  if not rsC.eof then
    intTemp = rsC("BOOKMARK_ID")
  end if
  set rsC = nothing
  chkIsBookmarked = intTemp
end function

function chkIsSubscribed(s_app,s_cat,s_sub,s_item,s_id)
  intTemp = 0
  cSQL = "SELECT SUBSCRIPTION_ID FROM "& strTablePrefix &"SUBSCRIPTIONS"
  cSQL = cSQL & " WHERE APP_ID=" & s_app & " AND M_ID=" & s_id
  'if s_cat <> 0 then
    cSQL = cSQL & " AND CAT_ID=" & s_cat
  'end if
  'if s_sub <> 0 then
    cSQL = cSQL & " AND SUBCAT_ID=" & s_sub
  'end if
  'if s_item <> 0 then
    cSQL = cSQL & " AND ITEM_ID=" & s_item
  'end if
  set rsC = my_Conn.execute(cSQL)
  if not rsC.eof then
    intTemp = rsC("SUBSCRIPTION_ID")
  end if
  set rsC = nothing
  chkIsSubscribed = intTemp
end function

function sendSubscriptionEmails(s_app,s_cat,s_sub,s_item,s_subject,s_msg)
if strEmail = 1 and intSubscriptions = 1 then
  s_app = clng(s_app)
  s_cat = clng(s_cat)
  s_sub = clng(s_sub)
  s_item = clng(s_item)

  sSQL = "SELECT " & strTablePrefix & "SUBSCRIPTIONS.APP_ID, " & strTablePrefix & "SUBSCRIPTIONS.M_ID, " & strTablePrefix & "SUBSCRIPTIONS.CAT_ID, " & strTablePrefix & "SUBSCRIPTIONS.SUBCAT_ID, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_ID, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_URL, " & strTablePrefix & "SUBSCRIPTIONS.ITEM_TITLE, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_RECMAIL, " & strTablePrefix & "APPS.APP_NAME, " & strTablePrefix & "APPS.APP_INAME, " & strTablePrefix & "APPS.APP_ACTIVE"

  sSQL = sSQL & " FROM (" & strTablePrefix & "SUBSCRIPTIONS LEFT JOIN " & strMemberTablePrefix & "MEMBERS ON " & strTablePrefix & "SUBSCRIPTIONS.M_ID = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID) LEFT JOIN " & strTablePrefix & "APPS ON " & strTablePrefix & "SUBSCRIPTIONS.APP_ID = " & strTablePrefix & "APPS.APP_ID"

  sSQL = sSQL & " WHERE (((" & strTablePrefix & "SUBSCRIPTIONS.APP_ID)=" & s_app & ") AND ((" & strMemberTablePrefix & "MEMBERS.M_STATUS)=1) AND ((" & strMemberTablePrefix & "MEMBERS.M_RECMAIL)=0) AND ((" & strTablePrefix & "APPS.APP_ACTIVE)=1) AND ((" & strTablePrefix & "SUBSCRIPTIONS.M_ID)<>" & strUserMemberID & "));"
	'response.Write(sSQL & "<br />")
  set rsSubs = my_Conn.execute(sSQL)
  if not rsSubs.eof then ' we have subscriptions to send
    do until rsSubs.eof
	 'if strUserMemberID <> rsSubs("M_ID") then
	  if lcase(rsSubs("APP_INAME")) <> "forums" then
        if rsSubs("CAT_ID") = 0 and rsSubs("SUBCAT_ID") = 0 and rsSubs("ITEM_ID") = 0 then
	      'send module subscriptions
		  sendOutEmail rsSubs("M_EMAIL"),s_subject,s_msg,2,0
			'response.Write("<br />App email sent to: " & rsSubs("M_EMAIL"))
	    else 'check for category and subcat subscriptions
	      if s_cat = rsSubs("CAT_ID") or (s_sub = rsSubs("SUBCAT_ID") and s_sub > 0) then
		    sendOutEmail rsSubs("M_EMAIL"),s_subject,s_msg,2,0
			'response.Write("<br />email sent to: " & rsSubs("M_EMAIL"))
	      end if
	    end if
	  else 'forum subscriptions
        if s_cat = 0 and s_sub = 0 then 'topic subscription
    		'response.Write("topic SUBSCRIPTIONS<br />")
	      'send topic reply subscriptions
		  if s_item = rsSubs("ITEM_ID") then
		    sSql = "SELECT FORUM_ID FROM "& strTablePrefix &"TOPICS WHERE TOPIC_ID="& s_item
		    set rsF = my_Conn.execute(sSql)
		    if chkForumAccess(rsSubs("M_ID"),rsF("FORUM_ID")) then
		      sendOutEmail rsSubs("M_EMAIL"),s_subject,s_msg,2,0
			end if
		    set rsF = nothing
		  end if
	    else 'check for category, subcat and item subscriptions
    '		response.Write("FORUM SUBSCRIPTIONS FOUND<br />")
	      if s_sub = rsSubs("SUBCAT_ID") then 
		    if chkForumAccess(rsSubs("M_ID"),s_sub) then
			  'response.Write("forum email about to be sent<br />")
		       sendOutEmail rsSubs("M_EMAIL"),s_subject,s_msg,2,0
			  'response.Write("forum email sent<br />")
			end if
	      end if
	    end if
	  end if 'app INAME = forums check
	 'end if 'member_id match
	 rsSubs.movenext
	loop
  else
    'response.Write("NO SUBSCRIPTIONS FOUND<br />")
  end if
  set rsSubs = nothing
end if
end function

sub sendOutEmail(mEmail,mTitle,mMsg,mFooter,mTyp)
	strRecipients = mEmail
	str_Subj = mTitle
	str_Msg = ""
	select case mFooter
	  case 1
	  	str_Msg = mMsg & vbCrLf & vbCrLf & vbCrLf
	  	str_Msg = str_Msg & txtThankYou & "!" & vbCrLf & vbCrLf
	  	str_Msg = str_Msg & "-" & txtEmlGenBy & " " & strSiteTitle & "-"
	  case 2
	  	str_Msg = mMsg & vbCrLf & vbCrLf & vbCrLf
	  	str_Msg = str_Msg & txtThankYou & "!" & vbCrLf & vbCrLf
	  	str_Msg = str_Msg & "-" & txtEmlGenBy & " " & strSiteTitle & "-" & vbCrLf
	  	str_Msg = str_Msg & txtEmlNoRespond & "" & vbCrLf
	  case else
	  	str_Msg = mMsg
	end select
%>
<!--#include file="includes/inc_mail.asp" -->
<%
end sub

function chkIsOnline(mem_name,ctype)
  ' ctypes
  ' 0 = returns boolian true/false
  ' 1 = returns text "Online!" or "Offline"
  ' 2 = returns an online/offline image
 if isArray(arrCurOnline) then
  select case ctype
    case 0
  	  tmpReturn = false
  	  for onlCh = 0 to ubound(arrCurOnline)
		if lcase(arrCurOnline(onlCh,0)) = lcase(mem_name) then
	  	  tmpReturn = true
		end if
  	  next
    case 1
  	  tmpReturn = txtOffline
  	  for onlCh = 0 to ubound(arrCurOnline)
		if lcase(arrCurOnline(onlCh,0)) = lcase(mem_name) then
		  tmpReturn = txtOnline
		end if
  	  next
    case 2
  	  tmpReturn = icon(icnOffline,txtOffline,"","","vspace=""3""")
  	  for onlCh = 0 to ubound(arrCurOnline)
		if lcase(arrCurOnline(onlCh,0)) = lcase(mem_name) then
		  tmpReturn = icon(icnOnline,txtOnline,"","","vspace=""3""")
		end if
  	  next
  end select
 else
   tmpReturn = ""
 end if
  chkIsOnline = tmpReturn
end function

function buildOnlineUsersArray()
'Builds an array of the current members that are online
	'redim arrCurOnline()
	'set rs = Server.CreateObject("ADODB.Recordset")
	strSql ="SELECT " & strMemberTablePrefix & "ONLINE.UserID, " & strMemberTablePrefix & "ONLINE.UserIP, " & strMemberTablePrefix & "ONLINE.DateCreated, " & strMemberTablePrefix & "ONLINE.LastChecked, " & strMemberTablePrefix & "ONLINE.CheckedIn"
	strSql = strSql & " FROM " & strMemberTablePrefix & "ONLINE"
	strSql = strSql & " where not UserID='" & txtGuest & "'"
	strSql = strSql & " ORDER BY " & strMemberTablePrefix & "ONLINE.DateCreated, " & strMemberTablePrefix & "ONLINE.CheckedIn DESC"
	
	Set tmpOnline=Server.CreateObject("ADODB.Recordset")
	on error resume next
	tmpOnline.Open strSql,my_Conn,1,1
	if err.number <> 0  then
	   set tmpOnline = nothing
	   response.Redirect("site_setup.asp?lerr=portal_online")
	end if
	on error goto 0
	userCount = clng(tmpOnline.recordcount)
	if userCount > 0 then
	  userCount = userCount-1
	else
	  userCount = 0
	end if
	redim arrCurOnline(userCount,1)
	'set tmpOnline = my_Conn.execute(strSql)
	If tmpOnline.EOF then
		'No online members
		'populate array
		arrCurOnline(0,0) = "Guest"
		arrCurOnline(0,1) = "Guest"
	Else
	  i = 0
	  do until tmpOnline.EOF
		'populate array
		arrCurOnline(i,0) = tmpOnline("UserID")
		arrCurOnline(i,1) = tmpOnline("UserIP")
		i = i + 1
		tmpOnline.movenext
	  loop
	end if
	tmpOnline.close
	set tmpOnline = nothing
end function

function getPageSkin(lev)
  if lev <> 3 then	'they are logged in
   if not uploadPg = true then
	if request("thm") <> "" then 'they just selected a new theme from the themechanger
		strTheme = chkString(request("thm"),"sqlstring")
			my_Conn.execute("UPDATE " & strMemberTablePrefix & "MEMBERS set THEME_ID = '" & strTheme & "' where M_NAME = '" & strDBNTUserName & "'")
	else 'member didn't select a new theme. Check for members personal theme
		set rs1 = my_Conn.execute("select THEME_ID from " & strMemberTablePrefix & "MEMBERS where M_NAME = '" & strDBNTUserName & "'")
		if not rs1.eof then 'they have selected a theme other than the default theme
			strTheme = rs1("THEME_ID")
			if strTheme = "0" or trim(strTheme) = "" or isNull(strTheme) then
			strTheme = strDefTheme
			end if
		else
			strTheme = strDefTheme
		end if
	end if
   else 'Check for members personal theme
		set rs1 = my_Conn.execute("select THEME_ID from " & strMemberTablePrefix & "MEMBERS where M_NAME = '" & strDBNTUserName & "'")
		if not rs1.eof then 'they have selected a theme other than the default theme
			strTheme = rs1("THEME_ID")
			if strTheme = "0" or trim(strTheme) = "" or isNull(strTheme) then
			strTheme = strDefTheme
			else
			  sSq = "SELECT C_INTSUBSKIN FROM " & strTablePrefix & "COLORS WHERE C_STRFOLDER='" & strTheme & "'"
			  set rsA = my_Conn.execute(sSq)
			  if not rsA.eof then
			    intSubSkin = rsA("C_INTSUBSKIN")
			  else
			    'intSubSkin = 0
			  end if
			  set rsA = nothing
			end if
		else
			strTheme = strDefTheme
		end if
   end if
  else	'they are a guest or not logged in
	if request("thm") <> "" then
		strTheme = chkString(request("thm"),"sqlstring")
		Response.Cookies(strCookieURL & strUniqueID & "guest")("Theme") = strTheme
	end if
	if Request.Cookies(strCookieURL & strUniqueID & "guest")("Theme") <> "" then
		strTheme = chkString(Request.Cookies(strCookieURL & strUniqueID & "guest")("Theme"),"sqlstring")
		if strTheme <> strDefTheme then
			strSQL = "SELECT * from " & strTablePrefix & "COLORS WHERE C_STRFOLDER='" & strTheme & "'"
			Set objRS2x = my_Conn.Execute(strSQL)
			if objRS2x.EOF then
				' theme level changed or theme's been deleted since guest selected the theme
				strTheme = strDefTheme
				Response.Cookies(strCookieURL & strUniqueID & "guest")("Theme") = strTheme
			end if
			Set objRS2x = nothing
		end if
	end if
  end if

  if strTheme <> strDefTheme then
	strSQL = "SELECT * FROM " & strTablePrefix & "COLORS WHERE C_STRFOLDER='" & strTheme & "'"
	Set objRS2x = my_Conn.Execute(strSQL)
	if not objRS2x.EOF then
		strTitleImage = objRS2x("C_STRTITLEIMAGE")
		intSubSkin = objRS2x("C_INTSUBSKIN")
	end if
	set objRS2 = nothing
  end if
end function

sub shoBreadCrumb(arg1,arg2,arg3,arg4,arg5,arg6)
  	Response.Write("<div class=""breadcrumb"">")
	Response.Write("<span style=""float:right;text-align:right;"">" & strCurDate & "</span>")
  if arg1 <> "" then
  	Response.Write(icon(icnNav,"","","","align=""middle""") & "&nbsp;<a href=""default.asp"">" & txtHome & "</a>")
  	writeCrumb(arg1)
	if arg2 <> "" then
  	  writeCrumb(arg2)
	  if arg3 <> "" then
  	    writeCrumb(arg3)
		if arg4 <> "" then
		  writeCrumb(arg4)
		  if arg5 <> "" then
		  	writeCrumb(arg5)
			if arg6 <> "" then
		  	  writeCrumb(arg6)
			end if
		  end if
		end if
	  end if
	end if
  elseif arg6 <> "" then
    Response.Write("&nbsp;" & arg6)
  end if
  Response.Write("</div>")
end sub

sub writeCrumb(v)
  Response.Write("&nbsp;&gt;&gt;&nbsp;")
	if instr(v,"|") > 0 then
  	  Response.Write("<a href=""" & split(v,"|")(1) & """>" & split(v,"|")(0) & "</a>")
	else
  	  Response.Write(v)
	end if
end sub

function getMenu(app)
	sSql = "SELECT APP_INAME FROM "& strTablePrefix & "APPS WHERE APP_ID = " & app
	set rsA = my_Conn.execute(sSql)
	if not rsA.eof then
	  strNm = rsA("APP_INAME")
	  execute("menu_" & strNm)
	else
	  response.Write(txtMnuNotFound)
	end if
	set rsA = nothing
end function

Function closeAndGo(where)
  if isObject(my_Conn) then
	my_Conn.close
	set my_Conn = nothing
	closeObjects()
  end if
	if where = "stop" then
	  response.End()
	else
	  response.Redirect(where)
	end if
End Function

function executeThis(SQLs)
	on error resume next
	Err.Clear	
		my_Conn.Execute(SQLs)
		dbHits = dbHits + 1
	if Err.number <> 0 then
	  Response.Write("<br /><br /><center><b>" & txtDBerror & "!</b><br />" & vbNewLine)
	  ErrorCount = ErrorCount + 1
	  Response.Write("<span class=""fAlert"">" & err.number & " | " & err.description & "</span><br />" & vbNewLine)
	  Response.Write("<br />" & SQLs & "<br /><br /></center><hr />")
	    closeAndGo("stop")
	end if
	Err.Clear
	on error goto 0
end function

Function GetRecordSet(byval stSQL)
	' -- given a valid SQL statement, returns a recordset with values
	const adUseClient = 3
	const adOpenForwardOnly = 0
	const adLockBatchOptimistic = 4

	Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset")
	' Setting the cursor location to client side is important
	' to get a disconnected recordset.
	objRS.CursorLocation = adUseClient
	objRS.Open stSQL, my_Conn, adOpenForwardOnly, adLockBatchOptimistic

	' Disconnect the recordset.
	Set objRS.ActiveConnection = Nothing
	Set GetRecordSet = objRS
End Function

function raiseHackAttempt(hMsg)
	browserIP = request.ServerVariables("REMOTE_ADDR")
	qryString1 = server.HTMLEncode(request.QueryString)
	qryString = request.QueryString
	PATH_TRANSLATED = request.ServerVariables("PATH_TRANSLATED")
	REQUEST_METHOD = request.ServerVariables("REQUEST_METHOD")
	HTTP_USER_AGENT = request.ServerVariables("HTTP_USER_AGENT")
	HTTP_COOKIE = request.ServerVariables("HTTP_COOKIE")
	hText = "Possible hack attempt.<br /><br />Your IP: " & browserIP & "<br />"
	hText = hText & "is being tracked and<br />an email has been sent to the site administrators<br />"
	
	'construct email message
	emailSubject = "Possible hack attempt at " & strSiteTitle
	'emailText = "THIS IS ONLY A TEST - DO NOT REPLY TO THIS EMAIL" & vbCrLf & vbCrLf
	emailText = emailText & "Possible hack attempt at " & strHomeURL & vbCrLf & vbCrLf
	if hMsg <> "" then
	  emailText = emailText & hMsg & vbCrLf & vbCrLf
	end if
	emailText = emailText & "Emails sent to:  " & strWebMaster & vbCrLf & vbCrLf
	emailText = emailText & DateAdd("h", strTimeAdjust , Now()) & vbCrLf & vbCrLf
	emailText = emailText & browserIP  & vbCrLf & vbCrLf
	emailText = emailText & PATH_TRANSLATED  & vbCrLf & vbCrLf
	emailText = emailText & qryString  & vbCrLf & vbCrLf
	'emailText = emailText & REQUEST_METHOD  & vbCrLf
	emailText = emailText & HTTP_USER_AGENT  & vbCrLf & vbCrLf
	emailText = emailText & HTTP_COOKIE  & vbCrLf & vbCrLf
	
	'send email to admin
  	for ish = 0 to ubound(tempArr)
	  sSQL = "select M_EMAIL from portal_MEMBERS where M_NAME='" & tempArr(ish) & "'"
	  set rsAdm = my_Conn.execute(sSQL)
	  if not rsAdm.EOF then   '"skydog @ insightbb.com"
	    sendOutEmail rsAdm("M_EMAIL"),emailSubject,emailText,2,0
	  end if
	  set rsAdm = nothing
  	next
	
	'display info to users browser
	response.Write("<table width=""350"" align=""center""><tr>")
	response.Write("<td align=""center"">")
	spThemeBlock1_open(intSkin)
	response.Write("<br /><center><b>" & hText & "</b></center><br /><br />")
	spThemeBlock1_close(intSkin)
	response.Write("</td></tr></table>")
end function

function OncePerDayChecks()
  dim tDate, dayCheck
  tDate = strCurDateString
  if left(strChkDate,8) <> left(tDate,8) then
	doOncePerDay()
  end if
end function

'this performs the daily routines for the OncePerDayChecks() function
function doOncePerDay()
  'this is the call to check to see if any reminders need sent
  'CheckEvents()
  
  'this calls the function in modules/custom_functions.asp
  checkOncePerDay()
  
  'check for PM purge
  purgePM()
  
  'reset the strChkDate to the new date	
    sSql = "update " & strTablePrefix & "CONFIG set C_ONEADAYDATE = '" & strCurDateString & "'"
	executeThis(sSql)
	
 	Application.Lock
	Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
	Application.UnLock
end function

function purgePM()
  pSQL = "SELECT APP_ID, APP_ACTIVE, APP_iDATA1, APP_iDATA2 FROM " & strTablePrefix & "APPS WHERE APP_INAME = 'PM'"
  set rsP = my_Conn.execute(pSQL)
  if not rsP.EOF then
    if rsP("APP_ACTIVE") = 1 and rsP("APP_iDATA1") = 1 and rsP("APP_iDATA2") <> 0 then
	  'get site admins
	  tmpAdmin = ""
	  mSQL = "select MEMBER_ID from " & strTablePrefix & "MEMBERS where M_LEVEL = 3"
	  set rsM = my_Conn.execute(mSQL)
	  tmpAdmin = rsM("MEMBER_ID")
	  rsM.movenext
	  do until rsM.EOF
	    tmpAdmin = tmpAdmin & "," & rsM("MEMBER_ID")
	    rsM.movenext
	  loop
	  arrAdmin = split(tmpAdmin,",")
	  set rsM = nothing
	  for ox = 0 to ubound(arrAdmin)
	    sqlAdmin1 = sqlAdmin1 & " AND M_TO <> " & arrAdmin(ox)
	    sqlAdmin2 = sqlAdmin2 & " AND M_FROM <> " & arrAdmin(ox)
	  next
	    sqlAdmin = sqlAdmin1 & sqlAdmin2
		targDate = left(DateToStr(DateAdd("d",-rsP("APP_iDATA2"),now())),8) & "000000"
		strSqL = "DELETE FROM " & strTablePrefix & "PM "
		strSql = strSql & "WHERE M_SAVED = 0 and M_SENT <= '" & targDate & "'" & sqlAdmin
		'response.Write(strSql & "<br />")
		executeThis(strSql)
	end if
  end if
  set rsP = nothing
end function

function addCSSfile(p)
  if p <> "" then
	cust_css = cust_css & "<link rel=""stylesheet"" href=""" & p & """ type=""text/css"" />" & vbCRLF
  end if
end function

function getCSSfile()
  if cust_css <> "" then
    response.Write(cust_css)
  end if
end function

function addRSSXML(t,u)
  if t <> "" and u <> "" then
    cust_xml = cust_xml & "<link rel=""alternate"" type=""application/rss+xml"" title=""" & t & """ href=""" & u & """ />" & vbCRLF
  end if
end function

function getRSSXML()
  if cust_xml <> "" then
    response.Write(cust_xml)
  end if
end function

function addJSfile(p)
  if p <> "" then
    cust_js = cust_js & "<script type=""text/javascript"" src=""" & p & """></script>" & vbCRLF
  end if
end function

function getJSFiles()
  if cust_js <> "" then
    response.Write(cust_js)
  end if
end function

function addToMeta(mt,t,c)
  if c <> "" then
    cust_meta = cust_meta & "<meta " & mt & "=""" & t & """ content=""" & c & """ />" & vbCRLF
  end if
end function

function getMetaTags()
    customMetaTags()
	'addToMeta "name","AUTHOR","SkyPortal www.SkyPortal.net"
	addToMeta "name","GENERATOR","SkyPortal www.SkyPortal.net"
	':: START - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SKYPORTAL EULA LICENSE AGREEMENT
	addToMeta "name","COPYRIGHT","SkyPortal code is Copyright (C)2005-2007 Tom Nance (SkyDogg) All Rights Reserved"
	':: END - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SKYPORTAL EULA LICENSE AGREEMENT
    if cust_meta <> "" then
      response.Write(cust_meta)
    end if
end function

function getSkin(iSkn,iCol)
  tSkn = 1
  select case iCol
    case 1 ':: left page column
	  select case iSkn
	    case 1
	      tSkn = 2
		case 2
	      tSkn = 1
		case 3
	      tSkn = 2
	  end select
	case 2 ':: main page column
	case 3 ':: right page column
	  select case iSkn
	    case 1
	      tSkn = 1
		case 2
	      tSkn = 3
		case 3
	      tSkn = 3
	  end select
  end select
  getSkin = tSkn
end function

function shoColumnBlocks(sCol)
    if right(sCol,1) = "," then
      sCol = left(sCol,len(sCol)-1)
    end if
    if instr(sCol,",") > 0 then
	  dim ar_col
      ar_col = split(sCol,",")
	else
	  redim ar_col(0)
      ar_col(0) = sCol
	end if
	
    for fp = 0 to ubound(ar_col)
	  fTemp = split(ar_col(fp),":")
      if ubound(fTemp) = 2 then
	    fFunct = fTemp(1) & "(""" & fTemp(2) & """)"
	  else
	    fFunct = fTemp(1)
	  end if
  	  execute(fFunct)
    next
end function

function shoBlocks(arrCol)
  if isArray(arrCol) then
    for fp = 0 to ubound(arrCol)
	  fTemp = split(arrCol(fp),":")
      if ubound(fTemp) = 2 then
	    fFunct = fTemp(1) & "(""" & fTemp(2) & """)"
	  else
	    fFunct = fTemp(1)
	  end if
  	  execute(fFunct)
    next
  else
    if arrCol <> "" then
  	  execute(arrCol)
	end if
  end if
end function

function icon(img,t,s,oc,o)
  ti = "<img src=""" & img & """ border=""0"" alt=""" & t & """ title=""" & t & """"
  if len(s & "x") > 1 then
    ti = ti & " style=""" & s & """"
  end if
  if len(oc & "x") > 1 then
    ti = ti & " onclick=""" & oc & """"
  end if
  if len(o & "x") > 1 then
    ti = ti & " " & o & ""
  end if
  ti = ti & " />"
  icon = ti
end function

function displayLink(e)
  displayLink = replace(e,"<a","<a rel=""nofollow""",1,-1,1)
end function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

sub deleteFile(p)
  if bFso then
    set oFs = new clsSFSO
    oFs.DeleteFile(p)
    set oFs = nothing
  end if
end sub

sub moveFile(sFrm,sTo)
  if bFso then
    set oFs = new clsSFSO
    oFs.MoveFile sFrm,sTo
    set oFs = nothing
  end if
end sub

sub deleteFolder(p)
 if bFso then
    set oFs = new clsSFSO
    oFs.DeleteFolder(p)
    set oFs = nothing
 end if
end sub

sub moveFolder(sFrm,sTo)
  if bFso then
    set oFs = new clsSFSO
    oFs.MoveFolder sFrm,sTo
    set oFs = nothing
  end if
end sub

sub writeToLog(m,p,txt)
  if bFso then
    set oFs = new clsSFSO
	if m <> "" then
	  oFs.Module = m
	end if
	if p <> "" then
	  oFs.LogFile = p
	end if
	oFs.WriteLogFile(txt)
    'oFs.MoveFolder(sFrm,sTo)
    set oFs = nothing
  end if
end sub

function chkValidUserName(n)
  bValid = true
  sInvalid = strInvalidUsernameChars
  if len(sInvalid) > 0 then
    if right(sInvalid,1) <> "," then
	  sInvalid = sInvalid & ",x"
	end if
    arInvalid = split(sInvalid,",")
    for v = 0 to ubound(arInvalid)-1
      if Instr(n,arInvalid(v)) > 0 then
	    bValid = false
	    exit for
	  end if
    next
  else
    bValid = false
  end if
  
  if Instr(n,",") > 0 then
	bValid = false
  end if
  chkValidUserName = bValid
end function		'

':: this is a test for installed components
Function isInstalled(obj)
	installed = False
	on error resume next
	Err = 0
	Dim chkObj
	Set chkObj = Server.CreateObject(obj)
	If 0 = Err Then installed = True
	Set chkObj = Nothing
	isInstalled = installed
	Err = 0
	on error goto 0
End Function

function EmailField(fTestString) 
	TheAt = Instr(2, fTestString, "@")
	if TheAt = 0 then 
		EmailField = 0
	else
		TheDot = Instr(cint(TheAt) + 2, fTestString, ".")
		if TheDot = 0 then
			EmailField = 0
		else
			if cint(TheDot) + 1 > Len(fTestString) then
				EmailField = 0
			else
				EmailField = -1
			end if
		end if
	end if
end function

function displayEmail(e)
  displayEmail = replace(e,"@",icon(icnAt,"","","","align=""middle"""),1,-1,1)
end function

function validate_email(e)
  Dim oEmlVal, tValid
  tValid = true
  if isInstalled("Pivo.EmailValidator") then
	on error resume next
    Set  oEmlVal = Server.CreateObject("Pivo.EmailValidator")
	If Err <> 0 Then
	else
      oEmlVal.Logging = 0
      oEmlVal.FromEmailAddress = strSender
      'Validation level can be one of the following
      '1 - Syntax
      '2 - Domain or MX Record
      '3 - SMTP
      '4 - Mailbox
      If oEmlVal.Validate(e,3) = 3 Then
	    'Response.Write("<br><br>Valid email Domain.")
	    tValid = true
      else
	    'Response.Write("<br><br>Invalid email address.")
	    tValid = false
      End If
      oEmlVal.ClearLog
      Set oEmlVal = nothing
	end if
	on error goto 0
  end if
  validate_email = tValid
end function

function chkNewItem(p,bn,u,bu)
  ':: checks if item passes is new or updated and writes an icon to the browser
  bTF = false
  lastVisit = Session(strUniqueID & "last_here_date")
 if bn then
  if len(p) = 14 then
    tdtSince = getDateDiff(strCurDateString,p)
    if lastVisit <= p then
	  response.Write icon(icnNew1,"New since last visit","","","hspace=""4"" align=""middle""")
	  bTF = true
    elseif tdtSince < 7 then
	  response.Write icon(icnNew2,"New in last 7 days","","","hspace=""4"" align=""middle""")
	  bTF = true
      elseif tdtSince < 14 then
	  response.Write icon(icnNew3,"New in last 14 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    end if
  end if
 end if
 if bu then
  if len(u) = 14 then
    tdtSince = getDateDiff(strCurDateString,u)
    if lastVisit <= u then
	  response.Write icon(icnUpdate1,"Updated since last visit","","","hspace=""4"" align=""middle""")
	  bTF = true
	elseif tdtSince < 7 then
	  response.Write icon(icnUpdate2,"Updated in last 7 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    elseif tdtSince < 14 then
	  response.Write icon(icnUpdate3,"Updated in last 14 days","","","hspace=""4"" align=""middle""")
	  bTF = true
    end if
  end if
 end if
  chkNewItem = bTF
end function

sub chkItemAttention(ob)
	if ob("BADLINK") <> 0 then
	  if bFull then
	    Response.Write "<a href=""" & app_page & "?cmd=22&amp;sid=" & ob("SUBCAT_ID") & """>"
	    Response.Write icon(icnAttention,"Item need attention","","","align=""middle""")
	    Response.Write "</a>&nbsp;"
	  else
	    if strDBNTUserName = ob("UPLOADER") then
	      Response.Write "<a href=""" & app_page & "?cmd=11&amp;item=" & ob("DL_ID") & """>"
	      Response.Write icon(icnAttention,"Item need attention","","","align=""middle""")
	      Response.Write "</a>&nbsp;"
		end if
	  end if
	end if
end sub

sub resetCoreConfig()
  Application(strCookieURL & strUniqueID & "ConfigLoaded")= ""
end sub

sub chkSessionMsg()
  tMsg = readSession("sMsg")
  If tMsg <> "" Then
	Response.Write("<span class=""fTitle"">")
	Response.Write(tMsg)
	Response.Write("</span><hr />")
	deleteSession("sMsg")
  End If
end sub

function isValidIP(sIPs)
  bValIP = true
  uIP = request.ServerVariables("REMOTE_ADDR")
  if sIPs <> "" then
    if instr(sIPs,",") > 0 then
      arInv = split(sIPs,",")
    else
      dim arInv(0)
	  arInv(0) = sIPs
    end if
    for iv = 0 to ubound(arInv)
      if left(uIP,len(arInv(iv))) = arInv(iv) then
        bValIP = false
	  end if
    next
  end if
  isValidIP = bValIP
end function

sub chkValidReferrer()
  sVal = false
  sRefer = replace(Request.ServerVariables("HTTP_REFERER"),"http://","")
  sHomUrl = replace(strHomeUrl,"http://","")
  if left(sRefer,len(sHomUrl)) = sHomUrl then
    sVal = true
  end if
  if not sVal then
    closeAndGo("error.asp?type=referrer")
  end if
end sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::        END SP v1.0 additions                 :::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%>
<!-- #include file="includes/inc_login_functions.asp" -->
<!-- #include file="includes/inc_group_functions.asp" -->
<!-- #include file="includes/inc_DateTime_functions.asp" -->
<!-- #include file="includes/inc_emails.asp" -->
<!-- #include file="includes/inc_theme.asp" -->
<!-- #include file="includes/inc_site_menus.asp" -->
<!-- #include file="includes/default_menu.asp" -->
<!-- #include file="includes/fp_core.asp" -->
<!-- #include file="fp_custom.asp" -->
<!-- #include file="Modules/pvt_msg/pm_functions.asp" -->
<!-- #include file="Modules/custom_functions.asp" -->
<script type="text/javascript" language="JavaScript1.2" runat=server>
function edit_hrefs(s_html, type){
    s_str = new String(s_html);
	if (type == 1) {
     	s_str = s_str.replace(/\b(http\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	} 
	if (type == 2) {

		s_str = s_str.replace(/\b(https\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 3) {
		s_str = s_str.replace(/\b(file\:\/\/\/\w\:\\[\w+\/\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 4) {

		s_str = s_str.replace(/\b(www\.[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
 		  "<a href=\"http://$1\" target=\"_blank\">$1</a>");
	}
	if (type == 5) {
		s_str = s_str.replace(/\b([\w+\-\'\#\%\.\_\,\$\!\+\*]+@[\w+\.?\-\'\#\%\~\_\.\;\,\$\!\+\*]*)/gi,
 		  "<a href=\"mailto\:$1\">$1</a>");
	}
		  	  
    return s_str;
}
</script>