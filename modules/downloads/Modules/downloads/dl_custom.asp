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
function app_LeftColumn()
    getMenu(intAppID)
	intShow = 5
	dl_small("new")
end function

function app_MainColumn_top()
	'modFeatures()
  	'intShow = 3
  	'intDir = 1
  '	dl_small("featured")
end function

function app_MainColumn_bottom()
	intShow = 6
  	'intDir = 1
	dl_large("new")
end function

function app_RightColumn()
	intShow = 3
	dl_small("featured")
	intShow = 3
	dl_small("rated")
	intShow = 3
	dl_small("top")
end function

function app_Footer()

end function

'::<><><><><><><><><><><><><><><><><><><><><><><><
':: 
'::		Start custom functions
'::

sub showInfoBlock(ob)
  stData1 = ob("TDATA1")
  stData2 = ob("TDATA2")
  stData3 = ob("TDATA3")
  stData4 = ob("TDATA4")
  stData5 = ob("TDATA5")
  stData6 = ob("TDATA6")
  stData7 = ob("TDATA7")
  stData8 = ob("TDATA8")
  stData9 = ob("TDATA9")
  stData10 = ob("TDATA10")

  strPoster = ob("UPLOADER")
  strDLEMAIL = ob("EMAIL")
  strPostDate = ChkDate2(ob("POST_DATE"))
  strUpdated = ob("UPDATED")
  'dateSince = getDateDiff(strCurDateString,ob("POST_DATE"))
  intHit = ob("Hit")
  if strUpdated <> "0" then
	strUpdated = ChkDate2(ob("UPDATED"))
  end if
  %>
	  <span class="infoblock">
	  Submitted by: <%=strPoster%><br />
  <% if strUpdated <> "0" then %>
      Updated:&nbsp;<%= strUpdated %><br />
  <% Else %>
      Date Added:&nbsp;<%= strPostDate %><br />
  <% End If %>
	  Downloaded:&nbsp;<%=intHit%><br />
	  <% 					
	if ob("VOTES") > 0 then
		intRating = Round(ob("RATING")/ob("VOTES"))
		Response.Write "Votes:&nbsp;"& ob("VOTES") &"<br/>"
		Response.Write "Rating:&nbsp;"& intRating &"<br/>"
	end if
  	  if len(Trim(stData1) & "x") > 1 and len(Trim(txtDlLabel1) & "x") > 1 then
		'Response.Write txtDlLabel1 & ":&nbsp;"& stData1 &"<br/>"
  	  end if
  	  if len(Trim(stData2) & "x") > 1 and len(Trim(txtDlLabel2) & "x") > 1 then
		Response.Write txtDlLabel2 & ":&nbsp;"& stData2 &"<br/>"
  	  end if
  	  if len(Trim(stData3) & "x") > 1 and len(Trim(txtDlLabel3) & "x") > 1 then
		Response.Write txtDlLabel3 & ":&nbsp;"& stData3 &"<br/>"
  	  end if
  	  if len(Trim(stData4) & "x") > 1 and len(Trim(txtDlLabel4) & "x") > 1 then
		'Response.Write txtDlLabel4 & ":&nbsp;"& stData4 &"<br/>"
  	  end if
  	  if len(Trim(stData4) & "x") > 1 and len(Trim(txtDlLabel4) & "x") > 1 then
		Response.Write txtDlLabel4 & ":&nbsp;"
		if len(Trim(stData5) & "x") > 1 then
		  Response.Write "<a rel=""nofollow"" href=""" & stData5 & """ target=""_blank"">"
		  Response.Write stData4 & "</a>"
		else
		  Response.Write stData4
		end if
		Response.Write "<br/>"
  	  end if
  	  if len(Trim(stData6) & "x") > 1 and len(Trim(txtDlLabel6) & "x") > 1 then
		Response.Write txtDlLabel6 & ":&nbsp;"& stData6 &"<br/>"
  	  end if
  	  if len(Trim(stData7) & "x") > 1 and len(Trim(txtDlLabel7) & "x") > 1 then
		Response.Write txtDlLabel7 & ":&nbsp;"& stData7 &"<br/>"
  	  end if
  	  if len(Trim(stData8) & "x") > 1 and len(Trim(txtDlLabel8) & "x") > 1 then
		Response.Write txtDlLabel8 & ":&nbsp;"& stData8 &"<br/>"
  	  end if
  	  if len(Trim(stData9) & "x") > 1 and len(Trim(txtDlLabel9) & "x") > 1 then
		Response.Write txtDlLabel9 & ":&nbsp;"& stData9 &"<br/>"
  	  end if
  	  if len(Trim(stData10) & "x") > 1 and len(Trim(txtDlLabel10) & "x") > 1 then
		Response.Write txtDlLabel10 & ":&nbsp;"& stData10 &"<br/>"
  	  end if
	%>
	</span>
  <%
end sub

sub shoDLGridInfo(o) 
  strDLName = o("NAME")
  intDLID = o("DL_ID")

  stData1 = o("TDATA1")
  stData2 = o("TDATA2")
  stData3 = o("TDATA3")
  stData4 = o("TDATA4")
  stData5 = o("TDATA5")
  stData6 = o("TDATA6")
  stData7 = o("TDATA7")
  stData8 = o("TDATA8")
  stData9 = o("TDATA9")
  stData10 = o("TDATA10")
  
  if len(Trim(stData1) & "x") > 1 and len(Trim(txtDlLabel1) & "x") > 1 then
	Response.Write "<tr><td width=""35%"" class=""fNorm"">" & txtDlLabel1 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData1 & "</td></tr>"
  end if
  
  if len(Trim(stData2) & "x") > 1 and len(Trim(txtDlLabel2) & "x") > 1 then
	Response.Write "<tr><td width=""35%"" class=""fNorm"">" & txtDlLabel2 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData2 & "</td></tr>"
  end if
  
  if len(Trim(stData3) & "x") > 1 and len(Trim(txtDlLabel3) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel3 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData3 & "</td></tr>"
  end if
  
  if len(Trim(stData4) & "x") > 1 and len(Trim(txtDlLabel4) & "x") > 1 then
	'Response.Write "<tr><td class=""fNorm"">" & txtDlLabel4 & "</td>"
	'Response.Write "<td class=""fNorm"">" & stData4 & "</td></tr>"
  end if
  
  if len(Trim(stData4) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel4 & "</td>"
	if len(Trim(stData5) & "x") > 1 then
	  Response.Write "<td class=""fNorm""><a rel=""nofollow"" href=""" & stData5 & """ target=""_blank"">"
	  Response.Write stData4 & "</a></td></tr>"
	else
	  Response.Write "<td class=""fNorm"">" & stData4 & "</td></tr>"
	end if
  end if
  
  if len(Trim(stData5) & "x") > 1 and len(Trim(txtDlLabel5) & "x") > 1 then
	'Response.Write "<tr><td>" & txtDlLabel5 & "</td>"
	'Response.Write "<td>" & stData5 & "</td></tr>"
  end if
  
  if len(Trim(stData6) & "x") > 1 and len(Trim(txtDlLabel6) & "x") > 1 then
	Response.Write "<tr><td>" & txtDlLabel6 & "</td>"
	Response.Write "<td>" & stData6 & "</td></tr>"
  end if
  
  if len(Trim(stData7) & "x") > 1 and len(Trim(txtDlLabel7) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel7 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData7 & "</td></tr>"
  end if
  
  if len(Trim(stData8) & "x") > 1 and len(Trim(txtDlLabel8) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel8 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData8 & "</td></tr>"
  end if
  
  if len(Trim(stData9) & "x") > 1 and len(Trim(txtDlLabel9) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel9 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData9 & "</td></tr>"
  end if
  
  if len(Trim(stData10) & "x") > 1 and len(Trim(txtDlLabel10) & "x") > 1 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlLabel10 & "</td>"
	Response.Write "<td class=""fNorm"">" & stData10 & "</td></tr>"
  end if
  
  if len(Trim(o("FILESIZE"))) > 0 then
	Response.Write "<tr><td class=""fNorm"">" & txtDlFileSize & "</td>"
	Response.Write "<td class=""fNorm"">" & o("FILESIZE") & "</td></tr>"
  end if
  %>
<tr>
<td class="fNorm">Date added</td>
<td class="fNorm"><%= strPostDate %>&nbsp;</td>
</tr>
<% If o("UPDATED") <> "0" Then %>
<tr>
<td class="fNorm">Last Updated</td>
<td class="fNorm"><%= chkDate(o("UPDATED")) %>&nbsp;</td>
</tr>
<% End If %>
<tr>
<td class="fNorm">Downloaded</td>
<td class="fNorm"><%=intHit%>&nbsp;<% if intHit > 1 then response.Write("times") end if %></td>
</tr>
<!-- <tr>
<td>Rating</td>
<td><% 'GetRating(intDLID) %>&nbsp;</td>
</tr> -->
<%
end sub

sub customFormElements() %>
  <tr valign="middle">
    <td align="right"><span class="fAlert">*</span> <%= txtDlUEml %>: </td>
    <td><input value="<%= strUserEmail %>" type="text" id="mail" name="mail" size="40" maxlength="90" /></td>
  </tr>
   <tr>
    <td align="right">
      <%= txtDlUName %>: 
    </td>
    <%if strDBNTUserName = "" then%>
    <td><input name="uploader" type="text" value="" size="90"></td>
    <%else%>
    <td><% =strDBNTUserName %><input type="hidden" value="<%=strDBNTUserName %>" name="uploader" /></td>
    <%end if%>
  </tr>
  <tr> 
	<td align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td align="right"><%= txtDlKeyWds %>: </td>
    <td><input type="text" name="key" size="40" maxlength="240" /></td>
  </tr>
  <!-- <tr>
    <td align="right"><%= txtDlFileSize %>: </td>
    <td><input type="text" name="filesize" size="40" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right">
      <%= txtDlLicense %>: 
    </td>
    <td><input type="text" name="license" size="40" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right"><%= txtDlLang %>: </td>
    <td><input type="text" name="language" size="40" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right"><%= txtDlPlat %>: </td>
    <td><input type="text" name="platform" size="40" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right">
      <%= txtDlPublisher %>: 
    </td>
    <td><input type="text" name="publisher" size="40" maxlength="240" /></td>
  </tr>
    <tr>
    <td align="right">
      <%= txtDlPubUrl %>: 
    </td>
    <td><input type="text" name="publisherurl" size="40" maxlength="240" value="http://" /></td>
  </tr>
  <tr> 
	<td align="center" colspan="2">&nbsp;</td>
  </tr> -->
  <%
'txtDlLabel1 = "Distribution Level"
'txtDlLabel2 = "Process Owner"
'txtDlLabel3 = "Original File Platform"
'txtDlLabel4 = "Rev#"
'txtDlLabel5 = "Rev Date"
'txtDlLabel6 = "Page Count"
'txtDlLabel7 = "Document Type"
'txtDlLabel8 = "Number"
'txtDlLabel9 = "Label 9"
'txtDlLabel10 = "Label 10"
  
  if txtDlLabel1 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel1 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA1"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel2 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel2 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA2"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel3 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel3 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA3"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel4 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel4 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA4"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel5 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel5 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA5"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel6 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel6 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA6"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel7 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel7 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA7"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel8 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel8 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA8"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel9 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel9 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA9"" size=""40"" maxlength=""240"" /></td></tr>"
  end if
  if txtDlLabel10 <> "" then
    Response.Write "<tr><td align=""right"">" & txtDlLabel10 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA10"" size=""40"" maxlength=""240"" /></td></tr>"
  end if

end sub

sub customEditFormElements(o)
  
  if isObject(o) then
    sTDATA1 = o("TDATA1")
    sTDATA2 = o("TDATA2")
    sTDATA3 = o("TDATA3")
    sTDATA4 = o("TDATA4")
    sTDATA5 = o("TDATA5")
    sTDATA6 = o("TDATA6")
    sTDATA7 = o("TDATA7")
    sTDATA8 = o("TDATA8")
    sTDATA9 = o("TDATA9")
    sTDATA10 = o("TDATA10")
  end if
  
  if txtDlLabel1 <> "" then
    Response.Write "<tr><td align=""right""><b>"&txtDlLabel1&":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA1"" size=""40"" maxlength=""240"" value=""" & sTDATA1 & """ /></td></tr>"
  end if
  if txtDlLabel2 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel2 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA2"" size=""40"" maxlength=""240"" value=""" & sTDATA2 & """ /></td></tr>"
  end if
  if txtDlLabel3 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel3 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA3"" size=""40"" maxlength=""240"" value=""" & sTDATA3 & """ /></td></tr>"
  end if
  if txtDlLabel4 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel4 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA4"" size=""40"" maxlength=""240"" value=""" & sTDATA4 & """ /></td></tr>"
  end if
  if txtDlLabel5 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel5 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA5"" size=""40"" maxlength=""240"" value=""" & sTDATA5 & """ /></td></tr>"
  end if
  if txtDlLabel6 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel6 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA6"" size=""40"" maxlength=""240"" value=""" & sTDATA6 & """ /></td></tr>"
  end if
  if txtDlLabel7 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel7 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA7"" size=""40"" maxlength=""240"" value=""" & sTDATA7 & """ /></td></tr>"
  end if
  if txtDlLabel8 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel8 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA8"" size=""40"" maxlength=""240"" value=""" & sTDATA8 & """ /></td></tr>"
  end if
  if txtDlLabel9 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel9 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA9"" size=""40"" maxlength=""240"" value=""" & sTDATA9 & """ /></td></tr>"
  end if
  if txtDlLabel10 <> "" then
    Response.Write "<tr><td align=""right""><b>" & txtDlLabel10 & ":</b></td>"
    Response.Write "<td><input type=""text"" name=""TDATA10"" size=""40"" maxlength=""240"" value=""" & sTDATA10 & """ /></td></tr>"
  end if

end sub

sub displayDL(ob)
	strPoster = ob("UPLOADER")		
	strDescription = ob("DESCRIPTION")
	strDLName = ob("NAME")
	strPostDate = ChkDate2(ob("POST_DATE"))
	'dateSince = getDateDiff(strCurDateString,ob("POST_DATE"))
	intHit = ob("Hit")
	intDLID = ob("DL_ID")
	strUpdated = ob("UPDATED")
	if strUpdated <> "0" then
	  strUpdated = ChkDate2(ob("UPDATED"))
	end if
	
  isOwner = false
  bFull = false
  if bAppFull or bSCatFull or bCatFull then
    bFull = true
  end if
  if bFull or (strDBNTUserName = strPoster) then
    isOwner = true
  end if

  spThemeBlock4_open()
  if sMode = 1 then
	pClass = "fAlert"
  else
	pClass = "fNorm"
  end if
%>
<table border="0" width="100%" cellspacing="1" cellpadding="6" align="center">
  <tr>
    <td width="100%">
      <div class="tSubTitle">
	  <% 
	  if isOwner then
	   if iPgType = 22 then
	    Response.Write "<a href=""" & sDLpage & "?cmd=22&amp;mode=321&amp;item="& intDLID &""">"
	   else
	    Response.Write "<a href=""" & sDLpage & "?cmd=23&amp;item="& intDLID &""">"
	   end if
	    Response.Write icon(icnEdit,"Edit Item","display:inline;","","align=""right""")
	    Response.Write "</a>"
	    Response.Write "<a href=""" & sDLpage & "?cmd=24&amp;item="& intDLID &""">"
	    Response.Write icon(icnDelete,"Delete Item","display:inline;","","align=""right""")
	    Response.Write "</a>"
	  end if %>
	  <a href="dl.asp?title=<%= strDLName %>&amp;cmd=6&amp;cid=<%=intDLID%>"><span class="fSubTitle"></span><b><%=strDLName%></b></a>
	   <% 
		n = ob("POST_DATE")
		u = ob("UPDATED")
		if u <> "0" then
		  n = "00"
		end if
		call chkNewItem(n,dl_chkNew,u,dl_chkUpdated)
	   'call chkNewItem(ob("POST_DATE"),dl_chkNew,ob("UPDATED"),dl_chkUpdated) %></div>
	   <span class="fSmall">Submitted by: 
	   <span class="<%= pClass %>"><b><%=strPoster%></b>
	   </span></span><br />
	  <span class="fSmall">
	  <% 
	  If strUpdated <> "0" and strUpdated <> "" Then %>
      (Updated: <%= strUpdated %>
	  <% Else %>
      (Added: <%= strPostDate %>
	  <% End If %>
	  &nbsp;&nbsp;&nbsp;Downloaded:<%=intHit%>
	  <%
	  if ob("VOTES") > 0 then
		intRating = Round(ob("RATING")/ob("VOTES"))
		Response.Write("&nbsp;&nbsp;&nbsp;Rating: " & intRating & "&nbsp;&nbsp;&nbsp;Votes: " & ob("VOTES"))
	  end if
	  %>
	  )</span><p>
      <%=strDescription%><br />
        <a href="dl.asp?title=<%= strDLName %>&amp;cmd=6&amp;cid=<%=intDLID%>">
		<span class="fSmall"><b>more...</b></span></a></p>
    </td>
  </tr>
</table>
<%
  spThemeBlock4_close()
end sub

%>