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

'::<><><><><><><><><><><><><><><><><><><><><><><><
':: 
'::		Start custom functions
'::

sub showInfoBlock(ob)
  memID = getMemberID(ob("POSTER"))
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
  %>
  <style type="text/css">
  .infoblock{
    float: left;
	margin-right: 15px;
	margin-bottom: 2px;
	font-family: Verdana, Arial, sans-serif;
    font-size: 10px;
	font-style: italic;
	/*font-weight: bold;*/
    padding: 8px;
    border: 1px dashed #05719F;
    background-color: #EFF8FF; 
  }
  .infoblock a:link, .infoblock a:hover, .infoblock a:visited{
	font-family: Verdana, Arial, sans-serif;
    font-size: 10px;
	font-style: normal;
	/*font-weight: bold;*/ 
  }
  </style>
  <%
	Response.Write("<span class=""infoblock"">")
    Response.Write("Date Added:&nbsp;"&chkDate2(ob("POST_DATE"))&"<br />")
  	if ob("UPDATED") <> "0" then
	  Response.Write("Updated:&nbsp;" & ob("UPDATED") & "<br />")
    End If
	Response.Write("Posted by:<br />")
    if strUserMemberID > 0 then
	  Response.Write("&nbsp;<a href=""cp_main.asp?cmd=8&amp;member=" & memID & """ title=""" & txtViewProf & """>")
	  Response.Write("<u><b>" & ob("POSTER") & "</b></u></a>")
	else
	  Response.Write("&nbsp;<b>" & ob("POSTER") & "</b>")
	end if
    if len(ob("POSTER_EMAIL") & "x") > 8 then
	  if strEmail = 1 and hasAccess(2) then
	    Response.Write("<br />&nbsp;<a href=""JavaScript:openWindow('pop_mail.asp?id=" & memID & "')"" title=""" & txtEmlMbr & """>" & displayEmail(ob("POSTER_EMAIL")) & "</a>")
	  else
		response.Write("<br />&nbsp;" & displayEmail(ob("POSTER_EMAIL")))
	  end if
    End If
	If len(Trim(stData1)) > 12 Then
	  Response.Write("<br />&nbsp;<a href=""" & stData1 & """ target=""_blank"" title=""Visit website""><b><u>")
	  Response.Write("Visit Website")
	  Response.Write("</u></b></a>")
	else
	  Response.Write(ob("TDATA1"))
	End If
    if len(Trim(stData2) & "x") > 3 then
	  Response.Write("<br /><br />Author/Source:<br />")
	  If len(Trim(stData4)) > 12 Then
	    Response.Write("&nbsp;<a href=""" & stData4 & """ target=""_blank"" title=""Visit website""><b><u>")
		Response.Write(stData2)
		Response.Write("</u></b></a>")
	  else
		Response.Write(stData2)
	  End If
      if len(Trim(stData3) & "x") > 3 then
	    Response.Write("<br />&nbsp;"&displayEmail(stData3))
      End If
	End If
  	  if len(Trim(stData5) & "x") > 1 then
		Response.Write txtDlLabel5 & ":&nbsp;"& stData5 &"<br/>"
  	  end if
  	  if len(Trim(stData6) & "x") > 1 then
		Response.Write txtDlLabel6 & ":&nbsp;"& stData6 &"<br/>"
  	  end if
  	  if len(Trim(stData7) & "x") > 1 then
		Response.Write txtDlLabel7 & ":&nbsp;"& stData7 &"<br/>"
  	  end if
  	  if len(Trim(stData8) & "x") > 1 then
		Response.Write txtDlLabel8 & ":&nbsp;"& stData8 &"<br/>"
  	  end if
  	  if len(Trim(stData9) & "x") > 1 then
		Response.Write txtDlLabel9 & ":&nbsp;"& stData9 &"<br/>"
  	  end if
  	  if len(Trim(stData10) & "x") > 1 then
		Response.Write txtDlLabel10 & ":&nbsp;"& stData10 &"<br/>"
  	  end if
	Response.Write("<br /><br />Viewed:&nbsp;")
	Response.Write(ob("HIT") & "&nbsp;times")
	if Comments > 0 and art_Comments then
	  Response.Write("<br />Comments:&nbsp;" & Comments)
    End If
    if ob("RATING") <> 0 and ob("VOTES") <> 0 and art_Rate then
	  Response.Write("<br />Votes:&nbsp;"& ob("VOTES"))
      Response.Write("<br />Rating:&nbsp;"& ob("RATING")/ob("VOTES"))
    End If
	Response.Write("</span>")
end sub

sub customFormElements(o)
  sfEmail = strUserEmail
  if isObject(o) then
    sfEmail = o("POSTER_EMAIL")
	sfKwrds = o("KEYWORD")
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
  
  Response.Write "<tr><td align=""right"" class=""fNorm"">"
  Response.Write "<span class=""fAlert"">*</span>" & txtArtUEml
  Response.Write ":</td>"
  Response.Write "<td><input type=""text"" name=""posteremail"" size=""40"" maxlength=""90"" value=""" & sfEmail & """ /></td></tr>"
  
  if txtArtLabel1 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">"&txtArtLabel1&":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA1"" size=""40"" maxlength=""240"" value=""" & sTDATA1 & """ /></td></tr>"
  end if
  Response.Write "<tr><td colspan=""2"" class=""fNorm"">&nbsp;</td></tr>"
  if txtArtLabel2 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel2 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA2"" size=""40"" maxlength=""240"" value=""" & sTDATA2 & """ /></td></tr>"
  end if
  if txtArtLabel3 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel3 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA3"" size=""40"" maxlength=""240"" value=""" & sTDATA3 & """ /></td></tr>"
  end if
  if txtArtLabel4 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel4 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA4"" size=""40"" maxlength=""240"" value=""" & sTDATA4 & """ /></td></tr>"
  end if
  if txtArtLabel5 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel5 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA5"" size=""40"" maxlength=""240"" value=""" & sTDATA5 & """ /></td></tr>"
  end if
  if txtArtLabel6 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel6 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA6"" size=""40"" maxlength=""240"" value=""" & sTDATA6 & """ /></td></tr>"
  end if
  if txtArtLabel7 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel7 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA7"" size=""40"" maxlength=""240"" value=""" & sTDATA7 & """ /></td></tr>"
  end if
  if txtArtLabel8 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel8 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA8"" size=""40"" maxlength=""240"" value=""" & sTDATA8 & """ /></td></tr>"
  end if
  if txtArtLabel9 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel9 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA9"" size=""40"" maxlength=""240"" value=""" & sTDATA9 & """ /></td></tr>"
  end if
  if txtArtLabel10 <> "" then
    Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtLabel10 & ":</td>"
    Response.Write "<td><input type=""text"" name=""TDATA10"" size=""40"" maxlength=""240"" value=""" & sTDATA10 & """ /></td></tr>"
  end if
  Response.Write "<tr><td colspan=""2"" class=""fNorm"">&nbsp;</td></tr>"
  Response.Write "<tr><td align=""right"" class=""fNorm"">" & txtArtKeyWds & ":</td>"
  Response.Write "<td><input type=""text"" name=""key"" size=""40"" maxlength=""240"" value=""" & sfKwrds & """ /></td></tr>"
end sub

sub displayArticle(ob)
  spThemeBlock4_open() '<hr /><div style="text-align:left;">
%>
<table border="0" width="96%" cellspacing="1" cellpadding="6" align="center">
  <tr>
    <td width="100%">
      <a href="<%= app_rpage %>title=<%=server.URLEncode(ob("Title"))%>&amp;item=<%= ob(item_fld) %>"><span class="fSubTitle"><%= ob("Title") %></span></a>
	   <% call chkNewItem(ob("POST_DATE"),art_chkNew,ob("UPDATED"),art_chkUpdated) %>
	  <br><span class="fSmall">Posted by: <%=ob("POSTER")%></span><br />
	  <span class="fSmall">
	  <%
	  If strUpdated <> "0" and strUpdated <> "" Then %>
      (Updated: <%= chkDate2(ob("UPDATED")) %>
	  <% Else %>
      (Added: <%= chkDate2(ob("POST_DATE")) %>
	  <% End If %>
	  &nbsp;&nbsp;&nbsp;Hits: <%=ob("HIT")%>
	  <%
	  if ob("VOTES") > 0 then
		intRating = Round(ob("RATING")/ob("VOTES"))
		Response.Write("&nbsp;&nbsp;&nbsp;Rating: " & intRating & "&nbsp;&nbsp;&nbsp;Votes: " & ob("VOTES"))
	  end if
	  %>
	  )</span><hr />
      <p><%=ob("SUMMARY")%>
	  <br />
	  <a href="<%= app_rpage %>title=<%=server.URLEncode(ob("Title"))%>&amp;item=<%= ob(item_fld) %>"><span class="fNorm"><b>read more...</b></span></a></p>
    </td>
  </tr>
</table>
<% '</div><hr />
  spThemeBlock4_close()
end sub

function app_Footer()

end function

':: <><><><><><><><><><><><><><><><><><><><><><><><><><><><><
%>