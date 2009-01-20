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

':: AVATAR functions used in cp_main.asp
  
':::::::::::::::::::::::::::::: display avatar :::::::::::::::::::::::::::::::::::
sub showAvatar()
	response.Write("<center><b>" & txtAvatar & "</b><br />")
		' - Get Avatar Settings from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR2.A_WSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_HSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"
		set rsav = my_Conn.Execute (strSql) %>
        <img src="<%= rsMem("M_AVATAR_URL") %>" align="absmiddle" width=<%= rsav("A_WSIZE") %> height=<%= rsav("A_HSIZE") %> border=<%= rsav("A_BORDER") %> hspace="0" />
        <% 
        set rsav = nothing
	response.Write("</center><br />")
end sub

sub showAVedit()
	showChooseOurs()
	showUploadProgress()
	showChooseYours()
end sub

sub showAvatarError()
msgText = Session.Contents("AVmsgText")
Session.Contents("AVmsgText") = ""
spThemeTitle = txtError
 spThemeBlock1_open(intSkin)
%>
 <table class="tPlain" cellpadding="10"><tr><td align="center">
<p><b><%= msgText %></b></p></td></tr></table>
<%
spThemeBlock1_close(intSkin) %>
<%
end sub

sub showUploadProgress() %>
	<div class="wait" id="wait" name="wait">
		<TABLE width="90%" border="0" cellspacing="1" cellpadding="5" class="tCellAlt0">
			<tr><td align="center" class="fNorm">
				<h4><span class="fAlert"><b><%= txtUpInProg %></b></span></h4>
			</td></tr>
		</TABLE>
	</div>
<%
end sub

sub showChooseYours()
spThemeTitle = txtChoseUown
spThemeBlock1_open(intSkin)
%>
            <TABLE width="90%" border="0" cellspacing="1" cellpadding="3">
              <TR> 
                <TD class="tCellAlt1" align="center"> 
<FORM name="avlink" method="post" action="cp_main.asp?cmd=2&mode=AvatarEditIt">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
<TD align="center">
          <br />
          <span class="fNorm"><b><%= txtLnkToAvatar %></b></span><br /><br />
<input class="textbox" type="text" id="url2" name="url2" size="30" value="HTTP://"> 
          </td>
</tr>
<TR> 
          <TD align="center"><br /><INPUT class="button" type="submit" value="<%= txtSubmitAVLnk %>" name="uploadBtn2" id="uploadBtn2"></TD>
</TR>
</table>
</FORM>
                </TD>
              </TR>
<% 
   If bFso = true and strAllowUploads = 1 then
		  	strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT, "
		  	strSQL = strSQL & "UP_NORM_MAX_W, UP_NORM_MAX_H "
		  	strSQL = strSQL & "from " & strTablePrefix & "UPLOAD_CONFIG where ID = 2"
			'set rsUload = my_Conn.execute(strSQL)
			'uActive = rsUload("UP_ACTIVE")
			'uAllowed = rsUload("UP_ALLOWEDGROUPS")
			'uSize = rsUload("UP_SIZELIMIT")
			'uExt = rsUload("UP_ALLOWEDEXT")
			'uID = rsUload("ID")
			'maxW = rsUload("UP_NORM_MAX_W")
			'maxH = rsUload("UP_NORM_MAX_H")
			'set rsUload = nothing
   		'If uActive = 1 and hasAccess(uAllowed) Then
			'session.Contents("uploadType") = uID
			'session.Contents("loggedUser") = strdbntusername
		':: get avatar config %>
              <TR>
                <TD class="tCellAlt1" align="center">
	<form name="form1" action="avatar_upload.asp" method="post" enctype="multipart/form-data"><br>
        <table border="0" width="100%" cellspacing="0" cellpadding="2" align="center">
          <tr> 
            <td class="fNorm" align="left">
              <CENTER> 
                <br />
                <B><%= txtAVToUpload %></B><br />
				<!-- - <%= txtAVSizeLmt %><br />
				- <%= txtAVTypeLmt %><br /> -->
                <br />
                 
                <input class="textbox" type="file" id="file1" name="file1" size="30" onchange="preview(this,'<%= maxW %>','<%= maxH %>')"><br />
			  <img alt="Avatar preview" id="previewField" src="images/spacer.gif">
                <br />
                <br />
                <INPUT class="button" type="Button" id="uploadBtn3" name="uploadBtn3" value="<%= txtAVNewUpl %>" onclick="send()">
                <br />
              </center>
            </td>
          </tr>
        </table>
		</form>
                </TD>
              </TR>
<% 		'End If %>
<% End If %>
            </TABLE>
<%
spThemeBlock1_close(intSkin)
end sub

sub showChooseOurs()
Session.contents("AVloggedUser") = ""
Session.Contents("AVAvatarUrl") = ""
Session.Contents("AVfileName") = ""
Session.Contents("AVmsgText") = ""
%>
<script type="text/javascript">
if (document.getElementById){ 
document.write('<style type="text/css">\n')
document.write('.wait{display: none;text-align:center;}\n')
document.write('</style>\n')
}

function CheckNav(Netscape, Explorer) {
  if ((navigator.appVersion.substring(0,3) >= Netscape && navigator.appName == 'Netscape') ||
      (navigator.appVersion.substring(0,3) >= Explorer && navigator.appName.substring(0,9) == 'Microsoft'))
    return true;
  else return false;
}

function send() {
document.getElementById('wait').style.display = "block";
//document.getElementById('blob').style.visibility = 'hidden';
document.getElementById('file1').style.visibility = 'hidden';
document.getElementById('uploadBtn1').style.visibility = 'hidden';
document.getElementById('uploadBtn2').style.visibility = 'hidden';
document.getElementById('uploadBtn3').style.visibility = 'hidden';
document.form1.submit();
}
</script>
<% 
Session.Contents("AVloggedUser") = strDBNTUserName
spThemeTitle = txtAVSelOurs
spThemeBlock1_open(intSkin)
%><FORM name="formEle" id="formEle" method="post" action="cp_main.asp?cmd=2&mode=AvatarEditIt">
<br />
                    <table border="0" width="100%" cellspacing="0" cellpadding="0" height="64" class="tCellAlt1">
                      <tr> 
                  		<td width="35%" height="25" align=right class="fNorm" valign=top nowrap><b><%= txtAvatar %>:&nbsp;</b></td>
                        <td width="15%" valign="top" align="left" nowrap> 
                          <select name="url2" size="6" onchange ="if (CheckNav(3.0,4.0)) URL.src=form.url2.options[form.url2.options.selectedIndex].value;">
                            <OPTION <% if IsNull(rsMem("M_AVATAR_URL")) or rsMem("M_AVATAR_URL") = "" or rsMem("M_AVATAR_URL") = " " or rsMem("M_AVATAR_URL") = "files/avatars/noavatar.gif" then %>selected value="files/avatars/noavatar.gif"><%= txtNone %></OPTION>
                            <%else%>value="<% =rsMem("M_AVATAR_URL")%>" selected="selected"> <%= txtCurrent %></OPTION>
                            <option value="files/avatars/noavatar.gif"> <%= txtNone %></OPTION>
                            <% end if %>
                            <%		' - Get Avatar Settings from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR2.A_HSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_WSIZE"
		strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"

		set rsavx = my_Conn.Execute (strSql)

		' - Get Avatars from DB
		strSql = "SELECT " & strTablePrefix & "AVATAR.A_ID" 
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_URL"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_NAME"
		strSql = strSql & ", " & strTablePrefix & "AVATAR.A_MEMBER_ID"
		strSql = strSql & " FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_MEMBER_ID = 0 "
		strSql = strSql & " OR " & strTablePrefix & "AVATAR.A_MEMBER_ID = " & rsMem("MEMBER_ID")
		strSql = strSql & " ORDER BY " & strTablePrefix & "AVATAR.A_ID ASC;"

		set rsAvs = Server.CreateObject("ADODB.Recordset")
		rsAvs.cachesize = 20
		rsAvs.open  strSql, my_Conn, 3

		if not(rsAvs.EOF or rsAvs.BOF) then  '## Avatars found in DB
			rsAvs.movefirst
			rsAvs.pagesize = strPageSize
			maxpages = cint(rsAvs.pagecount)
			howmanyrecs = 0
			rec = 1

			do until rsAvs.EOF '**
%>
                            <OPTION <% if rsAvs("A_URL") = rsMem("M_AVATAR_URL") then response.write("selected") %> VALUE="<% =rsAvs("A_URL") %>">&nbsp; 
                            <% =rsAvs("A_NAME") %>
                            </OPTION>
<%
			        rsAvs.MoveNext
	 			rec = rec + 1
			loop
		end if
		rsAvs.close
		set rsAvs = nothing
%>
                          </select>
                          <br /><a href="pop_portal.asp?cmd=8" target="_blank">
                          <span class="fNorm"><b><%= txtViewAllAV %></span></b></a> <br /><br /> </td>
                        <td valign="top" align="left"><img name="URL" src="<% if IsNull(rsMem("M_AVATAR_URL")) or rsMem("M_AVATAR_URL") = "" or rsMem("M_AVATAR_URL") = " " then %>files/avatars/noavatar.gif<% else %><% =rsMem("M_AVATAR_URL")%><% end if %>" border="<%= rsavx("A_BORDER") %>" width="<%= rsavx("A_WSIZE") %>" height="<%= rsavx("A_HSIZE") %>"></td>
                      </tr>
                <tr>
                  <td colspan="2" align="right" valign="middle">
                    <INPUT class="button" type="submit" name="uploadBtn1" id="uploadBtn1" value="<%= txtSubmit %>"><br />&nbsp; 
                  </td>
                  <td align=center valign=middle nowrap>&nbsp;</td>
                </tr>
                    </table>
                     
                    <% set rsavx = nothing %></FORM>
<%
spThemeBlock1_close(intSkin) %>
<%
end sub

sub editAvatar()
  AVATAR_URL = ""
  if trim(Request.Form("url2")) <> "" and trim(lcase(Request.Form("url2"))) <> "http://" then
    AVATAR_URL = Trim(chkString(Request.Form("url2"),"sqlstring"))
  end if
  If AVATAR_URL = "" Then
    AVATAR_URL = Session.Contents("AVAvatarUrl")
  End If

  msgText = Session.Contents("AVmsgText")
  filename = Session.Contents("AVfileName")
  AVurl = Session.Contents("AVAvatarUrl")
  If AVATAR_URL <> "" and (lcase(right(AVATAR_URL,3)) = "gif" or lcase(right(AVATAR_URL,3)) = "jpg") Then
    sSql = "UPDATE PORTAL_MEMBERS set M_AVATAR_URL ='"&AVATAR_URL&"' WHERE MEMBER_ID = "&strUserMemberID
    executeThis(sSql)
    erAVMsg = ""
  else
    erAVMsg = txtNoAVSelect
  End If

  Session.Contents("AVloggedUser") = ""
  Session.Contents("AVAvatarUrl") = ""
  Session.Contents("AVfileName") = ""
  Session.Contents("AVmsgText") = ""
  %>
  <center><span align="center" Style="width:500px;">
  <p><%
  If erAVMsg = "" Then
	tmpResult = ""
    If AVurl <> "" Then
 		tmpResult = tmpResult & txtAvUplSucc & "<br />"
    End If
	tmpResult = tmpResult & txtNewAVadded
	Else
	tmpResult = erAVMsg
  End If %></p>
  </span></center>
<%
end sub

%>