<!--#include file="config.asp" --><%
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
 uploadPg = true
pgType = "manager"

if Request.Querystring("mode") <> 1 and Request.Querystring("mode") <> 2 and Request.Querystring("mode") <> 6 then
  session.Contents("uploadType") = ""
end if 
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="includes/inc_clsUpload.asp" -->
<% 'errors.clear
   'err.clear %>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<% 
DIM intMode, intID, strName, strLink, intHits, strImage, intActive, strHover, intLocation, strUploadDate, btnMsg
intMode = 0
intID = 0
strName = ""
strLink = "http://"
intHits = 0
strImage = ""
intActive = 0
strHover = ""
intLocation = 1
strUploadDate = ""
btnMsg = ""
loc = 1

	'sSQL = "select * from " & strTablePrefix & "UPLOAD_CONFIG where ID = " & uploadType
	'set rsU = my_Conn.execute(sSQL)
	'remotePath = rsU("UP_FOLDER")
	'set rsU = nothing

' Check for valid querystring
if Request.QueryString("mode") <> "" then
	if IsNumeric(Request.QueryString("mode")) = True then
		intMode = clng(Request.QueryString("mode"))
	end if
end if 
if Request.QueryString("id") <> "" then
	if IsNumeric(Request.QueryString("id")) = True then
		intID = clng(Request.QueryString("id"))
	end if
end if 
if Request.QueryString("sort") <> "" then
	if IsNumeric(Request.QueryString("sort")) = True then
		sorter = clng(Request.QueryString("sort"))
	end if
end if 
if Request.QueryString("loc") <> "" then
	if IsNumeric(Request.QueryString("loc")) = True then
		loc = clng(Request.QueryString("loc"))
		if loc < 1 then
		  loc = 1
		end if
	end if
end if 

'if intMode <> "" then
	select case intMode
		case 1 'insert new banner
			'if len(filename) = 0 then
			if trim(filename) = "" then
			    sString = ""
			  if left(objUpload.Fields("banImage").Value,7) = "http://" or left(objUpload.Fields("banImage").Value,7) = "https://" then
				imgURL = trim(chkString(objUpload.Fields("banImage").Value,""))
			  else
				imgURL = strHomeUrl & remotePath & trim(chkString(objUpload.Fields("banImage").Value,"displayimage"))
			  end if
			else
			  if bFso = true then
			  imgURL = strHomeUrl & remotePath & filename
			  end if
			end if
			if left(objUpload.Fields("banLink").Value,7) = "http://" or left(objUpload.Fields("banLink").Value,8) = "https://" then
				imgLink = trim(chkString(objUpload.Fields("banLink").Value,"clean"))
			else
				imgLink = "http://" & trim(chkString(objUpload.Fields("banLink").Value,"clean"))
			end if
			
			loc = clng(objUpload.Fields("loc").Value)
						
			strSql = "INSERT INTO " & strTablePrefix & "BANNERS ("
			strSql = strSql & "B_NAME"
			strSql = strSql & ", B_LINKTO"
			strSql = strSql & ", B_HITS"
			strSql = strSql & ", B_IMAGE"
			strSql = strSql & ", B_ACTIVE"
			strSql = strSql & ", B_ACRONYM"
			strSql = strSql & ", B_LOCATION"
			strSql = strSql & ", B_ACTIVATED_DATE"
   			strSql = strSql & ") VALUES ("
			strSql = strSql & " '" & trim(replace(objUpload.Fields("banName").Value,"'","''")) & "'"
			strSql = strSql & ", '" & imgLink & "'"
			strSql = strSql & ", 0"
			strSql = strSql & ", '" & imgURL & "'"
			strSql = strSql & ", 1"
			strSql = strSql & ", '" & trim(replace(objUpload.Fields("banHover").Value,"'","''")) & "'"
			strSql = strSql & ", " & loc
			strSql = strSql & ", '" & strCurDateString & "')"
			executeThis(strSql)
			set objUpload=nothing
			
			'sString = filename
			if sString <> "" then
			  response.Write("<ul><li>" & sString & "</li></ul><br />")
			  closeAndGo("stop")
			else
			  closeAndGo("admin_banner_manager.asp?loc=" & loc)
			end if
			session.Contents("uploadType") = ""
		
		case 2 'edit banner
			sSQL = "SELECT * FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			set rs = my_Conn.execute(sSQL)
			strName = rs("B_NAME")
			strLink = rs("B_LINKTO")
			intHits = rs("B_HITS")
			strImage = trim(chkString(rs("B_IMAGE"),"displayimage"))
			intActive = rs("B_ACTIVE")
			strHover = rs("B_ACRONYM")
			loc = rs("B_LOCATION")
			strUploadDate = rs("B_ACTIVATED_DATE")
			strImpressions = rs("B_IMPRESSIONS")
			intID = rs("ID")
			btnMsg = " " & txtBMgr10 & " "
			deMode = 6
			set rs = nothing
		case 3 'delete banner
			if bFso = true then
			  sSQL = "SELECT B_IMAGE FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			  set rsImg = my_Conn.execute(sSQL)
			  imgURL = rsImg("B_IMAGE")
			  set rsImg = nothing
			  if inStr(imgURL, strHomeURL & remotePath) then
				imgPath = server.MapPath(replace(imgURL,strHomeURL,""))
				on error resume next
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				 if fso.fileexists(imgPath) then
				   fso.deletefile(imgPath)
				 end if
				set fso = nothing
				on error goto 0
			  end if
			end if
			
			sSQL = "DELETE FROM " & strTablePrefix & "BANNERS WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			executeThis(sSQL)
			closeAndGo("admin_banner_manager.asp?loc=" & loc)
		case 4 'de-activate banner
			sSQL = "UPDATE " & strTablePrefix & "BANNERS SET " & strTablePrefix & "BANNERS.B_ACTIVE=0 WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			executeThis(sSQL)
			closeAndGo("admin_banner_manager.asp?sort=2&loc=" & loc)
		case 5 'activate banner
			sSQL = "UPDATE " & strTablePrefix & "BANNERS SET " & strTablePrefix & "BANNERS.B_ACTIVE=1 WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			executeThis(sSQL)
			closeAndGo("admin_banner_manager.asp?sort=2&loc=" & loc)
		case 6 'update banner
			if len(filename) = 0 then
			'response.Write(request.form("banImage"))
			  if left(objUpload.Fields("banImage").Value,7) = "http://" or left(objUpload.Fields("banImage").Value,8) = "https://" then
				imgURL = trim(chkString(objUpload.Fields("banImage").Value,""))
			  else
				imgURL = strHomeUrl & remotePath & trim(chkString(objUpload.Fields("banImage").Value,"displayimage"))
			  end if
			else
			  imgURL = strHomeUrl & remotePath & filename
			  ximgURL = objUpload.Fields("xbanImage").Value
			   if inStr(ximgURL, strHomeURL & remotePath) then
				imgPath = server.MapPath(replace(ximgURL,strHomeURL,""))
				on error resume next
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				 if fso.fileexists(imgPath) then
				   fso.deletefile(imgPath)
				 end if
				set fso = nothing
				on error goto 0
			   end if
			end if
			if left(objUpload.Fields("banLink").Value,7) = "http://" or left(objUpload.Fields("banLink").Value,8) = "https://" then
				imgLink = trim(chkString(objUpload.Fields("banLink").Value,"clean"))
			else
				imgLink = "http://" & trim(chkString(objUpload.Fields("banLink").Value,"clean"))
			end if
			loc = clng(objUpload.Fields("loc").Value)
			sSQL = "UPDATE " & strTablePrefix & "BANNERS "
			sSQL = sSQL & "SET " & strTablePrefix & "BANNERS.B_NAME='" & trim(replace(objUpload.Fields("banName").Value,"'","''")) & "'"
			sSQL = sSQL & ", " & strTablePrefix & "BANNERS.B_LINKTO='" & imgLink & "'"
			sSQL = sSQL & ", " & strTablePrefix & "BANNERS.B_IMAGE='" & imgURL & "'"
			sSQL = sSQL & ", " & strTablePrefix & "BANNERS.B_ACRONYM='" & trim(replace(objUpload.Fields("banHover").Value,"'","''")) & "'"
			sSQL = sSQL & ", " & strTablePrefix & "BANNERS.B_LOCATION=" & loc
			sSQL = sSQL & " WHERE " & strTablePrefix & "BANNERS.ID=" & intID
			'response.Write(sSQL & "<br />")
			executeThis(sSQL)
			set objUpload=nothing
			session.Contents("uploadType") = ""
			closeAndGo("admin_banner_manager.asp?loc=" & loc)
		case 7 'add banner
			btnMsg = " " & txtBMgr11 & " "
			deMode = 1
	end select
'else

  
'end if
%>

<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol" valign="top">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtBMgr21
spThemeBlock1_open(intSkin)
  		bannerConfigMenu("1")
  		response.Write("<hr />")
  		menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtBanMgr & "|admin_banner_manager.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<% 
spThemeTitle = txtBanMgr
spThemeBlock1_open(intSkin)

select case intMode
  case 0
    displayAllBanners()
  case 2, 7
    showBannerForm()
  case else
    displayAllBanners()
end select

spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% 
else
  Response.Redirect "admin_login.asp?target=admin_banner_manager.asp"
end if

Sub writeFlashx(swfImg) %>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="468" height="60" id="Flash_Banner" align=""><param name=movie value="<%= swfImg %>?clickTAG=<%= strHomeUrl %>admin_banner_manager.asp?mode=2&id=<%= intID %>&txtStr=<%= server.urlencode(strName) %>"><param name=quality value=high><embed src="<%= swfImg %>?clickTAG=<%= strHomeUrl %>admin_banner_manager.asp?mode=2&id=<%= intID %>&txtStr=<%= server.urlencode(strName) %>" quality="high" name="Flash_Banner" height="60" width="468" pluginspage="http://www.macromedia.com/go/getflashplayer"></embed></object>
<% end sub

sub displayAllBanners()
	
	select case sorter
		case 1
		  orderby = ""
		case 2
		  orderby = " AND B_ACTIVE = 1"
		case 3
		  orderby = " AND B_ACTIVE = 0"
		case else
		  orderby = " AND B_ACTIVE = 1"
	end select
	sSQL = "SELECT * FROM " & strTablePrefix & "BANNERS where B_LOCATION = " & loc & orderby & " ORDER BY " & strTablePrefix & "BANNERS.ID DESC"
	set rs = my_Conn.execute(sSQL)
%>
        <p>
          <%= replace(txtBMgr22,"[%img%]",icon(icnDelete,txtDel,"","","")) %><br />
		  <%= txtBMgr23 %><br />
          <%= txtBMgr24 %><br />
          <%= txtBMgr25 %></p>
		  <p>
		  <form name="form23" id="form23" method="get" action="<%= Request.ServerVariables("URL") %>">
		  <!-- Type:&nbsp;&nbsp;
		  <select name="loc" onchange="submit()">
          	<option value="1"<% If loc = 0 or loc = 1 Then response.Write(" selected") else 'do nothing %>selected="selected"><%= txtALL %></option>
          	<option value="2"<% If loc = 2 Then response.Write(" selected") else 'do nothing %>><%= ucase(txtActive) %></option>
          </select> -->
		  <%= txtShow %>:&nbsp;&nbsp;
		  <select name="sort" onchange="submit()">
          	<option value="1"<% If sorter = "1" Then response.Write(" selected") else 'do nothing %>selected="selected"><%= txtALL %></option>
          	<option value="2"<% If sorter = "" or sorter = "2" Then response.Write(" selected") else 'do nothing %>><%= ucase(txtActive) %></option>
          	<option value="3"<% If sorter = "3" Then response.Write(" selected") else 'do nothing %>><%= ucase(txtIActive) %></option>
          </select>
              <input type="hidden" name="loc" value="<%= loc %>">
      	  </form>
		  </p>
<%  if not rs.eof then
  	 	do until rs.eof
			strName = rs("B_NAME")
			strLink = rs("B_LINKTO")
			intHits = rs("B_HITS")
			strImage = rs("B_IMAGE")
			intActive = rs("B_ACTIVE")
			strHover = rs("B_ACRONYM")
			intLocation = rs("B_LOCATION")
			strUploadDate = rs("B_ACTIVATED_DATE")
			strImpressions = rs("B_IMPRESSIONS")
			intID = rs("ID")		  
		  %>
        <table width="500" class="tCellAlt1" border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
          <tr>
            <td valign="top">
            <table class="tCellAlt2" width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td align="right" valign="top"><%= txtName %>:</td>
                <td align="left" valign="top"><b><%= strName %></b></td>
                <td align="center" nowrap><b><%= intHits %></b> <%= txtHits %>&nbsp;<a href="admin_banner_manager.asp?mode=3&id=<%= intID %>&loc=<%= loc %>"><%= icon(icnDelete,txtBMgr26,"","","") %></a></td>
              </tr>
              <tr> 
                <td align="right" valign="top"><%= txtLink %>:</td>
                <td align="left" valign="top"><a href="<%= strLink %>"><%= strLink %></a></td>
                <td width="17%" align="center" nowrap><b><%= strImpressions %></b> <%= txtBMgr27 %>
                  </td>
              </tr>
              <tr> 
                <td width="11%" align="right" valign="top"> 
                  <%= txtHover %>:</td>
                <td align="left" valign="top"><i><%= strHover %></i></td>
                <td width="17%" align="center" nowrap><div class="tBorder"><%= txtBMgr28 %>:<br />
                  <% If intActive = 1 Then %>
                  <a title="<%= txtBMgr29 %>" href="admin_banner_manager.asp?mode=4&id=<%= intID %>&sort=<%= sorter %>&loc=<%= loc %>"><b><%= txtActive %></b></a> 
                  <% Else %>
                  <a title="<%= txtBMgr30 %>" href="admin_banner_manager.asp?mode=5&id=<%= intID %>&sort=<%= sorter %>&loc=<%= loc %>"><b><%= txtNActive %></b></a> 
                  <% End If %>
				  </div></td>
              </tr>
              <tr align="center"> 
                <td colspan="3"><a href="admin_banner_manager.asp?mode=2&id=<%= intID %>&loc=<%= loc %>" target="_top"><% If right(strImage,4) = ".swf" Then writeFlashx(strImage) Else response.write("<img name=""bImage"" border=""0"" src=""" & strImage & """ title=""" & txtBMgr31 & """ alt=""" & txtBMgr31 & """>") end if %></a> 
                  <br />
                </td>
              </tr>
            </table>
            </td>
          </tr>
        </table><br />&nbsp;
<%   rs.movenext
	loop
  else
	response.Write("<br /><br /><p>" & txtBMgr32 & "</p><br /><br />&nbsp;")
  end if
end sub

sub showBannerForm()
  %>
<script type="text/javascript">
function checkfrm(){
 if (document.forms.banner.banName.value == "") {
 alert("<%= txtBMgr12 %>");
	document.forms.banner.banName.focus();
 return false;
 }
 if (!CheckName(document.forms.banner.banName.value)) {
 alert("<%= replace(txtBMgr13,"[%br%]","\n") %>: \\ / : *  \" < > |");
	document.forms.banner.banName.focus();
 return false;
 }
 if (document.forms.banner.banLink.value == "") {
 alert("<%= txtBMgr14 %>");
	document.forms.banner.banLink.focus();
 return false;
 }
 if (!CheckThis(document.forms.banner.banLink.value)) {
 alert("<%= replace(txtBMgr15,"[%br%]","\n") %>:  *  \" < > |");
	document.forms.banner.banLink.focus();
 return false;
 }
 if (document.forms.banner.banHover.value == "") {
 alert("<%= txtBMgr16 %>");
	document.forms.banner.banHover.focus();
 return false;
 }
 if (!chkInput(document.forms.banner.banHover.value,'/ \ \ < > |')) {
 alert("<%= replace(txtBMgr17,"[%br%]","\n") %>: \\ / * \" < > |");
	document.forms.banner.banHover.focus();
 return false;
 }
 if (document.forms.banner.banImage.value == "") {
<% If bFso Then %>
 	if (document.forms.banner.file1.value == "") {
 		alert("<%= txtBMgr18 %>");
		document.forms.banner.banImage.focus();
 		return false;
	}
<% Else %>
 		alert("<%= txtBMgr19 %>");
		document.forms.banner.banImage.focus();
<% End If %>
 }
 if (!CheckThis(document.forms.banner.banImage.value)) {
 alert("<%= replace(txtBMgr20,"[%br%]","\n") %>:  *  \" < > |");
	document.forms.banner.banImage.focus();
 return false;
 }
// return true;
 if (document.forms.banner.file1.value != ""){
 document.getElementById('wait').style.display = 'block';
 document.getElementById('file1').style.visibility = 'hidden';
 document.getElementById('button').style.visibility = 'hidden';
 }
 document.forms.banner.submit();
 }

function chkInput(strStr,params) {
var re = new RegExp("\.(" + params.replace(/,/gi,"|").replace(/\s/gi,"") + ")$","i");
    if(!re.test(strStr)) return false;
	else return true;
}
function CheckThis(str) {
	var re;
	re = /[*'"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
function CheckName(str) {
	var re;
	re = /[\\\/:*'?"<>|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
</script>
		 <br />
	  	<form name="banner" method="post" action="admin_banner_manager.asp?mode=<%= deMode %>&id=<%= intID %>" onSubmit="checkfrm();return false" enctype="multipart/form-data">
        <table width="500" class="tCellAlt2" border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
          <tr> 
            <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr align="center"> 
                  <td colspan="2">
				 <% locTxt = ""
				 if loc = 2 then 
				   locTxt = " " & txtAffiliate & ""
				 end if
				    if intMode = 2 then
					Response.Write("<br /><b>" & replace(txtBMgr40,"[%location%]",locTxt) & "</b><br /><br />")
					If right(strImage,4) = ".swf" Then writeFlashx(strImage) Else response.write("<img name=""bImage"" border=""1"" src=""" & strImage & """ title="""& strHover &""" alt="""& strHover &""">") end if
					'Response.Write("<img src="""& strImage &""" border=""1"" alt="""& strHover &"""><br /><br />")
					else
					Response.Write("<br /><b>" & replace(txtBMgr33,"[%location%]",locTxt) & "</b><br /><br />")
					If bFso = true Then
					Response.Write(txtBMgr34 & "<br /><br />")
					else
					Response.Write("Be sure the banner is uploaded to the '" & remotePath & "' folder first.<br /><br />")
					end if
					end if %>
				  </td>
                </tr>
                <tr> 
                  <td width="28%" align="right" valign="top"><%= txtBMgr41 %>: </td>
                  <td width="72%" valign="middle"> 
                    <input name="banName" type="text" class="textbox" id="banName" value="<%= strName %>" size="50">
                  </td>
                </tr>
                <tr> 
                  <td align="right" valign="top"><span class="fAlert">*</span> <%= txtBMgr42 %>: </td>
                  <td valign="middle"> 
                    <input name="banLink" type="text" class="textbox" id="banLink" value="<%= strLink %>" size="50">
                  </td>
                </tr>
                <tr> 
                  <td align="right" valign="top"><span class="fAlert">*</span> <%= txtBMgr43 %>: </td>
                  <td valign="middle"> 
                    <input name="banHover" type="text" class="textbox" id="banHover" value="<%= strHover %>" size="50">
                  </td>
                </tr>
			<% If bFso = true Then
				strSQL = "select ID from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'banner'"
				set rsUload = my_Conn.execute(strSQL)
				 banID = rsUload("ID")
				set rsUload = nothing
		  		session.Contents("uploadType") = banID
		  		session.Contents("loggedUser") = strdbntusername %>
                <tr> 
                  <td align="right" valign="top"><span class="fAlert">**</span> <%= txtBMgr44 %>: </td>
                  <td valign="middle"> 
                    <input name="banImage" type="text" class="textbox" id="banImage" value="<%= strImage %>" size="50">
					<input name="xbanImage" id="xbanImage" type="hidden" value="<%= strImage %>">
                  </td>
                </tr>
          <tr>
            <td align="right" class="tCellAlt1">
			  <span class="fAlert">**</span> <%= txtBMgr45 %>:&nbsp; </td>
            <td class="tCellAlt1">
              <input class="textbox" name="file1" id="file1" type="file" size="30">
            </td>
          </tr>
          <tr> 
            <td align="center" class="tCellAlt1" colspan="2">
			  <div id="wait" name="wait" style="display:none;"><center><span class="fAltSubTitle"><b><%= txtUpInProg %></b></span></center><br /></div></td>
          </tr>
		  <% Else %>
                <tr> 
                  <td align="right" valign="top"><span class="fAlert">*</span> <%= txtBMgr44 %>: </td>
                  <td valign="middle"> 
                    <input name="banImage" type="text" class="textbox" id="banImage" value="<%= strImage %>" size="50">
					<input class="textbox" name="file1" id="file1" type="hidden" value="">
                  </td>
                </tr>
		  <% End If %>
                <tr align="center"> 
                  <td colspan="2">
                    <input id="button" class="button" type="submit" name="Submit" value="<%= btnMsg %>">
                    <input name="banID" type="hidden" id="banID" value="<%= intID %>">
              		<input type="hidden" name="max" value="1">
              <input type="hidden" name="memID" value="<%= strUserMemberID %>">
              <input type="hidden" name="loc" value="<%= loc %>">
			  <br />Upload Type: <%= session.Contents("uploadType") %>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </form><br />
	<%
end sub

sub bannerConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_banners"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 1
    mnu.canMinMax = 0
    mnu.keepOpen = 1
    mnu.GetMenu()
  else
	if typ = 1 then
	  cls = "block"
	  icn = "min"
	  alt = txtCollapse
	else
	  cls = "none"
	  icn = "max"
	  alt = txtExpand
	end if
	 'onclick="javascript:mwpHSs('block12<%= typ ','0');" %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:location.reload();"><span style="margin: 2px;"><img name="blockFP<%= typ %>Img" id="blockFP<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b><%= txtBMgr35 %></b></div>
    <div class="menu" id="blockFP<%= typ %>" style="display: <%= cls %>;">
		<a href="admin_banner_manager.asp"><%= icn_bar %><%= txtBMgr36 %><br /></a>
		<a href="admin_banner_manager.asp?mode=7"><%= icn_bar %><%= txtBMgr37 %><br /></a>
		<a href="admin_banner_manager.asp?loc=2"><%= icn_bar %><%= txtBMgr38 %><br /></a>
		<a href="admin_banner_manager.asp?mode=7&loc=2"><%= icn_bar %><%= txtBMgr39 %><br /></a>
	</div>
  <%
  end if
end sub %>