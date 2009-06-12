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

icn_bar = "<img src=""images/icons/icon_bar.gif"" align=""absmiddle"" height=""15"" width=""15"" border=""0"">&nbsp;"

sub adminPndTasks()
	'dl_adminPndLink()
	'article_adminPndLink()
	'weblinks_adminPndLink()
	'pictures_adminPndLink()
	'classified_adminPndLink()
	
	sSql = "SELECT * FROM " & strTablePrefix & "MODS WHERE M_CODE='admTaskLnk'"
	set rsP = my_Conn.execute(sSql)
	if not rsP.eof then
	  do until rsP.eof
	    execute("Call " & rsP("M_VALUE"))
	    rsP.movenext
	  loop
	end if
	set rsP = nothing
	
end sub

sub menu_admin()
  
  if pgType = "SiteConfig" then
    mainConfigMenu("1")
  else
    mainConfigMenu("0")
  end if
  if pgType = "memberConfig" then
    memberConfigMenu("1")
  else
    memberConfigMenu("0")
  end if
  if pgType = "manager" then
    managerMenu("1")
  else
    managerMenu("0")
  end if

  response.Write("<hr />")
  
 if bFso then
    mnu.menuName = "m_admin"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = 0
    mnu.canMinMax = 1
    mnu.keepOpen = 1
    mnu.GetMenu()
 else

  if chkApp("forums","USERS") then
  	forumConfigMenu("0")
  end if
  if chkApp("events","USERS") then
  	eventConfigMenu("0")
  end if
  if chkApp("article","USERS") then
    articleConfigMenu("0")
  end if
  if chkApp("classifieds","USERS") then
    classifiedsConfigMenu("0")
  end if
  if chkApp("downloads","USERS") then
    downloadConfigMenu("0")
  end if
  if chkApp("links","USERS") then
    linksConfigMenu("0")
  end if
  if chkApp("pictures","USERS") then
    pictureConfigMenu("0")
  end if
 end if
  response.Write("<hr />&nbsp;")
end sub

sub mainConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_site_cfg"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = typ
    mnu.canMinMax = 1
    mnu.keepOpen = 0
    mnu.GetMenu()
  else
	if typ = 1 then
	  cls = "block"
	  icn = "min1"
	  alt = txtCollapse
	else
	  cls = "none"
	  icn = "max1"
	  alt = txtExpand
	end if %>	
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block1','2');"><span style="margin: 2px;"><img name="block1Img" id="block1Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" vspace="2" align="absmiddle" style="cursor:pointer;" alt="<%= alt %>"></span>
    <b><%= txtSiteConfig %></b></div>
      <div class="menu" id="block1" style="display: <%= cls %>; text-align:left;">
	<% if typ = 3 then   '  %>
		<a onclick="show('aa');hide('ab');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtAdminHome %><br /></a>
		<% if intIsSuperAdmin then %>
		   <a onclick="show('ab');hide('aa');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtGenSetting %><br /></a>
		<% end if %>
		<a onclick="show('ac');hide('aa');hide('ab');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtBWfilter %><br /></a>
		<a onclick="show('ad');hide('aa');hide('ab');hide('ac');hide('ae');hide('af');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtSvDtTm %><br /></a>
		<%if intIsSuperAdmin then%>
		  <%if strAuthType <> "db" then %>
        	<a onclick="show('aj');hide('aa');hide('ac');hide('ad');hide('ab');hide('af');hide('ag');hide('ah');hide('ai');hide('ae');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtNTfeatures %><br /></a>
		  <%End if %>
		<a onclick="show('ae');hide('aa');hide('ac');hide('ad');hide('ab');hide('af');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtEmlSrvr %><br /></a>
		<!-- <a onclick="show('af');hide('aa');hide('ac');hide('ad');hide('ae');hide('ab');hide('ag');hide('ah');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %>Check Installation<br /></a> -->
		<a href="admin_emaillist.asp"><%= icn_bar %><%= txtEmlMbrs %><br /></a>
		<a onclick="show('ah');hide('aa');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ab');hide('ai');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtSvrInfo %><br /></a>
		<a onclick="show('ai');hide('aa');hide('ac');hide('ad');hide('ae');hide('af');hide('ag');hide('ah');hide('ab');hide('aj');hide('zz');" href="javascript:;"><%= icn_bar %><%= txtSiteVars %><br /></a>
		<%end if
	   else %>
		<a href="admin_home.asp"><%= icn_bar %><%= txtAdminHome %><br /></a>
		<%
		if intIsSuperAdmin then%>
		<a href="admin_home.asp?cmd=1"><%= icn_bar %><%= txtGenSetting %><br /></a>
		<%end if%>
		<a href="admin_home.asp?cmd=2"><%= icn_bar %><%= txtBWfilter %><br /></a>
		<a href="admin_home.asp?cmd=3"><%= icn_bar %><%= txtSvDtTm %><br /></a>
		<%if intIsSuperAdmin then%>
		  <%if strAuthType <> "db" then %>
        	<a href="admin_home.asp?cmd=9"><%= icn_bar %><%= txtNTfeatures %><br /></a>
		  <%End if %>
		<a href="admin_home.asp?cmd=4"><%= icn_bar %><%= txtEmlSrvr %><br /></a>
		<!-- <a href="admin_site_setup.asp"><%= icn_bar %>Check Installation<br /></a> -->
		<a href="admin_emaillist.asp"><%= icn_bar %><%= txtEmlMbrs %><br /></a>
		<a href="admin_home.asp?cmd=7"><%= icn_bar %><%= txtSvrInfo %><br /></a>
		<a href="admin_home.asp?cmd=8"><%= icn_bar %><%= txtSiteVars %><br /></a>
		<%
	   end if
	 end if %>
		   </div><%
  end if
end sub

sub memberConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_mem_cfg"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = typ
    mnu.canMinMax = 1
    mnu.keepOpen = 0
    mnu.GetMenu()
  else
	if typ = 1 then
	  cls = "block"
	  icn = "min1"
	  alt = txtCollapse
	else
	  cls = "none"
	  icn = "max1"
	  alt = txtExpand
	end if %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block2','2');"><span style="margin: 2px;"><img name="block2Img" id="block2Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b><%= txtMembers %></b></div>
      <div class="menu" id="block2" style="display: <%= cls %>; text-align:left;">
		<a href="admin_config_members.asp"><%= icn_bar %><%= txtMemDet %><br /></a>
		<a href="admin_config_members.asp?cmd=1"><%= icn_bar %><%= txtMemRank %><br /></a>
		<%if intIsSuperAdmin then%>
		<a href="admin_accounts_pending.asp"><%= icn_bar %><%= txtMemPend %><br /></a>
		<a href="admin_config_members.asp?cmd=2"><%= icn_bar %><%= txtMemClean %><br /></a>
		<%end if%>
		   </div><%
  end if
end sub

sub managerMenu(typ)
  if bFso then
    mnu.menuName = "b_managers"
    mnu.template = 4
    mnu.thmBlk = 0
    mnu.title = ""
    mnu.shoExpanded = typ
    mnu.canMinMax = 1
    mnu.keepOpen = 0
    mnu.GetMenu()
  else
	if typ = 1 then
	  cls = "block"
	  icn = "min1"
	  alt = txtCollapse
	else
	  cls = "none"
	  icn = "max1"
	  alt = txtExpand
	end if %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block3','2');"><span style="margin: 2px;"><img name="block3Img" id="block3Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" vspace="2" align="absmiddle" style="cursor:pointer;" alt="<%= alt %>"></span>
    <b><%= txtManagers %></b></div>
      <div class="menu" id="block3" style="display: <%= cls %>; text-align:left;">
	  	<%if intIsSuperAdmin then%>
		<a href="admin_config_cp.asp"><%= icn_bar %><%= txtManLayout %><br /></a>
			  <a href="admin_config_modules.asp"><%= icn_bar %><%= txtModMgr %><br /></a>
		<%end if%>
			  <a href="admin_config_groups.asp"><%= icn_bar %><%= txtGrpMgr %><br /></a>
			  <a href="admin_menu.asp"><%= icn_bar %><%= txtMnuMgr %><br /></a>
			  <a href="admin_banner_manager.asp"><%= icn_bar %><%= txtBanMgr %><br /></a>
			  <a href="admin_skins_config.asp"><%= icn_bar %><%= txtThmMgr %><br /></a>
		<%if intIsSuperAdmin then%>
			  <a href="admin_config_uploads.asp"><%= icn_bar %><%= txtUplMgr %><br /></a>
			  <a href="admin_pm.asp"><%= icn_bar %><%= txtPMmgr %><br /></a>
			  <a href="admin_ipgate.asp"><%= icn_bar %><%= txtIPGmgr %><br /></a>
		  <% If strDBType = "access" Then %>
			  <a href="admin_db.asp"><%= icn_bar %><%= txtDBmgr %><br /></a>
		  <% End If %>
		<%end if%>
		<a href="admin_countries.asp"><%= icn_bar %><%= txtCtryFlg %><br /></a>
		<a href="admin_avatar_home.asp"><%= icn_bar %><%= txtAvMgr %><br /></a>
		<a href="admin_welcome.asp"><%= icn_bar %><%= txtWelcome %><br /></a>
		<a href="admin_announce.asp"><%= icn_bar %><%= txtAnnouncements %><br /></a>
		   </div>
  <%
  end if
end sub

sub pmConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_pm_cfg"
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
	  icn = "min1"
	  alt = txtCollapse
	else
	  cls = "none"
	  icn = "max1"
	  alt = txtExpand
	end if %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:mwpHSa('block11<%= typ %>','2');"><span style="margin: 2px;"><img name="block11<%= typ %>Img" id="block11<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b><%= txtPMmgr %></b></div>
      <div class="menu" id="block11<%= typ %>" style="display: <%= cls %>; text-align:left;">
	<a onclick="show('pbb');show('paa');hide('pcc');" href="javascript:;"><%= icn_bar %><%= txtPMconfig %><br /></a>
	<a onclick="show('pcc');hide('paa');hide('pbb');" href="javascript:;"><%= icn_bar %><%= txtPmNewUsrs %><br /></a>
		   </div>
  <%
  end if
end sub

sub fpConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_layout"
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
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:location.reload();"><span style="margin: 2px;"><img name="blockFP<%= typ %>Img" id="blockFP<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" title="<%= alt %>" alt="<%= alt %>"></span>
    <b><%= txtManLayout %></b></div>
    <div class="menu" id="blockFP<%= typ %>" style="display: <%= cls %>;">
	  <a href="admin_config_fp.asp?cmd=3"><%= icn_bar %><%= txtCFP36 %><br /></a>
	  <a href="admin_config_cp.asp"><%= icn_bar %><%= txtCFP38 %><br /></a>
	  <a href="admin_config_fp.asp?cmd=1"><%= icn_bar %><%= txtCFP33 %><br /></a>
	  <a href="admin_config_fp.asp?cmd=0"><%= icn_bar %><%= txtCFP34 %><br /></a>
	  <a href="admin_config_fp.asp?cmd=2"><%= icn_bar %><%= txtCFP35 %><br /></a>
<a href="admin_config_fp.asp?cmd=<%= iPgType %>&mode=5"><%= icn_bar %><%= txtCFP37 %><br /></a>
	</div>
  <%
  end if
end sub

sub getEmailComponents()
Dim arrComponent(10)
Dim arrValue(10)
Dim arrName(10)

' components
arrComponent(0) = "CDO.Message"
arrComponent(1) = "CDONTS.NewMail"
arrComponent(2) = "SMTPsvg.Mailer"
arrComponent(3) = "Persits.MailSender"
arrComponent(4) = "SMTPsvg.Mailer"
arrComponent(5) = "CDONTS.NewMail"
arrComponent(6) = "dkQmail.Qmail"
arrComponent(7) = "Geocel.Mailer"
arrComponent(8) = "iismail.iismail.1"
arrComponent(9) = "Jmail.smtpmail"
arrComponent(10) = "SoftArtisans.SMTPMail"

' component values
arrValue(0) = "cdosys"
arrValue(1) = "cdonts"
arrValue(2) = "aspmail"
arrValue(3) = "aspemail"
arrValue(4) = "aspqmail"
arrValue(5) = "chilicdonts"
arrValue(6) = "dkqmail"
arrValue(7) = "geocel"
arrValue(8) = "iismail"
arrValue(9) = "jmail"
arrValue(10) = "smtp"

' component names
arrName(0) = "CDOSYS (IIS 5/5.1/6)"
arrName(1) = "CDONTS (IIS 3/4/5)"
arrName(2) = "ASPMail"		'yes
arrName(3) = "ASPEMail"	'yes
arrName(4) = "ASPQMail"	'yes			'
arrName(5) = "Chili!Mail (Chili!Soft ASP)"	'
arrName(6) = "dkQMail"						'
arrName(7) = "GeoCel"						'
arrName(8) = "IISMail"					'
arrName(9) = "JMail"						'
arrName(10) = "SA-Smtp Mail"

Response.Write("<select name=""strMailMode"">") & vbcrlf
Dim ix
for ix=0 to UBound(arrComponent)
	if isInstalled(arrComponent(ix)) then
	  Response.Write("<option value=""" & arrValue(ix) & """")
		if lcase(strMailMode)=arrValue(ix) then
		  Response.Write(" selected")
		end if
	  Response.Write(">" & arrName(ix) & "</option>") & vbcrlf
	end if
next
Response.Write("</select>") & vbcrlf
end sub

sub getImageComponents()
Dim arrComponent(5)
Dim arrValue(5)
Dim arrName(5)

' components
arrComponent(0) = "Persits.Jpeg"
arrComponent(1) = "AspImage.Image"
if bAspNet then
arrComponent(2) = strXmlHttpComp
end if

' component values
arrValue(0) = "aspjpeg"
arrValue(1) = "aspimage"
if bAspNet then
arrValue(2) = "aspnet"
end if

' component names
arrName(0) = "AspJpeg"
arrName(1) = "AspImage"
if bAspNet then
arrName(2) = "AspNet"
end if

Response.Write("<select name=""imgComp"">") & vbcrlf
Response.Write("<option value=""none"">[" & txtNONE & "]</option>")
Dim ix
for ix=0 to UBound(arrComponent)
  if len(arrComponent(ix) & "x") > 1 then
	if isInstalled(arrComponent(ix)) then
	  Response.Write("<option value=""" & arrValue(ix) & """")
		if lcase(strImgComp)=arrValue(ix) then
		  Response.Write(" selected")
		end if
	  Response.Write(">" & arrName(ix) & "</option>") & vbcrlf
	end if
  end if
next
Response.Write("</select>") & vbcrlf
end sub		

function checkForDotNet(DotNetFile)
  Dim DotNetComp, ResizeComUrl, LastPath
	DotNetComp = ""
	ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
	LastPath = InStrRev(ResizeComUrl,"/")
	if LastPath > 0 then
		ResizeComUrl = left(ResizeComUrl,Lastpath)
	end if
	ResizeComUrl = ResizeComUrl & DotNetFile
	'Response.Write ResizeComUrl & "<br />"
	
	'Check for ASP.NET 1
	if chkDotNetComponent("Msxml2.ServerXMLHTTP.4.0", ResizeComUrl) = true then
		'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP.4.0<br />"
		DotNetComp = "DOTNET1"
	else
		if chkDotNetComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
			'Response.Write "FOUND: ASP.NET Msxml2.ServerXMLHTTP<br />"
			DotNetComp = "DOTNET2"
		else
			if chkDotNetComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
				'Response.Write "FOUND: ASP.NET Microsoft.XMLHTTP<br />"
				DotNetComp = "DOTNET3"
			else
				'Response.Write "NOT FOUND: ASP.NET Server Component<br />"
			end if
		end if
	end if
	on error goto 0  
	checkForDotNet = DotNetComp
end function

function chkDotNetComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
	'response.write("Checking "&DotNetObj&"<br />")
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
  	'response.write("Object "&DotNetObj&" created<br />")
    objHttp.open "GET", ResizeComUrl, false
		if err.number = 0 then
      objHttp.Send ""
			if (objHttp.status <> 200 ) then
				'Response.Write "An error has accured with ASP.NET component " & DotNetObj & "<br />"
				'Response.Write "Returned:<br />" & objHttp.responseText & "<br />"
				'Response.End
			end if
      if trim(objHttp.responseText) <> "" and trim(objHttp.responseText) = "DONE" then
        Detection = true
      end if
		end if
    Set objHttp = nothing
  End if
  on error goto 0
 	'response.write("Detection is "&Detection&"<br />")
  chkDotNetComponent = Detection
end function

sub ListFolderContents(path)
	on error resume next
	Err.Clear
     set fs = CreateObject("Scripting.FileSystemObject")
	if Err.Number=0 then
     set AVfolder = fs.GetFolder(path)
	 	if AVfolder.Files.Count = 1 then
			txtFile = "file"
		else
			txtFile = "files"
		end if
	   If ia = "1" then
	   Response.Write("<li><b>" & AVfolder.Name & "</b> - "   & AVfolder.Files.Count & " " & txtFile & ", ")
	   Else
     Response.Write("<li><a href=""admin_avatar_home.asp?mode=deletefolder&fpath=" & Server.HTMLEncode(AVfolder.name) & """>" & icon(icnDelete,txtDel,"","","") & "</a>&nbsp;&nbsp;<b>" & AVfolder.Name & "</b> - " & AVfolder.Files.Count & " " & txtFile & ", ")
	   End If
	 
     if AVfolder.SubFolders.Count > 0 then
     	if AVfolder.SubFolders.Count > 1 then
       		Response.Write(AVfolder.SubFolders.Count & " directories, ")
     	else	
       		Response.Write(AVfolder.SubFolders.Count & " directory, ")
		end if
     end if
     Response.Write(Round(AVfolder.Size / 1024) & " KB total.</li>" & vbCrLf)

     Response.Write("<ul>" & vbCrLf)
		ia = ia + 1
     for each item in AVfolder.SubFolders
		if recurse then
	     if item.name <> "tiny_mce" then
          ListFolderContents(item.Path)
	     end if
		else
	     if item.Files.Count = 1 then
			txtFile = "file"
		 else
			txtFile = "files"
		 end if
         Response.Write("<li><a href=""admin_avatar_home.asp?mode=deletefolder&fpath=" & Server.HTMLEncode(item.name) & """>" & icon(icnDelete,txtDel,"","","") & "</a>&nbsp;&nbsp;<b>" & item.Name & "</b> - " & item.Files.Count & " " & txtFile)
	 
         if item.SubFolders.Count > 0 then
     	   if item.SubFolders.Count > 1 then
       		Response.Write(", " & item.SubFolders.Count & " directories, ")
     	   else	
       		Response.Write(", " & item.SubFolders.Count & " directory, ")
		   end if
          end if
		end if
     next

     for each file in AVfolder.Files  '<a href=""javascript:openWindow4('" & url & "');"">
       url = MapURL(file.path)
	   Response.Write("<li><a href=""admin_avatar_home.asp?mode=deletefile&fpath=" & Server.urlEncode(AVfolder.Name & "\" & file.name) & "&id=" & AVfolder.name & """>" & icon(icnDelete,txtDel,"","","") & "</a>&nbsp;&nbsp;<a href=""" & url & """>" _
         & file.Name & "</a> - " _
         & file.Size & " bytes, " _
         & "modified " & file.DateLastModified & "." _
         & "</li>" & vbCrLf)
     next
     Response.Write("</ul>" & vbCrLf)
	else
		Response.Write("<br /><b>" & txtFSOnotEnabled & "</b><br />" & vbCrLf)
	end if
	on error goto 0
end sub

function MapURL(path)
     dim rootPath, url2
     rootPath = Server.MapPath("/")
     url2 = Right(path, Len(path) - Len(rootPath))
     MapURL = Replace(url2, "\", "/")
end function

	if Request.QueryString("cmd") <> "" and IsNumeric(Request.QueryString("cmd")) = True then
		shoMod = cLng(Request.QueryString("cmd"))
	else
		shoMod = 0
	end if
		aa = "none"
		ab = "none"
		ac = "none"
		ad = "none"
		ae = "none"
		af = "none"
		ag = "none"
		ah = "none"
		ai = "none"
		aj = "none"
	select case shoMod
	  case 1
		ab = "block"
	  case 2
		ac = "block"
	  case 3
		ad = "block"
	  case 4
		ae = "block"
	  case 5
		af = "block"
	  case 6
		ag = "block"
	  case 7
		ah = "block"
	  case 8
		ai = "block"
	  case 9
		aj = "block"
	  case else
		aa = "block"
	end select
 %>