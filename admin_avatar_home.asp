<!--#include file="config.asp" --><%
'response.Buffer = false
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
pgType = "manager"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<%

AVpath = "files/members/" 'root path to your member avatar folder
avatarDir = "files/avatars/"

ia = "1"
MSGstr = ""
recurse = true
folder = avatarDir '$$ Default folder for avatars

		aa = "block"
		ab = "block"
		ac = "block"
		ad = "block"
		ae = "block" 
	
if cLng(Request.QueryString("cmd")) = 10 then
	if Request.QueryString("avatar") <> "" and IsNumeric(Request.QueryString("avatar")) = True then
		Err_Msg = ""
		avatar = cLng(Request.QueryString("avatar"))
		strSql = "SELECT A_URL FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_ID = " & avatar
		set rsA = my_Conn.execute(strSql)
		  aUrl = rsA(0)
		set rsA = nothing
		
		strSql = "DELETE FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_ID = " & avatar
		executeThis(strSql)
                		
        ' - Update Members who had this Avatar to noavatar.gif
        strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
        strSql = strSql & "SET M_AVATAR_URL = '" & avatarDir & "noavatar.gif'"
        strSql = strSql & "WHERE M_AVATAR_URL = '" & aUrl & "'"
   		executeThis(strSql)
		
		Session.Contents("avatarHome") = "<li><span class=""fTitle"">" & replace(txtAH01,"[%url%]",aUrl) & "</span></li>"
		closeAndGo("admin_avatar_home.asp?mode=avrev")
	end if
end if

if Request.Form("Method_Type") = "updateAvatar" then
  ':: av review
		Err_Msg = ""

	txURL = ChkString(Request.Form("AvatarURL"),"url")
	txName = ChkString(Request.Form("AvatarName"),"")

	if trim(txURL) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtAH02 & "</li>"
	end if

	if trim(txName) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtAH03 & "</li>"
	end if

	if Err_Msg = "" then
		' - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "AVATAR "
		strSql = strSql & " SET A_URL = '" & txURL & "'"
		strSql = strSql & ",    A_NAME = '" & txName & "'"
		strSql = strSql & ",    A_MEMBER_ID = " & cLng(Request.Form("AvatarMemberID"))
		strSql = strSql & " WHERE A_ID = " & cLng(Request.Form("A_ID"))

		executeThis(strSql)

		Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtAH04 & "</span></li>"
	else
		Err_Msg1 = "<li><span class=""fSubTitle"">" & txtThereIsProb & "</span></li>"
		Session.Contents("avatarHome") = Err_Msg1 & Err_Msg
	end if
		closeAndGo("admin_avatar_home.asp?mode=avrev")
end if

If Request.QueryString("mode") = "deletefile" then
  if bFso = true then
	fPath = Server.MapPath(AVpath & Request.QueryString("fpath"))
	on error resume next
	Err.Clear
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(fPath) = true Then
	  objFSO.DeleteFile fPath
	End If
	Set objFSO = nothing
	on error goto 0
	sSQL = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_AVATAR_URL = '" & strHomeUrl & "files/avatars/noavatar.gif' WHERE MEMBER_ID = " & cLng(Request.QueryString("id"))
	executeThis(sSQL)
    Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtAH05 & "</span></li>"
	closeAndGo("admin_avatar_home.asp?cmd=4")
  else
    Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtFSOnotEnabled & "</span></li>"
	closeAndGo("admin_avatar_home.asp?mode=avupld")
  end if
End If

If Request.QueryString("mode") = "deletefolder" then
  Err_Msg = ""
  if bFso = true then
	fPath = Server.MapPath(AVpath & Request.QueryString("fpath"))
	on error resume next
	Err.Clear
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		  If objFSO.FolderExists(fPath) = true Then
			objFSO.DeleteFolder fPath
		  End If
	if Err.Number=0 then
		Set objFSO = nothing
		sSQL = "UPDATE " & strMemberTablePrefix & "MEMBERS SET M_AVATAR_URL = '" & strHomeUrl & "files/avatars/noavatar.gif' WHERE MEMBER_ID = " & Request.QueryString("fpath")
		executeThis(sSQL)
		Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtAH06 & "</span></li>"
	else
		Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtFSOnotEnabled & "</span></li>"
	end if
	on error goto 0
  else
    Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtFSOnotEnabled & "</span></li>"
  end if
  closeAndGo("admin_avatar_home.asp?mode=avupld")
End If
	
if Request.Form("Method_Type") = "avatarSettings" then
		strSql = "UPDATE " & strTablePrefix & "AVATAR2 "
		strSql = strSql & " SET A_HSIZE   = " & clng(Request.Form("AvatarHSize"))
		strSql = strSql & ",    A_WSIZE   = " & clng(Request.Form("AvatarWSize"))
		strSql = strSql & ",    A_BORDER = " & clng(Request.Form("AvatarBorder"))

		executeThis(strSql)
		Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtAH07 & "</span></li>"
		closeAndGo("admin_avatar_home.asp?mode=avset")
end if
	
if Request.Form("Method_Type") = "addAvatar" then 
	Session.Contents("avatarHome") = ""
	txURL = ChkString(Request.Form("AvatarURL"),"url")
	txName = ChkString(Request.Form("AvatarName"),"")
	
	if trim(txURL) = "" then 
		Session.Contents("avatarHome") = "<li>" & txtAH08 & "</li>"
	end if

	if trim(txName) = "" then 
		Session.Contents("avatarHome") = Session.Contents("avatarHome") & "<li>" & txtAH09 & "</li>"
	end if

	if Session.Contents("avatarHome") = "" then
		strSql = "INSERT INTO " & strTablePrefix & "AVATAR ("
		strSql = strSql & "A_URL"
		strSql = strSql & ", A_NAME"
		strSql = strSql & ", A_MEMBER_ID"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txURL & "'"
		strSql = strSql & ", '" & txName & "'"
		strSql = strSql & ", " & ChkString(Request.Form("AvatarMemberID"),"")	
		strSql = strSql & ")"
		executeThis(strSql)

		Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtAH10 & "</span></li>"
	else 
		Err_Msg1 = "<li><span class=""fSubTitle"">" & txtThereIsProb & "</span></li>"
		Session.Contents("avatarHome") = Err_Msg1 & Session.Contents("avatarHome")
	end if
		closeAndGo("admin_avatar_home.asp?mode=added")
end if

if Request.Form("Method_Type") = "Sync_Database" then
  if bFso then
	Dim folder
	Dim strPath
	Dim objFSO
	Dim objFolder
	Dim objItem
	Dim FNarray()
	Dim DBarray()
	Dim DUarray()
	Err_Msg = ""
	if len(Request.Form("folder")) >= 1 then
		'folder = Request.Form("folder") + "/"
	end if
	'strPath = "/" + folder
	strPath = folder

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))
	i = 0
	For Each objItem In objFolder.Files
		if (lcase(right(objItem.Name,4)) = ".gif" or lcase(right(objItem.Name,4)) = ".jpg") and (len(objItem.Name) > 4) then
			txURL = ChkString(folder & objItem.Name,"url")
			txName = ChkString(left(objItem.Name,len(objItem.Name)-4),"")
			ReDim Preserve FNarray(i)
			FNarray(i) = txURL
			i = i + 1

			'$$ Test for file already in database
			OKtoAdd = False
			strSql = "select * from " & strTablePrefix & "AVATAR where A_URL = '" & txURL & "'"
			set rs = Server.CreateObject("ADODB.Recordset")
			rs.cachesize = 20
			rs.open  strSql, my_Conn, 3
				if rs.EOF then OKtoAdd = True
			rs.close
			set rs = nothing
			set rsav = nothing

			'$$ Adds files not already in Database to Database
			if OKtoAdd then
				strSql = "INSERT INTO " & strTablePrefix & "AVATAR"
				strSql = strSql & " (A_URL, A_NAME, A_MEMBER_ID) VALUES"
				strSql = strSql & " ('" & txURL & "', '" & txName & "', " & 0 & ");"
				executeThis(strSql)
			end if
        	end if
	Next
	Set objItem = Nothing
	Set objFolder = Nothing
	Set objFSO = Nothing


	'$$ Tests contents of database against the list of files in the default folder
	i = 0
	len_folder = len(folder)
	strSql = "SELECT " & strTablePrefix & "AVATAR.A_ID"
	strSql = strSql & ", " & strTablePrefix & "AVATAR.A_URL"
	strSql = strSql & " FROM " & strTablePrefix & "AVATAR;"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
		do until rs.EOF
			avatar = rs("A_URL")
			if lcase(left(avatar,len_folder)) = lcase(folder) then
				found = false

				For each strAvatar in FNarray
					if strAvatar = avatar then found = true
				next
				if not found then
					Redim preserve DBarray(i)
					Redim preserve DUarray(i)
	 				DBarray(i) = rs("A_ID")
 					DUarray(i) = rs("A_URL")
 					i = i + 1
				end if
			end if
			rs.MoveNext
		loop
	end if
	rs.close
	set rs = nothing
	set rsav = nothing

	'$$ Removes from database files referenced as being in the specified directory that are not in the specified directory
	For each strID in DBarray
		strSql = "DELETE FROM " & strTablePrefix & "AVATAR "
		strSql = strSql & " WHERE " & strTablePrefix & "AVATAR.A_ID = " & strID
               	executeThis(strSql)
	next

	For each strURL in DUarray
		' - Update Members who had this Avatar to noavatar.gif
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & "SET M_AVATAR_URL = 'files/avatars/noavatar.gif'"
		strSql = strSql & "WHERE M_AVATAR_URL = '" & strURL & "'"
		executeThis(strSql)
	next
	
	Err_Msg = "<li><span class=""fSubTitle"">" & txtAH11 & "</span></li>"
	Session.Contents("avatarHome") = Err_Msg & "<li><span class=""fSubTitle"">" & txtAH12 & "</span></li>"
  else
    Session.Contents("avatarHome") = "<li><span class=""fSubTitle"">" & txtFSOnotEnabled & "</span></li>"
  end if
	closeAndGo("admin_avatar_home.asp?mode=avsync")
end if

 %>
<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
  <tr>
    <td class="leftPgCol">
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
  		avatarConfigMenu("1")
  		response.Write("<hr />")
  		menu_admin()
	spThemeBlock1_close(intSkin) %>
	</td>
    <td class="MainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtAvMgr & "|admin_avatar_home.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

  		if Session.Contents("avatarHome") <> "" or Err_Msg <> "" then %>
		<%
  		spThemeBlock1_open(intSkin) %>
		<table align=center border=0>
	  	  <tr>
	    	<td>
			<ul>
			<%= Err_Msg %>
			<% response.Write(Session.Contents("avatarHome"))
			   Session.Contents("avatarHome") = ""
			 %>
			</ul>
	    	</td>
	  	  </tr>
		</table>
		<%
  		spThemeBlock1_close(intSkin)
  		end if
		
	spThemeBlock1_open(intSkin) %>
   <table border="0" cellspacing="1" cellpadding="1">
	<tr>
	  <td class="tCellAlt1">
	  <%  
	  select case request.QueryString("mode")
	    case "avset"
	      avatarSettings()
	    case "added"
	  	  avatarAdd()
		case "avsync"
	  	  avatarSync()
		case "avrev"
	  	  avatarReview()
		case "avupld"
	  	  avatarUploaded()
	    case else
		  avatarSettings()
	  end select
	  %>	  
	  </td>
	</tr>
      </table>
	<% spThemeBlock1_close(intSkin) %>
    </td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp?target=admin_avatar_home.asp" %><% end if

sub avatarSettings() 
	' - Get Avatar Settings from DB
	strSql = "SELECT " & strTablePrefix & "AVATAR2.A_HSIZE"
	strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_WSIZE"
	strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
	strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"

	set rs = my_Conn.Execute(strSql)
%>
	<form action="admin_avatar_home.asp" method="post" id="PostTopic" name="PostTopic">
	<table border="0" cellspacing="0" cellpadding="0" align=center>
	  <tr>
	    <td class="tCellAlt2">
	<input type="hidden" name="Method_Type" value="avatarSettings">
	      <table width="100%" border="0" cellspacing="1" cellpadding="1">
		<tr valign="center">
		  <td align="center" class="tTitle" colspan="2"><b><%= txtAH14 %></b></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH30 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0">
		  <%= txtYes %>: <input type="radio" name="AvatarBorder" value="1"<% if rs("A_BORDER") <> "0" then Response.Write(" checked") %>>
		  <%= txtNo %>: <input type="radio" name="AvatarBorder" value="0"<% if rs("A_BORDER") = "0" then Response.Write(" checked") %>>
		  </td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH31 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0" align="left">
		    <table border="0" cellspacing="1" cellpadding="1">
		      <tr valign="top">
			<td><span class="fSmall"><b><%= txtHeight %></b></span></td>
			<td></td>
			<td><span class="fSmall"><b><%= txtWidth %></b></span><br /></td>
		      </tr>
		      <tr>
			<td>
			  
			  <select name="AvatarHSize" size=1>
			  	<option value="32"<% if rs("A_HSIZE") = "32" then Response.Write(" selected") %>>32</option>
			  	<option value="48"<% if rs("A_HSIZE") = "48" then Response.Write(" selected") %>>48</option>
		  		<option value="64"<% if rs("A_HSIZE") = "64" then Response.Write(" selected") %>>64</option>
		  		<option value="100"<% if rs("A_HSIZE") = "100" then Response.Write(" selected") %>>100</option>
			  </select></td>
			<td><b>X</b></td>
			<td>
			  
			  <select name="AvatarWSize" size=1>
			  	<option value=32<% if rs("A_WSIZE") = "32" then Response.Write(" selected") %>>32</option>
			  	<option value=48<% if rs("A_WSIZE") = "48" then Response.Write(" selected") %>>48</option>
		  		<option value=64<% if rs("A_WSIZE") = "64" then Response.Write(" selected") %>>64</option>
		  		<option value=100<% if rs("A_WSIZE") = "100" then Response.Write(" selected") %>>100</option>
			  </select></td>
		      </tr>
		    </table>
		  </td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
		</tr>
	      </table>
	    </td>
	  </tr>
	</table>
	</form>
<%
end sub

sub avatarAdd() %>
	<form action="admin_avatar_home.asp" method="post" id="formEle" name="PostTopic">
	<input type="hidden" name="Method_Type" value="addAvatar">
	<table border="0" cellspacing="0" cellpadding="0" align=center>
	  <tr>
	    <td class="tCellAlt2">
	      <table border="0" cellspacing="1" cellpadding="1">
		<tr valign="center">
		  <td align="center" class="tTitle" colspan="2"><b><%= txtAH15 %></b></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAVUrl %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="255" name="AvatarURL" value="<% =Trim(ChkString(txURL,"display")) %>" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH22 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="50" name="AvatarName" value="<% =Trim(ChkString(txName,"display")) %>" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH26 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0">
  	  	  <select name="AvatarMemberID" size="1">
  	  	  <OPTION SELECTED VALUE="0"><%= txtNone %></OPTION>
 		  <% ' - Get Members from DB
		  strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
		  strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME"
		  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
 	  	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
 	  	  strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"

		  set rs = Server.CreateObject("ADODB.Recordset")
		  rs.cachesize = 20
		  rs.open  strSql, my_Conn, 3

		  if rs.EOF or rs.BOF then  '## No replies found in DB
		  else
			  rs.movefirst
			  rs.pagesize = strPageSize
			  maxpages = cint(rs.pagecount)
			  howmanyrecs = 0
			  rec = 1

  			  do until rs.EOF '** %>
			  	<OPTION VALUE="<% =rs("MEMBER_ID") %>"><% =rs("M_NAME") %></OPTION>
			  <% rs.MoveNext
	 		  rec = rec + 1
			  loop
		  end if
		  rs.close
		  set rs = nothing %>
		  </select></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<%= txtSubmit %>" id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
		</tr>
	      </table>
	    </td>
	  </tr>
	</table>
	</form>
<%
end sub

sub avatarSync() %>
	<form action="admin_avatar_home.asp" method="post" id="PostTopic" name="PostTopic">
	<input type="hidden" name="Method_Type" value="Sync_Database">
	<table border="0" cellspacing="0" cellpadding="0" align=center width="450">
	  <tr>
	    <td class="tCellAlt2">
	      <table border="0" cellspacing="1" cellpadding="4">
		<tr valign="center">
		  <td class="tTitle" align="center" colspan="2">
		      <span class="fTitle"><%= txtAH27 %></span>
          </td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" colspan="2">
                  <%= replace(txtAH29,"[%avatarDir%]",avatarDir) %><br /><br />
                  <b><%= txtNote %>:</b><br />
                  &nbsp;&nbsp;&nbsp;<%= txtAH28 %>
          </td>
		</tr>
		  <td class="tCellAlt0" align="center" colspan="2">
                  <input type="submit" value="<%= txtAH16 %>" id="submit1" name="submit1" class="button"></td>
	      </table>
	    </td>
	  </tr>
	</table>
	</form>
<%
end sub

sub avatarReview() %>
<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2">
    <table width="100%" align="center" border="0" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" class="tTitle" colspan="5"><b><%= txtAH17 %></b></td>
      </tr>
      <tr class="tAltSubTitle">
        <td align="center"><b><%= txtPreview %></b></td>
        <td align="center"><b><%= txtAVUrl %></b></td>
        <td align="center"><b><%= txtAH22 %></b></td>
        <td align="center"><b><%= txtOwner %></b></td>
        <td align="center"><b></b></td>
      </tr>
<% 
	' - Get Avatar Settings from DB
	strSql = "SELECT " & strTablePrefix & "AVATAR2.A_HSIZE"
	strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_WSIZE"
	strSql = strSql & ", " & strTablePrefix & "AVATAR2.A_BORDER"
	strSql = strSql & " FROM " & strTablePrefix & "AVATAR2"

	set rsav = my_Conn.Execute (strSql)

	' - Get Avatars from DB
	strSql = "SELECT " & strTablePrefix & "AVATAR.A_ID" 
	strSql = strSql & ", " & strTablePrefix & "AVATAR.A_URL"
	strSql = strSql & ", " & strTablePrefix & "AVATAR.A_NAME"
	strSql = strSql & ", " & strTablePrefix & "AVATAR.A_MEMBER_ID"
	strSql = strSql & " FROM " & strTablePrefix & "AVATAR "
	strSql = strSql & " ORDER BY " & strTablePrefix & "AVATAR.A_ID ASC;"

    strPageSize = 20
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open  strSql, my_Conn, 3

	if rs.EOF or rs.BOF then
%>
      <tr>
        <td class="tCellAlt0" colspan="5"><span class="fSubTitle"><b><%= txtAH23 %></b></span></td>
      </tr>
<%
	else
		rs.movefirst
		rs.pagesize = strPageSize
		maxpages = cint(rs.pagecount)
		intI = 0 
		howmanyrecs = 0
		rec = 1

		do until rs.EOF '**
			if intI = 0 then
				CColor = "tCellAlt1"
			else
				CColor = "tCellAlt2"
			end if
 %>
      <tr>
        <td class="<% =CColor %>" valign="center" align="center" nowrap="nowrap"><img src="<% =rs("A_URL") %>" height=<% =rsav("A_HSIZE") %> width=<% =rsav("A_WSIZE") %> border=<% =rsav("A_BORDER") %> hspace=0 title="<% =rs("A_NAME") %>" alt="<% =rs("A_NAME") %>" /></td>
        <td class="<% =CColor %>" valign="center" align="center">
        	<% =rs("A_URL") %></td>
        <td class="<% =CColor %>" valign="center" align="center">
        	<% =rs("A_NAME") %></td>
        <td class="<% =CColor %>" valign="center" align="center">
        	<% if rs("A_MEMBER_ID") <> "0" then response.write (getMemberName(rs("A_MEMBER_ID")))  else response.write " - " end if %></td>
        <td class="<% =CColor %>" valign="center" align="center" nowrap="nowrap"><a href="javascript:mwpHSs('edit<% =rs("A_ID") %>','1');"><%= icon(icnEdit,txtAH24,"","","") %></a>
        <a href="admin_avatar_home.asp?cmd=10&avatar=<% =rs("A_ID") %>"><%= icon(icnDelete,txtAH25,"","","") %></a></td>
      </tr>
	  <tbody id="edit<% =rs("A_ID") %>" style="display:none;">
	  <tr><td colspan="5" class="<% =CColor %>">
	<form action="admin_avatar_home.asp" method="post" id="form<% =rs("A_ID") %>" name="form<% =rs("A_ID") %>">
	<input type="hidden" name="Method_Type" value="updateAvatar">
	<input type="hidden" name="A_ID" value="<% =rs("A_ID") %>">
	<table border="0" cellspacing="0" cellpadding="0" align=center>
	  <tr>
	    <td class="tCellAlt2">
              <table border="0" cellspacing="1" cellpadding="1">
		<tr valign="center">
		  <td align="center" class="tAltSubTitle" colspan="2"><b><%= txtEditAvatar %></b></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAVUrl %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="255" name="AvatarURL" value="<% =rs("A_URL") %>" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH22 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="50" name="AvatarName" value="<% =rs("A_NAME") %>" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtAH26 %>:</b>&nbsp;</td>
		  <td class="tCellAlt0">
  	  	  <select name="AvatarMemberID" size="1">
		  <% if IsNull(rs("A_MEMBER_ID")) or rs("A_MEMBER_ID") = " " or rs("A_MEMBER_ID") = "" or rs("A_MEMBER_ID") = 0 then %>
  	  	  <option value="0" selected>None</option>
		  <% Else %>
  	  	  <option value="0" selected>None</option>
		  <% end if %>
 		  <% ' - Get Member from DB
		  strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"
		  strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_NAME"
		  strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
 	  	  strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID=" & rs("A_MEMBER_ID")
 	  	  'strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
		  
		  set rsA = my_Conn.execute(strSql)
		  if not rsA.eof then
		    Response.Write "<option value=""" & rsA("MEMBER_ID") & """ selected>" & rsA("M_NAME") & "</option>"
		  end if
		  set rsA = nothing %>
		  </select></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value="<%= txtAH21 %>" id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
		</tr>
	      </table>
	    </td>
	  </tr>
        </table>
        </form><br />
	  
	  </td></tr></tbody>
<%
		    rs.MoveNext
		    intI = intI + 1
		    if intI = 2 then
				intI = 0
			end if
		    rec = rec + 1
		loop
	end if
	rs.close
	set rs = nothing
	set rsav = nothing
 %>
    </table></td>
  </tr>
</table>
<%
end sub

sub avatarUploaded() %>
<table width="550" border="0" align=center cellpadding="0" cellspacing="0">
  <tr>
    <td class="tCellAlt2">
      <table width="100%" border="0" cellpadding="1" cellspacing="1">
	<tr valign="top">
	      <td height="20" align="center" valign="middle" class="tTitle">
            <P><%= txtAH19 %></P>
          </td>
	</tr>
	<tr>
	      <td align="center" valign="top" class="tCellAlt1"> 
     <% if bFso = true then %>
            <table width="100%" border="0" cellspacing="0" cellpadding="10">
              <tr>
                <td align="left" valign="top">
				<b><%= txtAH20 %></b><br />
                  <hr align="center" width="90%">
				  <%	Call ListFolderContents(Server.MapPath(AVpath))  %>
                  <hr align="center" width="90%">
                  <B><%= txtNote %>:</B><br /></td>
              </tr>
            </table>
   <%  else
			Response.Write("<br /><br /><b>" & txtFSOnotEnabled & "</b><br /><br /><br />")
	   end if %>
          </td>
  		</tr>
	  </table>
	</td>
  </tr>
</table>
<%
end sub

sub avatarConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_avatar_cfg"
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
	  alt = "Collapse"
	else
	  cls = "none"
	  icn = "max"
	  alt = "Expand"
	end if
	 'onclick="javascript:mwpHSs('block12<%= typ ','0');" %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:location.reload();"><span style="margin: 2px;"><img name="block12<%= typ %>Img" id="block12<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b><%= txtAH13 %></b></div>
	  <% if typ = 1 then %>
      <div class="menu" id="block12<%= typ %>" style="display: <%= cls %>;">
	    <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH14 %><br /></a>
	    <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH15 %><br /></a>
	    <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH16 %><br /></a>
	    <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH32 %><br /></a>
	    <% If bFso = true Then %>
	    <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH18 %></a>
	    <% End If %>
	  <% else %>
      <div class="menu" id="block12<%= typ %>" style="display: <%= cls %>;">
	  <a href="admin_avatar_home.asp?mode="><%= icn_bar %><%= txtAH14 %><br /></a>
	  <a href="admin_avatar_home.asp"><%= icn_bar %><%= txtAH15 %><br /></a>
	  <a href="admin_avatar_home.asp"><%= icn_bar %><%= txtAH16 %><br /></a>
	  <a href="admin_avatar_home.asp"><%= icn_bar %><%= txtAH32 %><br /></a>
	  <% If bFso = true Then %>
	  <a href="admin_avatar_home.asp"><%= icn_bar %><%= txtAH18 %></a>
	  <% End If %>
	  <% end if %>
		   </div>
  <%
  end if
end sub
 %>
