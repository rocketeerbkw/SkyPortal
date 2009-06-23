<!-- #INCLUDE FILE="config.asp" --><%
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
%>
<!-- #INCLUDE file="lang/en/core_admin.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %> 
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/inc_admin_functions.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<%
pic_id = Request.form("picid")
status = Request.form("status")
update1 = strCurDateString
stMSG = ""

if status = "delete" then
	set rs = my_Conn.Execute ("DELETE FROM PIC WHERE PIC_ID = " & pic_id)
	stMSG = stMSG & "Picture " & pic_id & " deleted."

elseif status = "yes" then
	set rsapp = my_Conn.Execute ("UPDATE PIC set ACTIVE=1, POST_DATE='" & update1 & "' where PIC_ID=" & pic_id)
	stMSG = stMSG & "Picture " & pic_id & " added to the database.<br>"
	  
	  if intSubscriptions = 1 and strEmail = 1 then
		sSql = "SELECT APP_ID FROM "& strTablePrefix & "APPS WHERE APP_iNAME = 'pictures'"
		set rsAP = my_Conn.execute(sSql)
		if not rsAP.eof then
	  	  intAppID = rsAP("APP_ID")
	      sSql = "SELECT CATEGORY, PARENT_ID FROM PIC WHERE PIC_ID = " & pic_id
		  set rsA = my_Conn.execute(sSql)
		    parent = rsA("PARENT_ID")
		    cat = rsA("CATEGORY")
		  set rsA = nothing
	      'send subscriptions emails
	      eSubject = strSiteTitle & " - New Picture"
		  eMsg = "A new picture has been submitted at " & strSiteTitle & vbCrLf
		  eMsg = eMsg & "that you have a subscription for." & vbCrLf & vbCrLf
		  eMsg = eMsg & "You can view the picture by visiting " & strHomeUrl & vbCrLf
	      sendSubscriptionEmails intAppID,parent,cat,"0",eSubject,eMsg
		  'response.Write("<br>Email sent<br>" )
		end if
		set rsAP = nothing
	  end if

	if lcase(strEmail) = "1" then
	
		dim rsGetInfo
		set rsGetInfo = server.CreateObject("adodb.recordset")
		strSql = "Select * from pic where pic_id=" & pic_id
		rsGetInfo.Open strSql, my_Conn
		
		Poster=rsGetInfo("POSTER")
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID=" & getMemberID(POSTER)

		set rs = my_Conn.Execute (strSql)
		
		strRecipientsName = rsGetInfo("POSTER")
		strRecipients = rs("M_EMAIL")
		strFrom = strSender
		strFromName = strSiteTitle
		strsubject = "Your picture has been added!"
		strMessage = strForumURL & vbcrlf & vbcrlf
		strMessage = strMessage & "Thank you for visiting our site. We have approved and added the following picture into our picture collection:" & vbcrlf & vbcrlf
		strMessage = strMessage & "Picture title: " & vbtab & rsGetInfo("TITLE") & vbcrlf
		strMessage = strMessage & "Description: " & vbtab & rsGetInfo("DESCRIPTION") & vbcrlf
		strMessage = strMessage & "Keywords: " & vbtab & rsGetInfo("KEYWORD") & vbcrlf
		strMessage = strMessage & "URL: " & vbtab & rsGetInfo("URL") & vbcrlf
		strMessage = strMessage & "Feel free to contact us if this information is not accurate."
		rsGetInfo.Close
		set rsGetInfo = nothing
		sendOutEmail strRecipients,strsubject,strMessage,2,0
		stMSG = stMSG & "Approval Email has been sent."
	end if
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign=top class="leftPgCol">
    <!--include file="admin_pic_menu.asp" -->
	<% 
	intSkin = getSkin(intSubSkin,1)
	spThemeBlock1_open(intSkin)
	pictureConfigMenu("1")
	response.write("<hr />")
	menu_admin()
	spThemeBlock1_close(intSkin) %>
    </td>
    <td class="mainPgCol">
	<% 
	intSkin = getSkin(intSubSkin,2)
	'breadcrumb here
  	arg1 = "Admin Area|admin_home.asp"
  	arg2 = "Picture Admin|admin_pic_main.asp"
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""
  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

	  	   spThemeBlock1_open(intSkin)
		   if stMSG <> "" then
		     response.Write("<p>" & stMSG & "</p>")
		   end if %>
	<%
	strSQL = "SELECT count(PIC_ID) FROM PIC WHERE ACTIVE=0"
	Set RScount = my_Conn.Execute(strSQL)
	rcount = RScount(0)
	RScount.close
	set RScount = nothing ' added
	
	strSQL = "SELECT * FROM PIC WHERE ACTIVE=0 ORDER BY PIC_ID"
	set rs = my_Conn.Execute(strSQL)
	if rs.eof then
	%>
	<center>
    No new pictures to approve.
    </center>
    <%
	else
	%>
    <P><span class="fAlert"><%= rcount%> New Pictures to Approve:</span></P>
<%
do while not rs.eof
%>
<form action="admin_pic_main.asp" method=post> 
<table border="0"  width="90%" cellspacing="0" cellpadding="0" class="tCellAlt2">
  <tr>
    <td class="tTitle" width="100%">
      <b><%=getCategoryName(rs("PARENT_ID"))%> / <%= getSubCatName(rs("CATEGORY"))%></b>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" width="100%">
    	Title: <%=replace(rs("TITLE"), "''","'", 1, -1, 1)%> 
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" width="100%">
      Description: <%=replace(rs("DESCRIPTION"), "''","'", 1, -1, 1)%>
      <br><br>
    </td>
  </tr>
    <tr>
    <td class="tCellAlt1" width="100%">
      Keywords: <%=replace(rs("KEYWORD"), "''","'", 1, -1, 1)%>
      <br><br>
    </td>
  </tr>
    <tr>
    <td class="tCellAlt1" width="100%">
      Copyright: <%=replace(rs("COPYRIGHT"), "''","'", 1, -1, 1)%>
      <br><br>
    </td>
  </tr>
    <tr>
    <td class="tCellAlt1" width="100%">
      URL: <a href="<%=replace(rs("URL"), "''","'", 1, -1, 1)%>" target="_blank"><%=replace(rs("URL"), "''","'", 1, -1, 1)%></a>
      <br><br>
    </td>
  </tr>
   <tr>
    <td class="tCellAlt1" width="100%">
      Thumbnail URL: <a href="<%=replace(rs("TURL"), "''","'", 1, -1, 1)%>" target="_blank"><%=replace(rs("TURL"), "''","'", 1, -1, 1)%></a>
      <br><br>
    </td>
  </tr>
    <tr>
    <td class="tCellAlt1" width="100%">
      Poster: <a href="cp_main.asp?cmd=8&member=<%=getMemberID(rs("POSTER"))%>" target="_blank"><%=replace(rs("POSTER"), "''","'", 1, -1, 1)%></a>
      <br><br>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" width="100%">
      Owner: <%= rs("OWNER") %>
      <br><br>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" width="100%">
      Date Submitted: <%=ChkDate(rs("POST_DATE")) & ChkTime(rs("POST_DATE"))%>
      <br><br>
    </td>
  </tr>
  <tr>
    <td class="tCellAlt1" width="100%">
      <input type="radio"  name="status" value="yes" checked id="approve">Approve<input type="radio"  name="status" value="delete">Delete this picture
      <input type="submit" Value="submit" class="button">
      <input type="hidden" name="picid" value="<%=rs("pic_ID")%>"><a href="admin_pic_editpic.asp?id=<%=rs("pic_ID")%>">Edit this picture</a>
    </td>
  </tr>
</table>
</form><br>
<%

rs.movenext
loop
end if
spThemeBlock1_close(intSkin)
%>
		</td>
	</tr>
</table>
<%
rs.close
set rs = nothing
%>
<!-- #INCLUDE FILE="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End If %>

<%
function getCategoryName(cat_id)
	strSQL = "SELECT CAT_NAME FROM pic_CATEGORIES WHERE CAT_ID = " & cat_id
	dim rsTemp
	set rsTemp = server.CreateObject("adodb.recordset")
	rsTemp.open strSQL, my_Conn
	getCategoryName = rsTemp("CAT_NAME")
	rsTemp.Close
	set rsTemp = nothing
end function

function getSubCatName(subcat_id)
	strSQL = "SELECT SUBCAT_NAME FROM pic_SUBCATEGORIES WHERE SUBCAT_ID = " & subcat_id
	dim rsTemp
	set rsTemp = server.CreateObject("adodb.recordset")
	rsTemp.open strSQL, my_Conn
	getSubCatName = rsTemp("SUBCAT_NAME")
	rsTemp.Close
	set rsTemp = nothing
end function
%>