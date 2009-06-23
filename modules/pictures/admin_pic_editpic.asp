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
CurPageType = "pictures"
%>
<!-- #INCLUDE file="lang/en/core_admin.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %> 
<!-- #INCLUDE FILE="inc_functions.asp" -->
<!-- #INCLUDE file="includes/inc_admin_functions.asp" -->
<!-- #INCLUDE FILE="inc_top.asp" -->
<%
If request("mode") = 8 Then 
  webid = cLng(Request.Form("webid"))
  title = replace(ChkString(Request.Form("title"),"SQLString"), "'","''",1,-1,1)
  pdescription = ChkString(Request.Form("description"),"message")
  keyword = replace(ChkString(Request.Form("keyword"),"SQLString"), "'","''",1,-1,1)
  copyright = replace(ChkString(Request.Form("copyright"),"SQLString"),"'","''",1,-1,1)
  url = replace(ChkString(Request.Form("url"),"url"), "'","''",1,-1,1)
  turl = replace(ChkString(Request.Form("turl"),"url"), "'","''",1,-1,1)
  owner = replace(replace(Request.Form("owner"),"'","",1,-1,1),";","",1,-1,1)
  show = Cint(Request.Form("show"))

  executeThis("UPDATE PIC set TITLE='" & Title & "',DESCRIPTION ='" & pdescription & "',KEYWORD ='" & keyword & "',COPYRIGHT ='" & copyright & "',URL ='" & url & "',TURL ='" & turl & "',OWNER ='" & owner & "', ACTIVE='" & show & "' where PIC_ID =" & webid)
	
	closeandgo("admin_pic_admin.asp?cmd=13&cid=" & webid)
end if
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td valign=top class="leftPgCol">
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
pic = cLng(Request.QueryString("id"))

'breadcrumb here
  arg1 = "Admin Area|admin_home.asp"
  arg2 = "Picture Admin|admin_pic_main.asp"
  arg3 = "Edit Picture|admin_pic_editpic.asp?id=" & pic
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  spThemeBlock1_open(intSkin)

Set RS = my_Conn.Execute("select * from PIC where PIC_ID =" & pic)
boolShow = rs("ACTIVE")
%>
<b>Editing picture:</b>
<form action="admin_pic_editpic.asp" method="post">
<table border="0">
  <tr>
    <td align="right">
      <input type="hidden" name="mode" value="8">
      <input type="hidden" name="cmd" value="32">
      <input type="hidden" name="webid" value="<%=pic%>">
       
      <b>Title:</b></font></td><td><input type="text" value="<%=replace(ChkString(RS("TITLE"), "edit"), "''","'", 1, -1, 1)%>" name="title" size="40"></td>
	  <td rowspan="4" align="center" valign="middle"><%
			if instr(rs("TURL"),"_sm") > 0 then
			  stImg = "<img src=""" & rs("TURL") & """ border=""0"" alt=""Image"" title=""Image"" />"
			else
			  stImg = "<img src=""" & rs("URL") & """ border=""0"" width=""120"" alt=""image"" title=""Image"" />"
			end if 
  			response.Write(stImg)%>	  
	  </td></tr>
  <tr>
    <td align="right">
      
      <b>Description:</b></font></td><td><textarea name="description" cols=30 rows=4><%= ChkString(rs("DESCRIPTION"), "display") %></textarea></td></tr>
  <tr>
    <td align="right">
      
      <b>Keyword:</b></font></td><td><input type="text" name="keyword" value="<%=ChkString(RS("KEYWORD"), "edit")%>" size="40"></td></tr>
  <tr>
    <td align="right">
      
      <b>Copyright:</b></font></td><td><input type="text" name="copyright" value="<%=ChkString(RS("COPYRIGHT"), "edit")%>" size="40"></td></tr>
  <tr>
    <td align="right">
      
      <b>URL :</b></font></td><td><input type="text" value="<%=replace(ChkString(RS("URL"), "display"), "''","'", 1, -1, 1)%>" name="url" size="40"></td><td></td></tr>
  <tr>
    <td align="right">
      
      <b>Thumbnail URL :</b></font></td><td><input type="text" value="<%=replace(ChkString(RS("TURL"), "display"), "''","'", 1, -1, 1)%>" name="turl" size="40"></td><td></td></tr>
   <tr>
    <td align="right">
      
      <b>Poster:</b></font></td><td><input type="text" value="<%=ChkString(RS("POSTER"), "edit")%>" name="poster" size="40"></td><td> (Member Name)</td></tr>
   <tr>
    <td align="right">
      
      <b>Owners:</b></font></td><td><input type="text" value="<%=ChkString(RS("OWNER"), "edit")%>" name="owner" size="40"> </td><td>(Use 0 for 'public' or |MemberID|)</td></tr>      
  <tr>
    <td align="right">
      <b>Show :</b></font></td><td><input name="show" type="checkbox" value="1" <%=Chked(boolShow)%>></font></td></tr>
  <tr><td></td>
    <td align="left"><input type="submit" value="Update first step" class="button"><input type="reset" value="Cancel" class="button"></td><td></td></tr>
</table>
</form> 

<%

RS.Close 
set RS = nothing ' added
  spThemeBlock1_close(intSkin)
%>
		</td>
	</tr>
</table>
<!-- #INCLUDE FILE="inc_footer.asp" -->
<% 
Else
  where = server.URLEncode("admin_pic_editpic.asp?id=" & cLng(request.QueryString("id")))
  Response.Redirect "admin_login.asp?target=" & where & ""
End If %>