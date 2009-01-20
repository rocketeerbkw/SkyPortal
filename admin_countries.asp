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

pgType = "manager"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<%

if cLng(Request.QueryString("cmd")) = 10 then
  if len(Request.QueryString("coid")) = 2 then
	strSql = "SELECT CO_NAME FROM " & strTablePrefix & "COUNTRIES "
	strSql = strSql & " WHERE " & strTablePrefix & "COUNTRIES.CO_ABBREV = '" & Request.QueryString("coid") & "'"
	set rsA = my_Conn.execute(strSql)
	  coName = rsA(0)
	set rsA = nothing
	
	
	'##  Delete Country record
	strSql = "DELETE FROM " & strTablePrefix & "COUNTRIES "
	strSql = strSql & " WHERE " & strTablePrefix & "COUNTRIES.CO_ABBREV = '" & Request.QueryString("coid") & "'"
    executeThis(strSql)
                		
    '##  Update Members who had this Country to noavatar.gif
     strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
     strSql = strSql & "SET M_COUNTRY = ''"
     strSql = strSql & "WHERE M_COUNTRY = '" & coName & "'"
    executeThis(strSql)
	Session.Contents("countryHome") = "<li><span class=""fSubTitle"">" & txtCFDeleted & "</span></li>"
  end if
  closeAndGo("admin_countries.asp")
end if

if Request.Form("Method_Type") = "addCountry" then 
	Err_Msg = ""
	txURL = ChkString(Request.Form("FlagURL"),"url")
	txName = ChkString(Request.Form("CO_Name"),"")
	
	if trim(txURL) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtCFNoUrl & "</li>"
	end if

	if trim(txName) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtCFNoName & "</li>"
	end if
	
	'check if country is already in db
	sSql = "SELECT * FROM " & strTablePrefix & "COUNTRIES WHERE CO_NAME = '" & txName & "'"
	set rsChk = my_Conn.execute(sSql)
	if not rsChk.eof then
		Err_Msg = Err_Msg & replace(txtCFAlrInDB,"[%country%]",txName)
	end if
	set rsChk = nothing

	if Err_Msg = "" Then
		strSql = "INSERT INTO " & strTablePrefix & "COUNTRIES ("
		strSql = strSql & "CO_NAME"
		strSql = strSql & ", CO_ABBREV"
		strSql = strSql & ", CO_CCTLD"
		strSql = strSql & ", CO_FLAG"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & txName & "'"
		strSql = strSql & ", '" & ChkString(Request.Form("co_ABBREV"),"sqlstring") & "'"
		strSql = strSql & ", '" & ChkString(Request.Form("co_CCTLD"),"sqlstring") & "'"
		strSql = strSql & ", '" & txURL & "'"
		strSql = strSql & ")"
		executeThis(strSql)
		strMypage=Left(txName,1)

		Session.Contents("countryHome") = "<li><span class=""fSubTitle"">" & txtCFAdded & "</span></li>"
	else
		Err_Msg1 = "<li><span class=""fSubTitle"">" & txtThereIsProb & "</span></li>"
		Session.Contents("countryHome") = Err_Msg1 & "<br />" &  Err_Msg
	end if
  closeAndGo("admin_countries.asp")
end if


if Request.Form("Method_Type") = "editCountry" Then
	txURL = ChkString(Request.Form("FlagURL"),"url")
	txName = ChkString(Request.Form("C_Name"),"sqlstring")
	
	if trim(txURL) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtCFNoUrl & "</li>"
	end if

	if trim(txName) = "" then 
		Err_Msg = Err_Msg & "<li>" & txtCFNoName & "</li>"
	end if

	if Err_Msg = "" then
		strSql = "UPDATE " & strTablePrefix & "COUNTRIES "
		strSql = strSql & " SET CO_FLAG = '" & txURL & "'"
		strSql = strSql & ",    CO_NAME = '" & txName & "'"
		strSql = strSql & ",    CO_ABBREV = '" & ChkString(Request.Form("CO_ABBREV"),"sqlstring") & "'"
		strSql = strSql & ",    CO_CCTLD = '" & ChkString(Request.Form("CO_CCTLD"),"sqlstring") & "'"
		strSql = strSql & " WHERE CO_NAME = '" & ChkString(Request.Form("CO_NAME"),"sqlstring")&"'"
		my_Conn.Execute (strSql)
		strpage=Left(ChkString(Request.Form("CO_NAME"),"sqlstring"),1)

		Session.Contents("countryHome") = "<li><span class=""fSubTitle"">" & txtCFUpdated & "</span></li>"
	else
		Err_Msg1 = "<li><span class=""fSubTitle"">" & txtThereIsProb & "</span></li>"
		Session.Contents("countryHome") = Err_Msg1 & "<br />" & Err_Msg
	end if
  closeAndGo("admin_countries.asp")
end if
 %>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
  		flagConfigMenu("1")
  		response.Write("<hr />")
  		menu_admin()%>
<%
spThemeBlock1_close(intSkin) %>
	</td>
    <td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtCFMgr & "|admin_countries.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

  		if Session.Contents("countryHome") <> "" or Err_Msg <> "" then
		  strCRUresult = "<ul>"
		  if Err_Msg <> "" then
		    strCRUresult = strCRUresult & Err_Msg & "<br />" 
		  end if
		  if Session.Contents("countryHome") <> "" then
		    strCRUresult = strCRUresult & Session.Contents("countryHome") 
			Session.Contents("countryHome") = ""
		  end if
		  strCRUresult = strCRUresult & "</ul>"
		  call showMsgBlock(1,strCRUresult)
  		end if

 spThemeBlock1_open(intSkin) %>
	<div id="aa" style="display:<%= aa %>;">
    <table width="500" align="center" border="0" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" class="tTitle" colspan="6"><b><%= txtCFRevEdit %></b></td>
      </tr>
      <tr>
        <td align="center" class="tCellAlt0" colspan="6">
	<% 
	arrAlpha = split(txtAlphabet,",")
	response.Write("<a href=""admin_countries.asp"">" & txtAll & "</a>&nbsp;")
	for xa = 0 to ubound(arrAlpha)
	response.Write("&nbsp;<a href=""admin_countries.asp?C_NAME=" & arrAlpha(xa) & """>" & arrAlpha(xa) & "</a>")
	next
	%>
        </td>
      </tr>
      <tr>
        <td align="center" class="tSubTitle"><b><%= txtCFName %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtCFISO %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtCFccTLD %></b></td>
        <td align="center" class="tSubTitle"><b><%= txtFlag %></b></td>
        <td Colspan="2" align="center" class="tSubTitle"><b></b></td>
      </tr>
<% 
	' - Get Countries from DB
	C_NAME=Left(Request("C_NAME"),1)
	If Request("C_NAME")<> "" Then
	 strSql = "Select * FROM " & strTablePrefix & "COUNTRIES WHERE CO_NAME like '"& C_NAME &"%' ORDER by CO_NAME"
	
	Else
	strSql = "Select * FROM " & strTablePrefix & "COUNTRIES ORDER by CO_NAME"
	End If
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3

	if rs.EOF or rs.BOF then  '## No replies found in DB
%>
      <tr>
        <td class="tCellAlt0" colspan="6"><span class="fTitle"><b><%= txtCFNoFnd %></b></span></td>
      </tr>
<%
	Else
		rs.movefirst
		intI = 0 
		howmanyrecs = 0
		rec = 1

		do until rs.EOF '**
			if intI = 0 then
				CColor = "tCellAlt1"
			else
				CColor = "tCellAlt2"
			end If
 %>
      <tr>
        <td class="<% =CColor %>" valign="center" align="left" nowrap="nowrap">
        	<% =rs("CO_NAME") %></td>
        <td class="<% =CColor %>" valign="center" align="center">
        	<% =rs("CO_ABBREV") %></td>
        <td class="<% =CColor %>" valign="center" align="center">
        	<% =rs("CO_CCTLD") %></td>
         <td class="<% =CColor %>" valign="center" align="center" nowrap="nowrap"><img src="<% =rs("CO_FLAG") %>" border="0" hspace="0" title="<% =rs("CO_NAME") %>" alt="<% =rs("CO_NAME") %>"></td>
        <td class="<% =CColor %>" valign="center" align="center">&nbsp;
        	</td>
        <td class="<% =CColor %>" valign="center" align="center" nowrap="nowrap"><a onclick="javascript:mwpHSs('edit<% =rs("CO_ABBREV") %>','1');" href="javascript:;"><%= icon(icnEdit,txtCFEdit,"","","") %></a>
        <a href="admin_countries.asp?cmd=10&coid=<% =rs("CO_ABBREV")%>"><%= icon(icnDelete,txtCFDel,"","","") %></a></td>
      </tr>
	  <tr height="1">
        <td colspan="6" class="<% =CColor %>" valign="center" align="center" height="1">
	  <div id="edit<% =rs("CO_ABBREV") %>" style="display:none;">
	<form action="admin_countries.asp" method="post" id="PostTopic" name="PostTopic">
	<table border="0" cellspacing="0" cellpadding="0" align="center" height="1">
	  <tr>
	    <td class="tCellAlt2" valign="top" height="10">
		<table border="0" cellspacing="1" cellpadding="1" height="1" align="center">
		<tr valign="middle">
		  <td align="center" class="tAltSubTitle" colspan="3"><span class="fTitle"><%= txtCFEdit %> - <%= rs("CO_NAME") %></span>
	<input type="hidden" name="Method_Type" value="editCountry">
	<input type="hidden" name="CO_NAME" value="<%=rs("CO_NAME")%>"></td>
		</tr>
		<tr valign="middle">
		  <td class="tCellAlt0" align="right"><b><%= txtCFUrlFlag %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="255" id="FlagURL" name="FlagURL" value="<% = rs("CO_FLAG") %>" size="40" onchange ="if (CheckNav(3.0,4.0)) URL.src=form.FlagURL.value;"></td>
		  <td class="tCellAlt0"><img name="URL" src="<% if IsNull(rs("CO_FLAG")) or rs("CO_FLAG") = "" or rs("CO_FLAG") = " " then %>images/blank.gif<% else %><% =rs("CO_FLAG")%><% end if %>" border="0"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFName %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="50" name="C_NAME" value="<%= rs("CO_NAME") %>" size="40" ></td>
		  <td class="tCellAlt0">&nbsp;</td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFISO %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="4" name="CO_ABBREV" value="<%= rs("CO_ABBREV") %>" size="4"></td>
		  <td class="tCellAlt0">&nbsp;</td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFccTLD %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="4" name="CO_CCTLD" value="<%= rs("CO_CCTLD") %>" size="4"></td>
		  <td class="tCellAlt0">&nbsp;</td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" colspan="3" align="center"><input type="submit" value="<%= txtCFUpd %>" id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
		</tr>
	      </table>
	    </td>
	  </tr>
        </table>
        
  <br />
</form></div>
		</td>
	  </tr>
<%
		    rs.MoveNext
		    intI  = intI + 1
		    if intI = 2 then
				intI = 0
			end if
		    rec = rec + 1
		Loop
	end if
	rs.close
	set rs = Nothing

 %>
    </table>
	</div>
	
	<div id="ab" style="display:<%= ab %>;">
	<form action="admin_countries.asp" method="post" id="formEle" name="PostTopic">
	<input type="hidden" name="Method_Type" value="addCountry">
	<table border="0" cellspacing="0" cellpadding="0" align=center>
	  <tr>
	    <td class="tCellAlt2">
	      <table border="0" cellspacing="1" cellpadding="1">
		<tr valign="center">
		  <td align="center" class="tTitle" colspan="2"><b><%= txtCFAddNew %></b></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFName %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="50" name="CO_Name" value="" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFUrlFlag %>:</b><br />
		  <%= txtCFExample %></td>
		  <td class="tCellAlt0"><input maxLength="255" name="FlagURL" value="" size="40"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFISO %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="4" name="CO_ABBREV" value="" size="4"></td>
		</tr>
		<tr valign="center">
		  <td class="tCellAlt0" align="right"><b><%= txtCFccTLD %>:</b>&nbsp;</td>
		  <td class="tCellAlt0"><input maxLength="3" name="CO_CCTLD" value="" size="3"></td>
		</tr>
		
		<tr valign="center">
		  <td class="tCellAlt0" colspan="2" align="center"><input type="submit" value=" <%= txtSubmit %> " id="submit1" name="submit1" class="button"> <input type="reset" value="<%= txtReset %>" id="reset1" name="reset1" class="button"></td>
		</tr>
	      </table>
	    </td>
	  </tr>
	</table>
  <br />
</form>	
	</div>
	
	<div id="ac" style="display:<%= ac %>;">
	
	</div>
<% spThemeBlock1_close(intSkin) %>
	</td>
  </tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp?target=admin_countries.asp" %><% end if

sub flagConfigMenu(typ)
  if bFso then
    mnu.menuName = "b_flags"
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
	 'onclick="javascript:mwpHSs('block12<%= typ ','0');"    %>
    <div class="tCellAlt1" onmouseover="this.className='tCellHover';" onmouseout="this.className='tCellAlt1';" style="cursor:pointer; text-align:left;" onclick="javascript:location.reload();"><span style="margin: 2px;"><img name="blockFP<%= typ %>Img" id="blockFP<%= typ %>Img" src="Themes/<%= strTheme %>/icon_<%= icn %>.gif" align="absmiddle" style="cursor:pointer;" vspace="2" alt="<%= alt %>"></span>
    <b>Countries</b></div>
    <div class="menu" id="blockFP<%= typ %>" style="display: <%= cls %>;">
		<a href="admin_countries.asp"><%= icn_bar %><%= txtCFAll %><br /></a>
		<a onclick="show('ab');hide('aa');" href="javascript:;"><%= icn_bar %><%= txtCFAddNew %><br /></a>
	</div>
  <%
  end if
end sub %>
