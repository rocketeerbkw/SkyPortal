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
<%If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_functions.asp" -->
<script type="text/javascript">
<!--
function selectUsers()
{
	if (document.PostTopic.AuthUsers.length == 1)
	{
		document.PostTopic.AuthUsers.options[0].value = "";
		return;
	}
	if (document.PostTopic.AuthUsers.length == 2)
		document.PostTopic.AuthUsers.options[0].selected = true
	else
	for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
		document.PostTopic.AuthUsers.options[x].selected = true;
		
	if (document.PostTopic.grpLeader.length == 1)
	{
		document.PostTopic.grpLeader.options[0].value = "";
		return;
	}
	if (document.PostTopic.grpLeader.length == 2)
		document.PostTopic.grpLeader.options[0].selected = true
	else
	for (x = 0;x < document.PostTopic.grpLeader.length - 1 ;x++)
		document.PostTopic.grpLeader.options[x].selected = true;
	//selectLeaders()
}

function MoveWholeList(strAction)
{
	if (strAction == "Add")
	{
		if (document.PostTopic.AuthUsersCombo.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsersCombo.length - 1 ;x++)
			document.PostTopic.AuthUsersCombo.options[x].selected = true;
			InsertSelection("Add");
		}
	}
	else
	{
		if (document.PostTopic.AuthUsers.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
			document.PostTopic.AuthUsers.options[x].selected = true;
			InsertSelection("Del");
		}
	}
}

function InsertSelection2(strAction)
{
	var pos,user,mText;
	var count,finished;

	if (strAction == "Add")
	{
		pos = document.PostTopic.grpLeader.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				//alert("You must select someone from the 'Members List'")
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected)
			{
				document.PostTopic.grpLeader.length +=1;
				document.PostTopic.grpLeader.options[pos].value = document.PostTopic.grpLeader.options[pos-1].value;	
				document.PostTopic.grpLeader.options[pos].text = document.PostTopic.grpLeader.options[pos-1].text;
				document.PostTopic.grpLeader.options[pos-1].value = document.PostTopic.AuthUsers.options[count].value;	
				document.PostTopic.grpLeader.options[pos-1].text = document.PostTopic.AuthUsers.options[count].text;
				document.PostTopic.grpLeader.options[pos-1].selected = true;
			}
			pos = document.PostTopic.grpLeader.length;
			count += 1;
		}while (!finished); //finished adding
	}	

	if (strAction == "Del")
	{
		pos = document.PostTopic.AuthUsersCombo.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected)
			{
				document.PostTopic.AuthUsersCombo.length +=1;
				document.PostTopic.AuthUsersCombo.options[pos].value = document.PostTopic.AuthUsersCombo.options[pos-1].value;	
				document.PostTopic.AuthUsersCombo.options[pos].text = document.PostTopic.AuthUsersCombo.options[pos-1].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].value = document.PostTopic.AuthUsers.options[count].value;	
				document.PostTopic.AuthUsersCombo.options[pos-1].text = document.PostTopic.AuthUsers.options[count].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].selected = true;
			}
			count += 1;
			pos = document.PostTopic.AuthUsersCombo.length;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsers.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}	
}

function DeleteSelection()
{
	var user,mText;
	var count,finished;

		finished = false;
		count = 0;
		count = document.PostTopic.AuthUsers.length - 1;
		if (count<1) {
			return;
		}
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
}

function delLeader()
{
	var user,mText;
	var count,finished;

		finished = false;
		count = 0;
		count = document.PostTopic.grpLeader.length - 1;
		if (count<1) {
			return;
		}
		do //remove from source
		{	
			if (document.PostTopic.grpLeader.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.grpLeader.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.grpLeader.length-1;z++)
				{	
					document.PostTopic.grpLeader.options[z].value = document.PostTopic.grpLeader.options[z+1].value;	
					document.PostTopic.grpLeader.options[z].text = document.PostTopic.grpLeader.options[z+1].text;
				}
				document.PostTopic.grpLeader.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
}

function delGroup(grp,frmID){
  if (confirm("<%= txtCG10 %> '"+grp+"'?\n<%= txtCannotBeUndn %>")) {
   for (i=0; i<document.forms.length; i++) {
   if (document.forms[i].name == "fm_"+frmID) {
    document.forms[i].submit();
	}}
	
  }
}

function allowmembers() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=allowmember", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }
//-->
</script>
<%	
  strCGrResult = ""
	if request.querystring("mode") = 3 and request.querystring("id") <> "" then
		g_id = cLng(request.querystring("id"))
		strTemp = 0
		
		strSql = "SELECT G_ACTIVE FROM " & strTablePrefix & "GROUPS"
		strSql = strSql & " WHERE G_ID = " & g_id
		set rsGp = my_Conn.execute(strSql)
		if not rsGp.eof then
			strTemp = rsGp("G_ACTIVE")
		end if
		set rsGp = nothing
		
		if strTemp = 1 then
		  strSql = "DELETE FROM " & strTablePrefix & "GROUP_MEMBERS"
		  strSql = strSql & " WHERE G_GROUP_ID = " & g_id
		  executeThis(strSql)
		
		  strSql = "DELETE FROM " & strTablePrefix & "GROUPS"
		  strSql = strSql & " WHERE G_ID = " & g_id
		  executeThis(strSql)
		
		' Still to do...
		' remove from portal_apps
		  strCGrResult = "Group Deleted"
		end if
		session.Contents("strCGrResult") = strCGrResult
		closeandgo("admin_config_groups.asp")
	end if
	
	if Request.Form("Method_Type") = "add_new" then 
		Err_Msg = ""
			g_name = replace(Request.Form("g_name"),"'","")
			g_desc = replace(Request.Form("g_desc"),"'","")
			g_members = Request.Form("AuthUsers")
			g_leaders = Request.Form("grpLeader")
			g_modify = strCurDateString
			g_create = g_modify
		if g_members = "" then
		  Err_Msg = "The Group must have members"
		end if

		if Err_Msg = "" then
			'g_id = cInt(Request.Form("g_id"))
			
			'response.Write("Name: " & g_name & "<br />")
			'response.Write("Desc: " & g_desc & "<br />")
			'response.Write("Members: " & g_members & "<br />")
			'response.Write("Leaders: " & g_leaders & "<br />")

			strSql = "INSERT INTO " & strTablePrefix & "GROUPS"
			strSql = strSql & " (G_DESC,G_NAME,G_MODIFIED,G_CREATE,G_ACTIVE,G_ADDMEM)"
			strSql = strSql & " VALUES "
			strSql = strSql & "('" & g_desc & "','" & g_name & "','" & g_modify & "',"
			strSql = strSql & "'" & g_create & "',1,1);"
			executeThis(strSql)

			strSql = "SELECT G_ID FROM " & strTablePrefix & "GROUPS"
			strSql = strSql & " WHERE G_NAME = '" & g_name & "'"
			set rsTmp = my_Conn.execute(strSql)
			  g_id = rsTmp(0)
			set rsTmp = nothing
			
			if g_members <> "" then
			  if inStr(g_members,",") > 0 then
			    arrMembers = split(g_members,",")
				for g = 0 to ubound(arrMembers)
				  strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
				  strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
				  strSql = strSql & " (" & trim(arrMembers(g)) & "," & g_id & ",0)"
				  executeThis(strSql)				  
				next
			  else
				  strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
				  strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
				  strSql = strSql & " (" & trim(g_members) & "," & g_id & ",0)"
				  executeThis(strSql)			  
			  end if
			  
			  ' check and insert group leaders	
			  if inStr(g_leaders,",") > 0 then
			      arrLeaders = split(g_leaders,",")
				  for h = 0 to ubound(arrLeaders)
				    strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
					strSql = strSql & " SET G_GROUP_LEADER = 1"
					strSql = strSql & " WHERE G_MEMBER_ID = " & arrLeaders(h)
					strSql = strSql & " AND G_GROUP_ID = " & g_id
					executeThis(strSql)				  
				  next
				elseif len(g_leaders) > 0 then
				    strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
					strSql = strSql & " SET G_GROUP_LEADER = 1"
					strSql = strSql & " WHERE G_MEMBER_ID = " & g_leaders
					strSql = strSql & " AND G_GROUP_ID = " & g_id
					executeThis(strSql)
				end if		
			else 'g_members = ""
			
			end if 'g_members check
				closeandgo("admin_config_groups.asp")
			strCGrResult = "<span class=""fTitle"">" & txtCG01 & "</span>"
		else 
			strCGrResult = "<span class=""fTitle"">" & txtThereIsProb & "</span>"
			strCGrResult = strCGrResult & "<ul>" & Err_Msg & "</ul>"
			strCGrResult = strCGrResult & "<a href=""JavaScript:history.go(-1)"">" & txtGoBack & "</a>"
		end if
		session.Contents("strCGrResult") = strCGrResult
		closeandgo("admin_config_groups.asp")
	
	elseif Request.Form("Method_Type") = "modify_config" then 
		Err_Msg = ""

		if Err_Msg = "" then
			g_id = cInt(Request.Form("g_id"))
			g_name = replace(Request.Form("g_name"),"'","")
			g_desc = replace(Request.Form("g_desc"),"'","")
			g_members = Request.Form("AuthUsers")
			g_leaders = Request.Form("grpLeader")
			g_modify = strCurDateString
			
			if g_id = 1 then
			  'response.Write("Admin group")
			  sSql = "UPDATE PORTAL_MEMBERS SET M_LEVEL = 1 WHERE M_LEVEL = 3"
			  executeThis(sSql)
			  sSql = "UPDATE PORTAL_MEMBERS SET M_LEVEL = 3"
			  sSql = sSql & " WHERE MEMBER_ID IN (" & g_members & ")"
			  executeThis(sSql)
			else
			  'response.Write("Other group")
			end if

			strSql = "UPDATE " & strTablePrefix & "GROUPS"
			strSql = strSql & " SET G_DESC = '" & g_desc & "'"
			strSql = strSql & ", G_NAME = '" & g_name & "'"
			strSql = strSql & ", G_MODIFIED = '" & g_modify & "'"
			strSql = strSql & " WHERE G_ID = " & g_id
			executeThis(strSql)

			strSql = "DELETE FROM " & strTablePrefix & "GROUP_MEMBERS"
			strSql = strSql & " WHERE G_GROUP_ID = " & g_id
			executeThis(strSql)
			
			if g_members <> "" then
			  if inStr(g_members,",") > 0 then
			    arrMembers = split(g_members,",")
				for g = 0 to ubound(arrMembers)
				  strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
				  strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
				  strSql = strSql & " (" & trim(arrMembers(g)) & "," & g_id & ",0)"
				  executeThis(strSql)				  
				next
			  else
				  strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
				  strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
				  strSql = strSql & " (" & trim(g_members) & "," & g_id & ",0)"
				  executeThis(strSql)			  
			  end if
			  
			  ' check and insert group leaders	
			  if inStr(g_leaders,",") > 0 then
			      arrLeaders = split(g_leaders,",")
				  for h = 0 to ubound(arrLeaders)
				    strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
					strSql = strSql & " SET G_GROUP_LEADER = 1"
					strSql = strSql & " WHERE G_MEMBER_ID = " & arrLeaders(h)
					strSql = strSql & " AND G_GROUP_ID = " & g_id
					executeThis(strSql)				  
				  next
				elseif len(g_leaders) > 0 then
				    strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
					strSql = strSql & " SET G_GROUP_LEADER = 1"
					strSql = strSql & " WHERE G_MEMBER_ID = " & g_leaders
					strSql = strSql & " AND G_GROUP_ID = " & g_id
					executeThis(strSql)
				end if		
			else 'g_members = ""
			
			end if 'g_members check
			strCGrResult = "<span class=""fTitle"">" & txtCG01 & "</span>"
		else
			strCGrResult = "<span class=""fTitle"">" & txtThereIsProb & "</span>"
			strCGrResult = strCGrResult & "<ul>" & Err_Msg & "</ul>"
			strCGrResult = strCGrResult & "<a href=""JavaScript:history.go(-1)"">" & txtGoBack & "</a>"
		end if
		session.Contents("strCGrResult") = strCGrResult
		closeandgo("admin_config_groups.asp") %>
<%	else  %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td class="leftPgCol" width="190">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtCG02 & "|admin_config_groups.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  if session.Contents("strCGrResult") <> "" then
    call showMsgBlock(1,session.Contents("strCGrResult"))
	session.Contents("strCGrResult") = ""
  end if

spThemeBlock1_open(intSkin)
		select case cInt(request.QueryString("mode"))
			case 1 ' edit
				call displayForm("edit")
			case 2 ' add
				call displayForm("add")
			case else ' list all
				call listAll()
		end select
spThemeBlock1_close(intSkin) %>
</td></tr></table>
<%	end if %>
<!--#include file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp?target=admin_config_groups.asp" %>
<% End If

Sub listAll() %>
  <table class="tCellAlt2" border="1" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td align="center"> 
        <table border="0" cellspacing="1" cellpadding="3" align="center">
          <tr valign="top"> 
            <td class="tTitle" colspan="3"><span class="fTitle"><%= txtCG03 %></span></td>
          </tr>
    <tr align="center" valign="middle" class="tCellAlt0"> 
      <td width="100" nowrap><a href="admin_config_groups.asp?mode=2" title="<%= txtCG04 %>"><b><%= txtCG05 %></b></a></td>
      <td width="150" height="25"><%= txtCG06 %></td>
      <td width="100" height="25">&nbsp;</td>
    </tr>
 <% set rsUp = my_Conn.execute("select * from " & strTablePrefix & "GROUPS ORDER by G_ADDMEM, G_ACTIVE, G_NAME")
	if not rsUp.eof then
		do until rsUp.eof
		  grpName = chkString(rsUp("G_NAME"),"display")
		  grpID = cInt(rsUp("G_ID"))
		  grpAddMem = cInt(rsUp("G_ADDMEM"))
		  grpActive = cInt(rsUp("G_ACTIVE")) %>
    <tr align="center" class="tCellAlt1">
      <td>&nbsp;</td>
      <td><%= grpName %></td>
      <td align="left"><form name="fm_<%= grpID %>" action="admin_config_groups.asp" id="fm_<%= grpID %>" method="get">
<input type="hidden" name="mode" value="3">
<input type="hidden" name="id" value="<%= grpID %>">
	  &nbsp;&nbsp;&nbsp;<a href="javascript:;" onclick="openWindow3('pop_portal.asp?cmd=5&mode=2&cid=<%= grpID %>')"><%= icon(icnInfo,replace(txtCG07,"[%groupName%]",grpName),"","","") %></a>
	  <% If grpAddMem Then %>&nbsp;<a href="admin_config_groups.asp?mode=1&id=<%= grpID %>"><%= icon(icnEdit,replace(txtCG08,"[%groupName%]",grpName),"","","") %></a><% End If %>
	  <% If grpActive Then %>&nbsp;<a href="javascript:delGroup('<%= grpName %>','<%= grpID %>');"><%= icon(icnDelete,replace(txtCG09,"[%groupName%]",grpName),"","","") %></a><% End If %>
</form></td>
    </tr>
	<% 		rsUp.movenext
		loop 
		set rsUp = nothing %>
 <% Else %>
    <tr align="center" class="tCellAlt0"> 
      <td colspan="3"><b><%= txtCG11 %></b></td>
    </tr>
 <% End If %>
  </table></td></tr></table><br />&nbsp;
<%
End Sub 

sub displayForm(typ)
	btn = txtCG05
	method = "add_new"
	g_field = ""
	if typ = "edit" then
		set rsGrp = my_Conn.execute("select * from " & strTablePrefix & "GROUPS where G_ID = " & cInt(request.QueryString("id")))
		'set rsGrp = my_Conn.execute(sSQL)
		if not rsGrp.eof then
		  gid = rsGrp("G_ID")
		  gname = rsGrp("G_NAME")
		  gdesc = rsGrp("G_DESC")
		  gcreate = strtodate(rsGrp("G_CREATE"))
		  if trim(rsGrp("G_MODIFIED")) <> "" then
		    gmodify = strtodate(rsGrp("G_MODIFIED"))
		  end if
		  gactive = rsGrp("G_ACTIVE")
		  btn = txtSknUpdate
		  method = "modify_config"
		  g_field = "readonly"
		end if
		Set rsGrp = nothing
	end if %>
<form action="admin_config_groups.asp" method="post" id="PostTopic" name="PostTopic">
<input type="hidden" name="Method_Type" value="<%= method %>">
<input type="hidden" name="g_id" value="<%= gid %>">
<input type="hidden" name="groupOwner" value="0">
<table class="tCellAlt2" style="width:600px;" border="0" cellspacing="0" cellpadding="0" align=center>
    <tr> 
      <td> 
        <table border="0" cellspacing="1" cellpadding="3" class="tCellAlt1" width="100%">
          <tr> 
            <td class="tTitle"><span class="fTitle"><%= txtCG12 %></span></td>
          </tr>
		<tr><td align="center" class="tSubTitle">
			<%= uCase(typ) %>&nbsp;<%= txtGroup %></td>
          </tr>
		<tr><td align=left>
		<br />
		<table border=0 cellpadding=0 cellspacing=5 align=center>
			<tr><td><b><%= txtGrpNam %></b></td><td>&nbsp;<input name="g_name" type="text" value="<%= gname %>" maxLength=50 size=51 class=textbox <%= g_field %>></td></tr>
			<tr><td><b><%= txtDesc %></b></td><td>&nbsp;<input name="g_desc" type="text" value="<%= gdesc %>" maxLength=50 size=51 class=textbox></td></tr>
		</table>
		
      <br />
      <fieldset style='margin:10px;'>
		<legend><b><%= txtMbrshp %></b></legend>
      <table border=0 cellpadding=0 cellspacing=0 width=400>
        <tr><td align="center" valign="middle"><a href="JavaScript:allowmembers();" title="<%= txtCG14 %>"><b><%= txtAddMem %></b></a><br />
				<a href="JavaScript:DeleteSelection();" title="<%= txtCG15 %>"><b><%= txtRemMember %></b></a></td>
          <td align="center"><b><%= txtCG13 %></b><br />
            <select size="5" id="AuthUsers" name="AuthUsers" style="width:170;" multiple>
			  <% if cInt(request.QueryString("mode")) = 1 then
			  		getGroupMembers(gid)
				 end if %>
			  <option value="0"></option>
            </select>
            <input type="hidden" name="access" value="">
          </td>
        </tr>
        <tr> <td align="center" valign="middle"><br />
				
				<!-- <a href="javascript:InsertSelection2('Add');" title="<%= txtCG16 %>"> -->
				<a href="JavaScript:moveGroup('Add','PostTopic','AuthUsers','grpLeader');" title="<%= txtCG16 %>">
				<b><%= txtCG18 %></b></a><br />
				<a href="JavaScript:delLeader();" title="<%= txtCG17 %>">
				<b><%= txtCG19 %></b></a></td>
          <td align=center> <br />
            <br />
            <b><%= txtCG20 %></b><br />
            <select size="5" name="grpLeader" id="grpLeader" style="width:170;" multiple>
			  <% if cInt(request.QueryString("mode")) = 1 then
			  		getGroupLeaders(gid)
				 end if %>
			  <option value="0"></option>
            </select>
            <input name="groupLeader" type="hidden" value=""><br /><br />
          </td>
        </tr>
      </table>
		</fieldset>		
		<br /><div align=center><input name="Submit" type="submit" value="<%= btn %>" onclick="selectUsers()" class="button">&nbsp;&nbsp;<a href="admin_config_groups.asp"><%= txtGoBack %></a>
		<br /><br />
</div>

		</td></tr>
	</table>
		</td></tr>
	</table>
</form><br />
<% End Sub 
%>