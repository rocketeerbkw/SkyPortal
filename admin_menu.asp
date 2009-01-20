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
PINAME = ""
addList = ""
mTitle = ""
c_color = "tCellAlt0"
m_debug = false

	if request("CINAME") <> "" then
	  menu = request("CINAME")
	else  
	  if request("menu") <> "" then
		Menu = request("menu")
  	  else
    	Menu = "def_main"
  	  end if
	end if %>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<!--#include file="inc_top.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin Then %>
<!--#include file="includes/inc_admin_menu_edit.asp" -->
<!-- <script type="text/javascript" src="includes/scripts/admin_menu.js"></script> -->
<script type="text/javascript">
	function mode(cmd,Pid,Cid,title) {
		switch (cmd) {
			case "Cdelete":
				if (!confirm('<%= txtJSDelLnk %>\n<%= txtNoUndo %>')) return;
				break;
			case "Pdelete":
				if (!confirm('<%= txtJSDelLnk %>\n<%= txtJSDelMnu2 %>\n<%= txtNoUndo %>!')) return;
				break;
			case "Mdelete":
				if (!confirm('<%= txtJSDelMnu %>\n<%= txtNoUndo %>')) return;
				break;
			case "editparent":
				str = prompt("Enter a new name for this Link", Cid);
				if(!str) return;
				else if (!CheckName(str)) {alert("Name can not contain any of the\nfollowing characters: \\ / : * ' . , ? \" < > |"); return;}
				Cid = Cid + '|' + str;
				break;
			case "upPorderss":
				if (!confirm('<%= txtJSDelMnu %>\n<%= txtNoUndo %>')) return;
				break;
			default:
				document.forms.osBuffer.action = "admin_menu.asp";
		}
		document.forms.osBuffer.cmd.value = cmd
		document.forms.osBuffer.Pid.value = Pid;
		document.forms.osBuffer.Cid.value = Cid;
		document.forms.osBuffer.mnuTitle.value = title;
		document.forms.osBuffer.action = "admin_menu.asp";
		document.forms.osBuffer.submit();	
	}
	
function CheckMenu(frm) {
	var oFrm = document.getElementById(frm);
	var sT = oFrm.mnuTitle.value;
	//alert(sT);
//	var strName = document.[strFrm];
//	alert(strName);
	if (!CheckName(sT)) {alert("<%= txtJSBadChars %>: \\ / : * ? \" < > |");
	return false;
	}
	if(!sT) {alert("<%= txtJSNamBlnk %>");
	return false;
	}
	else 
	{ 
	return true;
	}
}

function upPorder(cont){
	var ord = cont + ':' + 'x';
	cont == cont++
	for (i=1; i < cont; i++){
		ord += "|";
		pord = "editPorder" + i;
		ord += document.forms.pord.Pid.value;
	alert(ord);
  	}
		pord = document.forms.editPorder[i].Pid[i].value;
		document.forms.osBuffer.cmd.value = "Porder";
		document.forms.osBuffer.mnuTitle.value = ord;
		document.forms.osBuffer.action = "admin_menu.asp";
//		document.forms.osBuffer.submit();
}

function popHelp(hlp){
	switch (hlp) {
		case "link":
			alert('<%= txtJSHpLnk %>');
			break;
		case "image":
			alert('<%= txtJSHpOpt %>\n\n<%= txtJSHpImg %>');
			break;
		case "onclick":
			alert('<%= txtJSHpOpt %>\n\n<%= txtJSHpOnClk %>');
			break;
		case "Function":
			alert('<%= txtJSHpOpt %>\n\n<%= txtJSHpFunct %>');
			break;
		case "grpAccess":
			alert('<%= txtJSHpOpt %>\n\n<%= txtJSHpGrp %>');
			break;
		case "target":
			alert('<%= txtJSHpTarg %>');
			break;
		case "addmenu":
			alert('<%= txtJSHpOpt %>\n\n<%= txtJSHpAddMnu %>');
			break;
		default:
			
	}
}
</script>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center"><tr>
<td class="leftPgCol" nowrap>
<% intSkin = getSkin(intSubSkin,1) %>
<% 
'pgType = ""
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
menu_admin()
spThemeBlock1_close(intSkin)

mnu.menuName = Menu
mnu.template = 4
mnu.thmBlk = 1
mnu.title = "Template 4"
mnu.shoExpanded = 0
mnu.canMinMax = 1
mnu.keepOpen = 0
'mnu.GetMenu()

mnu.title = Menu
mnu.template = 5
mnu.thmBlk = 1
mnu.title = "Template 5"
mnu.menuName = Menu
'mnu.GetMenu()

mnu.menuName = Menu
mnu.template = 4
mnu.thmBlk = 1
mnu.title = "Template 4"
mnu.shoExpanded = 0
mnu.canMinMax = 1
mnu.keepOpen = 1
'mnu.GetMenu()

mnu.menuName = Menu
mnu.template = 1
mnu.thmBlk = 1
mnu.title = "Template 1"
'mnu.GetMenu()

'mnu.template = 2
'mnu.title = "clickMenu"
'mnu.thmBlk = 1
'mnu.menuName = Menu
'mnu.GetMenu()

'spThemeTitle = txtMenu
'spThemeBlock1_open(intSkin)
	'menu_admin()
'spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<% intSkin = getSkin(intSubSkin,2) %>
<%
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtManagers & "|admin_home.asp"
  arg3 = txtMnuMgr & "|admin_menu.asp"
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<!-- <table border="0" cellspacing="4" cellpadding="0" align="center">
  <tr>
    <td width="50" height="50" class="tCellAlt0">tCellAlt0</td>
    <td width="50" height="50" class="tCellAlt1">tCellAlt1</td>
    <td width="50" height="50" class="tCellAlt2">tCellAlt2</td>
    <td width="50" height="50" class="tCellHover">tCellHover</td>
  </tr>
</table> -->

<% 
spThemeTitle = txtSkyMnuMgr
spThemeBlock1_open(intSkin)

mnu.title = Menu
mnu.template = 3
mnu.menuName = Menu
'mnu.GetMenu()

'mnu.template = 3
'mnu.menuName = "Main"
'mnu.GetMenu()

  chkSessionMsg()
%>
<table width="100%">
<tr><td width="100%">
<% 
editmenu()
%>
</td></tr></table>
<%
spThemeBlock1_close(intSkin)
shoMenuTemplates()
 %>
</td><td class="rightPgCol" nowrap>
<% 
intSkin = getSkin(intSubSkin,3)

spThemeTitle = txtRbldMnuFls
spThemeBlock1_open(intSkin)
%>
<table width="100%">
<tr><td width="100%">
<%= txtRbldMnuTxt %><br><br>
<form action="admin_menu.asp" method="post" name="rmenu">
<input name="mode" type="hidden" value="resetmenu">
<input name="submit" type="submit" value="<%= txtRbldMnuFls %>">
</form>
</td></tr></table>
<%
spThemeBlock1_close(intSkin)

mnu.menuName = "m_admin"
mnu.title = "Admin Menu"
mnu.template = 4
mnu.thmBlk = 1
mnu.shoExpanded = 0
mnu.canMinMax = 1
'mnu.GetMenu()

%>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_home.asp" %><% end if

sub shoMenuTemplates()
  Response.Write "<table border=""0"" cellpadding=""10"" cellspacing=""0"" width=""100%"">"
  Response.Write "<tr><td width=""33%"">"
mnu.menuName = "def_main"
mnu.template = 1
mnu.thmBlk = 1
mnu.title = "Template 1"
mnu.GetMenu()
  Response.Write "</td><td width=""33%"">"
mnu.menuName = "def_main"
mnu.template = 4
mnu.thmBlk = 1
mnu.title = "Template 4"
mnu.shoExpanded = 0
mnu.canMinMax = 1
mnu.keepOpen = 1
mnu.GetMenu()
  Response.Write "</td><td width=""33%"">"
mnu.menuName = "def_main"
mnu.template = 5
mnu.thmBlk = 1
mnu.title = "Template 5"
mnu.GetMenu()
  Response.Write "</td></tr>"
  Response.Write "</table>"



end sub

sub editmenu() 
  dim arrMenu() 
  dim arrSub()
  dim upOrd()
  sc = 0 %>
	<FORM name="osBuffer" method = "post" action = "admin_menu.asp">
    	<INPUT type="hidden" name="cmd" value="">
        <INPUT type="hidden" name="Pid" value="">
        <INPUT type="hidden" name="Cid" value="">
        <INPUT type="hidden" name="mnuTitle" value="">
    </FORM>
<style type="text/css">
	.menutitle{
	cursor:pointer;
	margin-bottom: 2px;
	background-color:#ECECFF;
	color:#000000;
	width:100px;
	padding:2px;
	text-align:center;
	border:1px solid #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 10px;
	}
	.submenu {
	  display: none;
	}
	</style>
<% 
set rsDist = my_Conn.Execute("SELECT distinct mnuTitle, INAME from Menu order by mnuTitle")
addList = "<select name=""AddMenu"">" & vbCRLF
addList = addList & "<option value=""""></option>" & vbCRLF

set rsPedit = my_Conn.Execute("SELECT *, (SELECT COUNT(*) FROM Menu Where Parent ='" & Menu & "' and INAME = '" & Menu & "') as Linkcnt from Menu Where Parent ='" & Menu & "' and INAME = '" & Menu & "' order by mnuOrder")
if not rsPedit.eof then
  CntLinks = rsPedit("Linkcnt") 
  PINAME = rsPedit("INAME")
  mTitle = rsPedit("mnuTitle")
  mINAME = rsPedit("INAME")
  app_id = rsPedit("app_id")
else
  Menu = "Main"
  set rsPedit = my_Conn.Execute("SELECT *, (SELECT COUNT(*) FROM Menu Where Parent ='" & Menu & "' and INAME = '" & Menu & "') as Linkcnt from Menu Where Parent ='" & Menu & "' and INAME = '" & Menu & "' order by mnuOrder")
  CntLinks = rsPedit("Linkcnt") 
  PINAME = rsPedit("INAME")
  mTitle = rsPedit("mnuTitle")
  mINAME = rsPedit("INAME")
  app_id = rsPedit("app_id")
end if %>
<div id="editdiv"> 
<!--form name="EditForm" method="post" action="admin_menu.asp"-->        
<table width="500" border="0" cellspacing="0" cellpadding="0" align="center" class="tCellAlt0">
  <tr align="center" valign="middle"> 
    <td height="20" colspan="2" class="tSubTitle">
	  <span class="fSubTitle"><%= txtEditSiteMenus %></span> </td>
  </tr>
  <tr> 
    <td> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="50" colspan="2" align="center" valign="middle"> 
              <form name="jmpMenu" method="post" action="admin_menu.asp">
                <img src="themes/<%= strTheme %>/mnu_icons/icon_Dogg.gif" width="15" height="15" title="SkyDogg is here!" alt="SkyDogg">&nbsp;&nbsp;&nbsp;&nbsp; 
                <%= txtSelMnu %>: 
                <select name="menu" onchange="submit()">
                  <% do while not rsDist.eof
				    If rsDist("INAME") <> "" Then
				     If rsDist("INAME") = Menu Then %>
                  <option value="<%= rsDist("INAME") %>" selected="selected"><%= rsDist("mnuTitle") %></option>
                  <% Else
				  	   if rsDist("INAME") <> "sadmin" and rsDist("INAME") <> "admin" and rsDist("INAME") <> "nav_main" and rsDist("INAME") <> "Main" then
				  		 addList = addList & "<option value=""" & rsDist("INAME") & """>" & rsDist("mnuTitle") & "</option>" & vbCRLF
					   end if
				  %>
                  <option value="<%= rsDist("INAME") %>"><%= rsDist("mnuTitle") %></option>
                  <% End If
				    End if
				  	 rsDist.movenext
				     loop 
				     addList = addList & "</select>" & vbCRLF
					 rsDist.close %>
                </select>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp; 
                <input class="button" onclick="SwitchMenu('editdiv','addmnu')" name="newMnu" type="button" id="newMnu" value="<%= txtCreateMnu %>">
                <!-- <input class="button" onclick="javascript:openJsLayer('addmnu','500','250');" name="newMnu" type="button" id="newMnu" value="Create Menu"> -->
              </form>
            </td>
          </tr>
          <tr valign="top"> 
            <td colspan="2"> 
              <% addMenuForm() %>
            </td>
          </tr>
          <tr> 
            <td colspan="2"> 
              <hr align="center" noshade />
            </td>
          </tr>
          <tr> 
            <td colspan="2"> &nbsp;&nbsp; <span class="fSubTitle"><b><%= txtMnuNam %>:</b>&nbsp;<span class="fAlert"><b><%= mTitle %></b></span></span>&nbsp;&nbsp; 
              <% If Menu <> "def_main" and Menu <> "nav_main" and Menu <> "cp_main" and Menu <> "sadmin" Then %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_trashcanL.gif" onclick="mode('Mdelete','','','<%= rsPedit("INAME") %>')" border="0" width="15" height="15" title="<%= txtDelMnu %> '<%= mTitle %>'" alt="<%= txtDelMnu %>" style="cursor:pointer;"> 
              <% End If %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkP.gif" onclick="SwitchMenu('editdiv','addparent')" border="0" width="15" height="15" alt="<%= txtAddLnkMnu %>" title="<%= txtAddLnkMnu %>: '<%= mTitle %>'" style="cursor:pointer;">
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkA.gif" onclick="SwitchMenu('editdiv','updateOrder')" border="0" width="15" height="15" alt="<%= txtUpdOrd %>" title="<%= txtUpdOrd %>" style="cursor:pointer;"> 
              &nbsp;<img src="themes/<%= strTheme %>/icons/view.gif" onclick="SwitchMenu('editdiv','viewCode')" border="0" alt="Menu Code" title="View Menu Code" style="cursor:pointer;"> 
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2"> 
              <%  
			    addLinkForm()
			    viewCode()
				UpdateOrderForm "Menu",0
			%>
            </td>
          </tr>
          <tr> 
            <td colspan="2"> 
              <hr align="center" noshade />
            </td>
          </tr>
<% do while not rsPedit.eof
		  ed = ed + 1
		  if rsPedit("mnuAdd") <> "" then
		    'arrMenu(0,11) = arrMenu(0,13)
	    	strSQL = "SELECT COUNT(*) FROM Menu Where Parent ='" & rsPedit("mnuAdd") & "' and INAME = '" & rsPedit("mnuAdd") & "'"
		  else
		    'arrMenu(0,13) = arrMenu(0,11)
	    	strSQL = "SELECT COUNT(*) FROM Menu Where Parent ='" & rsPedit("Name") & "' and INAME = '" & rsPedit("INAME") & "'"
		  end if
		set rsCount = my_Conn.Execute(strSQL)
		intCount = cLng(rsCount(0))
		set rsCount = nothing
 %>
          <tr> 
            <td width="27%" align="center" valign="bottom"> 
              <% If (trim(rsPedit("Link")) = "" and trim(rsPedit("onclick") = "")) and intCount <> 0 Then %>
              <div class="menutitle" onclick="SwitchMenu('editdiv','child<%= rsPedit("id") %>')"><%= replace(rsPedit("Name"),"''","'") %></div>
              <% Else %>
              <div class="menutitle"><%= replace(rsPedit("Name"),"''","'") %></div>
              <% End If %>
            </td>
            <td valign="bottom"> 
            <% 'If Menu <> "Main" and Menu <> "nav_main" and Menu <> "admin" and Menu <> "sadmin" Then
			   If rsPedit("mnuAdd") <> "" Then
			    %>
              	<img src="themes/<%= strTheme %>/mnu_icons/icon_trashcanL.gif" onclick="mode('Pdelete','<%= rsPedit("id") %>','<%= rsPedit("Name") %>','<%= Menu %>')" border="0" width="15" height="15" alt="<%= txtDelLnk %>" title="<%= txtDelLnk %>: '<%= rsPedit("Name") %>'" style="cursor:pointer;"> 
			<% Else %>
              	<img src="themes/<%= strTheme %>/mnu_icons/icon_trashcanL.gif" onclick="mode('Pdelete','<%= rsPedit("id") %>','<%= rsPedit("Name") %>','<%= Menu %>')" border="0" width="15" height="15" alt="<%= txtDelLnk %>" title="<%= txtDelLnk %>: '<%= rsPedit("Name") %>'" style="cursor:pointer;"> 
            <%  end if
			   If len(rsPedit("mnuAdd") & "x") = 1 Then %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_pencilL.gif" onclick="SwitchMenu('editdiv','parent<%= ed %>')" border="0" width="15" height="15" alt="<%= txtEditLnk %>" title="<%= txtEditLnk %>: '<%= rsPedit("Name") %>'" style="cursor:pointer;"> 
        <% 	   end if
		If rsPedit("Link") = "" and rsPedit("onclick") = "" Then
		  If len(rsPedit("mnuAdd") & "x") = 1 Then %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkP.gif" onclick="SwitchMenu('editdiv','addchild<%= rsPedit("id") %>')" border="0" width="15" height="15" alt="<%= txtAddChldLnk %>" title="<%= txtAddChldLnk %>: '<%= rsPedit("Name") %>'" style="cursor:pointer;"> 
     <%   end if
		   	If intCount <> 0 Then %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_childYL.gif" onclick="SwitchMenu('editdiv','child<%= rsPedit("id") %>')" border="0" width="15" height="15" alt="<%= txtView %>" title="<%= txtVECLnks %>: '<%= rsPedit("Name") %>'" style="cursor:pointer;"> 
            <%  end if %>
           <% If intCount > 1 Then %>
              &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkA.gif" onclick="SwitchMenu('editdiv','updateOrder<%= rsPedit("id") %>')" border="0" width="15" height="15" alt="<%= txtUpdOrd %>" title="<%= txtUpdOrd %>: '<%= rsPedit("Name") %>' " style="cursor:pointer;"> 
           <% End if
		End If %>
              <input type="hidden" name="Pid" id="Pid" value="<%= rsPedit("id") %>">
			<% 
		  reDim arrMenu(1,15)
		  arrMenu(0,0) = rsPedit("id")
	  	  arrMenu(0,1) = rsPedit("Parent")
		  If len(rsPedit("mnuAdd") & "x") = 1 Then
	  	   arrMenu(0,2) = rsPedit("Name")
		  else
	  	   arrMenu(0,2) = rsPedit("mnuAdd")
		  end if
		  arrMenu(0,3) = rsPedit("Link")
		  arrMenu(0,4) = rsPedit("mnuImage")
		  arrMenu(0,5) = rsPedit("onclick")
		  arrMenu(0,6) = rsPedit("Target")
		  arrMenu(0,7) = rsPedit("mnuOrder")
		  arrMenu(0,8) = rsPedit("mnuTitle")
		  arrMenu(0,9) = rsPedit("Linkcnt")
		  arrMenu(0,10) = rsPedit("ParentID")
		  arrMenu(0,11) = rsPedit("INAME")
		  arrMenu(0,12) = rsPedit("mnuFunction")
		  arrMenu(0,13) = rsPedit("mnuAdd")
		  arrMenu(0,14) = rsPedit("app_id")
		  arrMenu(0,15) = rsPedit("mnuAccess")
		   %>
            </td>
          </tr>
          <tr valign="top"> 
            <td colspan="2"> 
              <% call editLinkForm(rsPedit) %>
            </td>
          </tr>
          <% call UpdateOrderForm(arrMenu,0)
			 call addChildLinkForm(arrMenu,0)  'sub starts with <tr><td>
		  	 call showChildrenForm(arrMenu,0)		
  rsPedit.MoveNext
loop 
rsPedit.close %>
        </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr align="center" noshade />
    </td>
  </tr>
  <tr align="center"> 
    <td colspan="2">
        <hr align="center" noshade />
      </td>
  </tr>
</table><br /><br />
      <!--/form-->
</div>
<%
end sub

sub viewCode() %>
  <span class="submenu" id="viewCode">
  <table width="400" border="1" cellpadding="5" cellspacing="0" bordercolor="#FFFFFF" class="tCellAlt1">
    <tr> 
      <td align="center"><span class="fTitle"><b>Menu Code</b>
	  <br /></span></td>
	</tr>
    <tr> 
      <td align="center"><span class="fTitle"></span></td>
	</tr>
    <tr> 
      <td align="left">
	  <span style="text-align:left;">
	  &lt;%<br />
	  mnu.menuName = "<%= Menu %>"<br />
	  mnu.template = 4<br />
	  mnu.thmBlk = 1<br />
	  mnu.title = "Template 4"<br />
	  mnu.shoExpanded = 0<br />
	  mnu.canMinMax = 1<br />
	  mnu.keepOpen = 1<br />
	  mnu.GetMenu()<br />
	  %&gt;</span></td>
	</tr>
    <tr> 
      <td align="center"></td>
	</tr>
    <tr> 
      <td align="center">
        <input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= mTitle %>">
        <input name="CINAME" type="hidden" id="CINAME" value="<%= PINAME %>">
        <input name="Cparent" type="hidden" id="Cparent" value="<%= Menu %>">
        <input name="CappID" type="hidden" id="CappID" value="<%= app_id %>">
	  </td>
	</tr>
  </table></span>
  <%
end sub

sub addMenuForm() %>
<!-- ############## Start Add New Menu Form ############### -->
  <form id="AddMenus" name="AddMenus" method="post" action="admin_menu.asp" onSubmit="return CheckMenu('AddMenus');">
  <span class="submenu" id="addmnu">
  <table width="500" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" class="tCellAlt1">
    <tr> 
      <td align="center"><span class="fAlert"><b><%= txtCreateMnu %></b></span></td>
      <td colspan="5" align="center"><span class="fAlert"><b>*</b></span> <b><span class="fAlert"><%= txtMnuNam %>:</span>&nbsp; 
        <input name="mnuTitle" type="text" id="mnuTitle" value="" size="15">
      </td>
    </tr>
    <tr> 
      <td width="80" align="center"><b></b></td>
      <td width="74" align="center"><span class="fAlert"><b>* </b></span> <%= txtName %></td>
      <td width="74" align="center"><%= txtMnuLnk %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('link')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
      <td width="74" align="center"><%= txtMnuImg %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('image')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
      <td width="74" align="center"><%= txtMnuOc %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('onclick')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"></td>
      <td width="65" align="center"><%= txtMnuTarg %></td>
    </tr>
    <tr> 
      <td align="right"> 
        <input name="Corder" type="hidden" id="Corder" value="1">
        <%= txtMnuLnk %>:&nbsp; </td>
      <td align="center"> 
        <input class="textbox" name="Cname" type="text" id="Cname" size="12" value="">
      </td>
      <td align="center"> 
        <input class="textbox" name="Clink" type="text" id="Clink" size="12" value="">
      </td>
      <td align="center"> 
        <input class="textbox" name="CImage" type="text" id="CImage" size="12" value="">
      </td>
      <td align="center"> 
        <input class="textbox" name="Conclick" type="text" id="Conclick" size="12" value="">
      </td>
      <td align="center"> 
        <select class="textbox" name="Ctarget">
          <option value="_parent" selected="selected"><%= txtCurrent %></option>
          <option value="_blank"><%= txtNew %></option>
          <option value="_search"><%= txtSearch %></option>
        </select>
      </td>
    </tr>
        <tr> 
          <td align="right" colspan="2" valign="middle"><%= txtMnuFct %>: </td>
          <td align="left" colspan="4"> 
	<% If bFso Then %>
            <input class="textbox" name="Cfunct" type="text" id="Cfunct" size="12" value="">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('Function')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% Else %>
        <input type="hidden" name="Cfunct" id="Cfunct" value="">
	<% End If %>
		  </td>
        </tr>
    <tr> 
      <td height="33" align="right"><span class="fAlert"><b>*</b>&nbsp;<%= txtRequired %></span></td>
      <td align="center">&nbsp;</td>
      <td colspan="3" align="center"> 
        <input class="button" type="submit" name="Submit" value="<%= txtCreateMnu %>">
        <input name="cmd" type="hidden" id="cmd" value="addmenu">
      </td>
      <td align="center">&nbsp;</td>
    </tr>
  </table>
                </span> 
              </form>
<%
'<!-- ############## End Add New Menu ######################## -->
end sub

sub editLinkForm(arr)
'<!-- ############## Start Edit Parent Link ################### -->
 %>
              <form id="EditParent<%= ed %>" name="EditParent<%= ed %>" method="post" action="admin_menu.asp" onSubmit="selectAll('EditParent<%= ed %>','g_read');">
                <span class="submenu" id="parent<%= ed %>"> 
                
  <table width="95%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" class="tCellAlt1">
    <tr> 
      <td align="center"><span class="fAlert"><b><%= txtEditLnk %></b></span></td>
      <td width="74" align="center"><span class="fAlert"><b>*</b></span> <%= txtName %></td>
      <td width="74" align="center"><%= txtMnuLnk %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('link')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
      <td width="74" align="center"><%= txtMnuImg %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('image')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
      <td width="74" align="center"><%= txtMnuOc %> <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('onclick')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
      <td width="65" align="center"><%= txtMnuTarg %></td>
    </tr>
    <tr> 
      <td align="right"> 
        <input class="textbox" name="Pname" type="hidden" id="Pname" value="<%= arr("Name") %>">
      </td>
      <td align="center"> 
        <input class="textbox" name="PNname" type="text" id="PNname" size="12" value="<%= replace(arr("Name"),"''","'") %>">
      </td>
      <td align="center"> 
        <input class="textbox" name="Plink" type="text" id="Plink" size="12" value="<%= arr("Link") %>">
      </td>
      <td align="center"> 
        <input class="textbox" name="PImage" type="text" id="PImage" size="12" value="<%= arr("mnuImage") %>">
      </td>
      <td align="center"> 
        <input class="textbox" name="Ponclick" type="text" id="Ponclick" size="12" value="<%= arr("onclick") %>">
      </td>
      <td align="center"> 
        <select class="textbox" name="Ptarget">
          <% If arr("Target") = "_blank" Then %>
          <option value="_parent"><%= txtCurrent %></option>
          <option value="_blank" selected="selected"><%= txtNew %></option>
          <option value="_search"><%= txtSearch %></option>
          <% ElseIf arr("Target") = "_parent" Then %>
          <option value="_parent" selected="selected"><%= txtCurrent %></option>
          <option value="_blank"><%= txtNew %></option>
          <option value="_search"><%= txtSearch %></option>
          <% ElseIf arr("Target") = "_search" Then %>
          <option value="_parent"><%= txtCurrent %></option>
          <option value="_blank"><%= txtNew %></option>
          <option value="_search" selected="selected"><%= txtSearch %></option>
          <% End If %>
        </select>
	<% If bFso Then %>
      </td>
    </tr>
        <tr> 
          <td align="right" colspan="2" valign="middle"><%= txtMnuFct %>: </td>
          <td align="left" colspan="4"> 
            <input class="textbox" name="Cfunct" type="text" id="Cfunct" size="15" value="<%= arr("mnuFunction") %>">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('Function')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% Else %>
        <input type="hidden" name="Cfunct" id="Cfunct" value="">
	<% End If %>
			</td>
        </tr>
        <tr> 
          <td align="center" colspan="6" valign="middle">
		<% mnuAccessBlock "EditParent" & ed & "",arr("mnuAccess") %>
		  </td>
        </tr>
    <tr align="center" valign="middle">
      <td height="33" colspan="6">
        <input class="button" type="submit" name="Submit" value="<%= txtUpdLnk %>: '<%= replace(arr("Name"),"''","'") %>'">
        <input name="cmd" type="hidden" id="cmd" value="editparent">
        <input name="id" type="hidden" id="id" value="<%= arr("id") %>">
        <input name="Porder" type="hidden" id="Porder" value="<%= arr("mnuOrder") %>">
        <input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= arr("mnuTitle") %>">
        <input name="CINAME" type="hidden" id="CINAME" value="<%= arr("INAME") %>">
        <input name="CaddMenu" type="hidden" id="CaddMenu" value="<%= arr("mnuAdd") %>">
      </td>
    </tr>
  </table>
                </span> 
              </form>
<%'<!-- ############################ End Edit Parent Link ########################## -->
 end sub

sub addLinkForm() %>
        <span class="submenu" id="addparent">
<form name="ParentAdd" id="ParentAdd" method="post" action="admin_menu.asp" onSubmit="selectAll('ParentAdd','g_read');">
  <table width="95%" cellpadding="3" cellspacing="0" border="1" bordercolor="#6E3019" class="tCellAlt1" align="center">
	<tr>
      <td>
  <table width="95%" cellpadding="2" cellspacing="0" border="1" bordercolor="#FFFFFF" align="center">
	<tr>
      <td colspan="2" align="center" height="35"><b><%= txtAddLnkMnu %>: '<span class="fAlert"><%= mTitle %></span>'</b>
      </td>
    </tr>
	<tr>
      <td width="45%" align="right"><span class="fAlert">* </span><%= txtName %>:&nbsp;</td>
      <td> 
        <input class="textbox" name="Cname" type="text" id="Cname" size="12" value="">
      </td>
    </tr>
	<tr>
      <td align="right"><%= txtMnuLnk %>:&nbsp;</td>
      <td> 
        <input class="textbox" name="Clink" type="text" id="Clink" size="12" value="">
		&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('link')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
      </td>
    </tr>
	<tr>
      <td align="right"><%= txtMnuImg %>:&nbsp;</td>
      <td> 
        <input class="textbox" name="CImage" type="text" id="CImage" size="12" value="">
		&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('image')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
      </td>
    </tr>
	<tr>
      <td align="right"><%= txtMnuOc %>:&nbsp;</td>
      <td> 
        <input class="textbox" name="Conclick" type="text" id="Conclick" size="12" value="">
		&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('onclick')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% If bFso Then %>
      </td>
    </tr>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuFct %>: </td>
          <td align="left"> 
            <input class="textbox" name="Cfunct" type="text" id="Cfunct" size="15" value="">
            &nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('Function')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% Else %>
        <input type="hidden" name="Cfunct" id="Cfunct" value="">
	<% End If %>
		  </td>
        </tr>
	<tr>
      <td align="right"><%= txtMnuTarg %>:&nbsp;</td>
      <td> 
        <select class="textbox" name="Ctarget">
          <option value="_parent" selected="selected"><%= txtCurrent %></option>
          <option value="_blank"><%= txtNew %></option>
          <option value="_search"><%= txtSearch %></option>
        </select>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('target')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
      </td>
    </tr>
	<tr>
      <td align="right"><%= txtMnuAdd %>:&nbsp;</td>
      <td> 
        <%= addList %>
		&nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('addmenu')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
      </td>
    </tr>
        <tr> 
          <td align="center" colspan="2" valign="middle">
		<% mnuAccessBlock "ParentAdd","" %>
		  </td>
        </tr>
    <tr>
      <td colspan="2" align="center"><p>
        <input class="button" type="submit" name="Submit" value="<%= txtMnuAddLnk %>">
        <input name="Corder" type="hidden" id="Corder" value="<%= CntLinks + 1 %>">
        <input name="cmd" type="hidden" id="cmd" value="addparent">
        <input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= mTitle %>">
        <input name="CINAME" type="hidden" id="CINAME" value="<%= PINAME %>">
        <input name="Cparent" type="hidden" id="Cparent" value="<%= Menu %>">
        <input name="CappID" type="hidden" id="CappID" value="<%= app_id %>">
		</p>
				  <% 
				  if m_debug then
				  Response.Write("<br /><br />cmd: addparent")
				  Response.Write("<br />mnuTitle: " & mTitle)
				  Response.Write("<br />Cparent: " & Menu)
				  Response.Write("<br />CINAME: " & PINAME)
				  Response.Write("<br />CappID: " & app_id)
				  Response.Write("<br />Corder: " & CntLinks + 1)
				  end if
				   %>
      </td>
    </tr>
  </table>
      </td>
    </tr>
  </table>
  </form>
  <% 'spThemeBlock1_close(intSkin) %>
  </span> 
<% end sub

sub addChildLinkForm(nm,ct) %>
<!-- ############################# ADD SUB LINK FORM ############################### -->
          <tr> 
            <td colspan="4" align="center"> 
                <span class="submenu" id="addchild<%= nm(ct,0) %>"> 
              <form id="ChildAdd<%= nm(ct,0) %>" name="ChildAdd<%= nm(ct,0) %>" method="post" action="admin_menu.asp" onSubmit="selectAll('ChildAdd<%= nm(ct,0) %>','g_read');">
      <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#6E3019" class="tCellAlt1">
        <tr align="left"> 
          <td colspan="2" nowrap>
            <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
              <tr> 
                <td height="30" colspan="2" align="center" class="fSubTitle"><span class="fAlert"><b><%= txtMnuAddSubLnk %></b></span></td>
              </tr>
              <tr> 
                <td width="30%" align="right"><span class="fAlert"><b>* </b></span><%= txtName %>:&nbsp;</td>
                <td> 
                  <input class="textbox" name="Cname" type="text" id="Cname" size="15" value="">
                </td>
              </tr>
              <tr> 
                <td align="right" valign="middle"><%= txtMnuLnk %>:&nbsp; </td>
                <td> 
                  <input class="textbox" name="Clink" type="text" id="Clink" size="15" value="">
                  &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('link')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
              </tr>
              <tr> 
                <td align="right" valign="middle"><%= txtMnuImg %>:&nbsp; </td>
                <td> 
                  <input class="textbox" name="CImage" type="text" id="CImage" size="15" value="">
                  &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('image')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
              </tr>
              <tr> 
                <td align="right" valign="middle"><%= txtMnuOc %>:&nbsp; </td>
                <td> 
                  <input class="textbox" name="Conclick" type="text" id="Conclick" size="15" value="">
                  &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('onclick')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
              </tr>
	<% If bFso Then %>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuFct %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="Cfunct" type="text" id="Cfunct" size="15" value="">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('Function')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% Else %>
        <input type="hidden" name="Cfunct" id="Cfunct" value="">
			</td>
        </tr>
	<% End If %>
              <tr> 
                <td align="right" valign="middle"><%= txtMnuTarg %>:&nbsp;</td>
                <td width="145"> 
                  <select class="textbox" name="Ctarget">
                    <option value="_parent" selected="selected"><%= txtCurrent %></option>
                    <option value="_blank"><%= txtNew %></option>
          			<option value="_search"><%= txtSearch %></option>
                  </select>
				  &nbsp;&nbsp;&nbsp;&nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('target')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
                </td>
              </tr>
        <tr> 
          <td align="center" colspan="2" valign="middle">
		<% mnuAccessBlock "ChildAdd" & nm(ct,0) & "","" %>
		  </td>
        </tr>
              <tr> 
                <td height="50" colspan="2" align="center" valign="middle"> 
                  <input class="button" type="submit" name="Submit" value="<%= txtAddChldLnk %>: '<%= nm(ct,2) %>'">
                  <input name="Corder" type="hidden" id="Corder" value="1">
                  <input name="Cparent" type="hidden" id="Cparent" value="<%= nm(ct,2) %>">
                  <input name="CparentID" type="hidden" id="CparentID" value="<%= nm(ct,0) %>">
                  <input name="cmd" type="hidden" id="cmd" value="addchild">
                  <input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= nm(ct,8) %>">
                  <input name="CINAME" type="hidden" id="CINAME" value="<%= nm(ct,11) %>">
                  <input name="CaddMenu" type="hidden" id="CINAME" value="<%= nm(ct,13) %>">
        		  <input name="CMenu" type="hidden" id="CMenu" value="<%= menu %>">
				  <% 
				  if m_debug then
				  Response.Write("<br /><br />cmd: addchild")
				  Response.Write("<br />Cparent: " & nm(ct,2))
				  Response.Write("<br />CparentID: " & nm(ct,0))
				  Response.Write("<br />mnuTitle: " & nm(ct,8))
				  Response.Write("<br />CINAME: " & nm(ct,11))
				  Response.Write("<br />CappID: " & nm(ct,14))
				  Response.Write("<br />CaddMenu: " & nm(ct,13))
				  Response.Write("<br />Menu: " & menu)
				  end if
				   %>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
	  </form>
	  </span>
	  </td>
	  </tr>
<!-- ############################# END: ADD SUB LINK FORM ################################### -->
<% end sub

' ############################# SHOW SUB LINKS ########################################
sub showChildrenForm(nam,ct)
  if nam(ct,13) <> "" then
	set rsCedit = my_Conn.Execute("SELECT *,(select count(*) from Menu where Parent ='" & nam(ct,13) & "' and INAME = '" & nam(ct,13) & "') as subCnt from Menu Where Parent ='" & nam(ct,13) & "' and INAME = '" & nam(ct,13) & "' order by mnuOrder asc")
  else
	set rsCedit = my_Conn.Execute("SELECT *,(select count(*) from Menu where Parent ='" & nam(ct,2) & "' and INAME = '" & nam(ct,11) & "') as subCnt from Menu Where Parent ='" & nam(ct,2) & "' and INAME = '" & nam(ct,11) & "' order by mnuOrder asc")
  end if
	if not rsCedit.eof then 
	  fa = 0
	  reDim arrSub(clng(rsCedit("subCnt"))-1,16)
	  do while not rsCedit.eof
	  		if nam(ct,13) <> "" then
			  set rsCcnt = my_Conn.Execute("select count(*) from Menu where Parent ='" & nam(ct,13) & "' and INAME = '" & nam(ct,13) & "'")
			else
			  set rsCcnt = my_Conn.Execute("select count(*) from Menu where Parent ='" & rsCedit("Name") & "' and INAME = '" & nam(ct,11) & "'")
			end if
			arrSub(fa,16) = rsCcnt(0)
			set rsCcnt = nothing
			
			arrSub(fa,0) = rsCedit("id")
			arrSub(fa,1) = rsCedit("Parent")
		  	If len(rsCedit("mnuAdd") & "x") = 1 Then
	  	   	  arrSub(fa,2) = rsCedit("Name")
		  	else
	  	   	  arrSub(fa,2) = rsCedit("mnuAdd")
		  	end if
			arrSub(fa,2) = rsCedit("Name")
			arrSub(fa,3) = trim(rsCedit("Link"))
			arrSub(fa,4) = rsCedit("mnuImage")
			arrSub(fa,5) = trim(rsCedit("onclick"))
			arrSub(fa,6) = rsCedit("Target")
			arrSub(fa,7) = rsCedit("mnuOrder")
			arrSub(fa,8) = rsCedit("mnuTitle")
			arrSub(fa,9) = rsCedit("subCnt")
			if trim(arrSub(fa,9)) = "" then
				arrSub(fa,9) = 0
			end if
			arrSub(fa,10) = rsCedit("ParentID")
			arrSub(fa,11) = rsCedit("INAME")
			arrSub(fa,12) = rsCedit("mnuFunction")
		    arrSub(fa,13) = trim(rsCedit("mnuAdd"))
		    arrSub(fa,14) = rsCedit("app_id")
		    arrSub(fa,15) = rsCedit("mnuAccess")
	  fa = fa + 1
	  rsCedit.movenext
	  loop
	  rsCedit.close
'	  set rsCredit = nothing
%>
    <tr> 
      <td colspan="4" align="right">
		<span class="submenu" id="child<%= nam(ct,0) %>">
		<% esl = nam(ct,0)
		   if c_color = "tCellAlt0" then
		     c_color = "tCellAlt1"
		   else
		     c_color = "tCellAlt0"
		   end if %>
          <div id="editsublink<%= nam(ct,0) %>">
      <table width="95%" align="right" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" class="<%= c_color %>">
        <% 'cnt = 0
'			sc = sc + 1
		for cnt = 0 to uBound(arrSub)
'			cnt = cnt + 1 
		%>
        <tr> 
          <td valign="baseline" height="18">
            &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_bar.gif" width="15" height="15" align="middle" alt=""/>&nbsp;<input class="textbox" name="Cname<%= cnt+1 %>" type="text" id="Cname<%= cnt+1 %>" size="12" value="<%= arrSub(cnt,2) %>" readonly="true">
            &nbsp;&nbsp;<input class="textbox" name="Clink<%= cnt+1 %>" type="text" id="Clink<%= cnt+1 %>" size="12" value="<%= arrSub(cnt,3) %>" readonly="true">
            &nbsp;&nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_trashcanL.gif" onclick="mode('Cdelete','','<%= arrSub(cnt,0) %>','<%= Menu %>')" alt="<%= txtDelLnk %>" title="<%= txtDelLnk %>" width="15" height="15" border="0" style="cursor:pointer;"/> 
		<% 	If len(arrSub(cnt,13) & "x") = 1 Then %>
            &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_pencilL.gif" onclick="SwitchMenu('editsublink<%= esl %>','edit<%= arrSub(cnt,0) %>')" border="0" width="15" height="15" alt="<%= txtEditLnk %>" title="<%= txtEditLnk %>: '<%= arrSub(0,2) %>'" style="cursor:pointer;"/> 
		<% End If %>
         <% If len(arrSub(cnt,3)) = 0 and len(arrSub(cnt,5)) = 0 Then %>
            &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkP.gif" onclick="SwitchMenu('editsublink<%= esl %>','addchild<%= arrSub(cnt,0) %>')" border="0" width="15" height="15" alt="<%= txtMnuAddSubLnk %>" title="<%= txtMnuAddSubLnk %>" style="cursor:pointer;"/>
			<% If cLng(arrSub(cnt,16)) >= 1 Then %>
            &nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_childYL.gif" onclick="SwitchMenu('editsublink<%= esl %>','child<%= arrSub(cnt,0) %>')" border="0" width="15" height="15" alt="<%= txtView %>" title="<%= txtVECLnks %>&nbsp;'<%= arrSub(cnt,2) %>'" style="cursor:pointer;">
				<% If cLng(arrSub(cnt,16)) > 1 Then %>
            	&nbsp;<img src="themes/<%= strTheme %>/mnu_icons/icon_linkA.gif" onclick="SwitchMenu('editsublink<%= esl %>','updateOrder<%= arrSub(cnt,0) %>')" border="0" width="15" height="15" alt="<%= txtUpdOrd %>" title="<%= txtUpdOrd %>:&nbsp;'<%= arrSub(cnt,2) %>' " style="cursor:pointer;">
				<% End If
               End if
			End If %>
          </td>
        </tr>
        <tr> 
          <td align="right"> 
              <table align="right" width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#000000">
              <% call addChildLinkForm(arrSub,cnt)
		   		 call showEditLinkForm(arrSub,cnt)
        		if trim(arrSub(cnt,3)) = "" and cLng(arrSub(cnt,16)) > 0 then
		   		 call showChildrenForm(arrSub,cnt)
				end if
				 if cLng(arrSub(cnt,16)) > 1 then
				   call UpdateOrderForm(arrSub,cnt)
				 end if %>
              </table>
          </td>
        </tr>
        <% 		
		 next %>
      </table>
            </div><br /></span>
            </td>
          </tr>
<%
	end if 
' ############################# END: SHOW SUB LINK FORM ###################################
end sub

sub UpdateOrderForm(ord,ct)
  if not isArray(ord) then
	sSql = "SELECT *,(select count(*) from Menu where Parent ='" & menu & "' and INAME = '" & menu & "') as subCnt from Menu Where Parent ='" & menu & "' and INAME = '" & menu & "' order by mnuOrder asc"
  else
    if ord(ct,13) <> "" then
	  sSql = "SELECT *,(select count(*) from Menu where Parent ='" & ord(ct,2) & "' and INAME = '" & ord(ct,13) & "') as subCnt from Menu Where Parent ='" & ord(ct,2) & "' and INAME = '" & ord(ct,13) & "' order by mnuOrder asc"
	else
	  sSql = "SELECT *,(select count(*) from Menu where Parent ='" & ord(ct,2) & "' and INAME = '" & ord(ct,11) & "') as subCnt from Menu Where Parent ='" & ord(ct,2) & "' and INAME = '" & ord(ct,11) & "' order by mnuOrder asc"
	end if
  end if
  set rsUpOrd = my_Conn.Execute(sSql)
	if not rsUpOrd.eof then 
	  up = 0
	  reDim upOrd(cLng(rsUpOrd("subCnt"))-1,12)
	  do while not rsUpOrd.eof
			upOrd(up,0) = rsUpOrd("id")
			upOrd(up,1) = rsUpOrd("Parent")
			upOrd(up,2) = rsUpOrd("Name")
			upOrd(up,3) = rsUpOrd("Link")
			upOrd(up,4) = rsUpOrd("mnuImage")
			upOrd(up,5) = rsUpOrd("onclick")
			upOrd(up,6) = rsUpOrd("Target")
			upOrd(up,7) = rsUpOrd("mnuOrder")
			upOrd(up,8) = rsUpOrd("mnuTitle")
			upOrd(up,9) = rsUpOrd("subCnt")
			upOrd(up,10) = rsUpOrd("ParentID")
			upOrd(up,11) = rsUpOrd("INAME")
			upOrd(up,12) = rsUpOrd("mnuFunction")
	  up = up + 1
	  rsUpOrd.movenext
	  loop
	  rsUpOrd.close
'	  set rsUpOrd = nothing
		
%>
          <tr> 
            <td colspan="4" align="center">
			<% If not isArray(ord) Then %>
                <span class="submenu" id="updateOrder">
              <form name="UpdatOrder" method="post" action="admin_menu.asp">
			<% Else %>
                <span class="submenu" id="updateOrder<%= ord(ct,0) %>">
              <form name="UpdatOrder<%= ord(ct,0) %>" method="post" action="admin_menu.asp">
			<% End If %>               
      <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#6E3019" class="tCellAlt1">
        <tr> 
          <td colspan="2" nowrap>
            <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#ffffff">
              <tr align="center"> 
                <td height="30" colspan="2" nowrap><span class="fAlert"><b><%= txtUpdOrd %></b></span></td>
              </tr>
			  <% for uo = 0 to ubound(upOrd) %>
              <tr> 
                <td width="30%" align="center" valign="middle"> 
                  <input class="textbox" name="Cname<%= uo+1 %>" type="text" id="Cname<%= uo+1 %>" size="18" value="<%= upOrd(uo,2) %>">
                </td>
                <td align="left">
                  <select name="mnuOrder<%= uo+1 %>">
                    <% for Ucnt = 1 to cLng(upOrd(0,9))
						if Ucnt = upOrd(uo,7) then %>
                    <option value="<%= Ucnt %>" selected="selected"><%= Ucnt %></option>
                    <% Else %>
                    <option value="<%= Ucnt %>"><%= Ucnt %></option>
                    <% End If
						next %>
                  </select>
                  <input name="id<%= uo+1 %>" type="hidden" id="id<%= uo+1 %>" value="<%= upOrd(uo,0) %>">
                </td>
              </tr>
			  <% next %>
              <tr align="center"> 
                <td height="50" colspan="2"> 
                  <input class="button" type="submit" name="Submit" value="<%= txtUpdOrd %>">
                  <input name="count" type="hidden" id="count" value="<%= upOrd(0,9) %>">
                  <input name="ParentID" type="hidden" id="ParentID" value="<%= upOrd(0,10) %>">
                  <input name="cmd" type="hidden" id="cmd" value="updateOrder">
                  <input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= upOrd(0,8) %>">
        		  <input name="CINAME" type="hidden" id="CINAME" value="<%= upOrd(0,11) %>">
        		  <input name="Cfunct" type="hidden" id="Cfunct" value="<%= upOrd(0,12) %>">
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
              </form>
                </span> 
            </td>
          </tr>
<% end if
end sub

sub showEditLinkForm(nam,ct)
%>
          <tr> 
            <td colspan="4" align="center">
              <form id="ChildEdit<%= nam(ct,0) %>" name="ChildEdit<%= nam(ct,0) %>" method="post" action="admin_menu.asp" onSubmit="selectAll('ChildEdit<%= nam(ct,0) %>','g_read');">
                <span class="submenu" id="edit<%= nam(ct,0) %>"> 
                
      <table width="90%" border="1" cellpadding="0" cellspacing="0" bordercolor="#6E3019" class="tCellAlt1">
        <tr> 
          <td colspan="2" nowrap>
      <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#ffffff">
        <tr align="center"> 
          <td height="30" colspan="2" class="fSubTitle"><span class="fAlert"><b><%= txtEditLnk %></b></span></td>
        </tr>
        <tr> 
          <td width="30%" align="right" valign="middle"><span class="fAlert"><b>* </b></span><%= txtName %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="Cname" type="text" id="Cname" size="15" value="<%= nam(ct,2) %>">
          </td>
        </tr>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuLnk %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="Clink" type="text" id="Clink" size="15" value="<%= nam(ct,3) %>">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('link')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/> 
          </td>
        </tr>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuImg %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="CImage" type="text" id="CImage" size="15" value="<%= nam(ct,4) %>">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('image')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
        </tr>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuOc %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="Conclick" type="text" id="Conclick" size="15" value="<%= replace(nam(ct,5),"''","'") %>">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('onclick')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></td>
        </tr>
	<% If bFso Then %>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuFct %>:&nbsp;</td>
          <td align="left"> 
            <input class="textbox" name="Cfunct" type="text" id="Cfunct" size="15" value="<%= nam(ct,12) %>">
            &nbsp; <img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('Function')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/>
	<% Else %>
        <input type="hidden" name="Cfunct" id="Cfunct" value="">
			</td>
        </tr>
	<% End If %>
        <tr> 
          <td align="right" valign="middle"><%= txtMnuTarg %>: </td>
          <td align="left"> 
            <select name="Ctarget">
              <% If nam(ct,6) = "_blank" Then %>
              <option value="_parent"><%= txtCurrent %></option>
              <option value="_blank" selected="selected"><%= txtNew %></option>
          	  <option value="_search"><%= txtSearch %></option>
          	  <% ElseIf nam(ct,6) = "_parent" Then %>
          	  <option value="_parent" selected="selected"><%= txtCurrent %></option>
          	  <option value="_blank"><%= txtNew %></option>
          	  <option value="_search"><%= txtSearch %></option>
          	  <% ElseIf nam(ct,6) = "_search" Then %>
          	  <option value="_parent"><%= txtCurrent %></option>
          	  <option value="_blank"><%= txtNew %></option>
          	  <option value="_search" selected="selected"><%= txtSearch %></option>
              <% End If %>
            </select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('target')" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/> 
          </td>
        </tr>
        <tr> 
          <td align="center" colspan="2" valign="middle">
		<% mnuAccessBlock "ChildEdit" & nam(ct,0) & "",nam(ct,15) %>
		  </td>
        </tr>
        <tr align="center"> 
          <td height="50" colspan="2"> 
            <input class="button" type="submit" name="Submit" value="<%= txtUpdLnk %>">
            <input name="ParentID" type="hidden" id="ParentID" value="<%= nam(ct,10) %>">
            <input name="id" type="hidden" id="id" value="<%= nam(ct,0) %>">
            <input name="cmd" type="hidden" id="cmd" value="editchild">
        	<input name="mnuTitle" type="hidden" id="mnuTitle" value="<%= nam(ct,8) %>">
        	<input name="CINAME" type="hidden" id="CINAME" value="<%= nam(ct,11) %>">
        	<input name="CaddMenu" type="hidden" id="CaddMenu" value="<%= nam(ct,13) %>">
        	<input name="CMenu" type="hidden" id="CMenu" value="<%= menu %>">
				  <% 
				  if m_debug then
				  Response.Write("<br /><br />cmd: editchild")
				  Response.Write("<br />Cparent: " & nam(ct,2))
				  Response.Write("<br />CparentID: " & nam(ct,10))
				  Response.Write("<br />mnuTitle: " & nam(ct,8))
				  Response.Write("<br />CINAME: " & nam(ct,11))
				  Response.Write("<br />CappID: " & nam(ct,14))
				  Response.Write("<br />CaddMenu: " & nam(ct,13))
				  Response.Write("<br />Menu: " & menu)
				  end if
				   %>
          </td>
        </tr>
      </table>
          </td>
        </tr>
      </table>
                </span> 
              </form>
            </td>
          </tr>
<%
end sub

sub mnuAccessBlock(f,g) %>
  <fieldset style='margin:10px;'>
  <legend><b>Group Access</b>&nbsp;&nbsp;<img src="<%= icnHelp %>" width="15" height="15" onclick="popHelp('grpAccess');" alt="<%= txtHelp %>" title="<%= txtHelp %>" style="cursor:pointer;"/></legend><br />
  <table border=0 cellpadding=0 cellspacing=0>
	  <!-- <tr><td colspan="2" align="center"><b>Groups with access</b></td></tr> -->
  	  <tr><td align="right" valign="middle" width="50%" nowrap>
  		<a href="JavaScript:allowgroups('<%= f %>','g_read','<%= gLst %>');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:removeGroup('<%= f %>','g_read');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:eGroup('<%= f %>','g_read');" title="<%= txtCM10 %>"><b><%= txtEditGrp %></b></a>&nbsp;&nbsp;<br />&nbsp;
		</td>
          <td align="left"><p>
            <select size="5" name="g_read" style="width:120;" multiple>
			  <% if trim(g) <> "" then
			  		getOptGroups(g)
				 end if %>
			  <option value="0"></option>
            </select><br />&nbsp;</p>
          </td>
        </tr>
  </table>
  </fieldset>
<%
end sub

%>