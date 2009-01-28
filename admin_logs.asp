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
sURL = "files/sp_logs/"

%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then

	select case request("mode")
	  case "delcur"
	    'Response.Write "Delete current: " & sURL & request("log")
		set oFso = new clsSFSO
		sFileToDelete = server.MapPath(sURL & request("log"))
		oFso.DeleteFile(sFileToDelete)
		set oFso = nothing
		closeAndGo("admin_logs.asp")
	  case "delall"
	    Response.Write "Delete All"
		set oFso = new clsSFSO
		sFolderToDelete = server.MapPath(sURL)
		oFso.DeleteFolder(sFolderToDelete)
		set oFso = nothing
		call writeToLog("","","All logs deleted")
		closeAndGo("admin_logs.asp")
	end select

 %>
<script type="text/javascript">
var cFile = ""
function shoLogfile(idx){
if (document.getElementById){
document.getElementById("tabiframe").src="files/sp_logs/" + idx;
cFile = idx;
return false;
}
else
return true;
}
function delCurLog(){
  var t = document.getElementById("logview").options.value;
  //alert("Delete Current\n" + t);
var stM
stM = "Are you sure you want to\ndelete this log file?\n";
stM += "\n" + t + "\n\nThis cannot be undone\n";
var del=confirm(stM);
//alert(del);
if (del==true){
  window.location='admin_logs.asp?mode=delcur&log=' + t;
  }
}
function delAllLog(){
var stM
stM = "Are you sure you want to\ndelete ALL log files?\n";
stM += "\nThis cannot be undone\n";
var del=confirm(stM);
//alert(del);
if (del==true){
  window.location='admin_logs.asp?mode=delall';
  }
}
</script>
<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
<tr><td class="leftPgCol">
<% intSkin = getSkin(intSubSkin,1) %>
<% 
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
	menu_admin()
spThemeBlock1_close(intSkin) %>
</td>
<td class="mainPgCol">
<% intSkin = getSkin(intSubSkin,2) %>
<%
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = "Log files"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6

set oFso = new clsSFSO
sFile = ""
spThemeTitle = "Site Log Files"
spThemeBlock1_open(intSkin)
pURL = server.MapPath(sURL)
allFinfo = oFso.GetAllFolderInformation(pURL)

if sFile = "" and isArray(allFinfo) then
  sFile = allFinfo(0).name
end if
'Response.Write "cFile: " & cFile
'Response.Write(ubound(allFinfo))
Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""4"" width=""100%"" align=""center"">"
Response.Write "<tr><td align=""center"" colspan=""4"">"
Response.Write "<input id=""delcur"" name=""delcur"" type=""button"" onclick=""delCurLog();"" value=""Delete Current Log"" />&nbsp;&nbsp;&nbsp;"
Response.Write "<select id=""logview"" name=""logview"" onChange=""shoLogfile(this.options[this.selectedIndex].value)"">"
if isArray(allFinfo) then
  for x = 0 to ubound(allFinfo)
    Response.Write "<option value=""" & allFinfo(x).name & """" & chkSelect(sFile,allFinfo(x).name) & ">"
    Response.Write allFinfo(x).name
    Response.Write "</option>"
  next
else
    Response.Write "<option value="""">No logs found"
    Response.Write "</option>"
end if
Response.Write "</select>&nbsp;&nbsp;&nbsp;"
Response.Write "<input name=""delall"" type=""button"" onClick=""delAllLog();"" value=""Delete All Logs"" />"
Response.Write "</td></tr>"
Response.Write "<tr><td align=""left"" colspan=""4"">"


Response.Write "</td></tr>"
Response.Write "</table>"
':: display log file
sURL = "files/sp_logs/" & sFile
'pURL = server.MapPath(sURL)
'oFso.WriteTextFile pURL,Date()
'oFso.AppendTextFile server.MapPath(oFso.LogFile),"[]"
oFso.DisplayLog sURL,"300","600"
 
spThemeBlock1_close(intSkin)
set oFso = nothing %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %><% Response.Redirect "admin_login.asp?target=admin_logs.asp" %><% end if %>