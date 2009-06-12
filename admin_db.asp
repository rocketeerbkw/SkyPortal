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
<!--#include file="includes/inc_admin_functions.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" and intIsSuperAdmin and strDBType = "access" Then %>
<% 
 'on error resume next
 fBackup = False
 bError = False
 'dbOrigFile = server.mappath("db/skyportal.mdb")
 dbOrigFile = strDBPath 
 'sBackUpFile = left(dbOrigFile, InStr(dbOrigFile, ".mdb")-1) & "_" & strCurDateString & ".bak"
 sBackUpFile = left(dbOrigFile, InStr(dbOrigFile, ".mdb")-1) & "_" & strCurDateString & ".mdb"
 sBackUpFile1 = left(dbOrigFile, InStr(dbOrigFile, ".mdb")-1) & "_" & strCurDateString & ".bak"
 sFileTmp = left(dbOrigFile, InStr(dbOrigFile, ".mdb")-1) & "_" & strCurDateString & ".comp.mdb"
 sFileTmp1 = left(dbOrigFile, InStr(dbOrigFile, ".mdb")-1) & "_" & strCurDateString & ".comp.bak"

 
 'sCompFrom = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & dbOrigFile
 sCompFrom = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sBackUpFile
 sCompTo   = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & sFileTmp
 %>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr><td width="190" class="leftPgCol">
<% 
	intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
  		'bannerConfigMenu("1")
  		response.Write("<hr />")
  		menu_admin()
spThemeBlock1_close(intSkin) %>
<script type="text/javascript">
function hideme(obj){
 if (document.getElementById(obj)){
  var el = document.getElementById(obj);
  el.style.display='none';
  return;
 } 
}
</script>
</td>
<td class="mainPgCol" valign="top">
<%
	intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtDBmgr & "|admin_db.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
<% 
spThemeTitle = txtDBmgr
spThemeBlock1_open(intSkin)
%>
<table width="100%">
<tr><td width="100%" valign="top"><br />
<% 
   set my_Conn = nothing
if lcase(strDBType) = "access" then
 if request("cmd")="1" and bFso then
  if intIsSuperAdmin then
   bakupFile dbOrigFile,sBackUpFile
   'bakupFile sBackUpFile,sBackUpFile1
   If fBackup = true Then
    set jro = server.createobject("jro.JetEngine")
    jro.CompactDatabase sCompFrom, sCompTo
   
    if err <> 0 then
     bError = True
     response.write "<div class=""fSubTitle"">Jet Error:<br />" & err.description & "</div>"
	else
     'response.write "<br />Jet: Compact complete. No error<br />"
    end if
	set jro = nothing
	if not bError then
      swapComp()
	else
      'response.write "<div class=""fSubTitle"">fso: unknown error</div>"
	end if
	'err = 0
   end if
   
   
   if err <> 0 then
    bError = True
    'response.write "<div class=""fSubTitle"">X Error:</div>" & err.description
   end if
   
	if fBackup then
     response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupSucc & "</div>"
	else
     response.write "<br /><div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupNoSucc & "</div>"
	end if
	
   if not bError then
    response.write "<br /><div class=""fSubTitle"" style=""text-align:center;"">" & txtDBCompSucc & "</div>"
   else
    response.write "<br /><div class=""fSubTitle"" style=""text-align:center;"">" & txtDBCompNoSucc & "</div>"
   end if
   
  else
   response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBNoPerm & "</div>"
  end if
  response.Write("<hr />")
  ShowForm
 elseif request("cmd")="2" and bFso then
   if intIsSuperAdmin then
     bakupFile dbOrigFile,sBackUpFile1
     if fBackup then
       response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupSucc & "</div>"
     else
       response.write "<br /><div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupNoSucc & "</div>"
     end if
   else
     response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBNoPerm & "</div>"
   end if
  
   response.Write("<hr />")
   ShowForm
 elseif request("cmd")="3" and bFso then
   if intIsSuperAdmin then
	 eFile = request("db")
     Dim fso
     Set fso = CreateObject("Scripting.FileSystemObject")
	   if fso.FileExists(eFile) = true then
	     fso.DeleteFile eFile,true
	     if fso.FileExists(eFile) = true then
     		response.Write("<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupNoDel & "</div>")
		 else
     		response.Write("<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupDel & "</div>")
		 end if
	   else
     	 response.Write("<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBBkupNoFnd & "</div>")
	   end if
	 Set fso = nothing
   else
     response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBNoPerm & "</div>"
   end if
   response.Write("<hr />")
   ShowForm
 elseif bFso then
   ShowForm
 else
   response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBfsoNoFnd & "</div>"
 end if
else
   response.write "<div class=""fSubTitle"" style=""text-align:center;"">" & txtDBOnlyAccessDB & "</div>"
end if

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear
	my_Conn.Open strConnString
%>
</td></tr></table>
<%
spThemeBlock1_close(intSkin) %>
</td></tr>
</table>
<!--#include file="inc_footer.asp" -->
<% else %>
<% Response.Redirect "admin_home.asp" %>
<% end if %>
<% 
sub ShowForm
 %>
  <p><%= replace(txtDBPara1,"[%dbname%]","""<b>" & dbOrigFile & "</b>""") %></p>
  <p><%= txtDBPara2 %></p>
  <p><%= txtDBPara3 %></p>
   <p><%= txtDBClkOnce %></p>
   <p><div id="c_bkup" style="display:block;"><center>
  <a href="admin_db.asp?cmd=1" onclick="hideme('c_bkup');"><%= txtDBBakComp %></a>&nbsp;|&nbsp;
  <a href="admin_db.asp?cmd=2" onclick="this.href.value='javascript:;'"><%= txtDBBakOnly %></a></center></div></p><br />
  <hr />
  <div class="fSubTitle" style="padding:4px;"><%= txtDBCurDB %></div>
 <%
    Dim fso, f, fo
    Set fso = CreateObject("Scripting.FileSystemObject")
 	set f=fso.GetFile(dbOrigFile)
	dName = left(f.name, InStr(f.name, ".mdb")-1)
	dSize = f.Size
	'dName = right(dName, InStrrev(dName, "\"))
	Response.write("<div style=""padding:2px;"">")
	Response.Write(txtDBName & ": " & f.name & " - Size: " & FormatNumber(dSize/1000,0) & " kb")
	Response.write("</div><br />")
	'Response.Write(f.ParentFolder)
	
	set fo=fso.GetFolder(f.ParentFolder)
	'Response.Write("Get folder " & f.ParentFolder & "<br />")
  	Response.Write("<div class=""fSubTitle"" style=""padding:4px;"">" & txtDBCurBkup & "</div>")
	for each x in fo.files
	'Response.Write("recursing folder<br />")
  	  if (lcase(left(x.Name,len(dName))) = lcase(dName)) and (lcase(right(x.Name,4)) = ".bak" or lcase(right(x.Name,9)) = ".comp.mdb") then
	    if right(x.Name,9) = ".comp.bak" then
	      bDate = left(x.Name, InStr(x.Name, ".comp.bak")-1)
		elseif right(x.Name,4) = ".bak" then
	      bDate = left(x.Name, InStr(x.Name, ".bak")-1)
		elseif right(x.Name,9) = ".comp.mdb" then
	      bDate = left(x.Name, InStr(x.Name, ".comp.mdb")-1)
		end if
		arDate = split(bDate,"_")
	    bDate = strToDate(arDate(ubound(arDate)))
	    Response.write("<div style=""padding:2px;"">")
	    Response.write("<a href=""admin_db.asp?cmd=3&db=" & f.ParentFolder & "\" & x.Name & """ title=""" & txtDel & """>")
        Response.write(icon(icnDelete,txtDel,"","","") & "</a>&nbsp;")
        Response.write(x.Name & " - " & bDate & " - size: " & FormatNumber(x.Size/1000,0) & " kb</div>")
	  end if
	next
	'Response.Write("finished recursing folder<br />")
	set fo=nothing
	set f=nothing
	set fso = nothing

end sub

 Sub swapComp()
    Dim fso
	on error resume next
    Set fso = CreateObject("Scripting.FileSystemObject")
	'fso.DeleteFile sOrg,true
	'sBackUpFile
    fso.MoveFile sBackUpFile,sBackUpFile1
    if err <> 0 then
     bError = True
     response.write "<br />swap Error: " & err.description
    end if
	err.clear
	
    fso.CopyFile sFileTmp, sFileTmp1
    if err <> 0 then
     bError = True
     response.write "<br />swap Error: " & err.description
    end if
	err.clear
	
	set fso = nothing
	on error goto 0
 End Sub

 Sub bakupFile(bFrom,bTo)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile bFrom, bTo
	if fso.FileExists(bTo) = true then
	  ':: all is good, backup created
     'response.write "<br /><br />fso: backup created"
	  fBackup = True
	else
	  ':: Backup not created
     'response.write "<br />fso Error: backup NOT created"
	  fBackup = False
      bError = True
	end if
    if err <> 0 then
     bError = True
     response.write "<br />bak Error:<br />" & err.description & "<br />"
    end if
	set fso = nothing
 End Sub
 %>