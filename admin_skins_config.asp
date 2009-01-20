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
'<> <a href="/">http://www.SkyPortal.net</a> in the footer of the pages MUST
'<> remain visible when the pages are viewed on the internet or intranet.
'<>
'<> Support can be obtained from support forums at:
'<> <a href="/">http://www.SkyPortal.net</a>
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~ Change: Skin Levels
'~ Date: 02/06/2006
'~ Change By: wingflap <a href="mailto:admin@wingflap.com">admin@wingflap.com</a> <a href="http://www.wingflap.com/">http://www.wingflap.com</a>
'~                                        <a href="http://www.planetloser.com/">http://www.planetloser.com</a>
'~
'~ Description: Added Skin Level to database
'~              Changed Skin Add to add skin as Admin level only
'~              Added Change Skin Level functionality
'~              Changed Delete skin so you can not delete default skin
'~              Added FSO lookup for available skin folders to add
'~              Changed Theme Changer so that only skins appropriate 
'~                  for member's level show
'~              Changed Theme Changer so that if there are more member skins
'~                  than guest, guest is invited to register.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pgType = "manager"
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<% If Session(strCookieURL & "Approval") = "256697926329" Then %>
<% 
rfMethod_Type = Request.Form("Method_Type")
rfMethod_Args = Request.Form("Method_Args")

'wf-FSO SiteLogo Dropdown - boolean flag
blnLogoDropdown=true  'show the dropdown
'blnLogoDropdown=false  'don't show the dropdown

' strLogoFolders is a ';' delimited string of folders to check for logos
' "/" is the skin folder itself (themes/[skin name]) - only finds .gif, .jpg, .png with 'logo' in the name
' "/;logos" will look in the skin folder and in the folder 'logos' under the skins folder.
' "/;logos;logos2" will look in the skin folder and in the folders 'logos' and 'logos2' under the skins folder.
'strLogoFolders = "/" ' Just show logos in the skins's folder
strLogoFolders = "/;logos;logo"

 if Session.Contents("thmLogoImage") <> "" then
   ' We've just gotten the skin info from inc_theme_colors.asp and are ready to add COLORS record.
   if request("tName") <> "" and request("tFolder") <> "" then
     newThm = replace(replace(request("tName"),"<",""),">","")
     thmFolder = replace(replace(request("tFolder"),"<",""),">","")
  whereto = "admin_skins_color.asp?tName=" & newThm & "&tFolder=" & thmFolder
  closeAndGo(whereto)
   end if
 end if

' We've requested to add a New Skin.  We need to get the skin info from inc_theme_colors.asp.
 if rfMethod_Type = "newtheme" then
 newMsg = ""
   if request("tName") <> "" and request("tFolder") <> "" then
   newThm = replace(replace(request("tName"),"<",""),">","")
   thmFolder = replace(replace(request("tFolder"),"<",""),">","")
   where = "themes/" & thmFolder & "/inc_theme_colors.asp?tName=" & newThm & "&tFolder=" & thmFolder
   closeAndGo(where)
   else
  Session("strMsg") = txtSknReqNameFldr
   end if
 end if
 
' Reset Skins for all members to the current Default Skin
 if rfMethod_Type = "resetskins" then
  strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
  strSql = strSql & " SET THEME_ID = '0'"
  strSql = strSql & " WHERE THEME_ID <> ''"
  executeThis(strSql)
  Session("strMsg") = txtSknAllMemSkinsReset
  closeAndGo("admin_skins_config.asp")
 end if
 
 
' Set the Default Skin for the site.
 if rfMethod_Type = "setdefault" then
'  strSQL = "SELECT * FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & Request.Form("defthm") & ""
  strSQL = "SELECT * FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & rfMethod_Args & ""
  'response.Write(strSQL & "<br />")
  set rs = my_conn.execute (strSQL)
   if not rs.eof then
  
  strSql = "UPDATE " & strTablePrefix & "CONFIG"
  strSql = strSql & " SET C_STRDEFTHEME = '" & rs("C_STRFOLDER") & "', C_INTSUBSKIN = " & rs("C_INTSUBSKIN")
  strSql = strSql & ", C_STRTITLEIMAGE = '" & rs("C_STRTITLEIMAGE") & "' WHERE CONFIG_ID = " & 1
  executeThis(strSql)
  set rs = nothing
  Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
  Session("strMsg") = txtSknNewDefSet
   else
  Session("strMsg") = txtSknNoSelect
   end if
  closeAndGo("admin_skins_config.asp")
 end if
 
' Delete a skin
 if rfMethod_Type = "delskin" then
'  strSQL = "SELECT C_STRFOLDER FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & Request.Form("defthm") & ""
  strSQL = "SELECT C_STRFOLDER FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & rfMethod_Args & ""
  set rsS = my_conn.execute (strSQL)
    thmName = rsS("C_STRFOLDER")
  set rsS = nothing
  
   if lcase(rfMethod_Args) <> lcase(strDefTheme) then
  set rs1 = my_Conn.execute("select * from " & strTablePrefix & "CONFIG where C_STRDEFTHEME = '" & thmName & "'")
  if not rs1.eof then
   Session("strMsg") = txtSknNoDelDef & "<br />" & txtSknSetNewDef
  else
   set rs3 = my_Conn.execute("select count(*) from " & strTablePrefix & "COLORS")
   if rs3(0) = 1 then
    Session("strMsg") = txtSknOneSkinOnly & "<br />" & txtSknAddSkinSetDef & "<br />" & txtSknBefDel
    set rs3 = nothing
   else
    set rs2 = my_Conn.execute("select * from " & strTablePrefix & "COLORS where CONFIG_ID = " & rfMethod_Args & "")
    if not rs2.eof then
     executeThis("delete from " & strTablePrefix & "COLORS where CONFIG_ID = " & rfMethod_Args & "")
     executeThis("UPDATE " & strMemberTablePrefix & "MEMBERS set THEME_ID = '" & strDefTheme & "' where THEME_ID = '" & thmName & "'")
     Session("strMsg") = thmName & txtSknDeleted
    else
     Session("strMsg") = txtSknThrIsNo & " '" & thmName & "' "  & txtSknNoneToDel
    end if
    set rs2 = nothing
   end if
  end if
  set rs1 = nothing
   else
    Session("strMsg") = txtSknNoDelDef & "<br />" & txtSknSetNewDef
   end if
  closeAndGo("admin_skins_config.asp")
 end if
 
' Reset a Skin's Level
 if rfMethod_Type = "skinlevel" then
  thmID=rfMethod_Args
 if not isnumeric(thmID) then
  Session("strMsg") = txtSknNoLevChg
 else
  thmID = cint(thmID)
  strSQL = "SELECT C_STRFOLDER FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & rfMethod_Args & ""
  set rsS = my_conn.execute (strSQL)
    thmName = rsS("C_STRFOLDER")
  set rsS = nothing
  
   if lcase(thmName) <> lcase(strDefTheme) then
  set rs1 = my_Conn.execute("select * from " & strTablePrefix & "CONFIG where C_STRDEFTHEME = '" & thmName & "'")
  if not rs1.eof then
   Session("strMsg") = txtSknNoLevChgDef & "<br />" & txtSknSetNewDef
  else
   set rs3 = my_Conn.execute("select count(*) from " & strTablePrefix & "COLORS")
   if rs3(0) = 1 then
    Session("strMsg") = txtSknOneSkinOnly & "<br />" & txtSknAddSkinSetDef & "<br />" & txtSknBefResetLev
    set rs3 = nothing
   else
    set rs2 = my_Conn.execute("select * from " & strTablePrefix & "COLORS where CONFIG_ID = " & thmID & "")
    if not rs2.eof then
     executeThis("UPDATE " & strTablePrefix & "COLORS set C_SKINLEVEL = " & Request.Form("tSkinLevel" & thmID) & " WHERE CONFIG_ID = " & thmID)
     Select Case Request.Form("tSkinLevel" & thmID)
     Case 3
      executeThis("UPDATE " & strMemberTablePrefix & "MEMBERS set THEME_ID = '" & strDefTheme & "' where THEME_ID = '" & thmName & "' and M_LEVEL < 3")
     Case 4
     executeThis("UPDATE " & strMemberTablePrefix & "MEMBERS set THEME_ID = '" & strDefTheme & "' where THEME_ID = '" & thmName & "' and M_LEVEL < 4")
     end select
     Session("strMsg") = thmName & txtSknChanged
    else
     Session("strMsg") = txtSknThrIsNo & " '" & thmName & "' " & txtSknNonToChange
    end if
    set rs2 = nothing
   end if
  end if
  set rs1 = nothing
   else
    Session("strMsg") = txtSknNoLevChgDef & "<br />" &  txtSknSetNewDef
   end if
 end if
  closeAndGo("admin_skins_config.asp")
 end if

' updateskin - Update the Skin Data
 if rfMethod_Type = "updateskin" then
  'get the form vars.
  rftxtSkin = trim(request.form("txtSkin" & rfMethod_Args))
  rftxtSkinDesc = trim(request.form("txtSkinDesc" & rfMethod_Args))
  rftxtSubSkin = cint(request.form("txtSubSkin" & rfMethod_Args))
  rftxtSkinLogo = trim(request.form("txtSkinLogo" & rfMethod_Args))
  rftxtSkinFolder = trim(request.form("txtSkinFolder" & rfMethod_Args))
  app_users = chkGrpAdmin(Request.Form("g_read" & rfMethod_Args))
  'error check the form
  if rftxtSkin = "" or rftxtSkinDesc = "" or rftxtSubSkin = "" or rftxtSkinLogo = "" then
   Session("strMsg") = Session("strMsg") & "<br />" & txtSknAllFldReq
  end if
  if not isnumeric(rftxtSubSkin) then
   Session("strMsg") = Session("strMsg") & "<br />" & txtSknSubSknNum
  end if

  ' if FSO enabled, just go look for the file by path/filename to see if it exists.
  if bFso = true then
   strSkinFolder = "themes/" & rftxtSkinFolder
   fPath = Server.MapPath(strSkinFolder)
   on error resume next
   Err.Clear
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   If objFSO.FileExists(fPath & "\" & rftxtSkinLogo) = false Then
'wf-FSO SiteLogo Dropdown - adjust the error message to reflect the logo folder
'    Session("strMsg") = Session("strMsg") & "<br />" & "themes/" & rftxtSkinFolder & "/" & rftxtSkinLogo & " " & txtSknDoesNotExist
    Session("strMsg") = Session("strMsg") & "<br />" & fPath & "/" & rftxtSkinLogo & " " & txtSknDoesNotExist
   End If
   Set objFSO = nothing
  else
  ' edit the filename to make sure there's a valid extension and no path info
   strFileExt = lcase(right(rftxtSkinLogo,4))
   if left(strFileExt,1) <> "." or _
       (strFileExt <> ".gif" and strFileExt <> ".jpg" and strFileExt <> ".png" and strFileExt <> ".swf") or _
       (instr(1,rftxtSkinLogo,"/") <> 0) or _
       (instr(1,rftxtSkinLogo,"\") <> 0) then
     Session("strMsg") = Session("strMsg") & "<br />" & rftxtSkinLogo & " " & txtSknNotValLogoFilename
   end if
  end if
  
  ' if no edit errors get the existing record to see what's changed.
  if Session("strMsg") = "" then
   strSQL = "SELECT C_TEMPLATE, C_STRDESCRIPTION, C_STRTITLEIMAGE, C_INTSUBSKIN FROM " & strTablePrefix & "COLORS WHERE CONFIG_ID=" & rfMethod_Args & ""
   set rsS = my_conn.execute (strSQL)
   if not rsS.eof then
    sqlUpdate = ""
    if rftxtSkin <> rsS("C_TEMPLATE") then 
     sqlUpdate = sqlUpdate & " C_TEMPLATE = '" & rftxtSkin & "'"
    end if
    if rftxtSkinDesc <> rsS("C_STRDESCRIPTION") then
     if sqlUpdate <> "" then sqlUpdate = sqlUpdate & ", "
     sqlUpdate = sqlUpdate & " C_STRDESCRIPTION = '" & rftxtSkinDesc & "'"
    end if
    if cint(rftxtSubSkin) <> cint(rsS("C_INTSUBSKIN")) then
     if sqlUpdate <> "" then sqlUpdate = sqlUpdate & ", "
     sqlUpdate = sqlUpdate & "C_INTSUBSKIN = " & rftxtSubSkin
    end if
    if rftxtSkinLogo <> rsS("C_STRTITLEIMAGE") then
     if sqlUpdate <> "" then sqlUpdate = sqlUpdate & ", "
     sqlUpdate = sqlUpdate & "C_STRTITLEIMAGE = '" & rftxtSkinLogo & "'"
    end if
    
     if sqlUpdate <> "" then sqlUpdate = sqlUpdate & ", "
     sqlUpdate = sqlUpdate & "C_SKINLEVEL = '" & app_users & "'"
   end if
   rsS.Close
   set rsS = nothing
 
   if sqlUpdate <> "" then
    strSql = "UPDATE "  & strTablePrefix & "COLORS SET "
    strSQL = strSQL & sqlUpdate
    strSql = strSql & " WHERE CONFIG_ID=" & rfMethod_Args
    executeThis(strSql)

    'also create the sql to update the portal_config table
    'if the logo is updated.
    if instr(1,sqlUpdate,"C_STRTITLEIMAGE") > 0 and rftxtSkinFolder = strDefTheme then
     strSql = "UPDATE " & strTablePrefix & "CONFIG"
     strSql = strSql & " SET C_STRTITLEIMAGE = '" & rftxtSkinLogo & "' WHERE CONFIG_ID = " & 1
     executeThis(strSql)
     Application(strCookieURL & strUniqueID & "ConfigLoaded") = ""
    end if
    Session("strMsg") = Session("strMsg") & "<br />" & txtSknDataFor & " " & rftxtSkinFolder & " " & txtSknFldrUpdSucc
   else
    Session("strMsg") = Session("strMsg") & "<br />" & txtSknNoChgNoUpdt
   end if
  end if
  closeAndGo("admin_skins_config.asp")
end if
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr> 
    <td valign="top" class="leftPgCol">
<% 
 intSkin = getSkin(intSubSkin,1)
spThemeTitle = txtMenu
spThemeBlock1_open(intSkin)
 menu_admin()
spThemeBlock1_close(intSkin) %>
 </td>
    <td valign="top" class="mainPgCol">
<%
 intSkin = getSkin(intSubSkin,2)
'breadcrumb here
  arg1 = txtAdminHome & "|admin_home.asp"
  arg2 = txtSknManager & "|admin_skins_config.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
%>
 
<%' Skin Level Mod%>
<% 
spThemeTitle = txtSknCurDefInfo
spThemeBlock1_open(intSkin)
 if Session("strMsg") <> "" then %>
<span class="fSubTitle"><b><%= Session("strMsg") %><% Session("strMsg") = "" %></b></span>
<% end if %>
<!-----------------Beginning of Skin Info------------------->
     <table width="95%" border="0" cellspacing="0" cellpadding="3">
              <tr align="center" valign="middle"> 
    <td width="50%" height="25" class="tSubTitle" align="left"> 
      <span class="fTitle"><b><%=txtSknCurSkn%></b></span></td>
    <td width="50%" height="25" class="tSubTitle" align="left"> 
      <span class="fTitle"><b><%=txtSknDefSkn%></b></span></td>
              </tr>
     <tr>
     <td align="left" valign="top">
<%
 'Current Skin
 strSQL = "SELECT * FROM " & strTablePrefix & "COLORS WHERE C_STRFOLDER='" & strTheme & "'"
 set rs = my_conn.execute (strSQL)
 if rs.eof then
  Response.write("---No Info for Current Skin ---")
 else
  response.write(txtSknSkinName & ": <b>" & rs("C_TEMPLATE") & "</b><br />")
  response.write(txtSknFolder & ": <b>" & rs("C_STRFOLDER") & "</b><br />")
  response.write(txtSknAuthor & ": <b>" & rs("C_STRAUTHOR") & "</b><br />")
  response.write(txtSknDescription & ": <b>" & rs("C_STRDESCRIPTION") & "</b><br />")
  response.write(txtSknSiteLogo & ": <b>" & rs("C_STRTITLEIMAGE") & "</b><br />")
  response.write(txtSknSubSkinValue & ": <b>" & rs("C_INTSUBSKIN") & "</b>")
 end if
 rs.close
 set rs = nothing
%>
   </td>
   <td align="left" valign="top">
<%
'Default Skin
 strSQL = "SELECT * FROM " & strTablePrefix & "COLORS WHERE C_STRFOLDER='" & strDefTheme & "'"
 set rs = my_conn.execute (strSQL)
 if rs.eof then
  Response.write("---No Info for Default Skin ---")
 else
  response.write(txtSknSkinName & ": <b>" & rs("C_TEMPLATE") & "</b><br />")
  response.write(txtSknFolder & ": <b>" & rs("C_STRFOLDER") & "</b><br />")
  response.write(txtSknAuthor & ": <b>" & rs("C_STRAUTHOR") & "</b><br />")
  response.write(txtSknDescription & ": <b>" & rs("C_STRDESCRIPTION") & "</b><br />")
  response.write(txtSknSiteLogo & ": <b>" & rs("C_STRTITLEIMAGE") & "</b><br />")
  response.write(txtSknSubSkinValue & ": <b>" & rs("C_INTSUBSKIN") & "</b>")
 end if
 rs.close
 set rs = nothing

%>
     </td>
 </tr>
</table>
<% spThemeBlock1_close(intSkin) %>
<!-----------------End of Skin Info------------------->
<script type="text/javascript">
function submitS(){
document.formSkinLevel.submit();
}

function chkSkinGrps(fom,ob){
  //alert(fom + ' : ' + ob);
    var bFound = 0
    var mFound = 0
 var oFrm = document[fom];
 for (x = 0;x < oFrm[ob].length ;x++){
   //alert(oFrm[ob].options[x].value);
   if(oFrm[ob].options[x].value == '2'){
  mFound = 1;
   }
   if(oFrm[ob].options[x].value == '3'){
  bFound = 1;
   }
 }
 if (bFound != 1){
   alert("Members and Guests must have READ access to this skin in order to make it the default skin");
   return false;
 }else{
   if (mFound != 1){
     alert("Members and Guests must have READ access to this skin in order to make it the default skin");
     return false;
   }else{
     return true;
   }
   //return true;
 }
}

function removeSGroup(fm,ob,id,df){
 var user,mID;
 var count,finished;
 var oFrm = document[fm];
 finished = false;
 count = 0;
 count = oFrm[ob].length - 1;
 if (count<1) {
  return;
 }
 do //remove from source
 { 
  if (oFrm[ob].options[count].text == ""){
   --count;
   continue;
  }
  if (oFrm[ob].options[count].selected ){
    mID = oFrm[ob].options[count].value
    for ( z = count ; z < oFrm[ob].length-1;z++){
      if (df == 'True'){
    if (!finished){
     if (oFrm[ob].options[z].value != '3'){
       if (oFrm[ob].options[z].value != '2'){
         oFrm[ob].options[z].value = oFrm[ob].options[z+1].value; 
         oFrm[ob].options[z].text = oFrm[ob].options[z+1].text;
        //oFrm[ob].length -= 1;
       } else {
      finished = true;
         alert("You cannot remove MEMBERS from the default skin.\n\nYou must make another skin the Default skin in order to remove this group.");
    }
     } else {
      finished = true;
         alert("You cannot remove GUESTS from the default skin.\n\nYou must make another skin the Default skin in order to remove this group.");
     }
    }
   } else {
     oFrm[ob].options[z].value = oFrm[ob].options[z+1].value; 
     oFrm[ob].options[z].text = oFrm[ob].options[z+1].text;
       //oFrm[ob].length -= 1;
   } 
    }
    if (!finished){
    oFrm[ob].length -= 1;
    }
  }
  --count;
  if (count < 0)
   finished = true;
 }while(!finished) //finished removing
 
  //return;
}
//wf-FSO SiteLogo Dropdown - function swapPreviewImage -  Begin
//replaces the skin's preview image with the one selected
function swapPreviewImage(pImgPath,pImage) {
  var field=document.getElementById(pImage);
  field.style.display="";
  field.src=pImgPath;
}
//wf-FSO SiteLogo Dropdown - function swapPreviewImage -  End
</script>
<!---------------- Beginning of Manage Skins ----------------->
<% 
spThemeTitle = txtSknManageSkins
spThemeBlock1_open(intSkin) %>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td valign="top">
   <table border="0" cellspacing="1" cellpadding="1" width="100%">
     <!-- <tr align="center" valign="middle">
    <td width="100%" height="25" class="tTitle">
   <b><%=txtSknManageSkins%></b></td>
  </tr> -->
  <tr align="center" valign="middle">
    <td>
    <form name="formSkinLevel" method="post" action="<%= Request.ServerVariables("URL") %>" id="formSkinLevel">
    <table id="skinCat" width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="bottom"> 
      <tr>
      <td width="10%" class="tSubTitle" align="center"><b><%=txtCurrent%></b></td>
      <td width="10%" class="tSubTitle" align="center"><b><%=txtSknDefault%></b></td>
      <td width="40%" class="tSubTitle" align="center"><b><%=txtSknSkinName%>/<%= txtSknFolder%> (<%=txtSknAuthor%>)<br />-<%=txtSknDescription%>-</b></td>
      <td width="20%" class="tSubTitle" align="center"><b><%=txtSknMembersUsing%></b></td>
      <!-- <td width="20%" class="tSubTitle" align="center"><b><%=txtSknMinLevForUse%></b></td> -->
      <td width="10%" class="tSubTitle" align="center"><b><%=txtSknDelSkin%></b></td>
                    </tr>
<%     
  strSQL = "SELECT * FROM " & strTablePrefix & "COLORS ORDER BY C_TEMPLATE"
  set rsS = my_conn.execute (strSQL)
  iSkinCounter = 1
  if rsS.eof then%>
   <tr>
   <td width="100%" class="tCellAlt0" align="center" colspan="6"><b><%=txtSknNoSkinsFound%></b></td>
   </tr>
<%  else
   Do while not rsS.eof
   strCellStyle = "tCellAlt" & (iSkinCounter mod 2)+1

   catHide = "none"
   catImg = "max"
   catAlt = txtCollapse
   if request.Cookies(strUniqueID & "hide")("skinCat" & rsS("CONFIG_ID") & "") <> "" then
       if request.Cookies(strUniqueID & "hide")("skinCat" & rsS("CONFIG_ID") & "") = "1" then
        catHide = "none"
        catImg = "max"
        catAlt = txtExpand
       end if
      end if
%>
   <tr>
     <td width="10%" class="<%=strCellStyle%>" align="center"><img name="skinCat<%=rsS("CONFIG_ID")%>Img" id="skinCat<%=rsS("CONFIG_ID")%>Img" src="Themes/<%=strTheme%>/icon_<%=catImg%>.gif" onclick="javascript:mwpHS('skinCat','<%=rsS("CONFIG_ID")%>','tbody');" style="cursor:pointer;" title="<%=catAlt%>" alt="<%=catAlt%>" hspace="3" class="spThemeblock1MinMax" />
   <!--<td width="10%" class="<%=strCellStyle%>" align="center"> -->
<%   if rsS("C_STRFOLDER") = strTheme then%>
   <input type="radio" style:"background-color: transparent;" name="thm" value="<%=rsS("C_STRFOLDER")%>" checked="checked">
<%    else%>
   <input type="radio" name="thm" value="<%=rsS("C_STRFOLDER")%>" onclick="submit()">
<%    end if%>
   </td>
   <td width="10%" class="<%=strCellStyle%>" align="center">
<%   if rsS("C_STRFOLDER") = strDefTheme then%>
    <input type="radio" name="tSkinDefault" value="<%=rsS("C_STRFOLDER")%>" checked="checked">
<%    else
    if rsS("C_SKINLEVEL") <> "0" then%>
     <input type="radio" name="tSkinDefault" value="<%=rsS("C_STRFOLDER")%>" onMouseup="chkSkinGrps('formSkinLevel','g_read<%=rsS("CONFIG_ID")%>');" onclick="javascript:document['formSkinLevel']['method_type'].value='setdefault';document['formSkinLevel']['method_args'].value='<%=rsS("CONFIG_ID")%>';submitS();">
<%    else%>
     <input type="radio" name="tSkinDefault" value="<%=rsS("C_STRFOLDER")%>" disabled="disabled">   
<%    end if
   end if%>
   </td>

   <td width="40%" class="<%=strCellStyle%>" align="center">
    <%=rsS("C_TEMPLATE")%>/<%=rsS("C_STRFOLDER")%> (<%=rsS("C_STRAUTHOR")%>)<br />-<%=rsS("C_STRDESCRIPTION")%>-
   </td>
   <td width="20%" class="<%=strCellStyle%>" align="center">
<%
   strSQL = "SELECT COUNT(*)  as SKINCOUNT FROM " & strTablePrefix & "MEMBERS"
   strSQL = strSQL & " WHERE THEME_ID = '" & rsS("C_STRFOLDER") & "'"
   set rsCount = my_conn.execute (strSQL)

   strCount = ""
   if rsCount.eof then
    strCount="0"
   else
    strCount=cstr(rsCount("SKINCOUNT"))
   end if
   rsCount.Close
   set rsCount = nothing
   if rsS("C_STRFOLDER") = strDefTheme then
    strSQL = "SELECT COUNT(*)  as SKINCOUNT FROM " & strTablePrefix & "MEMBERS"
    strSQL = strSQL & " WHERE THEME_ID = '0' OR THEME_ID = ''"
    set rsCount = my_conn.execute (strSQL)
    If rsCount.eof then
     'do nothing
    else
     strCount = strCount & "<br />(" & cstr(rsCount("SKINCOUNT")) & "&nbsp;" & txtSknByDefault & ")"
    end if
    rsCount.Close
    set rsCount = nothing
   end if
   response.write(strCount)
%>
   </td>
   <td width="10%" class="<%=strCellStyle%>" align="center">
   <%if rsS("C_STRFOLDER") = strDefTheme then %>
   <img src="images/icons/icon_donothing.gif" alt="<%=txtSknCantDel%>&nbsp;<%=rsS("C_TEMPLATE")%> - <%=txtSknIsCurDefSkn%>" border="0">
    <%else%>
   <img style="cursor:pointer;" src="images/icons/icon_delskin.gif" alt="<%=txtDel%>&nbsp;<%=rsS("C_TEMPLATE")%>" border="0" onclick="javascript:document['formSkinLevel']['method_type'].value='delskin';document['formSkinLevel']['method_args'].value='<%=rsS("CONFIG_ID")%>';submitS();">
   <%end if%>
   </td>
   </tr>
   <tbody id="skinCat<%=rsS("CONFIG_ID")%>" class="<%=strCellStyle%>" style="display:<%=catHide%>;">
   <tr>
   <td width="100%" valign="top" align="left" colspan="5"><hr />
   <div align="justify" class="<%=strCellStyle%>">
   
    <table width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr><td align="right" width="20%">
    <%=txtSknSkinName%>:&nbsp;
    </td><td align="left" width="40%">
    <input type="text" name="txtSkin<%=rsS("CONFIG_ID")%>" id="txtSkin<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_TEMPLATE")%>"><br />
    </td>
    <td align="left" valign="top" rowspan="4" width="40%">
    <fieldset style='margin:10px;padding:5px;'>
  <legend><b><%= txtGrpsRead %></b></legend>
      <table border="0" cellpadding="0" cellspacing="0">
   <!-- <tr><td colspan="2" align="center"><b></b><br />&nbsp;</td></tr> -->
     <tr><td align="right" valign="middle" width="50%" nowrap>
    <a href="JavaScript:allowgroups('formSkinLevel','g_read<%=rsS("CONFIG_ID")%>','<%= gLst %>');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
  <a href="JavaScript:removeSGroup('formSkinLevel','g_read<%=rsS("CONFIG_ID")%>','<%=rsS("CONFIG_ID")%>','<%= rsS("C_STRFOLDER") = strDefTheme %>');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;<br />
  <a href="JavaScript:eGroup('formSkinLevel','g_read<%=rsS("CONFIG_ID")%>');" title="<%= txtCM10 %>"><b><%= txtEditGrp %></b></a>&nbsp;&nbsp;
  </td>
          <td align="left"><p>
            <select size="5" name="g_read<%=rsS("CONFIG_ID")%>" style="width:120;" multiple>
     <% 'if gRead <> "" then
       getOptGroups(rsS("C_SKINLEVEL"))
     'end if %>
     <option value="0"></option>
            </select></p>
          </td>
        </tr></table></fieldset>
  
    </td>
    </tr>
    <tr><td align="right">
    <%=txtSknDescription%>:&nbsp;
    </td>
    <td align="left">
    <input type="text" name="txtSkinDesc<%=rsS("CONFIG_ID")%>" id="txtSkinDesc<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_STRDESCRIPTION")%>"><br />
    </td></tr>
    <tr><td align="right">
    <%=txtSknSubSkinValue%>:&nbsp;
    </td>
    <td align="left" valign="middle">
    <input type="text" name="txtSubSkin<%=rsS("CONFIG_ID")%>" id="txtSubSkin<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_INTSUBSKIN")%>"><br />
    </td></tr>
    <tr><td align="right">
    <%=txtSknSiteLogo%>:&nbsp;
    </td>
    <td align="left" valign="middle">
<!--'wf-FSO SiteLogo Dropdown - add form elements for the dropdown BEGIN-->
<!--   <input type="text" name="txtSkinLogo<%=rsS("CONFIG_ID")%>" id="txtSkinLogo<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_STRTITLEIMAGE")%>">&nbsp;&nbsp;&nbsp; -->
<%strSkinFolder="themes/"&cstr(rsS("C_STRFOLDER")) %>
    <input type="text" name="txtSkinLogo<%=rsS("CONFIG_ID")%>" id="txtSkinLogo<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_STRTITLEIMAGE")%>" onchange="swapPreviewImage('<%=strSkinFolder%>/'+this.form.txtSkinLogo<%=rsS("CONFIG_ID")%>.value, 'imgLogoPreview<%=rsS("CONFIG_ID")%>');">&nbsp;&nbsp;&nbsp;
<%if blnLogoDropdown then%>
                <input type="hidden" name="txtFullSkinLogo<%=rsS("CONFIG_ID")%>" id="txtFullSkinLogo<%=rsS("CONFIG_ID")%>" value="<%=strSkinFolder%>/<%=rsS("C_STRTITLEIMAGE")&chr(34)%>"><br />
    </td></tr>
    <tr>
    <td align="right" valign="top">
    <%="Available Logos:&nbsp;"%>
    </td>
    <td align="left" valign="middle">
    <% response.write(DoFSOLogoDropDown(strSkinFolder, strLogoFolders, rsS("C_STRTITLEIMAGE"), rsS("CONFIG_ID"),bFSO))%>
    <img alt="Logo preview" name="imgLogoPreview<%=rsS("CONFIG_ID")%>" id="imgLogoPreview<%=rsS("CONFIG_ID")%>" src="<%=strSkinFolder%>/<%=rsS("C_STRTITLEIMAGE")%>">
<%end if%>
<!--'wf-FSO SiteLogo Dropdown - add form elements for the dropdown END-->
    <input type="hidden" name="txtSkinFolder<%=rsS("CONFIG_ID")%>" id="txtSkinFolder<%=rsS("CONFIG_ID")%>" value="<%=rsS("C_STRFOLDER")%>">
    <input type="submit" value="<%=txtSknUpdate%>" id="btnUpdateSkin<%=rsS("CONFIG_ID")%>" name="btnUpdateSkin<%=rsS("CONFIG_ID")%>" onclick="javascript:document['formSkinLevel']['method_type'].value='updateskin';document['formSkinLevel']['method_args'].value='<%=rsS("CONFIG_ID")%>';selectAll('formSkinLevel','g_read<%=rsS("CONFIG_ID")%>');">
    </td>
    </tr>
    </table>
   </div><hr />
   
   </td>
   </tr>
   </tbody>
<%   iSkinCounter = iSkinCounter + 1
   rsS.movenext
   loop
  end if
 rsS.Close
 set rsS=Nothing%>
                  </table>
            <input name="method_type" type="hidden" id="method_type" value="">
   <input name="method_args" type="hidden" id="method_args" value="">
      </form>
    </td>
  </tr>
   </table>
 </td>
  </tr>
</table>
 
<%  spThemeBlock1_close(intSkin) 
%>
<!---------------- Beginning of RESET/ADD Skin ----------------->
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td valign="top" width="50%">
<% 
spThemeTitle = txtSknResetMemSkins
spThemeBlock1_open(intSkin) %>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr>
    <td valign="top"><img src="images/spacer.gif" height="2">
 </td>
  </tr>
  <tr>
    <td valign="top">
   <table border="0" cellspacing="1" cellpadding="1" width="100%">
  <tr align="center" valign="middle">
     <td align="center"><br />
  <b><%=txtSknClickToReset%>:</b><br /><br />
     <form name="resetskin" method="post" action="<%= Request.ServerVariables("URL") %>">
      <input class="button" type="submit" name="Submit" value="<%=txtSknResetMemSkins%>">
      <input name="method_type" type="hidden" id="method_type" value="resetskins">
     </form><br />&nbsp;
  </td>
  </tr>
   </table>
 </td>
  </tr>
</table>
<% spThemeBlock1_close(intSkin) %>
    </td><td valign="top" width="50%">
<% 
spThemeTitle = txtSknAddASkin
spThemeBlock1_open(intSkin) %>
<%if bFso then%>
    <b><%=txtSknAddSkinFSO%></b>
<%else%>
    <b><%=txtSknAddSkinNoFSO%></b>
<%end if%>
    <br />
                  <form name="AddNewTheme" method="post" action="admin_skins_config.asp">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="56%">&nbsp;</td>
                        <td width="44%">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td align="right">New 
                          <%=txtSknSkinName%>:&nbsp; </td>
                        <td> 
                          <input class="textbox" name="tName" type="text" value="" size="15">
                        </td>
                      </tr>
                      <tr> 
                        <td align="right">Skin 
                          <%=txtSknFolder%>:&nbsp; </td>
                        <td> 
<%'stop%>
      <%if bFso = true then%>
      <%=DoFSOFolderDropDown("themes","C_STRFOLDER",strTablePrefix & "COLORS",0,blnAnyToAdd)%>
      <%else%>
                          <input class="textbox" name="tFolder" type="text" value="" size="15">
      <%blnAnyToAdd=True
        end if%>
                        </td>
                      </tr>
                      <tr> 
                        <td align="right">&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr align="center"> 
                        <td colspan="2" height="25" align="top">
                          <input class="button" type="submit" name="Submit" value="<%=txtSknAddSkn%>"
        <%if blnAnyToAdd = False then response.write(" disabled=""True""")%>>
                          <input name="method_type" type="hidden" id="method_type" value="newtheme">
                        </td>
                      </tr>
                    </table>
                  </form>
<% spThemeBlock1_close(intSkin) %>
 </td>
  </tr>
</table>
<!---------------- End of RESET/ADD Skin ----------------->
    </td>
  </tr></table>
<!--#include file="inc_footer.asp" -->
<%
Else
 scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
 Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
End IF
%>
<%
'  Function DoFSODropDown
'   This function creates a dropdown box with a list of the folder names
'   that do not belong
'   sStartPath is the folder that contains the folders you want in the list
Function DoFSOFolderDropDown(sStartPath, sCompField, sCompTable, sMatch, bFoundAny)

 'assume you're looking for discrepancies
 '0=folders that don't match db.
 '1=folders that match db.
 if not isnumeric(sMatch) then
  sMatch = 0
 else
  if sMatch <> 1 then
   sMatch = 0
  end if
 end if
 bFoundAny = False

 on error resume next
 ' Get all the Skin Names and Folder Names from the COLORS table
 strSQL = "SELECT " & sCompField & " FROM "  & sCompTable
 strSQL = strSQL & " ORDER BY " & sCompField
 set rsFldr = my_conn.execute (strSQL)
 if err.number <> 0 then
  DoFSOFolderDropDown = txtDBerror
  exit function
 end if
 'Make a list of all folders in the db sandwiched by "|"s
 strFldrs = ""
 do while not rsFldr.eof
  if strFldrs = "" then
   strFldrs = strFldrs & "|"
  end if
  strFldrs = strFldrs & lcase(rsFldr(sCompField)) & "|"
  if err.number <> 0 then
   DoFSOFolderDropDown = txtDBerror
   exit function
  end if
  
  rsFldr.movenext
 loop
 rsFldr.Close
 set rsFldr = nothing
 
 on error resume next
 Set objFSO = CreateObject("Scripting.FileSystemObject") 
 set mainfolder=objFSO.GetFolder(Server.MapPath(sStartPath))
 if err.number <> 0 then
  DoFSOFolderDropDown = txtSknFolderNotFnd
  exit function
 end if
 Set foldercollection = mainfolder.SubFolders

 'build the <options> for the <select>
 strOptions=""
 For Each folder In foldercollection
 
 if sMatch = 0 then
  if instr(1, strFldrs, "|" & lcase(folder.Name) & "|") = 0 then
   if strOptions <> "" then
    strOptions = strOptions & "<br />"
   end if
   strOptions = strOptions & "<option value=" & chr(34) & folder.Name & chr(34) & ">" & folder.Name & "</option>"
  end if
 else
  if instr(1, strFldrs, "|" & lcase(folder.Name) & "|") > 0 then
   if strOptions <> "" then
    strOptions = strOptions & "<br />"
   end if
   strOptions = strOptions & "<option value=" & chr(34) & folder.Name & chr(34) & ">" & folder.Name & "</option>"
  end if
 end if

 Next
 set objFSO = nothing

 if strOptions = "" then
  DoFSOFolderDropDown = "<b>" & txtSknNoNewToAdd & "</b>"
 else
  strSelect = "<select name=" & chr(34) & "tFolder" & chr(34) & " id=" & chr(34) & "tFolder" & chr(34)& ">" & vbcrlf
  strSelect = strSelect & strOptions & vbcrlf
  strSelect = strSelect & "</select>" & vbcrlf
  DoFSOFolderDropDown = strSelect
  bFoundAny = True
 end if
end function

'  wf-FSO SiteLogo Dropdown - Function to render the Dropdown
Function DoFSOLogoDropDown(sStartPath,strLogoFolders,sCurrLogo,iConfigID,bFSO)
 if bFso = false then
  DoFSOLogoDropDown = "FSO not enabled on your server"
  exit function
 end if

' initialize the string to hold the 'options' tags.
 strOptions = ""
' make an array of the folders
 arLogoFolders = split(strLogoFolders,";")
' get the logo files for the specified folder
 for i = 0 to ubound(arLogoFolders)
  bGoodFolder = true
  sShortPath = arLogoFolders(i)
  if sShortPath = "/" then
   sShortPath = ""
  end if
  on error resume next
  Set objFSO = CreateObject("Scripting.FileSystemObject") 
  set mainfolder=objFSO.GetFolder(Server.MapPath(sStartPath & "/" & sShortPath))
  if err.number <> 0 then
    bGoodFolder = false
'    DoFSOLogoDropDown = "Can't find the Skin Folder"
'    exit function
  end if
  Set filesObject = mainfolder.Files 
  if err.number <> 0 then
   bGoodFolder = false
'    DoFSOLogoDropDown = "No logo files found "
'    exit function
  end if

' Now add the "/" to the end of the non-blank Short Path
  if sShortPath <> "" then
   sShortPath = sShortPath & "/"
  end if
  if bGoodFolder then
   For Each file In filesObject 
    strFileName = lcase(sShortPath & file.name)
    if instr(1,strFileName, ".jpg") > 0 or instr(1,strFileName, ".gif") > 0 or instr(1,strFileName, ".png") > 0 then
     if (instr(1, strFileName, "logo") > 0 and strFileName <> "logout.gif") or (sShortPath <> "") then
      if strOptions <> "" then
       strOptions = strOptions & "<br>"
      end if
      
      strOptions = strOptions & "<option value=" & chr(34) & sShortPath & file.Name & chr(34)
      if strFileName = lcase(sCurrLogo) then
       strOptions = strOptions & " selected"
      end if
      strOptions = strOptions & ">" 
      strOptions = strOptions & sShortPath & file.Name & "</option>"
     end if
    end if
   Next ' file
  end if 'good folder
  set objFSO = nothing
 Next ' folder
 if strOptions = "" then
  DoFSOLogoDropDown = "<b>" & txtSknNoNewToAdd & "</b>"
 else 
   strSelect = "<select name=" & chr(34) & "selSkinLogo" & iConfigID & chr(34) & " id=" & chr(34) & "selSkinLogo" & iConfigID & chr(34)
  ' for debugging - uncomment the alert and comment out the assign to see the value changes.
  '  strSelect = strSelect & " ONCHANGE=" & chr(34) & "alert('Index: ' + this.selectedIndex + '\nValue: ' + this.options[this.selectedIndex].value+'\ntextbox: '+this.form.txtSkinLogo" & iConfigID & ".value);" & chr(34)
  strSelect = strSelect & " ONCHANGE=" & chr(34) & "this.form.txtSkinLogo" & iConfigID & ".value=this.options[this.selectedIndex].value;"
  strSelect = strSelect & "this.form.txtFullSkinLogo" & iConfigID & ".value='" & sStartPath &"/'+this.options[this.selectedIndex].value;"
  strSelect = strSelect & "swapPreviewImage(this.form.txtFullSkinLogo" & iConfigID & ".value , 'imgLogoPreview"& iConfigID & "')" & ";"
  strSelect = strSelect & chr(34) 
  strSelect = strSelect & ">" & vbcrlf
  strSelect = strSelect & strOptions & vbcrlf
  strSelect = strSelect & "</select>" & vbcrlf
  DoFSOLogoDropDown = strSelect
  bFoundAny = True
 end if
end function
%>
