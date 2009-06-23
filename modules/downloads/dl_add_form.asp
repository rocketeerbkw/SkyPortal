<!-- #include file="config.asp" --><%
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
CurPageType = "downloads"
%>
<!-- #include file="Lang/en/downloads_lang.asp" -->
<!-- #include file="inc_functions.asp" -->
<!-- #include file="includes/core_module_functions.asp" -->
<!-- #include file="modules/downloads/dl_functions.asp" -->
<% 
CurPageInfoChk = "1"
function CurPageInfo ()
	PageName = txtDownload
	PageAction = txtSubmitting & "<br />" 
	PageLocation = "dl_add_form.asp"
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function

  hasEditor = true  
  strEditorType = "advanced"
  strEditorElements = "Message"
  editorFull = true 
%>
<!-- #include file="inc_top.asp" -->
<script type="text/javascript">
function send() {
	if (document.forms.dlForm.file1.value != ""){
    document.getElementById('wait').style.display = 'block';
    document.getElementById('file1').style.display = 'none';
    document.getElementById('button').style.display = 'none';
	}
	document.dlForm.submit();
}
</script>

<%

':: set default module permissions
setAppPerms CurPageType,"iName"

if strDBNTUserName = "" then
	doNotLoggedInForm
else
if Request.QueryString("parent_id") <> "" or  Request.QueryString("parent_id") <> " " then
	if IsNumeric(Request.QueryString("parent_id")) = True then
		parentID = cLng(Request.QueryString("parent_id"))
	else
		closeAndGo("dl.asp")
	end if
end if
if Request.QueryString("cat_id") <> "" or  Request.QueryString("cat_id") <> " " then
	if IsNumeric(Request.QueryString("cat_id")) = True then
		cat = cLng(Request.QueryString("cat_id"))
	else
		closeAndGo("dl.asp")
	end if
end if
	
'parent = chkString(Request.QueryString("parent_name"), "sqlstring")
'catname = chkString(Request.QueryString("Cat_name"), "sqlstring")
  arg1 = txtDownloads & "|dl.asp"
  arg2 = txtSubDL & "|dl_add_form.asp"
  arg3 = ""
  arg4 = ""
  arg5 = ""
  arg6 = ""
  
if parentID <> "" and parentID > 0 then
  sSQL = "SELECT CAT_ID, SUBCAT_NAME FROM " & strTablePrefix & "M_SUBCATEGORIES WHERE SUBCAT_ID = " & cat
  set rsT = my_Conn.execute(sSQL)
    'cat = rsT(0)
    catname = rsT(1)
    arg3 = catname & "|dl.asp?cmd=2&amp;cid=" & parentID & "&amp;sid=" & cat
  set rsT = nothing
  sSQL = "SELECT CAT_NAME FROM " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID = " & parentID
  set rsT = my_Conn.execute(sSQL)
    parent = rsT(0)
    arg2 = parent & "|dl.asp?cmd=1&amp;cid=" & parentID
  set rsT = nothing
end if
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
<tr>
<td class="leftPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,1)
app_LeftColumn() %>
<script type="text/JavaScript">
function js_val_dladd(){
var c_title = $F("name");
var c_sdes = $F("sdes");
var c_msg = $F("Message");
var c_url = $F("url");
var c_file1 = $F("file1");
var c_email = $F("mail");
var c_subcat = $F("subcat");
//var c_cat = $F("cat");

var alMsg = "";
submitOK="true";

if (c_subcat == 0){
 alMsg += "Please select a category";
 alert(alMsg);
  $('subcat').activate();
  submitOK="false";
  return false;
 }
if (c_title.length<1){
 alMsg += "The Title cannot be empty";
 alert(alMsg);
  $('name').activate();
  submitOK="false";
  return false;
 }
 if (!CheckSql(c_title)){
 alMsg += "The Title cannot contain any of the following characters:\n  \\ / * \" < > | [ ] ;";
 alert(alMsg);
  $('name').activate();
 submitOK="false";
  return false;
 }
if (c_email.length<1){
 alMsg += "The Submitter Email cannot be empty";
 alert(alMsg);
  $('mail').activate();
  submitOK="false";
  return false;
 }
 if (c_file1.length<1 && c_url.length<12){
 alMsg += "Choose a file to upload or link to a URL";
 alert(alMsg);
  $('file1').activate();
  submitOK="false";
  return false;
 }
 if (c_sdes.length<1){
 alMsg += "The Short Description cannot be empty";
 alert(alMsg);
  $('Message').activate();
  submitOK="false";
  return false;
 }
 if (submitOK=="false"){
 alert(alMsg);
 return false;
 }
}
function CheckSql(str) {
	var re;
	re = /[\\\/\]\[:\;*?"<>%|]/gi;
	if (re.test(str)) return false;	
	else return true;
}
 
</script>
</td>
<td class="mainPgCol" valign="top">
<% 
intSkin = getSkin(intSubSkin,2)
dim intDir, intSho
  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  app_MainColumn_top()
%>
<form method="post" action="dl_add_url.asp" id="dlForm" name="dlForm" enctype="multipart/form-data" onsubmit="return js_val_dladd()">
<%
spThemeTitle= txtSubmitDld
spThemeBlock1_open(intSkin) %>
<div id="frmDiv"></div>
  <table border="0" cellpadding="0" cellspacing="4" width="100%"><tr>
    <td align="right" width="30%">
	  <%= txtSubCat %>:<input type="hidden" name="memID" value="<%= strUserMemberID %>" />
	  
	</td><td>
	<% mod_selectCatSubcat cat,"WRITE" %>
	</td>
  </tr>
  <tr> 
	<td align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td align="right">
	  <span class="fAlert">*</span>&nbsp;<%= txtDlTitle %>: 
	</td>
	<td><input type="text" id="name" name="name" size="40" maxlength="90" /></td>
  </tr>
  <% customFormElements() %>
  <tr> 
	<td align="center" colspan="2">&nbsp;</td>
  </tr>
		  <%
		  	strSQL = "select ID, UP_ACTIVE, UP_ALLOWEDGROUPS, UP_SIZELIMIT, UP_ALLOWEDEXT from " & strTablePrefix & "UPLOAD_CONFIG where UP_LOCATION = 'download'"
			set rsUload = my_Conn.execute(strSQL)
			uActive = rsUload("UP_ACTIVE")
			uUpGrps = rsUload("UP_ALLOWEDGROUPS")
			uSize = rsUload("UP_SIZELIMIT")
			uExt = rsUload("UP_ALLOWEDEXT")
			uID = rsUload("ID")
			set rsUload = nothing
		  	session.Contents("uploadType") = uID
		  	session.Contents("loggedUser") = strdbntusername
		  If bFSO = true and strAllowUploads = 1 and uActive = 1 and hasAccess(uUpGrps) Then
		    ast = "**"
			btxt = "<span class=""fAlert"">**</span> = " & txtLnkOrUpld & "<br />" %>
          <tr>
            <td align="right">&nbsp;</td>
            <td align="left" valign="top">
			  <br /><%= txtMaxUpldSize %> <b><%= uSize %> kb</b><br />
			  <%= txtAllowExt %> <b><%= uExt %></b><br />
            </td>
          </tr>
          <tr>
            <td align="right">
			  <span class="fAlert">**</span> <%= txtUpldFile %>:&nbsp; </td>
            <td><input type="hidden" name="max" value="1" />
              <input class="textbox" name="file1" id="file1" type="file" size="30" />
            </td>
          </tr>
		  <% Else
		  	   ast = "*"
			   btxt = "" %>
		  		<input class="textbox" name="file1" id="file1" type="hidden" value="" />
		  <% End If %>
  <tr>
    <td align="right">
      <span class="fAlert"><%= ast %></span> <%= txtUrlOfFile %>: 
    </td>
    <td><input type="text" name="url" size="40" value="http://" maxlength="190" /></td>
  </tr>
  <tr> 
	<td align="center" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td align="right" valign="top"><span class="fAlert">*</span> <%= txtDlSummary %>: <br /><br /><span id="charLeft1">250 <%= txtCharLeft %>&nbsp;</span> </td>
    <td><textarea rows="10" name="sdes" id="sdes" cols="50" wrap="virtual" onKeyUp="cntChar('sdes','charLeft1','{CHAR} <%= txtCharLeft %>.',250);"></textarea></td>
  </tr>
  <tr> 
	<td align="center" colspan="2">&nbsp;
</td>
  </tr>
  <%
  If strAllowHtml = 1 Then 
  	displayHTMLeditor "Message", "<span class=""fAlert"">*</span> " & txtDlLngDesc & ": ",""
  else
  	displayPLAINeditor 1,""
  end if
  if intSecCode <> 0 then
  %>
  <tr>
    <td></td>
    <td><% shoSecurityImg %></td>
  </tr>
  <% 
  End If %>
  <tr>
	<td>&nbsp;</td>
    <td><div id="wait" class="fTitle" style="display:none;"><b><%= txtUpldInProg %></b><br /></div><input id="button" type="submit" value="<%= txtSubmit %>" name="B1" accesskey="s" title="<%= txtSubmit %>" class="button" /></td>
  </tr></table>
<%spThemeBlock1_close(intSkin)%>
</form>
<center>
<span class="fAlert">*</span> = <%= txtReqFld %><br /><%= btxt %><br />
<%= txtNoSeeCat %>,<br />
<a href="Javascript:openWindowPM('pm_pop.asp?mode=2&cid=0&sid=<%= getMemberID(split(strwebmaster,",")(0)) %>');"><u><span class="fAlert"><%= txtContactUs %></span></u></a>, <%= txtBHapConsider %><br />
<br />
</center>
 
<br />
<%  app_MainColumn_bottom() %>
</td>
</tr>
</table>
<% End If
  app_Footer() %>
<!-- #include file="inc_footer.asp" -->
<%
%>