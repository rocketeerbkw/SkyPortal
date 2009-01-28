<!--#include file="config.asp" --><%
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
CurPageType = "core"
PageName = ""
hasEditor = true
strEditorElements = ""
%>
<!-- #include file="lang/en/core_admin.asp" -->
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top.asp" -->
<!--#include file="includes/inc_admin_functions.asp" -->
<%
if request("cmd") <> "" then
iCmd = chkString(request("cmd"),"numeric")
else
iCmd = 0
end if

If Request("pagesize") = "" Then
	iPageSize = 10
Else
	iPageSize = clng(Request("pagesize"))
End If

If Request("sortme") = "" Then
	iSort = 0
Else
	iSort = clng(Request("sortme"))
End If

If Request("MAilAllMembers") = "" Then
	iMailAll = 0
Else
	iMailAll = Request("MAilAllMembers")
End If

%>
<script type="text/javascript">
<!-- hide from JavaScript-challenged browsers
function selectAll(formObj, isInverse) 
{ 
with (formObj) 
{ 
for (var i=0;i < length;i++) 
{ 
fldObj = elements[i]; 
if(isInverse) 
{ 
if (fldObj.name != 'inverse') 
{ 
if (fldObj.name == 'selectall') 
fldObj.checked = false; 
else 
fldObj.checked = (fldObj.checked) ? false : true; 
} 
else fldObj.checked = true; 
} 
else 
{ 
fldObj.checked = true; 
if (fldObj.name == 'inverse') fldObj.checked = false; 
} 
} 
} 
}
function ChangePage(){
	document.PageNum.submit();
}


function js_togeditor(fe,fte){
  //alert('new function');
  toggleEditor(fte);
  return true;
}
function js_tagdata(fe,fte){
  var htm = $F('html');
  var tgdta = $F('tagdata');
  if(htm == 1){
    tinyMCE.getInstanceById(fte).execCommand('mceInsertContent',false,document.getElementById('tagdata').value);
    //tinyMCE.getInstanceById('Message').execCommand('mceInsertContent',false,document.getElementById('tagdata').value);
  }else{
	AddText(fe,fte,tgdta);
  }
}

function getActiveText(selectedtext) { 
	text = (document.all) ? document.selection.createRange().text : document.getSelection();
		if (selectedtext.createTextRange) {	
   			selectedtext.caretPos = document.selection.createRange().duplicate();	
  		}
		return true;
}
function AddText(fr,el,NewCode) {
var ele=document[fr]
var fele=ele[el]
if (fele.createTextRange && fele.caretPos) {
var caretPos = fele.caretPos;
caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? NewCode + ' ' : NewCode;
}
else {
fele.value+=NewCode
}
fele.focus();
}
// done hiding -->
</script>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
<td valign="top" class="leftPgCol">
	<% intSkin = getSkin(intSubSkin,1)
spThemeBlock1_open(intSkin)
		menu_admin()
spThemeBlock1_close(intSkin) 
%>
</td>
<td valign="top" class="mainPgCol">
<% intSkin = getSkin(intSubSkin,2)

	select case iCmd
	case 1
  		arg1 = txtAdmin & "|admin_home.asp"
  		arg2 = txtemUserEmailList & "|admin_emaillist.asp"  
  		arg3 = "Sending Message"
  		arg4 = ""
  		arg5 = ""
  		arg6 = ""
	case 4
  		arg1 = txtAdmin & "|admin_home.asp"
  		arg2 = txtemUserEmailList & "|admin_emaillist.asp"  
  		arg3 = "Sending Group Message"
  		arg4 = ""
  		arg5 = ""
  		arg6 = ""  		
		
	case else
  		arg1 = txtAdmin & "|admin_home.asp"
  		arg2 = txtemUserEmailList & "|admin_emaillist.asp"  
  		arg3 = ""
  		arg4 = ""
  		arg5 = ""
  		arg6 = ""
	end select

  shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  
  spThemeBlock1_open(intSkin) %>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td align="center" valign="middle">
		<%
		adm_topMnu()
		
	select case iCmd
		
		case 0 'default
				adm_listMembers()
		case 1 'select email review
			set oFormVars=GetFormObject()
			if oFormVars("ID") = "" and iMailAll = 0 then
			adm_error(txtemNoRecipSel)
			adm_listMembers()
			else
			sendEmailSelect()
			end if
			
		case 2 'send selected email
			'response.write("M_ID: " & request("m_id") & " MailAll: " & iMailAll)
			set oFormVars=GetFormObject()
			if oFormVars("m_id") = "" and iMailAll = 0 then
			adm_error(txtemNoRecipSel)
			adm_listMembers()
			else
			sendSelectedMail()
			end if
			
		case 3' grp emailing
			adm_listGroups()
			
		case 4 'group message review
			set oFormVars=GetFormObject()
			if oFormVars("ID") = "" then
			adm_error(txtemNoRecipSel)
			adm_listGroups()
			else
			'response.write("groups: " & oFormVars("ID"))
			sendEmailSelectG()
			end if
			
		case 5 'sending selectG mail
			set oFormVars=GetFormObject()
			if oFormVars("m_id") = "" and iMailAll = 0 then
			adm_error(txtemNoRecipSel)
			adm_listGroups()
			else
			sendSelectedMailG()
			end if
			
		case 6 'create new message
			adm_CreateMessage()
			
		case 7 'submit new message
			if request("Message") = "" or request("subject") = "" then
				adm_error("Both Subject And Message Are Required!")
				intError = true
				adm_CreateMessage()
			else
				save = "INSERT INTO PORTAL_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) "
				save = save & " VALUES ("
				save = save & "'" & chkString(request("subject"),"sqlstring") & "', '" & request("message") & "', '" & strCurDateString & "', " & request.form("ARCHIVE") & ")"
				'response.write(save)
				executeThis(save)
				adm_CreateMessage()
			end if
			
		case 8 'manage messages
			adm_ManageMessages()
			
		case else 'default
				adm_listMembers()
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
<%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::: PAGE FUNCTIONS :::::::::::::::::::::::::::::::::::::::::::::::::::::::::
function adm_listMembers()
  iPageCurrent = 1
If Request("pageno") = "" Then
  iPageCurrent = 1
Else
  iPageCurrent = cLng(Request("pageno"))
End If

  strSql = "SELECT * FROM PORTAL_MEMBERS WHERE MEMBER_ID<>0"
  
if iSort <> 0 then

	select case iSort
	
	case 1
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 3
	case 2
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 2
	case 3
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 1
	case 4
		My_Last = DateToStr(DateAdd("m", -6, now()))
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE < '" & My_Last & "'"
	case 5
		My_Last = DateToStr(DateAdd("yyyy", -1, now()))
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE < '" & My_Last & "'"
	case 6
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & 0 
	case 7
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 1
	
	end select


end if

  Set objPagingRS = Server.CreateObject("ADODB.Recordset")
  objPagingRS.PageSize = iPageSize
  objPagingRS.CacheSize = iPageSize
  objPagingRS.Open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  reccount = objPagingRS.recordcount
  iPageCount = objPagingRS.PageCount

  If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  If iPageCurrent < 1 Then iPageCurrent = 1
  
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#000000">
  <form action="admin_emaillist.asp" method="post">
    <tr>
      <td width="100%" class="tCellAlt1">
      <table border="0" width="100%" cellspacing="1" cellpadding="4">
        <input type="hidden" name="cmd" value="1">
        <tr>
          <td colspan="4">
          <p align="center" style="margin-top: 0; margin-bottom: 0"><%emailListBox()%>
          </p>
          <p align="center" style="margin-top: 0; margin-bottom: 0"><%=txtemSendAll%> <input type="checkbox" name="MAilAllMembers" value="1"><input type="submit" name="action" value="<%=txtemSendMsgToSelected%>" class="button"></p>
          <p align="center" style="margin-top: 0; margin-bottom: 0"><%
  						':::::::::: PAGING BROKEN ::::::::::::::
  						%> </p>
          </td>
        </tr>
        <tr>
          <td class="tTitle">
          <input type="checkbox" name="CheckAll" value="1" onclick="selectAll(this.form,1)"><%= " " & txtemMail%></td>
          <td class="tTitle"><%=txtUsrNam%></td>
          <td class="tTitle"><%=txtEmlAdd%></td>
          <td class="tTitle">&nbsp;</td>
        </tr>
        <% If iPageCount = 0 or objPagingRS.eof or objPagingRS.bof Then %>
        <tr>
          <td class="tCellAlt1" colspan="4">
          <p align="center"><b><%=txtNoMemFnd%></b></p>
          </td>
        </tr>
        <%
 Else
 	objPagingRS.AbsolutePage = iPageCurrent
 
 	iRecordsShown = 0
 	CColor = "tCellAlt2"
 
 	Do While iRecordsShown < iPageSize And Not objPagingRS.EOF

 		if CColor = "tCellAlt1" then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if
%>
        <tr class="<%=CColor%>">
          <td><% '--------Does user want spam?
if objPagingRS("M_RECMAIL") = "1" then
%> <img src="select_nocheck.gif" border="0"> <% else %>
          <input type="checkbox" name="ID" value="<% =objPagingRS("MEMBER_ID") %>">
          <input type="hidden" name="Mail_ALL" value="<% =objPagingRS("MEMBER_ID") %>">
          <% end if %> </td>
          <td><font class="fBold">
          <a href="cp_main.asp?cmd=8&member=<%=objPagingRS("MEMBER_ID")%>"><%=objPagingRS("M_NAME")%></a></font></td>
          <td><font class="fBold"><a href="mailto:<%=objPagingRS("M_EMAIL")%>"><%=objPagingRS("M_EMAIL")%></a></font></td>
          <td>
          <a href="cp_main.asp?cmd=10&mode=Modify&ID=<% =objPagingRS("MEMBER_ID") %>&name=<% =objPagingRS("M_NAME") %>">
          <%= icon(icnEdit,txtemEditMember,"","","") %></a>
          <a href="JavaScript:openWindow('pop_portal.asp?cmd=1&cid=<% =objPagingRS("MEMBER_ID") %>')">
          <%= icon(icnDelete,txtemDelMember,"","","") %></a> </td>
        </tr>
        <%
   iRecordsShown = iRecordsShown + 1
   objPagingRS.MoveNext
 Loop
  end if
  
  objPagingRS.Close
  Set objPagingRS = Nothing
%>
      </table>
      </td>
    </tr>
  </form>
</table>
<%
end function

function adm_topMnu()
%>   
 <div align="center">
      <center>
<table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width="80%" style="border-collapse: collapse">
  <tr class="tCellAlt2">
    <td width="100%">
    <table border="0" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <tr>
        <td width="100%" class="tSubTitle"><%=txtemOptionBar%></td>
      </tr>
      <tr>
        <td width="100%">
<p align="center" style="margin-top: 0; margin-bottom: 0"><a href="admin_emaillist.asp?cmd=6"><%=txtemCreateNewMessage%></a> | <a href="admin_emaillist.asp?cmd=8"><%=txtemManageMessages%></a> | 
<a href="admin_emaillist.asp?cmd=3"><%=txtemGrpEmail%></a> | <a href="javascript:;" onclick="show('pageFilter')"><%=txtemPgFilter%></a> 
| <a href="admin_emaillist.asp"><%=txtemUserEmailList%></a><div align="center" style="display:none" id="pageFilter">
  <center>
  &nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="40%">
    <form methd="POST" action="admin_emaillist.asp">
    <tr class="tCellAlt1">
      <td align="right"><%=txtemPerPage%></td>
      <td align="left">
    
    <% if request("cmd") <> "" then %><input type="hidden" name="cmd" value="<%=chkString(request("cmd"),"numeric")%>"><% end if %>
<select name="pagesize" size="1">
  <!-- <option value="<%= iPageSize %>" selected="selected">&nbsp;<%= PageSize2 %></option> -->
  <option value="10" selected="selected">&nbsp;10</option>
  <option value="25"<%= chkSelect(iPageSize,25) %>>&nbsp;25</option>
  <option value="50"<%= chkSelect(iPageSize,50) %>>&nbsp;50</option>
  <option value="100"<%= chkSelect(iPageSize,100) %>>&nbsp;100</option>
  <option value="200"<%= chkSelect(iPageSize,200) %>>&nbsp;200</option>
  <option value="500"<%= chkSelect(iPageSize,500) %>>&nbsp;500</option>
  <!-- <option value="1000"<%'= chkSelect(iPageSize,1000) %>>&nbsp;1000</option>
  <option value="2000"<%'= chkSelect(iPageSize,2000) %>>&nbsp;2000</option>
  <option value="5000"<%'= chkSelect(iPageSize,5000) %>>&nbsp;5000</option> -->
  </select>
      </td>
      <td align="left"><input type="submit" value="<%=txtGo%>" class="button"></td>
    </tr>
    <tr class="tCellAlt1">
      <td align="right"><%=txtemSelUserGroup%></td>
      <td align="left">
 <select name="sortme" size="1">
  <!-- <option value="<%=My_Sort%>" selected="selected">&nbsp;<% =Sort_Name%></option> -->
  <option value="0">&nbsp;<%=txtemAllUsers%></option>
  <option value="3">&nbsp;<%=txtemGenUsersOnly%></option>
  <!-- <option value="2">&nbsp;<%'=txtemModOnly%></option> -->
  <option value="1">&nbsp;<%=txtemAdminOnly%></option>
  <option value="4">&nbsp;<%=txtemInactive6Mo%></option>
  <option value="5">&nbsp;<%=txtemInactive1Yr%></option>
  <option value="6">&nbsp;<%=txtemNeverPosted%></option>
  <option value="7">&nbsp;<%=txtemRefuseEmail%></option>
  </select>
      </td>
      <td align="left"><input type="submit" value="<%=txtGo%>" class="button"></td>
    </tr>
    <tr class="tCellAlt1">
      <td align="right">&nbsp;</td>
      <td align="left">&nbsp;</td>
      <td align="left">&nbsp;</td>
    </tr>
    </form>
  </table>
  </center>
</div>
<p align="center" style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</td>
      </tr>
    </table>

    </td>
  </tr>
</table>
      </center>
    </div>
&nbsp;
<%
end function

function adm_ManageMessages()

select case chkString(request("mode"),"numeric")
	
	case 1 'delete
		executeThis("Delete * From PORTAL_SPAM WHERE ID=" & chkString(request("msg_id"),"numeric"))
		strMsg = "<font class=""fTitle"">Message Removed!</font>"
	case 2 'edit
	
	case 3 'save edited message
		sql = "UPDATE PORTAL_SPAM SET ARCHIVE=" & chkString(request("archive"),"numeric") & ", SUBJECT = '" & chkString(request("subject"),"sqlstring") & "', MESSAGE = '" & request("message") & "', F_SENT = '" & strCurDateString & "'"
		sql = sql & " WHERE ID=" & chkString(request("MSG"),"numeric")
		executeThis(sql)
		strMsg = "<font class=""fTitle"">Message Saved!</font>"
end select
%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width="90%" style="border-collapse: collapse">
    <tr class="tCellAlt2">
      <td width="100%">
      <table border="0" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%" class="tSubTitle"><%=txtemManageMessages%></td>
        </tr>
        <tr>
          <td><% if strMsg <> "" then %>
          <p align="center" style="margin-top: 0; margin-bottom: 0"><%=strMsg%> <% end if %>
          </p>
          <p align="center" style="margin-top: 0; margin-bottom: 0"><% 
strSql = "SELECT * FROM " & strTablePrefix & "SPAM ORDER BY ARCHIVE ASC"
set rs = Server.CreateObject("ADODB.Recordset")
rs.open  strSql, My_Conn, 3
if request("mode") <> 2 then
%> </p>
          <table BORDER="1" class="grid" CELLSPACING="0" align="center" width="100%">
            <tr ALIGN="CENTER">
              <td class="tTitle"><b><%=txtStatus%></b> &nbsp;</td>
              <td class="tTitle"><b><%=txtemMsgTitle%></b> &nbsp;</td>
              <td class="tTitle"><b><%=txtemComposed%></b> &nbsp;</td>
              <td class="tTitle"><a href="admin_emaillist.asp?cmd=6">
              <img src="images/icons/icon_folder_new_topic.gif" alt="<%=txtemAddNewMsg%>" title="<%=txtemAddNewMsg%>" border="0" hspace="0">
              <%=txtemAddNewMsg%></a></td>
            </tr>
            <%
if RS.eof or RS.bof then
  response.write "<b>No messages found!</b>"
else
  RS.MoveFirst
  do while Not RS.eof                       
  ARCHIVED = rs("ARCHIVE")
  if ARCHIVED = "1" then
  ARCHIVED = "ARCHIVED"
  else
  ARCHIVED = "LIVE"
  end if
  if rs("F_SENT") <> "" then
  F_SENT = ChkDate(rs("F_SENT"))
  else
  F_SENT = "-" 
end if
 %>
            <tr VALIGN="TOP">
              <td class="tCellAlt1"><%= ARCHIVED %> &nbsp;</td>
              <td class="tCellAlt1">
              <input type="hidden" name="ID" value="<%=RS("ID")%>">
              <a href="admin_emaillist.asp?cmd=8&mode=2&msg_id=<% =rs("ID") %>">
              <%=RS("SUBJECT")%></a> &nbsp;</td>
              <td ALIGN="CENTER" class="tCellAlt1"><% =F_SENT %> &nbsp;</td>
              <td class="tCellAlt1" align="right">
              <a href="admin_emaillist.asp?cmd=8&mode=2&msg_id=<% =rs("ID") %>">
              <%= icon(icnEdit,txtemEditMsg,"","","") %></a>
              <a href="admin_emaillist.asp?cmd=8&mode=1&msg_id=<% =rs("ID") %>">
              <%= icon(icnDelete,txtemDelMsg,"","","") %></a> &nbsp;</td>
            </tr>
            <%
RS.MoveNext
loop
end if
%>
          </table>
          <% 
set rs = nothing 
else
sql = "SELECT * FROM PORTAL_SPAM WHERE ID=" & chkString(request("msg_id"),"numeric")
set daMsg = my_conn.execute(sql)
	if NOT daMsg.eof and NOT daMsg.bof then
%>
          <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <form method="POST" action="admin_emaillist.asp" name="emCreate" id="emCreate">
              <input type="hidden" name="cmd" value="8">
              <input type="hidden" name="mode" value="3">
              <input type="hidden" name="MSG" value="<%=chkString(request("msg_id"),"numeric")%>">
              <tr>
                <td align="right"><b><%=txtSubject & ": &nbsp;"%></b></td>
                <td>
                <input class="textbox" name="SUBJECT" size="50" style="float: left" value="<%=daMsg("SUBJECT")%>"></td>
              </tr>
              <tr>
                <td align="right"><b><%=txtemTagInsert%> :</b></td>
                <td>
                <% tagSelect() %>
                </td>
              </tr>
              <tr>
                <td align="right"><b><%=txtMsg & ": &nbsp;"%></b></td>
                <td>
                <%adm_showEditor(daMsg("MESSAGE")) %>
                </td>
              </tr>
        <tr>
          <td align="right">
          <b>Create HTML Message:</b></td>
          <td>
          <input type="checkbox" name="html" value="1" onmouseup="toggleEditor('Message');">
          </td>
        </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td>
                <p align="left"><select name="ARCHIVE" size="1">
                <option value="0" selected="selected">&nbsp;<%=txtemLiveList%>
                </option>
                <option value="1">&nbsp;<%=txtemArchive%></option>
                </select> </p>
                </td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td><input type="submit" value="<%=txtemEditMsg%>" class="button"></td>
              </tr>
            </form>
          </table>
          <% 
end if
	set daMsg = nothing
end if 
%> </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>
&nbsp; 
<%
end function

function adm_CreateMessage()
%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width="90%" style="border-collapse: collapse">
    <tr class="tCellAlt2">
      <td width="100%">
      <table border="0" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%" class="tSubTitle"><%=txtemCreateNewMessage%></td>
        </tr>
        <% if iCmd = 7 and NOT intError then %>
        <tr>
          <td width="100%">
          <p align="center"><font class="fTitle"><%=txtemMsgSaved%></font></p>
          </td>
        </tr>
        <% end if %>
        <tr>
          <td width="100%">
          <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
            <form method="post" action="admin_emaillist.asp" name="emCreate" id="emCreate">
              <input type="hidden" name="cmd" value="7">
              <tr>
                <td align="right" width="25%"><b><%=txtSubject & ": &nbsp;"%></b></td>
                <td align="right">
                <input class="textbox" name="SUBJECT" size="50" style="float: left"></td>
              </tr>
              <tr>
                <td align="right"><b><%=txtemTagInsert%> :</b></td>
                <td>
                <% tagSelect() %>
                </td>
              </tr>
              <tr>
                <td align="right"><b><%=txtMsg & ": &nbsp;"%></b></td>
                <td>
                <%adm_showEditor("") %>
                </td>
              </tr>
        <tr>
          <td align="right">
          <b>Create HTML Message:</b></td>
          <td>
          <input type="checkbox" name="html" value="1" onmouseup="js_togeditor('emCreate','Message');">
          </td>
        </tr>
              <tr>
                <td align="right">Save to:</td>
                <td align="right">
                <p align="left"><select name="ARCHIVE" size="1">
                <option value="0" selected="selected">&nbsp;<%=txtemLiveList%>
                </option>
                <option value="1">&nbsp;<%=txtemArchive%></option>
                </select> </p>
                </td>
              </tr>
              <tr>
                <td align="right">&nbsp;</td>
                <td>
                <input type="submit" value="<%=txtemSaveMsg%>" class="button"></td>
              </tr>
            </form>
          </table>
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>
&nbsp; 
<%
end function

function getFormObject ()
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
set getFormObject=Request.QueryString
else
set getFormObject=Request.Form
end if
end function

function adm_error(msg)
%>
<div align="center">
  <center>
  <table border="1" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="80%">
  <tr>
  <td class="tSubTitle"><%=txtemTitleErr%> &nbsp;</td></tr>
    <tr>
      <td width="100%" class="tCellAlt1">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="center"><font class="fAlert"><%=msg%></font></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>
&nbsp;
<%
end function

function adm_alert(msg)
%>
<div align="center">
  <center>
  <table border="1" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="80%">
  <tr>
  <td class="tSubTitle"><%=txtemTitleCom%> &nbsp;</td></tr>
    <tr>
      <td width="100%" class="tCellAlt1">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="100%">
          <p align="center"><font class="fTitle"><%=msg%></font></td>
        </tr>
      </table>
      </td>
    </tr>
  </table>
  </center>
</div>
&nbsp;
<%
end function

function sendEmailSelect()
mySUBJECT = ""
myMESSAGE = ""
if request("MSG") <> "" then
	sql = "SELECT * FROM PORTAL_SPAM WHERE ID =" & chkString(request("MSG"),"numeric")
	set rsE = my_conn.execute(sql)

		if NOT rsE.eof and NOT rsE.bof then
			mySUBJECT = rsE("SUBJECT")
			myMESSAGE = rsE("MESSAGE")
			myMSG = rsE("ID")
		else
			mySUBJECT = ""
			myMESSAGE = ""
		end if

end if
%>
<div align="center">
  <center>
  <table border="1" cellpadding="3" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="90%">
  <tr>
  <td class="tSubTitle"><%=txtSndMsg%> &nbsp;</td></tr>
    <tr>
      <td width="100%" class="tCellAlt1">
      <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
      <form method="POST" action="admin_emaillist.asp" name="emCreate" id="emCreate">
        <tr>
          <td width="25%" align="right">
          <b><%=txtSubject%>: </b>
          </td>
          <td>
          <input class="textbox" type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemTagInsert%> :</b></td>
          <td>
          <% tagSelect() %>
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtMsg%>:</b></td>
          <td>
          <%adm_showEditor(myMESSAGE) %>
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemHTMLSend%>:</b></td>
          <td>
          <input type="checkbox" name="html" value="1" onmouseup="toggleEditor('Message');">
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemSaveThisMessage%>?</b></td>
          <td>
          <input type="checkbox" name="save" value="1">
           <select name="ARCHIVE" size="1">
  			<option value="0" selected="selected">&nbsp;<%=txtemLiveList%></option>
  			<option value="1">&nbsp;<%=txtemArchive%></option>
		   </select>
          </td>
        </tr>
        <tr>
          <td align="right">&nbsp;
          </td>
          <td>
          <input type="hidden" name="m_id" value="<%=request("id")%>">
          <input type="hidden" name="cmd" value="2">
          <input type="hidden" name="MSG" value="<%=MyMSG%>">
          <input type="hidden" name="MailAllMembers" value="<%=iMAilAll%>">
          <input type="submit" value="<%=txtemSendMsg%>" name="B1" class="button"> <input type="reset" value="Reset" name="B2" class="button"></td>
        </tr>
        <tr>
          <td width="100%" align="right" colspan="2" class="tSubTitle">
          <p align="center"><%=txtemSendTo%></td>
        </tr>
<%
	select case iMailAll
	
		case 0
			strSql = "select * from " & strMemberTablePrefix & "Members where MEMBER_ID in (" & request.form("id") & ")"
		case 1
			strSql = "select * from " & strMemberTablePrefix & "Members WHERE MEMBER_ID<>0"
		end select
	
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"

	set SendTo = my_Conn.execute(strSql)

	CColor = "tCellAlt2"
	if iMailAll = 0 then

do until SendTo.eof or SendTo.bof

 	if CColor = "tCellAlt1" then 
		CColor = "tCellAlt2"
	else
		CColor = "tCellAlt1"
	end if
%>
        <tr class="<%=CColor%>">
          <td align="center">
          <p align="center"><a href="cp_main.asp?cmd=8&member=<%=SendTo("MEMBER_ID")%>"><%=SendTo("M_NAME")%></a></td>
          <td>
          <a href="mailto:<%=SendTo("M_EMAIL")%>"><%=SendTo("M_EMAIL")%></a></td>
        </tr>
<%
	SendTo.MoveNext
	loop
	
SendTo.Close
set SendTo = nothing
else
%>
        <tr class="<%=CColor%>">
          <td align="center" colspan="2">
          <p align="center"><font class="fTitle"><%=txtemSendAllMembers%></font></td>
        </tr>
<% end if %>
  </form>
      </table>
      </td>
    </tr>

  </table>
  </center>
</div>
&nbsp;
<%
end function

function sendEmailSelectG()
mySUBJECT = ""
myMESSAGE = ""
	if request("MSG") <> "" then
		sql = "SELECT * FROM PORTAL_SPAM WHERE ID =" & chkString(request("MSG"),"numeric")
		set rsE = my_conn.execute(sql)

		if NOT rsE.eof and NOT rsE.bof then
			mySUBJECT = rsE("SUBJECT")
			myMESSAGE = rsE("MESSAGE")
			myMSG = rsE("ID")
		else
			mySUBJECT = ""
			myMESSAGE = ""
		end if
	end if
%>
<div align="center">
  <center>
  <table border="1" cellpadding="4" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="80%">
  <tr>
  <td class="tSubTitle"><%=txtSndMsg%> &nbsp;</td></tr>
    <tr>
      <td width="100%" class="tCellAlt2">
      <form method="post" action="admin_emaillist.asp" name="emCreate" id="emCreate">
      <table border="0" cellpadding="0" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
        <tr>
          <td width="25%" align="right">
          <b><%=txtSubject%>:</b>
          </td>
          <td>
          <input class="textbox" type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemTagInsert%> :</b></td>
          <td>
			<% tagSelect() %>
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtMsg%>:</b></td>
          <td>
          <%adm_showEditor(myMESSAGE) %>
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemHTMLSend%>:</b></td>
          <td>
          <input type="checkbox" name="html" value="1" onmouseup="toggleEditor('Message');">
          </td>
        </tr>
        <tr>
          <td align="right">
          <b><%=txtemSaveThisMessage%>?</b></td>
          <td>
          <input type="checkbox" name="save" value="1">
           <select name="ARCHIVE" size="1">
  			<option value="0" selected="selected">&nbsp;<%=txtemLiveList%></option>
  			<option value="1">&nbsp;<%=txtemArchive%></option>
		   </select>
          </td>
        </tr>
        <tr>
          <td align="right">&nbsp;
          </td>
          <td>
          <input type="hidden" name="m_id" value="<%=request("id")%>">
          <input type="hidden" name="cmd" value="5">
          <input type="hidden" name="MSG" value="<%=MyMSG%>">
          <input type="submit" value="<%=txtemSendMsg%>" name="B1" class="button"> <input type="reset" value="Reset" name="B2" class="button"></td>
        </tr>
        <tr>
          <td align="center" colspan="2" class="tSubTitle">
          <p><%=txtemSendToG%></p></td>
        </tr>
<%
  strSql = "select * from " & strMemberTablePrefix & "GROUPS where G_ID in (" & request.form("id") & ")"
  set grpInfo = my_Conn.execute(strSql)
  do until grpInfo.eof
    %>
        <tr>
          <td align="right" colspan="2">
          <p align="left"><font class="fTitle"><%=grpInfo("G_NAME")%><font></td>
        </tr>
    <%
	'strSql = "select * from " & strMemberTablePrefix & 
	'"GROUP_MEMBERS where G_GROUP_ID=" & grpInfo("G_ID")	
	'set grpSend = my_Conn.execute(strSql)	
	'do until grpSend.eof or grpSend.bof
	'strSql = "SELECT * FROM PORTAL_MEMBERS WHERE MEMBER_ID=" & grpSend("G_MEMBER_ID")
	'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1"
	'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = 1"
	'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
	
    CColor = "tCellAlt2"

	strSql = "SELECT PORTAL_MEMBERS.*"
	strSql = strSql & " FROM PORTAL_GROUP_MEMBERS INNER JOIN PORTAL_MEMBERS"
	strSql = strSql & " ON PORTAL_GROUP_MEMBERS.G_MEMBER_ID = PORTAL_MEMBERS.MEMBER_ID"
	strSql = strSql & " WHERE (((PORTAL_GROUP_MEMBERS.G_GROUP_ID)=" & grpInfo("G_ID") & ")"
	strSql = strSql & " AND ((PORTAL_MEMBERS.M_RECMAIL)=0) AND ((PORTAL_MEMBERS.M_STATUS)=1));"
	set SendTo = my_Conn.execute(strSql)
	if not SendTo.eof then
      do until SendTo.eof
 		if CColor = "tCellAlt1" then 
		  CColor = "tCellAlt2"
		else
		  CColor = "tCellAlt1"
		end if
		%>
        <tr class="<%=CColor%>">
          <td align="center">
          <p align="center"><a href="cp_main.asp?cmd=8&member=<%=SendTo("MEMBER_ID")%>"><%=SendTo("M_NAME")%></a></td>
          <td>
          <a href="mailto:<%=SendTo("M_EMAIL")%>"><%=SendTo("M_EMAIL")%></a></td>
        </tr>
	    <%
	    SendTo.MoveNext
	  loop
	end if
    SendTo.close
    set SendTo = nothing
	
    grpInfo.MoveNext
  loop
  grpInfo.close
  set grpInfo = nothing
  %>
      </table>
        </form>
      </td>
    </tr>
  </table>
  </center>
</div>
&nbsp;
<%
end function

function emailListBox()

strSql = "SELECT * FROM " &strMemberTablePrefix & "SPAM WHERE ARCHIVE = '0'"
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql, my_Conn, 3

if NOT rsSP.EOF and NOT rsSP.BOF then 
%>
<select name="MSG" size="1">
<%
do until rsSP.EOF 
%>
<option value="<% =rsSP("ID") %>">&nbsp;<% =Server.HTMLEncode(rsSP("Subject")) %></option>
<%
rsSP.MoveNext
loop 
%>
</select>&nbsp;&nbsp;
<% 
end if 
set rsSP = nothing
end function

function sendSelectedMail()
	select case iMailAll
	
	case 0
		strSql = "select * from " & strMemberTablePrefix & "Members where MEMBER_ID in (" & oFormVars("m_id") & ")"
	case 1
		strSql = "select * from " & strMemberTablePrefix & "Members WHERE MEMBER_ID<>0"
	end select
	
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
	
	set SendTo = my_Conn.execute(strSql)
	
	cnter = 0
do until SendTo.eof or SendTo.bof
	cnter = cnter + 1
	strRecipientsName = SendTo("M_NAME")
	strRecipients = SendTo("M_EMAIL")
	strFrom = strSender
	strSubject = request.form("SUBJECT")

	select case request.form("html")
	
	case 1
	 	
	 	strMessage = replaceTagData(request.form("MESSAGE"),SendTo("MEMBER_ID"))
 	 	sendOutEmail strRecipients,strSubject,strMessage,0,1
 	 	
	case else
	 	
	 	strMessage = replaceTagData(request.form("MESSAGE"),SendTo("MEMBER_ID"))
 		sendOutEmail strRecipients,strSubject,strMessage,2,0
 		
	end select

SendTo.MoveNext
loop

if request.form("save") = 1 then

strSubject = replace(request.form("SUBJECT"),"'","''")

strMessage = request.form("MESSAGE")

if request("MSG") <> "" then
strMsgID = chkString(request("MSG"),"numeric")
else
strMsgID = 0
end if

	if strMsgID <> 0 then 'update message
	
		sql = "SELECT * FROM PORTAL_SPAM WHERE ID=" & request("MSG")
		set rsMsg = my_conn.execute(sql)
	
		if rsMsg.eof or rsMsg.bof then 'insert new
			save = "INSERT INTO PORTAL_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) "
			save = save & " VALUES ("
			save = save & "'" & strSubject & "', '" & strMessage & "', '" & strCurDateString & "', " & request.form("ARCHIVE") & ")"
			'response.write(save)
			executeThis(save)
	
		else 'update
			sql = "UPDATE PORTAL_SPAM SET SUBJECT = '" & strSubject & "', MESSAGE = '" & strMessage & "', F_SENT = '" & strCurDateString & "'"
			sql = sql & " WHERE ID=" & chkString(request("MSG"),"numeric")
			executeThis(sql)
		end if
	
	else 'create new message
			save = "INSERT INTO PORTAL_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) "
			save = save & " VALUES ("
			save = save & "'" & strSubject & "', '" & strMessage & "', '" & strCurDateString & "', " & request.form("ARCHIVE") & ")"
			'response.write(save)
			executeThis(save)
	
	end if

end if
'let the user know it was sent
adm_alert(txtemMessageSent & " (" & cnter & " Emails Sent)")

end function

function sendSelectedMailG()

  strSql = "select * from " & strMemberTablePrefix & "GROUPS where G_ID in (" & oFormVars("m_id") & ")"
  set grpInfo = my_Conn.execute(strSql)

  do until grpInfo.eof
		
		'sql = "SELECT * FROM PORTAL_GROUP_MEMBERS WHERE G_GROUP_ID=" & grpInfo("G_ID")
		'set grpMembers = my_Conn.execute(sql)
	'do until grpMembers.eof or grpMembers.bof
		'strSql = "SELECT * FROM PORTAL_MEMBERS WHERE MEMBER_ID=" & grpMembers("G_MEMBER_ID")
		'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
		'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_RECMAIL = " & 0
		'strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"

	  strSql = "SELECT PORTAL_MEMBERS.*"
	  strSql = strSql & " FROM PORTAL_GROUP_MEMBERS INNER JOIN PORTAL_MEMBERS"
	  strSql = strSql & " ON PORTAL_GROUP_MEMBERS.G_MEMBER_ID = PORTAL_MEMBERS.MEMBER_ID"
	  strSql = strSql & " WHERE (((PORTAL_GROUP_MEMBERS.G_GROUP_ID)=" & grpInfo("G_ID") & ")"
	  strSql = strSql & " AND ((PORTAL_MEMBERS.M_RECMAIL)=0) AND ((PORTAL_MEMBERS.M_STATUS)=1));"
		set SendTo = my_Conn.execute(strSql)
		cnter = 0
		
	  do until SendTo.eof
		
		cnter = cnter + 1
		strRecipientsName = SendTo("M_NAME")
		strRecipients = SendTo("M_EMAIL")
		strFrom = strSender
		strSubject = request.form("SUBJECT")
		
		select case request.form("html")
		  case 1
		    strMessage = replaceTagData(request.form("MESSAGE"),SendTo("MEMBER_ID"))
 	 		sendOutEmail strRecipients,strSubject,strMessage,0,1	
		  case else
	 		strMessage = replace(request.form("MESSAGE"),"<br />",vbcrlf)
	 		strMessage = replaceTagData(strMessage,SendTo("MEMBER_ID"))
 			sendOutEmail strRecipients,strSubject,strMessage,2,0		
		end select	
		
	  SendTo.MoveNext
	loop
	SendTo.Close
	set SendTo = nothing
		
	grpInfo.MoveNExt
  loop
  grpInfo.close
  set grpInfo = nothing


if request.form("save") = 1 then

strSubject = replace(request.form("SUBJECT"),"'","''")

strMessage = request.form("MESSAGE")

	if request("MSG") <> 0 then 'update message
	
		sql = "SELECT * FROM PORTAL_SPAM WHERE ID=" & request("MSG")
		set rsMsg = my_conn.execute(sql)
	
		if rsMsg.eof or rsMsg.bof then 'insert new
			save = "INSERT INTO PORTAL_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) "
			save = save & " VALUES ("
			save = save & "'" & strSubject & "', '" & strMessage & "', '" & strCurDateString & "', " & request.form("ARCHIVE") & ")"
			'response.write(save)
			executeThis(save)
	
		else 'update
			sql = "UPDATE PORTAL_SPAM SET SUBJECT = '" & strSubject & "', MESSAGE = '" & strMessage & "', F_SENT = '" & strCurDateString & "'"
			sql = sql & " WHERE ID=" & chkString(request("MSG"),"numeric")
			executeThis(sql)
		end if
	
	else 'create new message
			save = "INSERT INTO PORTAL_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) "
			save = save & " VALUES ("
			save = save & "'" & strSubject & "', '" & strMessage & "', '" & strCurDateString & "', " & request.form("ARCHIVE") & ")"
			'response.write(save)
			executeThis(save)
	
	end if

end if
'let the user know it was sent
adm_alert(txtemMessageSent & " (" & cnter & " Emails Sent)")

end function

function adm_listGroups()
  'set page size
If Request("pagesize") = "" Then
	iPageSize = 10
Else
	iPageSize = clng(Request("pagesize"))
End If
  
  iPageCurrent = 1
  If Request("pageno") = "" Then
  iPageCurrent = 1
Else
  iPageCurrent = cLng(Request("pageno"))
End If

  strSql = "SELECT * FROM PORTAL_GROUPS WHERE G_ID<>3 AND G_ID<>2"
  Set objPagingRS = Server.CreateObject("ADODB.Recordset")
  objPagingRS.PageSize = iPageSize
  objPagingRS.CacheSize = iPageSize
  objPagingRS.Open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

  reccount = objPagingRS.recordcount
  iPageCount = objPagingRS.PageCount

  If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
  If iPageCurrent < 1 Then iPageCurrent = 1
  
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" bordercolor="#000000">

  <tr>
    <td width="100%" class="tCellAlt1">
 <table border="0" width="100%" cellspacing="1" cellpadding="4">  
 <form action="admin_emaillist.asp" method="post">
 <input type="hidden" name="cmd" value="4">
 <tr><td colspan="4">
   <p align="center" style="margin-top: 0; margin-bottom: 0">
   <%emailListBox()%>
   <p align="center" style="margin-top: 0; margin-bottom: 0">
	<input type="submit" name="action" value="<%=txtemSendMsgToSelected%>" class="button"><p align="center" style="margin-top: 0; margin-bottom: 0">
  	<%
  	':::::::::: PAGING BROKEN ::::::::::::::
  if iPageCount > 1 then
		if Request("pageno") = "" then
			sPageNumber = 1
		else
			sPageNumber = chkString(Request("pageno"),"numeric")
		end if
		'if Request.QueryString("method") = "" then
		'	sMethod = "postsdesc"
		'else
		'	sMethod = chkString(Request.QueryString("method"),"sqlstring")
		'end if

		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<form name=""PageNum"" action=""admin_emaillist.asp"">")
		Response.Write("<select name=""pageno"" size=""1"" onchange=""ChangePage()"">")
		for counter = 1 to iPageCount
			if counter <> cint(sPageNumber) then   
				Response.Write "<OPTION VALUE=""" & counter &  """>" & "Goto Page " & counter
			else
				Response.Write "<OPTION SELECTED VALUE=""" & counter &  """>" & "Page Number " & counter
			end if
		next
		Response.Write("</select></form>")
  end if
  %>
   </td></tr>
  <tr>
    <td class="tTitle"><input type="checkbox" name="CheckAll" value="1" onclick="selectAll(this.form,1)"><%= " " & txtAll%></td>
    <td class="tTitle"><%=txtemGTitleG%></td>
    <td class="tTitle"><%=txtemGTitleD%></td>
    <td class="tTitle">&nbsp;</td>
  </tr>
<%
  
  If iPageCount = 0 or objPagingRS.eof or objPagingRS.bof Then
%>
<tr>
  <td class="tCellAlt1" colspan="4">
  <p align="center"><b><%=txtNoMemFnd%></b></td>
</tr>
<%
  Else
 objPagingRS.AbsolutePage = iPageCurrent
 
 iRecordsShown = 0
 CColor = "tCellAlt2"
 
 Do While iRecordsShown < iPageSize And Not objPagingRS.EOF

 		if CColor = "tCellAlt1" then 
			CColor = "tCellAlt2"
		else
			CColor = "tCellAlt1"
		end if
%> 
  <tr class="<%=CColor%>">
    <td>
<input type="checkbox" name="ID" value="<% =objPagingRS("G_ID") %>"><input type="hidden" name="Mail_ALL" value="<% =objPagingRS("G_ID") %>">
    </td>
    <td><font class="fBold"><%=objPagingRS("G_NAME")%></font></td>
    <td><font class="fBold"><%=objPagingRS("G_DESC")%></font></td>
    <td>
   <a href="admin_config_groups.asp?mode=1&id=<% =objPagingRS("G_ID") %>"><%= icon(icnEdit,txtEditGrp,"","","") %></a>
    </td>
  </tr>
  
<%

   iRecordsShown = iRecordsShown + 1
   objPagingRS.MoveNext
 Loop
  end if
  
  objPagingRS.Close
  Set objPagingRS = Nothing
%>
</form>
</table>
    </td>
  </tr>
</table>
<%
end function

function replaceTagData(data,memberID)

sql = "SELECT * FROM PORTAL_MEMBERS WHERE MEMBER_ID=" & chkString(memberID,"numeric")
set tDm = my_conn.execute(sql)

	strMessage = data

	 	strMessage = replace(strMessage,"[%member%]",tDm("M_NAME"))
	 	strMessage = replace(strMessage,"[%memberip%]",tDm("M_LAST_IP"))
	 	strMessage = replace(strMessage,"[%memberlastlogin%]",strToDate(tDm("M_LASTHEREDATE")))
	 	strMessage = replace(strMessage,"[%memberemail%]",tDm("M_EMAIL"))
	 	strMessage = replace(strMessage,"[%sitetitle%]",strSiteTitle)

set tDM = nothing
		replaceTagData = strMessage
end function

function tagSelect()
%>
		<div id="tagSelect" align ="left">
          <!-- <select size="1" name="tagdata" id="tagdata" onchange="tinyMCE.getInstanceById('Message').execCommand('mceInsertContent',false,document.getElementById('tagdata').value);"> -->
          <select size="1" name="tagdata" id="tagdata" onchange="js_tagdata('emCreate','Message');">
          <option>- select -</option>
          <option value="[%memberlastlogin%]">Member's Last Login</option>
          <option value="[%member%]">Member's Username</option>
          <option value="[%memberemail%]">Member's Email</option>
          <option value="[%memberip%]">Member's Last IP</option>
          <option value="[%sitetitle%]">Site Title</option>
          </select>
        </div>
<%
end function

function adm_showEditor(msg)
%>
          <textarea name="Message" id="Message" cols="70" rows="20" onfocus="getActiveText(this)" onkeyup="getActiveText(this)" onselect="getActiveText(this)" onclick="getActiveText(this)" onchange="getActiveText(this)"><%= msg %></textarea>
<%    
end function
%>