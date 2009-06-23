<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Classic ASP Form Creator v1.0
' Copyright David Angell, http://www.angells.com/FormCreator
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'/**
' * SkyPortal Forms Module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://www.opensource.org/licenses/mit-license.php MIT License
' */

dim formID
dim fldCaption, fldFieldType, fldValidation, fldRequired, fldWidth, fldHeight, fldOrder, fldDefault, fldOptions, fldInfo

sub InsertBlankFormField
    strSql = "INSERT INTO " & strTablePrefix & "FORMFIELDS (FLDLINKFORMID, FLDCAPTION, FLDFIELDTYPE, FLDVALIDATION, FLDREQUIRED, FLDWIDTH, FLDHEIGHT, FLDORDER, FLDDEFAULT, FLDOPTIONS) VALUES (" & formID & ", ' ', ' ', ' ', 'N', 0, 0, 0, ' ', ' ');"
	Set dbtable = my_Conn.Execute(strSql)
end sub

sub GetFieldData (source, FieldID)
	if source = "form" then
		fldCaption = ChkString(Request.Form("fldcaption"&FieldID),"sqlstring")
		fldFieldType = ChkString(Request.Form("fldfieldtype"&FieldID),"sqlstring")
		fldValidation = ChkString(Request.Form("fldvalidation"&FieldID),"sqlstring")
		fldRequired = ChkString(Request.Form("fldrequired"&FieldID),"sqlstring")
		fldWidth = ChkString(Request.Form("fldwidth"&FieldID),"sqlstring")
		fldHeight = ChkString(Request.Form("fldheight"&FieldID),"sqlstring")
		fldOrder = ChkString(Request.Form("fldorder"&FieldID),"sqlstring")
		fldDefault = ChkString(Request.Form("flddefault"&FieldID),"sqlstring")
		fldOptions = ChkString(Request.Form("fldoptions"&FieldID),"sqlstring")
		fldInfo = ChkString(Request.Form("fldinfo"&FieldID),"message")
	else
		Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " AND ID = " & FieldID & ";")
		fldCaption = dbtable.fields("FLDCAPTION")
		fldFieldType = dbtable.fields("FLDFIELDTYPE")
		fldValidation = dbtable.fields("FLDVALIDATION")
		fldRequired = dbtable.fields("FLDREQUIRED")
		fldWidth = dbtable.fields("FLDWIDTH")
		fldHeight = dbtable.fields("FLDHEIGHT")
		fldOrder = dbtable.fields("FLDORDER")
		fldDefault = dbtable.fields("FLDDEFAULT")
		fldOptions = replace(dbtable.fields("FLDOPTIONS"),"[|]",vbcrlf)
		fldInfo = dbtable.fields("FLDOPTIONS")
		set dbtable = nothing
	end if
end sub

sub DefineFields
  msg = ""
  errFlds = ","
	if request.form("save") = "submit" then
		Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " ORDER BY FLDORDER, ID;")
		If not dbtable.EOF Then
			while not dbtable.EOF

				FieldID = dbtable.fields("ID")
				if request.form("DeleteThisField"&FieldID) = "Y" then
					Set dbtable2 = my_Conn.Execute("DELETE FROM " & strTablePrefix & "FORMFIELDS WHERE ID=" & FieldID & ";")
				else
					GetFieldData "form", FieldID
					strSql = "UPDATE " & strTablePrefix & "FORMFIELDS SET"

					msg = msg & requiredfield(fldCaption,"Field " & FieldID & " Caption")
					if len(fldCaption) > 0 then
						if len(fldCaption) > 250 then
							msg = msg & "<li>Field " & FieldID & " Caption cannot be longer than 250 characters.</li>"
						else
							strSql = strSql & " FLDCAPTION = '" & fldCaption & "',"
						end if
					end if

					msg = msg & requiredfield(fldFieldType,"Field " & FieldID & " Type")
					strSql = strSql & " FLDFIELDTYPE = '" & fldFieldType & "',"

					msg = msg & requiredfield(fldValidation,"Field " & FieldID & " Validation")
					strSql = strSql & " FLDVALIDATION = '" & fldValidation & "',"

				if isEmpty(fldRequired) then
					strSql = strSql & " FLDREQUIRED = 'N',"
				else
					strSql = strSql & " FLDREQUIRED = 'Y',"
				end if

					if fldFieldType = "Text Area" or fldFieldType = "Text Field" or fldFieldType = "Check Box" or fldFieldType = "Radio Button" then
						msg = msg & requiredfield(fldWidth,"Field " & FieldID & " Width")
						if len(fldWidth) > 0 then
							if not IsNumeric(fldWidth) then
								msg = msg & "<li>Field " & FieldID & " Width must be numeric.</li>"
							else
								fldWidth = cint(fldWidth)
								if fldWidth < 1 then
									msg = msg & "<li>Field " & FieldID & " Width must be greater then zero.</li>"
								else
									 strSql = strSql & " FLDWIDTH = " & fldWidth & ","
								end if
							end if
						end if
					else
						strSql = strSql & " FLDWIDTH = 0,"
					end if

					if fldFieldType = "Text Area" then
						msg = msg & requiredfield(fldHeight,"Field " & FieldID & " Height")
						if len(fldHeight) > 0 then
							if not IsNumeric(fldHeight) then
								msg = msg & "<li>Field " & FieldID & " Height must be numeric.</li>"
							else
								fldHeight = cint(fldHeight)
								if fldHeight < 1 then
									msg = msg & "<li>Field " & FieldID & " Height must be greater then zero.</li>"
								else
									strSql = strSql & " FLDHEIGHT = " & fldHeight & ","
								end if
							end if
						end if
					else
						strSql = strSql & " FLDHEIGHT = 0,"
					end if

					msg = msg & requiredfield(fldOrder,"Field " & FieldID & " Order")
					if len(fldOrder) > 0 then
						if not IsNumeric(fldOrder) then
							msg = msg & "<li>Field " & FieldID & " Order must be numeric.</li>"
						else
							fldOrder = cint(fldOrder)
							if fldOrder < 1 then
								msg = msg & "<li>Field " & FieldID & " Order must be greater then zero.</li>"
							else
								strSql = strSql & " FLDORDER = " & fldOrder & ","
							end if
						end if
					end if

					if len(fldDefault) < 1 then fldDefault = " "
					if len(fldDefault) > 250 then
						msg = msg & "<li>Field " & FieldID & " Default cannot be longer than 250 characters.</li>"
					else
						if fldFieldType = "Check Box" then
							'if fldDefault <> "Y" and fldDefault <> "N" and fldDefault <> " " then
								'msg = msg & "Field " & FieldID & " Default values must be either ""Y"" or ""N"" or blank for check boxes.<br />"
							'else
								strSql = strSql & " FLDDEFAULT = '" & fldDefault & "',"
							'end if
						else
							strSql = strSql & " FLDDEFAULT = '" & fldDefault & "',"
						end if
					end if

					'if len(fldOptions) < 1 then fldOptions = " "
					if len(trim(fldOptions)) > 0 then
						Dim NewOptions
						NewOptions = split(fldOptions, vbcrlf)
						fldOptions = ""
						for i = 0 to ubound(NewOptions)
							fldOptions = fldOptions & NewOptions(i) & "[|]"
						next
						fldOptions = left(fldOptions,len(fldOptions)-3)
						strSql = strSql & " FLDOPTIONS = '" & fldOptions & "'"
					else
						strSql = strSql & " FLDOPTIONS = '" & fldInfo & "'"
					end if


					strSql = strSql & " WHERE ID = " & FieldID & ";"
					'response.write strsql
					'response.end
					Set dbtable2 = my_Conn.Execute(strSql)
				end if

				dbtable.movenext
			wend

			msg = trim(msg&" ")
			if len(msg) < 1 then
			  response.redirect "admin_forms.asp"
			else
			  msg = "We encountered errors in the information you submitted.<br /><ul>" & msg & "</ul>"
			end if

		end if
	end if

	Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER WHERE ID = " & formID & ";")
	fldFormName = dbtable.fields("FLDFORMNAME")
	%>
	<script type="text/javascript">
    /**
     * SkyPortal Forms Module
     *
     * LICENSE: You may copy, modify and redistribute this work,
     *          provided that you do not remove this copyright notice
     *
     * @copyright  2008 Brandon Williams. Some Rights Reserved.
     * @license    http://www.opensource.org/licenses/mit-license.php MIT License
     */


  	function dispFields(qType, qID) {
  		var html = '';

  		var htmlWidth   = '<label for="fldWidth'+qID+'">Width</label>\n<input type="text" name="fldWidth'+qID+'" id="fldWidth'+qID+'" size="4" value=""><br />'
	  var htmlHeight  = '<label for="fldHeight'+qID+'">Height</label>\n<input type="text" name="fldHeight'+qID+'" id="fldHeight'+qID+'" size="4" value=""><br />'
	  var htmlDefault = '<label for="fldDefault'+qID+'">Default</label>\n<input type="text" name="fldDefault'+qID+'" id="fldDefault'+qID+'" size="20" value=""><br />'
	  var htmlOptions = '<label for="fldOptions'+qID+'">Options</label>\n<textarea name="fldOptions'+qID+'" id="fldOptions'+qID+'" rows="4" cols="40"></textarea> Place each item on a seperate line.<br />'
	  var htmlInfo    = '<label for="fldInfo'+qID+'">Message</label>\n<textarea name="fldInfo'+qID+'" id="fldInfo'+qID+'" rows="4" cols="40"></textarea><br />'

  		switch(qType) {
  			case 'Text Field':
  				html = htmlWidth + htmlDefault;
  				break;

  			case 'Drop Down List':
		  html = htmlWidth + htmlOptions + htmlDefault;
		  break;

		case 'Check Box':
		  html = htmlWidth + htmlOptions + htmlDefault;
		  break;

		case 'Radio Button':
		  html = htmlWidth + htmlOptions + htmlDefault;
		  break;

		case 'Text Area':
		  html = htmlWidth + htmlHeight + htmlDefault;
		  break;

		case 'Info':
		  html = htmlInfo;
		  break;

		case 'State':
		  html = htmlDefault;
		  break;

		case 'Country':
		  html = htmlDefault;
		  break;

		case 'Month':
		  html = htmlDefault;
		  break;

		case 'Day of Week':
		  html = htmlDefault;
		  break;

		case 'Year':
		  html = htmlDefault;
		  break;

		case 'Date 31':
		  html = htmlDefault;
		  break;

		case 'Date 30':
		  html = htmlDefault;
		  break;

		case 'Date 29':
		  html = htmlDefault;
		  break;

		case 'Date 28':
		  html = htmlDefault;
		  break;

  		}
  		document.getElementById('tbody' + qID).innerHTML = html;
  	}

  	function deleteField(obj, fID) {
	  var origText = obj.innerHTML;
  		obj.innerHTML = '<% =ajaxLoadingIcon %>';
	  ajaxDelete = new Ajax.Request('admin_forms.asp?ajax=DeleteFormField&field=' + fID, {
		method: 'post',
		onSuccess: function(transport) {
		  if (transport.responseText.match(/ffielddeleted/)) {
		  	document.getElementById('delete' + fID).style.display = 'none';
		  	reClassFields();
		  }
		  else {
			alert('There was an error processing your request.\nError: '+transport.responseText);
			obj.innerHTML = origText;
		  }
		},
		onFailure: function(transport) {
		  alert('There was an error processing your request.\nError: '+transport.responseText);
		  obj.innerHTML = origText;
		}
	  });
	}

	function addField() {
	  document.getElementById('AddField').disabled = true;
	  document.getElementById('submit').disabled = true;
	  document.getElementById('ajaxNotify').innerHTML = '<% =ajaxLoadingIcon %>';
	  ajaxAdd = new Ajax.Request(
		'admin_forms.asp?ajax=AddFormField&form=<%=formID%>',
		{
		  method: 'post',
		  onSuccess: function(transport) {
			newID = transport.responseText;

			if (newID.length > 0 && isNumeric(newID)) {
			  var readroot = document.getElementById('readroot').childNodes;
			  for (var z=0;z<readroot.length;z++) {
				if (readroot[z].id == 'delete') {
				  newFields = readroot[z].cloneNode(true);
				}
			  }
			  newFields.id = newFields.id + newID;

			  var newField = newFields.getElementsByTagName('input');
			  for (var i=0;i<newField.length;i++) {
				var theName = newField[i].name;
				if (theName)
				newField[i].name = theName + newID;
				var theID = newField[i].id;
				if (theID)
				newField[i].id = theID + newID;
			  }
			  newField = newFields.getElementsByTagName('select');
			  for (var i=0;i<newField.length;i++) {
				var theName = newField[i].name
				if (theName)
				newField[i].name = theName + newID;
				var theID = newField[i].id
				if (theID)
				newField[i].id = theID + newID;
			  }
			  newField = newFields.getElementsByTagName('a');
			  for (var i=0;i<newField.length;i++) {
				newField[i].name = newField[i].name + newID;
				newField[i].innerHTML = newID + ' Delete';
			  }
			  newField = newFields.getElementsByTagName('span');
			  for (var i=0;i<newField.length;i++) {
				newField[i].id = newField[i].id + newID;
			  }
			  var insertHere = document.getElementById('writeroot');
			  insertHere.parentNode.insertBefore(newFields,insertHere);
			  document.getElementById('AddField').disabled = false;
			  document.getElementById('submit').disabled = false;
			  document.getElementById('ajaxNotify').innerHTML = '';
			  reClassFields();
			}
			else {
			  alert('There was an error processing your request.\nError: '+transport.responseText);
			  document.getElementById('AddField').disabled = false;
			  document.getElementById('submit').disabled = false;
			  document.getElementById('ajaxNotify').innerHTML = '';
			}
		  },
		  onFailure: function(transport) {
			alert('There was an error processing your request.\nError: '+transport.responseText);
			document.getElementById('AddField').disabled = false;
			document.getElementById('submit').disabled = false;
			document.getElementById('ajaxNotify').innerHTML = '';
		  }
		}
	  );
	}
	
	function isNumeric(vTestValue) {
	  for(var x=0; x < vTestValue.length; x++) {
		if(vTestValue.charAt(x) >= 0 && vTestValue.charAt(x) <= 9) {
		}
		else {
		  return false;
		}
	  }
	  return true;
	}
	
	function reClassFields() {
	  var flds = document.getElementById('fields').childNodes;
	  var class1 = 'tCellAlt0';
	  var class2 = 'tCellAlt1';
	  var curClass = class1;
	  
	  for (i=0;i<flds.length;i++) {
		if (flds[i].id && flds[i].id.substr(0,6) == 'delete') {
		  if (flds[i].style.display != 'none') {
			flds[i].className = curClass;
			curClass == class1 ? curClass = class2:curClass = class1;
		  }
		}
	  }

	}
	</script>
	<style type="text/css">
    /**
     * SkyPortal Forms Module
     *
     * LICENSE: You may copy, modify and redistribute this work,
     *          provided that you do not remove this copyright notice
     *
     * @copyright  2008 Brandon Williams. Some Rights Reserved.
     * @license    http://www.opensource.org/licenses/mit-license.php MIT License
     */


	#FORMFIELDS label{
	  float: left;
	  font-weight: bold;
	  padding-right: 10px;
	  text-align: right;
	  width: 100px;
	}

	#FORMFIELDS input,
	#FORMFIELDS textarea,
	#FORMFIELDS select{
	  margin-bottom: 5px;
	  width: 180px;
	}

	#FORMFIELDS textarea{
	  height: 150px;
	  width: 250px;
	}

	#FORMFIELDS .boxes{
	  width: 1em;
	}

	#FORMFIELDS br{
	  clear: left;
	}
  </style>
	<form method="post" name="FORMFIELDS" id="FORMFIELDS" action="admin_forms.asp?action=DefineFields&form=<%=formID%>">
		<table cellspacing=0>
			<tr>
				<td colspan=5 align="center"><b><%=fldFormName%></b></td>
			</tr>
			<tr>
			  <td colspan="5">&nbsp;</td>
			</tr>
			<tr>
			  <td colspan="5">Changes submitted are immediate.  Be careful editing live forms.</td>
			</tr>
			<tr>
			  <td colspan="5">&nbsp;</td>
			</tr>
			<tr>
			  <td colspan="5" align="left" class="fAlert"><% = msg %></td>
			</tr>
			<tr>
			  <td colspan="5">&nbsp;</td>
			</tr>
			<tr>
			<td id="fields">
			<%
			Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " ORDER BY FLDORDER, ID;")
			If not dbtable.EOF Then
				while not dbtable.EOF
					FieldID = dbtable.fields("ID")
					if request.form("save") = "Submit" then
						GetFieldData "form", FieldID
					else
						GetFieldData "db", FieldID
					end if
					if fClass = "tCellAlt0" then
						fClass = "tCellAlt1"
					else
						fClass = "tCellAlt0"
					end if

					htmlWidth = "<label for=""fldWidth" & FieldID & """>Width</label>" &_
								"<input type=""text"" name=""fldWidth" & FieldID & """ id=""fldWidth" & FieldID & """ size=""4"" value=""" & fldWidth & """><br />"

					htmlHeight = "<label for=""fldHeight" & FieldID & """>Height</label>" &_
								 "<input type=""text"" name=""fldHeight" & FieldID & """ id=""fldHeight" & FieldID & """ size=""4"" value=""" & fldHeight & """><br />"

					htmlDefault = "<label for=""fldDefault" & FieldID & """>Default</label>" &_
								  "<input type=""text"" name=""fldDefault" & FieldID & """ id=""fldDefault" & FieldID & """ size=""20"" value=""" & fldDefault & """><br />"

					htmlOptions = "<label for=""fldOptions" & FieldID & """>Options</label>" &_
								  "<textarea name=""fldOptions" & FieldID & """ id=""fldOptions" & FieldID & """ rows=""4"" cols=""40"">" & fldOptions & "</textarea> Place each item on a seperate line.<br />"

					htmlInfo = "<label for=""fldInfo" & FieldID & """>Message</label>" &_
							   "<textarea name=""fldInfo" & FieldID & """ id=""fldInfo" & FieldID & """ rows=""4"" cols=""40"">" & fldOptions & "</textarea><br />"


										Response.Write "<div id=""delete" & FieldID & """ class=""" & fClass & """><a href=""javascript:;"" onClick=""deleteField(this, '" & FieldID & "')"">" & FieldID & " Delete</a><br />"
					Response.Write "<label for=""fldRequired" & FieldID & """>Required</label><input type=""checkbox"" name=""fldRequired" & FieldID & """ id=""fldRequired" & FieldID & """ class=""boxes"" value=""Y"" " & chkCheckbox(fldRequired, "Y", true) & "><br />"
					Response.Write "<label for=""fldOrder" & FieldID & """>Order</label><input type=""text"" name=""fldOrder" & FieldID & """ id=""fldOrder" & FieldID & """ size=""3"" value=""" & fldOrder & """><br />"

										Response.Write "<label for=""fldCaption" & FieldID & """>Caption</label>"
										Response.Write "<input type=""text"" name=""fldCaption" & FieldID & """ id=""fldCaption" & FieldID & """ size=""50"" value=""" & fldCaption & """><br />"


					Response.Write "<label for=""fldFieldType" & FieldID & """>Type</label>"
						Response.Write "<select size=""1"" name=""fldFieldType" & FieldID & """ id=""fldFieldType" & FieldID & """ onChange=""dispFields(this.value, " & FieldID & ")"">"
							for i = 0 to ubound(FieldTypeArray)
								Response.Write "<option value=""" & FieldTypeArray(i) & """"
									if fldFieldType = FieldTypeArray(i) then response.write " SELECTED"
									Response.Write ">" & FieldTypeArray(i) & "</option>"
							next
						Response.Write "</select><br />"

					Response.Write "<label for=""fldValidation" & FieldID & """>Validation</label>"
						Response.Write "<select size=""1"" name=""fldValidation" & FieldID & """ id=""fldValidation" & FieldID & """>"
							for i = 0 to ubound(FieldValidationArray)
								Response.Write "<option value=""" & FieldValidationArray(i) & """"
									if fldValidation = FieldValidationArray(i) then response.write " SELECTED"
									Response.Write ">" & FieldValidationArray(i) & "</option>"
							next

						Response.Write "</select><br />"

										Response.Write "<span id=""tbody" & FieldID & """>"

					if fldFieldType <> "" then
					  Select Case fldFieldType
						case "Text Field"
						  Response.Write htmlWidth & htmlDefault

						case "Drop Down List"
						  Response.Write htmlWidth & htmlOptions & htmlDefault

						case "Check Box"
						  Response.Write htmlWidth & htmlOptions & htmlDefault

						case "Radio Button"
						  Response.Write htmlWidth & htmlOptions & htmlDefault

						case "Text Area"
						  Response.Write htmlWidth & htmlHeight & htmlDefault

						case "Info"
						  Response.Write htmlInfo

						case "State"
						  Response.Write htmlDefault

						case "Country"
						  Response.Write htmlDefault

						case "Day of Week"
						  Response.Write htmlDefault

						case "Month"
						  Response.Write htmlDefault

						case "Year"
						  Response.Write htmlDefault

						case "Date 31"
						  Response.Write htmlDefault

						case "Date 30"
						  Response.Write htmlDefault

						case "Date 29"
						  Response.Write htmlDefault

						case "Date 28"
						  Response.Write htmlDefault

						case else
						  Response.Write htmlWidth & htmlDefault 'Show info for a text field

					  End Select
					else
					  Response.Write htmlWidth & htmlDefault
					end if

				Response.Write "</span></div>"

					dbtable.MoveNext
				wend
			end if
			%>
			<span id="writeroot" style="display:none;"></span><span style="display:none;"></span></td></tr>
			<tr>
			   <td colspan="5">&nbsp;</td>
			</tr>
			<tr>
				<td colspan=5 align="left">
				  <input type="hidden" name="save" id="save" value="submit" />
				  <button type="submit" class="button" name="submit" id="submit">Submit</button>&nbsp;
				  <button type="button" onClick="addField()" name="AddField" id="AddField" class="button">Add Field</button>&nbsp;<span id="ajaxNotify"></span>
				</td>
			</tr>
		</table>
	</form>
	<div id="readroot" style="display:none;">
	  <div id="delete" class="tCellAlt0">
		<a href="javascript:;" onClick="deleteField(this, this.name)"></a><br />

		<label for="fldRequired">Required</label>
		<input type="checkbox" name="fldRequired" id="fldRequired" class="boxes" value="Y" ><br />

		<label for="fldOrder">Order</label>
		<input type="text" name="fldOrder" id="fldOrder" size="3" value=""><br />

		<label for="fldCaption">Caption</label>
		<input type="text" name="fldCaption" id="fldCaption" size="50" value=""><br />

		<label for="fldFieldType">Type</label>
		<select size="1" name="fldFieldType" id="fldFieldType" onChange="dispFields(this.value, this.id.substring(12, this.id.length))">
		  <%
			for i = 0 to ubound(FieldTypeArray)
				Response.Write "<option value=""" & FieldTypeArray(i) & """"
				Response.Write ">" & FieldTypeArray(i) & "</option>"
			next
		  %>
		</select><br />

		<label for="fldValidation">Validation</label>
		<select size="1" name="fldValidation" id="fldValidation">
		  <%
			for i = 0 to ubound(FieldValidationArray)
				Response.Write "<option value=""" & FieldValidationArray(i) & """"
				Response.Write ">" & FieldValidationArray(i) & "</option>"
		   next
		  %>
		</select><br />

		<span id="tbody">
		  <label for="fldWidth">Width</label>
		  <input type="text" name="fldWidth" id="fldWidth" size="4" value=""><br />

		  <label for="fldDefault">Default</label>
		  <input type="text" name="fldDefault" id="fldDefault" size="20" value=""><br />
		</span>
	  </div>
	</div>
<%End Sub

sub NewForm
	if request.form("save") = "Submit" then
		if Request.Form("fldActive") = "1" then
		  fldActive = 1
		else
		  fldActive = 0
		end if
		if Request.Form("advSendPM") = "1" then
		  fldSendPM = 1
		else
		  fldSendPM = 0
		end if
		if Request.Form("advSendEmail") = "1" then
		  fldSendEmail = 1
		else
		  fldSendEmail = 0
		end if
		fldFormName = ChkString(Request.Form("fldFormName"),"sqlstring")
		fldRecipientEmail = ChkString(Request.Form("fldRecipientEmail"),"sqlstring")
		fldSendTo = ChkString(Request.Form("SendTo"),"sqlstring")
		fldEmailSubject = ChkString(Request.Form("fldEmailSubject"),"sqlstring")
		fldNumFields = trim(request.form("fldNumFields"))
		fldIntroText = ChkString(Request.Form("fldIntroText"),"message")
		fldThankYou = ChkString(Request.Form("fldThankYou"),"message")
		fldInactiveText = ChkString(Request.Form("fldInactiveText"),"message")
		if len(fldIntroText) < 1 then fldIntroText = " "
		msg = ""
		msg = msg & requiredfield(fldFormName,"Form Name")
		'msg = msg & requiredfield(fldRecipientEmail,"Recipient E-mail")
		msg = msg & requiredfield(fldEmailSubject,"Subject Line")
		msg = msg & requiredfield(fldNumFields,"# of Fields")
		msg = msg & requiredfield(fldThankYou,"Thank You Text")
		msg = msg & requiredfield(fldInactiveText,"Inactive Text")
		if len(fldIntroText) < 1 then fldIntroText = " "
		if len(fldFormName) > 50 then msg = msg & "<li>Form name cannot be more than 50 characters.</li>"
		if len(fldRecipientEmail) > 0 then if not IsValidEmail(fldRecipientEmail) then msg = msg & "<li>Invalid e-mail address.</li>"
		if len(fldRecipientEmail) > 250 then msg = msg & "<li>Recipient e-mail cannot be more than 250 characters.</li>"
		if (fldSendEmail = 0) and (fldSendPM = 0) then msg = msg & "<li>You must send a PM or Email.</li>"
		if (fldSendPM = 1) and (len(fldSendTo) < 1) then msg = msg & "<li>You must choose a member to PM</li>"
		if (fldSendEmail = 1) and (len(fldSendTo) < 1) and (len(fldRecipientEmail) < 1) then msg = msg & "<li>You must choose a member to email, or enter in an additional email address</li>"
		if (len(fldSendTo) > 0) then
			arrSendTo = split(fldSendTo, ",")
			for i=0 to uBound(arrSendTo)
				if getMemberID(trim(arrSendTo(i))) = 0 then
					msg = msg & "<li>" & arrSendTo(i) & " is not a member</li>"
				end if
				i = i + 1
			next
		end if
		if len(fldNumFields) > 0 then
			if not IsNumeric(fldNumFields) then
				msg = msg & "<li># of fields must be numeric.</li>"
			else
				fldNumFields = cint(fldNumFields)
				if fldNumFields < 1 then msg = msg & "<li># of fields must be greater then zero.</li>"
			end if
		end if
		msg = trim(msg&" ")
		if len(msg) < 1 then
			strSql = "INSERT INTO " & strTablePrefix & "FORMHEADER (FLDFORMNAME, FLDRECIPIENTEMAIL, FLDEMAILSUBJECT, FLDINTROTEXT, FLDTHANKYOU, FLDINACTIVETEXT, ACTIVE, SENDEMAIL, SENDPM, SENDTO) VALUES ('" & fldFormName & "', '" & fldRecipientEmail & "', '" & fldEmailSubject & "', '" & fldIntroText & "', '" & fldThankYou & "', '" & fldInactiveText & "', " & fldActive & ", " & fldSendEmail & ", " & fldSendPM & ", '" & fldSendTo & "');"
			Set dbtable = my_Conn.Execute(strSql)
			Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER ORDER BY ID DESC;")
			formID = dbtable.fields("ID")
			for i = 1 to fldNumFields
				InsertBlankFormField
			next
			response.redirect "admin_forms.asp?action=DefineFields&form=" & formID
		end if
	else
	  fldInactiveText = "<div align=""center"">We&#39;re sorry, this form is no longer available.<br /></div>"
	end if
	response.write "<div><ul>" & msg & "</ul></div>"
	%>
	<script type="text/javascript">
	function memberlist() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=pm", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }
	</script>
	<form method="post" name="PostTopic" action="admin_forms.asp?action=NewForm">
		<table border="0">
			<tr>
				<td width="100" align="right"><label for="fldActive">Active</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="checkbox" name="fldActive" id="fldActive" value="1" <% =chkCheckbox(fldActive, 1, true) %>></td>
			</tr>
			<tr>
				<td align="right"><label for="fldFormName">Form Name</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldFormName" id="fldFormName" size="30" value="<%=fldFormName%>"></td>
			</tr>
			<tr>
				<td align="right">Send</td>
				<td>&nbsp;</td>
				<td align="left"><input type="checkbox" name="advSendPM" id="advSendPM" value="1" <% =chkCheckbox(fldSendPM, 1, true) %> />&nbsp;<label for="advSendPM">PM</label>&nbsp;<input type="checkbox" name="advSendEmail" id="advSendEmail" value="1" <% =chkCheckbox(fldSendEmail, 1, true) %> />&nbsp;<label for="advSendEmail">Email</label></td>
			</tr>
			<tr>
				<td align="right">To</td>
				<td>&nbsp;</td>
				<td align="left">
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>
								<a href="JavaScript:memberlist();"><%=strSiteTitle%>&nbsp;Member(s)</a>:<br />
								<input type="text" name="SendTo" id="SendTo" size="30" value="<% =fldSendTo %>" /><br />
								<label for="SendTo">(seperate members with a comma)</label>
							</td>
						</tr>
						<tr>
							<td align="center"><br />and/or<br /><br /></td>
						</tr>
						<tr>
							<td align="left">
							<label for="fldRecipientEmail">Additional e-mail addresses:<br /></label>
							<input type="text" name="fldRecipientEmail" id="fldRecipientEmail" size="30" value="<%=fldRecipientEmail%>">
						</tr>
				</td>

					</table>
				</td>
			</tr>
			<tr>
				<td align="right"><label for="fldEmailSubject">Subject Line</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldEmailSubject" id="fldEmailSubject" size="30" value="<%=fldEmailSubject%>"></td>
			</tr>
			<tr>
				<td align="right"><label for="fldNumFields"><b># of Fields</b></label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldNumFields" id="fldNumFields" size="30" value="<%=fldNumFields%>"></td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><label for="fldIntroText"><b>Introduction Text</b></label><br>
					<textarea name="fldIntroText" id="fldIntroText" rows="15" cols="70"><%=fldIntroText%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;<br><label for="fldThankYou"><b>Thank You Text</b></label><br>
					<textarea name="fldThankYou" id="fldThankYou" rows="15" cols="70"><%=fldThankYou%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;<br><label for="fldInactiveText"><b>Inactive Text</b></label><br>
					<textarea name="fldInactiveText" id="fldInactiveText" rows="15" cols="70"><%=fldInactiveText%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right"><input type="submit" value="Submit" name="save" id="save" class="button" /></td>
			</tr>
		</table>
	</form>
<%End Sub

sub EditForm
	if len(formID) < 1 then response.redirect "admin_forms.asp"
	if request.form("save") = "Submit" then
		if Request.Form("fldActive") = "1" then
		  fldActive = 1
		else
		  fldActive = 0
		end if
		if Request.Form("advSendPM") = "1" then
		  fldSendPM = 1
		else
		  fldSendPM = 0
		end if
		if Request.Form("advSendEmail") = "1" then
		  fldSendEmail = 1
		else
		  fldSendEmail = 0
		end if
		fldFormName = ChkString(Request.Form("fldFormName"),"sqlstring")
		fldRecipientEmail = ChkString(Request.Form("fldRecipientEmail"),"sqlstring")
		fldSendTo = ChkString(Request.Form("SendTo"),"sqlstring")
		fldEmailSubject = ChkString(Request.Form("fldEmailSubject"),"sqlstring")
		fldIntroText = ChkString(Request.Form("fldIntroText"),"message")
		fldThankYou = ChkString(Request.Form("fldThankYou"),"message")
		fldInactiveText = ChkString(Request.Form("fldInactiveText"),"message")
		msg = ""
		msg = msg & requiredfield(fldFormName,"Form Name")
		'msg = msg & requiredfield(fldRecipientEmail,"Recipient E-mail")
		msg = msg & requiredfield(fldEmailSubject,"Subject Line")
		msg = msg & requiredfield(fldThankYou,"Thank You Text")
		msg = msg & requiredfield(fldInactiveText,"Inactive Text")
		if len(fldIntroText) < 1 then fldIntroText = " "
		if len(fldFormName) > 50 then msg = msg & "<li>Form name cannot be more than 50 characters.</li>"
		if len(fldRecipientEmail) > 0 then if not IsValidEmail(fldRecipientEmail) then msg = msg & "<li>Invalid e-mail address.</li>"
		if len(fldRecipientEmail) > 250 then msg = msg & "<li>Recipient e-mail cannot be more than 250 characters.</li>"
		if (fldSendEmail = 0) and (fldSendPM = 0) then msg = msg & "<li>You must send a PM or Email.</li>"
		if (fldSendPM = 1) and (len(fldSendTo) < 1) then msg = msg & "<li>You must choose a member to PM</li>"
		if (fldSendEmail = 1) and (len(fldSendTo) < 1) and (len(fldRecipientEmail) < 1) then msg = msg & "<li>You must choose a member to email, or enter in an additional email address</li>"
		if (len(fldSendTo) > 0) then
			arrSendTo = split(fldSendTo, ",")
			for i=0 to uBound(arrSendTo)
				if getMemberID(trim(arrSendTo(i))) = 0 then
					msg = msg & "<li>" & arrSendTo(i) & " is not a member</li>"
				end if
				i = i + 1
			next
		end if
		msg = trim(msg&" ")
		if len(msg) < 1 then
			strSql = "UPDATE " & strTablePrefix & "FORMHEADER SET "
			strSql = strSql & "FLDFORMNAME = '" & fldFormName & "', "
			strSql = strSql & "FLDRECIPIENTEMAIL = '" & fldRecipientEmail & "', "
			strSql = strSql & "FLDEMAILSUBJECT = '" & fldEmailSubject & "', "
			strSql = strSql & "FLDINTROTEXT = '" & fldIntroText & "', "
			strSql = strSql & "FLDTHANKYOU = '" & fldThankYou & "', "
			strSql = strSql & "FLDINACTIVETEXT = '" & fldInactiveText & "', "
			strSql = strSql & "ACTIVE = " & fldActive & ", "
			strSql = strSql & "SENDEMAIL = " & fldSendEmail & ", "
			strSql = strSql & "SENDPM = " & fldSendPM & ", "
			strSql = strSql & "SENDTO = '" & fldSendTo & "' "
			strSql = strSql & "WHERE ID = " & formID & ";"
			Set dbtable = my_Conn.Execute(strSql)
			response.redirect "admin_forms.asp?action=DefineFields&form=" & formID
		end if
	else
		Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER WHERE ID = " & formID & ";")
		fldActive = dbtable.fields("ACTIVE")
		fldSendEmail = dbtable.fields("SENDEMAIL")
		fldSendPM = dbtable.fields("SENDPM")
		fldSendTo = dbtable.fields("SENDTO")
		fldFormName = dbtable.fields("FLDFORMNAME")
		fldRecipientEmail = dbtable.fields("FLDRECIPIENTEMAIL")
		fldEmailSubject = dbtable.fields("FLDEMAILSUBJECT")
		fldIntroText = dbtable.fields("FLDINTROTEXT")
		fldThankYou = dbtable.fields("FLDTHANKYOU")
		fldInactiveText = dbtable.fields("FLDINACTIVETEXT")
	end if
	response.write "<div><ul>" & msg & "</ul></div>"
	response.write ""
	%>
	<script type="text/javascript">
	function memberlist() { var MainWindow = window.open ("pop_memberlist.asp?pageMode=pm", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }
	</script>
	<form method="post" name="PostTopic" action="admin_forms.asp?action=EditForm&form=<%=formID%>">
		<table border="0">
			<tr>
				<td width="100" align="right"><label for="fldActive">Active</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="checkbox" name="fldActive" id="fldActive" value="1" <% =chkCheckbox(fldActive, 1, true) %>></td>
			</tr>
			<tr>
				<td align="right"><label for="fldFormName">Form Name</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldFormName" id="fldFormName" size="30" value="<%=fldFormName%>"></td>
			</tr>
			<tr>
				<td align="right">Send</td>
				<td>&nbsp;</td>
				<td align="left"><input type="checkbox" name="advSendPM" id="advSendPM" value="1" <% =chkCheckbox(fldSendPM, 1, true) %> />&nbsp;<label for="advSendPM">PM</label>&nbsp;<input type="checkbox" name="advSendEmail" id="advSendEmail" value="1" <% =chkCheckbox(fldSendEmail, 1, true) %> />&nbsp;<label for="advSendEmail">Email</label></td>
			</tr>
			<tr>
				<td align="right"><label for="SendTo">To</label></td>
				<td>&nbsp;</td>
				<td align="left">
					<input type="text" name="SendTo" id="SendTo" size="30" value="<% =fldSendTo %>" />&nbsp;<a href="JavaScript:memberlist();">Select Member(s)</a><br />
					(seperate members with a comma)
				</td>
			</tr>
			<tr>
				<td align="right"><label for="fldRecipientEmail">Additional e-mail addresses</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldRecipientEmail" id="fldRecipientEmail" size="30" value="<%=fldRecipientEmail%>"></td>
			</tr>
			<tr>
				<td align="right"><label for="fldEmailSubject">Subject Line</label></td>
				<td>&nbsp;</td>
				<td align="left"><input type="text" name="fldEmailSubject" id="fldEmailSubject" size="30" value="<%=fldEmailSubject%>"></td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="3"><label for="fldIntroText"><b>Introduction Text</b></label><br>
					<textarea name="fldIntroText" id="fldIntroText" rows="15" cols="70"><%=fldIntroText%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;<br><label for="fldThankYou"><b>Thank You Text</b></label><br>
					<textarea name="fldThankYou" id="fldThankYou" rows="15" cols="70"><%=fldThankYou%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3">&nbsp;<br><label for="fldInactiveText"><b>Inactive Text</b></label><br>
					<textarea name="fldInactiveText" id="fldInactiveText" rows="15" cols="70"><%=fldInactiveText%></textarea>
				</td>
			</tr>
			<tr>
				<td colspan="3" align="right"><input type="submit" value="Submit" name="save" id="save" class="button" /></td>
			</tr>
		</table>
	</form>
<%End Sub

sub DeleteForm
	if request.querystring ("Confirm") = "Y" then
		Set dbtable = my_Conn.Execute("DELETE FROM " & strTablePrefix & "FORMHEADER WHERE ID=" & FormID & ";")
		Set dbtable = my_Conn.Execute("DELETE FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID=" & FormID & ";")
		response.redirect "admin_forms.asp"
	else
		Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER WHERE ID = " & formID & ";")
		Response.write "<p><b>" & dbtable.fields("FLDFORMNAME") & "</b> results to go to <b>" & dbtable.fields("FLDRECIPIENTEMAIL") & "</b>.</p>"
		response.write "<p>Are you certain you want to delete this form?<br>"
		response.write "<a href=""admin_forms.asp?action=DeleteForm&form=" & formID & "&Confirm=Y"">Yes, delete it.</a>"
		Response.Write " • <a href=""admin_forms.asp"">No, return to Forms Manager.</a></p>"
	end if

end sub

sub CopyForm
	Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER WHERE ID = " & formID & ";")
	strSql = "INSERT INTO " & strTablePrefix & "FORMHEADER (FLDFORMNAME, FLDRECIPIENTEMAIL, FLDEMAILSUBJECT, FLDINTROTEXT, FLDTHANKYOU, FLDINACTIVETEXT, ACTIVE) VALUES ('Copy of " & dbtable.fields("FLDFORMNAME") & "', '" & dbtable.fields("FLDRECIPIENTEMAIL") & "', '" & dbtable.fields("FLDEMAILSUBJECT") & "', '" & dbtable.fields("FLDINTROTEXT") & "', '" & dbtable.fields("FLDTHANKYOU") & "', '" & dbtable.fields("FLDINACTIVETEXT") & "', " & dbtable.fields("ACTIVE") & ");"
	Set dbtable = my_Conn.Execute(strSql)

	Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER ORDER BY ID DESC;")
	NewformID = dbtable.fields("ID")

	Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " ORDER BY FLDORDER,ID;")
	while not dbtable.eof
		strSql = "INSERT INTO " & strTablePrefix & "FORMFIELDS ("
		strSql = strSql & "FLDLINKFORMID, FLDCAPTION, FLDFIELDTYPE, FLDVALIDATION, FLDREQUIRED, FLDWIDTH, FLDHEIGHT, FLDORDER, FLDDEFAULT, FLDOPTIONS"
		strSql = strSql & ") VALUES ("
		strSql = strSql & NewformID & ", '" & dbtable.fields("FLDCAPTION") & "', '" & dbtable.fields("FLDFIELDTYPE") & "', '" & dbtable.fields("FLDVALIDATION") & "', "
		strSql = strSql & "'" & dbtable.fields("FLDREQUIRED") & "', " & dbtable.fields("FLDWIDTH") & ", " & dbtable.fields("FLDHEIGHT") & ", " & dbtable.fields("FLDORDER") & ", "
		strSql = strSql & "'" & dbtable.fields("FLDDEFAULT") & "', '" & dbtable.fields("FLDOPTIONS") & "'"
		strSql = strSql & ");"
		Set dbtable2 = my_Conn.Execute(strSql)
		dbtable.movenext
	wend

	response.redirect "admin_forms.asp?action=EditForm&form=" & NewformID
end sub


sub ShowForms
	Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER ORDER BY ID;")
	If not dbtable.EOF Then
		Response.Write "<center><table cellpadding=""2"" cellspacing=""0"" width=""100%"">"
		Response.Write "<tr><td colspan=""2"">Please select a form...</td></tr>"
		Response.Write "<tr><td>Form ID</td><td>Form Name</td></tr>"
		while not dbtable.EOF
			if fClass = "tCellAlt0" then
				fClass = "tCellAlt1"
			else
				fClass = "tCellAlt0"
			end if
			Response.Write "<tr class=""" & fClass & """>"
			Response.Write "<td>" & dbtable.fields("ID") & "</td>"
			Response.Write "<td><a href=""admin_forms.asp?action=" & Request.Querystring("next") & "&form=" & dbtable.fields("ID") & """>" & dbtable.fields("FLDFORMNAME") & "</a></td>"
			Response.Write "</tr>"
			dbtable.MoveNext
		wend
		Response.Write "</table></center>"
	else
		response.write "<center>No Forms</center>"
	end if
end sub

sub FormSummary
  fClass = "tCellAlt1" %>
  <script type="text/javascript">
    /**
     * SkyPortal Forms Module
     *
     * LICENSE: You may copy, modify and redistribute this work,
     *          provided that you do not remove this copyright notice
     *
     * @copyright  2008 Brandon Williams. Some Rights Reserved.
     * @license    http://www.opensource.org/licenses/mit-license.php MIT License
     */
	 
	function deleteForm(obj,id,txt) {
	  var origText = obj.innerHTML;
	  var areYouSure = confirm('Do you really want to delete "' + txt + '"?');
	  if (areYouSure) {
		obj.innerHTML = '<% =ajaxLoadingIcon %>';
		ajaxDelete = new Ajax.Request('admin_forms.asp?ajax=DeleteForm&Form=' + id, {
			 method: 'post',
			 onSuccess: function(transport) {
				  if (transport.responseText.match(/fdeleted/)) {
				  	document.getElementById('ftr' + id).style.display = 'none';
				  	reClassTable('FormSummary','tCellAlt0','tCellAlt1',true);
				  }
				  else {
					alert('There was an error processing your request.\nError: '+transport.responseText);
					obj.innerHTML = origText;
				  }
			 },
			 onFailure: function(transport) {
			  alert('There was an error processing your request.\nError: '+transport.responseText);
			  obj.innerHTML = origText;
			}
		});
	  }
	}

	function reClassTable(tblID,class1,class2,skipFirstRow) {
	  var table = document.getElementById(tblID);
	  var rows = table.getElementsByTagName('tr');
	  var curClass = class1;

	  skipFirstRow ? i=1:i=0;

	  for (i;i<rows.length;i++) {
		if (rows[i].style.display != 'none' && rows[i].style.display != 'collapse') {
		  rows[i].className = curClass;
		  curClass == class1 ? curClass = class2:curClass = class1;
		}
	  }
	}
  </script>
  <p><a href="admin_forms.asp?action=NewForm"><% = icon(icnPlus,txtAdd,"vertical-align:text-bottom;","","")%>&nbsp;Add New Form</a></p>
  <%
  Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER ORDER BY ID;")
  If not dbtable.EOF Then

	Response.Write "<center><table cellpadding=""2"" cellspacing=""0"" width=""100%"" id=""FormSummary"">"
	Response.Write "<tr><th width=""100"">&nbsp;</th><th>Form ID</th><th>Form Name</th></tr>" '<th>Hits</th><th>Submits</th></tr>"
	while not dbtable.EOF
		if fClass = "tCellAlt0" then
			fClass = "tCellAlt1"
		else
			fClass = "tCellAlt0"
		end if
		Response.Write "<tr class=""" & fClass & """ id=""ftr" & dbtable.fields("ID") & """>"

		Response.write "<td>" &_
						"<a href=""#"" onclick=""deleteForm(this," & dbtable.fields("ID") & ",'" & Replace(dbtable.fields("FLDFORMNAME"),"&#39;","\'") & "'); return false;"">" & icon(icnDelete,txtDel,"","","") & "</a>" &_
						"<a href=""admin_forms.asp?action=EditForm&form=" & dbtable.fields("ID") & """>" & icon(icnEdit,txtEdit,"","","") & "</a>" &_
						"<a href=""admin_forms.asp?action=CopyForm&form=" & dbtable.fields("ID") & """>" & icon(icnCopy,txtCopy,"","","") & "</a>" &_
						"<a href=""form.asp?form=" & dbtable.fields("ID") & """ target=""_blank"">" & icon(icnBinoc,txtView,"","","") & "</a>" &_
						"</td>"
		Response.Write "<td>" & dbtable.fields("ID") & "</td>"
		Response.Write "<td>" & dbtable.fields("FLDFORMNAME") & "</td>"
		'Response.Write "<td>0</td>"
		'Response.Write "<td>0</td>"
		Response.Write "</tr>"
		dbtable.MoveNext
	wend
	Response.Write "</table></center>"
  else
		response.write "<center>No Forms</center>"
	end if
end sub

sub config_forms()
end sub

Dim statename, countryname, fm_monthname, dayofweekname, yearname, datename31, datename30, datename29, datename28, markRequired, FieldTypeArray, FieldValidationArray, ajaxLoadingIcon
FieldTypeArray = array ("Text Field", "Text Area", "Info", "Drop Down List", "Check Box", "Radio Button", "State", "Country", "Date Picker", "Month", "Day of Week", "Year", "Date 31", "Date 30", "Date 29", "Date 28")
FieldValidationArray = array ("None", "Numeric", "E-mail", "Date", "Phone Number", "Zip Code")
statename = Array("Alaska", "Alabama", "Arkansas", "Arizona", "California", "Colorado", "Connecticut", "District of Columbia", "Delaware", "Florida", "Georgia", "Hawaii", "Iowa", "Idaho", "Illinois", "Indiana", "Kansas", "Kentucky", "Louisiana", "Massachusetts", "Maryland", "Maine", "Michigan", "Minnesota", "Missouri", "Mississippi", "Montana", "North Carolina", "North Dakota", "Nebraska", "New Hampshire", "New Jersey", "New Mexico", "Nevada", "New York", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Vermont", "Washington", "Wisconsin", "West Virginia", "Wyoming", "American Samoa", "Federated States of Micronesia", "Guam", "Marshall Islands", "Northern Mariana Islands", "Palau", "Puerto Rico", "Virgin Islands", "Armed Forces Africa", "Armed Forces Americas", "Armed Forces Canada", "Armed Forces Europe", "Armed Forces Middle East", "Armed Forces Pacific", "Outside USA")
countryname = Array("Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei", "Bulgaria", "Burkina Faso", "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde", "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo (Brazzaville)", "Congo (Democratic Republic)", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica", "Dominican Republic", "East Timor", "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Ethiopia", "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Honduras", "Hungary", "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Korea (North)", "Korea (South)", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia", "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Macedonia (Former Yugoslav Republic)", "Madagascar", "Malawi", "Malaysia", "Maldives", "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia (Federated States)", "Moldova", "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar (Burma)", "Namibia", "Nauru", "Nepal", "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "Norway", "Oman", "Pakistan", "Palau", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia", "Saint Vincent and The Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia", "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands", "Somalia", "South Africa", "Spain", "Sri Lanka", "Sudan", "Suriname", "Swaziland", "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Togo", "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom", "United States of America", "Uruguay", "Uzbekistan", "Vanuatu", "Vatican City", "Venezuela", "Vietnam", "Western Sahara", "Yemen", "Zambia", "Zimbabwe")
fm_monthname = Array(txtJanuary, txtFebruary, txtMarch, txtApril, txtMay, txtJune, txtJuly, txtAugust, txtSeptember, txtOctober, txtNovember, txtDecember)
dayofweekname = Array(txtSunday, txtMonday, txtTuesday, txtWednesday, txtThursday, txtFriday, txtSaturday)
yearname = Array("2007", "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016")
datename31 = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")
datename30 = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30")
datename29 = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29")
datename28 = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28")
markRequired = "<span style=""color:#FF000A;"">*&nbsp;</span>"
ajaxLoadingIcon = "<img src=""images/icons/icon_ajax_loading.gif"" alt=""loading..."" title=""loading..."" border=""0""/>"

incFormsFP = true
%>
