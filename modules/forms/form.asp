<!--#include file="config.asp" -->
<%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2007 Dogg Software All Rights Reserved
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

CurPageType = "core"

':: modify one of the 2 values below
':: sPage_id is the id of the record in the database
':: and is the preferred way to call the recordset.
sPage_iName = "form"
'sPage_id = 0

':::::::::::::::::::::::::::::::::::::::::::::::::

pgname = "ERROR!"
CurPageInfoChk = "1"
%>
<!--#include file="inc_functions.asp" -->
<!-- #include file="modules/forms/form_functions.asp" -->
<%

'''''''''''''''''''''''''''''''''''''''''''''
'' Get some data about the form that helps display this skypage
'''''''''''''''''''''''''''''''''''''''''''''
formID = GetID("form")
if len(formID) < 1 then response.redirect "default.asp"
If Request.Form("save") = "Cancel" Then Response.Redirect "default.asp"
Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMHEADER WHERE ID = " & formID & ";")
If dbtable.EOF Then response.redirect "default.asp"
fldACTIVE = dbtable.fields("ACTIVE")
FLDFORMNAME = chkString(dbtable.fields("FLDFORMNAME"),"sqlstring")
FLDRECIPIENTEMAIL = dbtable.fields("FLDRECIPIENTEMAIL")
FLDEMAILSUBJECT = ChkString(dbtable.fields("FLDEMAILSUBJECT"),"display")
FLDINTROTEXT = ChkString(dbtable.fields("FLDINTROTEXT"),"out")
FLDTHANKYOU = ChkString(dbtable.fields("FLDTHANKYOU"),"out")
FLDINACTIVETEXT = ChkString(dbtable.fields("FLDINACTIVETEXT"),"out")
FLDSENDEMAIL = dbtable.fields("SENDEMAIL")
FLDSENDPM = dbtable.fields("SENDPM")
FLDSENDTO = dbtable.fields("SENDTO")


'get the default layout 
if sPage_id = 0 then
  cpSQL = "SELECT * FROM PORTAL_PAGES WHERE P_INAME = '" & sPage_iName & "'"
else
  cpSQL = "SELECT * FROM PORTAL_PAGES WHERE P_ID = " & sPage_id & ""
end if
set rsCPs = my_Conn.execute(cpSQL)
if not rsCPs.eof then
  pgtitle = rsCPs("P_TITLE")
  pgname = rsCPs("P_NAME")
  if rsCPs("P_ACONTENT") <> "" then
    pgbody = replace(rsCPs("P_ACONTENT"),"''","'")
  else
    if rsCPs("P_CONTENT") <> "" then
      pgbody = replace(rsCPs("P_CONTENT"),"''","'")
    end if
  end if
  left_Col = rsCPs("P_LEFTCOL")
  maint_Col = rsCPs("P_MAINTOP")
  mainb_Col = rsCPs("P_MAINBOTTOM")
  right_Col = rsCPs("P_RIGHTCOL")
	  
  m_title = FLDFORMNAME
  addToMeta "NAME","Description",rsCPs("P_META_DESC")
  addToMeta "NAME","Keywords",rsCPs("P_META_KEY")
  addToMeta "HTTP-EQUIV","Expires",rsCPs("P_META_EXPIRES")
  addToMeta "NAME","Rating",rsCPs("P_META_RATING")
  addToMeta "NAME","Distribution",rsCPs("P_META_DIST")
  addToMeta "NAME","Robots",rsCPs("P_META_ROBOTS")
	'm_description = rsCPs("P_META_DESC")
	'm_keywords = rsCPs("P_META_KEY")
	'm_expires = rsCPs("P_META_EXPIRES")
	'm_rating = rsCPs("P_META_RATING")
	'm_distribution = rsCPs("P_META_DIST")
	'm_robots = rsCPs("P_META_ROBOTS")
end if
set rsCPs = nothing

PageTitle = m_title

function CurPageInfo () 
	PageName = FLDFORMNAME 
	PageAction = txtViewing & "<br />" 
	PageLocation = "form.asp?form="&formID
	CurPageInfo = PageAction & "<a href=" & PageLocation & ">" & PageName & "</a>"
end function 
%>
<!--#include file="inc_top.asp" -->
<%
setAppPerms "forms","iName"

if not chkAppACTIVE("forms") and not hasAccess("1") then
    closeAndGo("default.asp")
end if
'end if

  cont = 0
  bLeft = false
  bMaint = false
  bMainb = false
  bRight = false
  
  if trim(left_Col) <> "" then
    if right(left_Col,1) = "," then
      left_Col = left(left_Col,len(left_Col)-1)
    end if
    if instr(left_Col,",") > 0 then
      l_col = split(left_Col,",")
	else
	  dim l_col(0)
      l_col(0) = left_Col
	end if
    bLeft = true
    cont = cont + 1
  end if
  if trim(maint_Col) <> "" then
    if right(maint_Col,1) = "," then
      maint_Col = left(maint_Col,len(maint_Col)-1)
    end if
    if instr(maint_Col,",") > 0 then
      mt_col = split(maint_Col,",")
	else
	  dim mt_col(0)
      mt_col(0) = maint_Col
	end if
    bMaint = true
    cont = cont + 1
  end if
  if trim(mainb_Col) <> "" then
    if right(mainb_Col,1) = "," then
      mainb_Col = left(mainb_Col,len(mainb_Col)-1)
    end if
    if instr(mainb_Col,",") > 0 then
      mb_col = split(mainb_Col,",")
	else
	  dim mb_col(0)
      mb_col(0) = mainb_Col
	end if
    bMainb = true
    cont = cont + 1
  end if
  if trim(right_Col) <> "" then
    if right(right_Col,1) = "," then
      right_Col = left(right_Col,len(right_Col)-1)
    end if
    if instr(right_Col,",") > 0 then
      r_col = split(right_Col,",")
	else
	  dim r_col(0)
      r_col(0) = right_Col
	end if
    bRight = true
    cont = cont + 1
  end if
  
  response.Write("<table class=""content"" border=""0"" width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr>")
  if bLeft then
    response.Write("<td class=""leftPgCol"" valign=""top"" nowrap=""nowrap"">")
	intSkin = getSkin(intSubSkin,1)
	 shoBlocks(l_col)
    response.Write("</td>")
  end if

    response.Write("<td class=""mainPgCol"" valign=""top"">")  
	intSkin = getSkin(intSubSkin,2)
	
	':: Breadcrumb values
  	arg1 = PageTitle & "|form.asp?form=" & formID
  	arg2 = ""
  	arg3 = ""
  	arg4 = ""
  	arg5 = ""
  	arg6 = ""

  
  	shoBreadCrumb arg1,arg2,arg3,arg4,arg5,arg6
  if bMaint then
	 shoBlocks(mt_col)
  end if

	'START SPECIAL FORM SKYPAGE

		
		spThemeTitle = FLDFORMNAME
	    spThemeBlock1_open(intSkin)
		%>
			<table width="100%" border="0"><tr><td>
			<%
      If fldACTIVE then
  			showform = True
  			FormMessage = ""
  			if request.form("save") = "Send" then
  				Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " ORDER BY FLDORDER, ID;")
  				If dbtable.EOF Then response.redirect "default.asp"

  				while not dbtable.EOF
  					FieldID = dbtable.fields("ID")
  					FLDCAPTION = dbtable.fields("FLDCAPTION")
  					FLDFIELDTYPE = dbtable.fields("FLDFIELDTYPE")
  					FLDVALIDATION = dbtable.fields("FLDVALIDATION")
  					FLDREQUIRED = dbtable.fields("FLDREQUIRED")
  					FieldInput = trim(chkString(request.form("Field"&FieldID), "sqlstring"))
  					if FLDREQUIRED = "Y" then  msg = requiredfield(FieldInput,FLDCAPTION)
  					if len(FieldInput) >= 1 then
  						select case FLDVALIDATION
  							case "Numeric"
  								if not isnumeric(FieldInput) then msg = msg & "<li>" & FLDCAPTION & " must be numeric.</li>"
  							case "E-mail"
  								if not IsValidEmail(FieldInput) then msg = msg & "<li>" & FLDCAPTION & " is not a valid e-mail address.</li>"
  							case "Date"
  								If Not Isdate(FieldInput) Then
  									msg = msg & "<li>" & FLDCAPTION & " is not a valid date.</li>"
  								Else
  									FieldInput = DanDate(FieldInput, "%b %d, %Y")
  								End If
  							Case "Phone Number"
  								If FormatPhoneNumber(FieldInput) = False Then
  									msg = msg & "<li>" & FLDCAPTION & " is not a valid phone number.</li>"
  								Else
  									FieldInput = FormatPhoneNumber(FieldInput)
  								End If
  							Case "Zip Code"
  								If FormatZipCode(FieldInput) = False Then
  									msg = msg & "<li>" & FLDCAPTION & " is not a valid zip code.</li>"
  								Else
  									FieldInput = FormatZipCode(FieldInput)
  								End If
  						end select
  					end if
  					FormMessage = FormMessage & "<p><b>" & FLDCAPTION & "</b><br>" & FieldInput & "</p>&nbsp;<br>"
  					dbtable.movenext
  				wend
  				
  				if intSecCode = 1 and  not DoSecImage(chkString(Request.Form("secCode"),"sqlstring")) then
					msg = msg & "<li>Your Security Code didn't match.</li>"
				end if

  				msg = trim(msg&" ")
  				if len(msg) < 1 then
					blParseSENDTO = false
					if len(FLDSENDTO) > 0 then
						arrSENDTO = split(FLDSENDTO, ",")
						blParseSENDTO = true
					end if
					if FLDSENDPM = 1 then
						if blParseSENDTO then
							for i=0 to uBound(arrSENDTO)
                                asdf = getMemberID(trim(arrSENDTO(i)))
                                if asdf > 0 then
                                    senderName = split(strWebMaster,",")
                                	strSql = "INSERT INTO " & strTablePrefix & "PM ("
                                	strSql = strSql & " M_SUBJECT"
                                	strSql = strSql & ", M_MESSAGE"
                                	strSql = strSql & ", M_TO"
                                	strSql = strSql & ", M_FROM"
                                	strSql = strSql & ", M_SENT"
                                	strSql = strSql & ", M_MAIL"
                                	strSql = strSql & ", M_READ"
                                	strSql = strSql & ", M_OUTBOX"
                                	strSql = strSql & ") VALUES ("
                                	strSql = strSql & " '" & trim(FLDEMAILSUBJECT) & " - " & FLDFORMNAME & "'"
                                	strSql = strSql & ", '" & formmessage & "'"
                                	strSql = strSql & ", " & asdf
                                	strSql = strSql & ", " & getMemberID(senderName(0))
                                	strSql = strSql & ", '" & strCurDateString & "'"
                                	strSql = strSql & ", " & "0"
                                	strSql = strSql & ", " & "0"
                                	strSql = strSql & ", '0')"
                                	executeThis(strSql)
                                end if
							next
						end if
					end if
					if FLDSENDEMAIL = 1 then
						strRecipients = ""
						if blParseSENDTO then
							for i=0 to uBound(arrSENDTO)
								'get list of emails to send to
								strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
								strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
								strSql = strSql & " WHERE M_NAME = '" & trim(arrSENDTO(i)) & "'"

								Set rsGetMemEmail = my_Conn.Execute(strSql)
								If Not ( rsGetMemEmail.BOF and rsGetMemEmail.EOF ) Then
									strRecipients = strRecipients & rsGetMemEmail("M_EMAIL") & "; "
								End If
								set rsGetMemEmail = nothing
							next
						end if
						
						strRecipients = strRecipients & FLDRECIPIENTEMAIL
						strSender = strSender
						strSubject = trim(FLDEMAILSUBJECT) & " - " & FLDFORMNAME
						strMessage = formmessage
						sendOutEmail strRecipients,strSubject,strMessage,0,1
						if len(Err_Msg) < 1 then
	  						'response.write "&nbsp;<p><div align=center>" & FLDTHANKYOU & "</div>&nbsp;<br>&nbsp;"
	  					else
	  						errMsg1 = errMsg1 & Err_Msg & "<br />"
	  					end if
					end if
					
					if len(Trim(errMsg1&" ")) < 1 then
					   response.write "&nbsp;<p><div align=""center"">" & FLDTHANKYOU & "</div>&nbsp;<br />&nbsp;"
					else
					   response.write errMsg1
					end if
					
					showform = False
  				else
					msg = "We encountered errors in the information you submitted.<br /><table border=""0"" align=""center""><tr><td><ul>" & msg & "</ul></td></tr></table>"
				end if
  			end if

  			if showform then
  				%>
  				<form method="post" action="form.asp?form=<%=FormID%>" name="aspFormCreator" id="aspFormCreator">
  				<table border="0" cellpadding="3" cellspacing="0" align="center" width="100%">
            <tr>
              <td colspan="2" align="center" class="fAlert"><% = msg %></td>
            </tr>
            <tr>
              <td colspan="2">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="2" align="center"><p><% = FLDINTROTEXT %></p></td>
            </tr>
            <tr>
              <td colspan="2">&nbsp;</td>
            </tr>
            <tr>
              <td colspan="2" align="center"><%=txtReg1a%>&nbsp;<%=markRequired%>&nbsp;<%=txtReg1b%></td>
            </tr>
            <tr>
              <td colspan="2">&nbsp;</td>
            </tr>
  						<%
  						Set dbtable = my_Conn.Execute("SELECT * FROM " & strTablePrefix & "FORMFIELDS WHERE FLDLINKFORMID = " & formID & " ORDER BY FLDORDER, ID;")
  						If dbtable.EOF Then response.redirect "default.asp"

  						while not dbtable.EOF
  							FieldID = dbtable.fields("ID")
  							FLDCAPTION = dbtable.fields("FLDCAPTION")
  							FLDFIELDTYPE = dbtable.fields("FLDFIELDTYPE")
  							FLDREQUIRED = dbtable.fields("FLDREQUIRED")
  							FLDWIDTH = dbtable.fields("FLDWIDTH")
  							FLDHEIGHT = dbtable.fields("FLDHEIGHT")
  							FLDDEFAULT = dbtable.fields("FLDDEFAULT")
  							FLDOPTIONS = dbtable.fields("FLDOPTIONS")
  							If Request.Form("save") = "Send" Then FLDDEFAULT = ChkString(Request.Form("Field"&FieldID),"in")

                Select case FLDFIELDTYPE
                  case "Info"
                    response.write "<tr><td colspan=""2"" valign=""top"" align=""center""><table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""50%""><tr><td>"
                  case "Check Box"
                    response.write "<tr><td align=""right"" valign=""top"" width=""45%"">"
                    if FLDREQUIRED = "Y" then Response.Write "&nbsp;" & markRequired
                    response.write FLDCAPTION
                    response.write "</td><td align=""left"" valign=""top"">"
                  case "Radio Button"
                    response.write "<tr><td align=""right"" valign=""top"" width=""45%"">"
                    if FLDREQUIRED = "Y" then Response.Write "&nbsp;" & markRequired
                    response.write FLDCAPTION
                    response.write "</td><td align=""left"" valign=""top"">"
                  case else
                    response.write "<tr><td align=""right"" valign=""top"" width=""45%"">"
                    Response.Write "<label for=""Field" & FieldID & """>"
    								if FLDREQUIRED = "Y" then Response.Write "&nbsp;" & markRequired
    								response.write FLDCAPTION
    								Response.Write "</label>" & vbcrlf
    								response.write "</td><td align=""left"" valign=""top"">"
                end select

  								select case FLDFIELDTYPE
  									case "Text Field"
  										Response.Write "<input type=""text"" name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""" & FLDWIDTH & """ value=""" & trim(FLDDEFAULT) & """>"
  									case "Text Area"
  										Response.Write "<textarea name=""Field" & FieldID & """ id=""Field" & FieldID & """ rows=""" & FLDHEIGHT & """ cols=""" & FLDWIDTH & """>" & trim(FLDDEFAULT) & "</textarea>"
  									case "Info"
  									  Response.Write "<span class=""behave"">" & FLDOPTIONS & "</span>"
  									case "Drop Down List"
  										Response.Write "<select size=""1"" name=""Field" & FieldID & """ id=""Field" & FieldID & """>"
  										NewOptions = split(FLDOPTIONS, "[|]")
  										for i = 0 to ubound(NewOptions)
  											Response.Write "<option value=""" & NewOptions(i) & """"
  												if FLDDEFAULT = NewOptions(i) then response.write " SELECTED"
  												Response.Write ">&nbsp;" & NewOptions(i) & "</option>"
  										next
  										Response.Write "</select>"
  									case "Check Box"
                      Dim chkCurColumn
                      chkCurColumn = 1
                      Response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0""><tr>"
                        CheckOptions = split(FLDOPTIONS, "[|]")
                        for i = 0 to ubound(CheckOptions)
                          Response.Write "<td><input type=""checkbox"" class=""boxes"" name=""Field" & FieldID & """ id=""Field" & FieldID & i & """ value=""" & CheckOptions(i) & """"
    											if instr( "," & FLDDEFAULT & ",",CheckOptions(i)) then response.write " checked"
    											Response.Write ">&nbsp;<label for=""Field" & FieldID & i & """>" & CheckOptions(i) & "</label></td>"

    											if chkCurColumn/FLDWIDTH - Round(chkCurColumn/FLDWIDTH) = 0 then
                            response.write "</tr><tr>"
                          end if
                          chkCurColumn = chkCurColumn + 1
                        next
                      Response.write "</tr></table>"
  									case "Radio Button"
                      Dim radioCurColumn
                      radioCurColumn = 1
                      Response.write "<table border=""0"" cellpadding=""2"" cellspacing=""0""><tr>"
                        RadioOptions = split(FLDOPTIONS, "[|]")
                        for i = 0 to ubound(RadioOptions)
                          Response.Write "<td><input type=""radio"" class=""boxes"" name=""Field" & FieldID & """ id=""Field" & FieldID & i & """ value=""" & RadioOptions(i) & """"
    											if instr( "," & FLDDEFAULT & ",",RadioOptions(i)) then response.write " checked"
    											Response.Write ">&nbsp;<label for=""Field" & FieldID & i & """>" & RadioOptions(i) & "</label></td>"

    											if radioCurColumn/FLDWIDTH - Round(radioCurColumn/FLDWIDTH) = 0 then
                            response.write "</tr><tr>"
                          end if
                          radioCurColumn = radioCurColumn + 1
                        next
                      Response.write "</tr></table>"
  									case "State"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										for i = 0 to ubound(statename)
  											response.write "<option value=""" & statename(i) & """"
  												if FLDDEFAULT = statename(i) then response.write " SELECTED"
  											response.write ">&nbsp;" & statename(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Country"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(countryname)
  											Response.Write "<option value=""" & countryname(i) & """"
  												If FLDDEFAULT = countryname(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & countryname(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Date Picker"
                      Response.Write "<script type=""text/javascript"">addCalendar('Cal" & FieldID & "', 'Select Date', 'Field" & FieldID & "', 'aspFormCreator')</script>"
                      Response.Write "<input type=""text"" name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""12"" value="""" readonly>&nbsp;<a href=""javascript:showCal('Cal" & FieldID & "')""><img border=""0"" src=""images/icons/SmallCalendar.gif"" width=""16"" height=""16""></a>"
  									Case "Month"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(fm_monthname)
  											Response.Write "<option value=""" & fm_monthname(i) & """"
  												If FLDDEFAULT = fm_monthname(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & fm_monthname(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Day of Week"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(dayofweekname)
  											Response.Write "<option value=""" & dayofweekname(i) & """"
  												If FLDDEFAULT = dayofweekname(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & dayofweekname(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Year"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(yearname)
  											Response.Write "<option value=""" & yearname(i) & """"
  												If FLDDEFAULT = yearname(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & yearname(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Date 31"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(datename31)
  											Response.Write "<option value=""" & datename31(i) & """"
  												If FLDDEFAULT = datename31(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & datename31(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Date 30"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(datename30)
  											Response.Write "<option value=""" & datename30(i) & """"
  												If FLDDEFAULT = datename30(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & datename30(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Date 29"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(datename29)
  											Response.Write "<option value=""" & datename29(i) & """"
  												If FLDDEFAULT = datename29(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & datename29(i) & " </option>"
  										next
  										response.write "</select>"
  									Case "Date 28"
  										response.write "<select name=""Field" & FieldID & """ id=""Field" & FieldID & """ size=""1"">"
  										For i = 0 To UBound(datename28)
  											Response.Write "<option value=""" & datename28(i) & """"
  												If FLDDEFAULT = datename28(i) Then Response.Write " SELECTED"
  											Response.Write ">&nbsp;" & datename28(i) & " </option>"
  										next
  										response.write "</select>"
  								end select

                If FLDFIELDTYPE = "Info" then
                  'response.write "<label>&nbsp;</label>"
                  response.write "</td></tr></table></td></tr>"
                else
                  response.write "</td></tr>"
                end if

  							dbtable.MoveNext
  						wend
  						
  						if intSecCode <> 0 then
              %>
              <tr>
                <td colspan="2" align="center"><% shoSecurityImg %></td>
              </tr>
              <%
              End If
  						%>
  						<tr><td colspan="2" align="center" valign="bottom">
  						<p><input type="hidden" name="save" value="Send" /><button type="submit" name="submit" class="button">Send</button>&nbsp;<button type="reset" name="reset" class="button">Reset</button>&nbsp;<button type="submit" name="cancel" class="button" onClick="document.forms['aspFormCreator'].action='default.asp'">Cancel</button></p>
  						</td></tr>
          </table>
  				</form>
  			<%
			  end if
      else 'fldACTIVE
        response.write FLDINACTIVETEXT
      end if
			%>
			</td></tr></table>
    <%
	'END SPECIAL FORM SKYKPAGE
    spThemeBlock1_close(intSkin)

  if bMainb then
	 shoBlocks(mb_col)
  end if
    response.Write("</td>")
  
  if bRight then
    if cont = 3 then
      response.Write("<td class=""rightPgCol"" valign=""top"" width=""195"">")
	else
      response.Write("<td class=""rightPgCol"" valign=""top"">")
	end if
	intSkin = getSkin(intSubSkin,3)
	shoBlocks(r_col)
    response.Write("</td>")
  end if
  response.Write("</tr></table>")

 %>
<!--#include file="inc_footer.asp" -->
