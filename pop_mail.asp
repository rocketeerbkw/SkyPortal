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

%>
<!--#include file="inc_functions.asp" -->
<!--#include file="inc_top_short.asp" -->
<!--#include file="includes/inc_emails.asp" -->
<% 
	if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
		intMemberID = cLng(Request.QueryString("ID"))
	else
		intMemberID = 0
	end if

	'
	strSql = "SELECT M.M_EMAIL, M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
	strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

	set rs = my_Conn.Execute (strSql)
	if Request.QueryString("mode") <> "DoIt" then
%>
      <p><span class="fTitle"><%= txtSndEmlMsg %></span></p>
<%	end if
	if lcase(strEmail) = "1" then
		if Request.QueryString("mode") = "DoIt" then
			Err_Msg = ""
			if Request.Form("YName") = "" then 
				Err_Msg = Err_Msg & "<li>" & txtErNoUNam & "</li>"
			end if
			if Request.Form("YEmail") = "" then 
				Err_Msg = Err_Msg & "<li>" & txtErNoEmlAdd & "</li>"
			else
				if EmailField(Request.Form("YEmail")) = 0 then 
					Err_Msg = Err_Msg & "<li>" & txtErValEml & "</li>"
				end if
			end if
			if Request.Form("Name") = "" then 
				Err_Msg = Err_Msg & "<li>" & txtErRecNam & "</li>"
			end if
			if Request.Form("Msg") = "" then 
				Err_Msg = Err_Msg & "<li>" & txtErNoMsg & "</li>"
			end if
			if (Err_Msg = "") then
				strRecipientsName = chkString(Request.Form("Name"),"sqlstring")
				strRecipients = rs("M_EMAIL")
				strFrom = chkString(Request.Form("YEmail"),"sqlstring")
				strFromName = chkString(Request.Form("YName"),"sqlstring")
				strSubject = ""
				strMessage = ""
				getEmailFromMemberTxt()
				sendOutEmail strRecipients,strSubject,strMessage,2,0
%>

      <p><span class="fTitle"><%= txtEmlSent %></span></p>
<%
			else
%>
      <p><span class="fTitle"><%= txtEmlProb %></span></p>
      <table>
        <tr>
          <td><ul><% =Err_Msg %></ul></td>
        </tr>
      </table>
      <p><a href="JavaScript:history.go(-1)"><%= txtGoBackData %></a></p>
<%
			end if
		else 
			Err_Msg = ""
			if rs("M_EMAIL") <> " " then
				strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
				strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & strDBNTUserName & "'"

				set rs2 = my_conn.Execute (strSql)
				YName = ""
				YEmail = ""

				if (rs2.EOF or rs2.BOF)  then
					if strLogonForMail <> "0" then 
						Err_Msg = Err_Msg & "<li>" & txtLgnEml & "</li>"
%>
      <table>
        <tr>
          <td><ul><% =Err_Msg %></ul></td>
        </tr>
      </table>
	  
      <p><a href="JavaScript:onclick= window.close()"><%= txtCloseWin %></a></p>
<%
						Response.End
					end if
				else
					YName = Trim("" & rs2("M_NAME"))
					YEmail = Trim("" & rs2("M_EMAIL"))
				end if
				rs2.close
				set rs2 = nothing
%>
      <form action="pop_mail.asp?mode=DoIt&id=<% =intMemberID %>" method="Post" id="Form1" name="Form1">
      <input type="hidden" name="Page" value="<% =chkString(Request.QueryString("page"),"sqlstring") %>">
<%
spThemeBlock1_open(intSkin)
%>
<table class="tPlain" width="100%">
              <tr>
                <td class="tCellAlt0" align="right" nowrap><b><%= txtSndToNam %>:</b></td>
                <td class="tCellAlt0"><% =rs("M_NAME") %><input type="hidden" name="Name" value="<% =rs("M_NAME") %>"></td>
              </tr>
              <tr>
                <td class="tCellAlt0" align="right" nowrap><b><%= txtUNam %>:</b></td>
                <td class="tCellAlt0"><input name="YName" type="<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %>" value="<% = YName %>" size="25"> <% if YName <> "" then Response.Write(YName) end if %></td>
              </tr>
              <tr>
                <td class="tCellAlt0" align="right" nowrap><b><%= txtUEml %>:</b></td>
                <td class="tCellAlt0"><input name="YEmail" type="<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %>" value="<% = YEmail %>" size="25"> <% if YEmail <> "" then Response.Write(YEmail) end if %></td>
              </tr> 
              <tr>
                <td class="tCellAlt0" colspan="2"><b><%= txtMsg %>:</b></td>
              </tr>
              <tr>
                <td class="tCellAlt0" colspan="2" align="center"><textarea name="Msg" cols="40" rows="5"></textarea></td>
              </tr>                    
              <tr>
                <td class="tCellAlt0" colspan="2" align="center"><input type="Submit" value="<%= txtSend %>" id="Submit1" name="Submit1" class="button"></td>
              </tr></table>
<%
spThemeBlock1_close(intSkin)%>
      </form>
<%
			else
%>
      <p><span class="fSubTitle"><%= txtNoEmlAvail %>.</span></p>
<%
			end if
		end if
	else
%>
      <p><%= txtClkSend %>&nbsp;<a href="mailto:<% =rs("M_EMAIL")%>"><% =rs("M_NAME")%></a>&nbsp;<%= txtAnEml %></p>
<%
	end if
set rs = nothing
%><!--#include file="inc_footer_short.asp" -->