<%
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'<> Copyright (C) 2005-2006 Dogg Software All Rights Reserved
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

'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|              Coded by Brandon Williams.             |'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'
'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'|'

'User Configurable Variables
intDaysToKeep = 8 'How many days to save drafts in the database - 0 means keep forever

'Some Multi-Language Stuff
if intDaysToKeep > 0 then draftsKept = "Drafs are kept for " & intDaysToKeep & " days." end if


' DO NOT EDIT BELOW THIS LINE UNLESS YOU KNOW WHAT YOU'RE DOING!
function delDraftsDaemon()
	if NOT intDaysToKeep = 0 then
		set rsA = nothing
		deleteDate = datetostr2(dateadd("d", - intDaysToKeep, now()))
		strSql = "DELETE FROM " & strTablePrefix & "DRAFTS WHERE DRAFT_ENTRYDATE < '" & deleteDate & "' OR DRAFT_LASTDATE < '" & deleteDate & "'"
		executeThis(strSql)
	end if
end function

function delDraft()
response.clear
	if isNumeric(Request("dID")) then
		'I decided to use some AJAX because it's cool.  We are checking to see if the draft ID was authored by the currently logged in user to make sure
		'other users can't delete drafts that are not their own just by visiting a URL
		strSql = "SELECT DRAFT_ID, DRAFT_ENTRYUSER FROM " & strTablePrefix & "DRAFTS WHERE DRAFT_ID = " & Request("dID")
		'response.write(strsql)
		set rs = my_Conn.Execute(strsql)
		if rs.EOF then 'Couldn't find any drafts, maybe it was already deleted?
			response.write("draftmodxxxdoesnotexistdraftmodxxx")
		else 'draft exists in DB
			if rs("DRAFT_ENTRYUSER") = strUserMemberID then 'the user is the author, OK to delete
				strSql = "DELETE FROM " & strTablePrefix & "DRAFTS WHERE DRAFT_ID = " & Request("dID") & " and DRAFT_ENTRYUSER = " & strUserMemberID
				'response.write(strsql)
				ExecuteThis(strsql)
				response.write("draftmodxxxdeleteddraftmodxxx")
			else
				response.write("draftmodxxxnotyoursdraftmodxxx")
			end if
		end if
	else
		raiseHackAttempt("Member, " & getMemberName(strUserMemberID) & ", may have tried to delete another users draft.  Please analyze hack attempt to make sure.")
	end if
	response.end
	set rs = nothing
end function

function saveDraft()
	if NOT Request("newDraft") <> "" or NOT Request("newDraft") <> " " then
		saveDraft = "Text Box Empty"
	else
		draftText = ChkString(Request("newDraft"),"message")
	
		strSql = "INSERT INTO " & strTablePrefix & "DRAFTS ("
		strSql = strSql & "DRAFT_TEXT, DRAFT_ENTRYUSER, DRAFT_ENTRYDATE, DRAFT_LASTUSER, DRAFT_LASTDATE) VALUES ("
		strSql = strSql & "'" & draftText & "', " & strUserMemberID & ", '" & datetostr2(now()) & "', " & strUserMemberID & ", '" & datetostr2(now()) & "')"
		'response.write(strSql)
		set rs = my_Conn.execute(strSql)
		
		saveDraft = "Draft Saved"
	end if
	set rs = nothing
end function

function saveDraftAJAX()
response.clear
	if NOT Request("newDraft") <> "" or NOT Request("newDraft") <> " " then
		response.write("draftmodxxxtextboxemptydraftmodxxx")
	else
		draftText = ChkString(Request("newDraft"),"message")
	
		strSql = "INSERT INTO " & strTablePrefix & "DRAFTS ("
		strSql = strSql & "DRAFT_TEXT, DRAFT_ENTRYUSER, DRAFT_ENTRYDATE, DRAFT_LASTUSER, DRAFT_LASTDATE) VALUES ("
		strSql = strSql & "'" & draftText & "', " & strUserMemberID & ", '" & datetostr2(now()) & "', " & strUserMemberID & ", '" & datetostr2(now()) & "')"
		'response.write(strSql)
		set rs = my_Conn.execute(strSql)
		
		response.write("draftmodxxxdraftsaveddraftmodxxx")
	end if
	set rs = nothing
end function

function showDrafts(memID)
%>
<script type="text/javascript">
//Delete drafts AJAX
function deleteDraft(dID) {

	new Ajax.Request('drafts.asp?cmd=2&dID=' + dID, {
		method: 'post',
		onSuccess: function(transport) {
			if (transport.responseText.match(/draftmodxxxdeleteddraftmodxxx/)) { //Draft Deleted successfully
				//Hide the deleted draft
				new Effect.Fade($(dID));
				//Give user msg notifiying of deletion
				$('message').update('Draft Deleted');
				//Show or highlight the message
				if ($('message').visible()) {
					new Effect.Highlight($('message'));
				}
				else {
					$('message').show();
				}
			}
			else if(transport.responseText.match(/draftmodxxxdoesnotexistdraftmodxxx/)) { //Draft not in DB
				$('message').update('Draft already deleted!')
				//Show or highlight the message
				if ($('message').visible()) {
					new Effect.Highlight($('message'));
				}
				else {
					$('message').show();
				}
			}
			else if(transport.responseText.match(/draftmodxxxnotyoursdraftmodxxx/)) { //Logged in user not author of draft being deleted
				//Give user msg notifiying of error
				$('message').update('There was an Error:&nbsp;Could not deleted draft, you are not the author');
				//Show or highlight the message
				if ($('message').visible()) {
					new Effect.Highlight($('message'));
				}
				else {
					$('message').show();
				}
			}
		}
	});
}
</script>
<%
	strSql = "SELECT * FROM " & strTablePrefix & "DRAFTS WHERE DRAFT_ENTRYUSER = " & memID & " OR DRAFT_LASTUSER = " & memID
	'response.write(strSql)
	set rs = my_Conn.execute(strSql)
	
	spThemeBlock1_open(intSkin) %>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" align=center>
		<tr align=center><td><p aling="center"><%=draftsKept%></p><div id="message" class="tCellAlt0" style="display:none;"></div><p align="center">
<%
	if NOT rs.EOF then
		response.write("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
		rs.MoveFirst
		Do While NOT rs.EOF %>
			<div id="<% =rs("DRAFT_ID") %>">
			<div class="tCellAlt1">
				<a href="javascript:;" onClick="deleteDraft('<% =rs("DRAFT_ID") %>');">
					<img src="images/icons/icon_delete_reply.gif" alt="delete draft" style="border: 0px;" />
				</a>
				<b>Saved On:</b>&nbsp;<% =chkDate(rs("DRAFT_ENTRYDATE")) %>
			</div>
			<div>
				<% =rs("DRAFT_TEXT") %>
			</div>
		</div>
		<%	rs.MoveNext
		Loop
		response.write("</table>")
	else	
		response.write("<span class=""fSubTitle"">No drafts found.</span><br /><br />")
	end if
%>
	<div id="draftBox"><textarea name="newDraft" id="newDraft" cols="85" rows="15"></textarea></p><p><button type="submit" class="btnLogin" value="Save Draft" >Save Draft</button></p></div>
	<p>&nbsp;</p>
		</td></tr></table>
<%  	spThemeBlock1_close(intSkin)
set rs = nothing
end function


function popDrafts(memID)
	strSql = "SELECT * FROM " & strTablePrefix & "DRAFTS WHERE DRAFT_ENTRYUSER = " & memID & " OR DRAFT_LASTUSER = " & memID
	'response.write(strSql)
	set rs = my_Conn.execute(strSql)
	%> <div style="height: 500px; width:100%; overflow:scroll"> <%
	if NOT rs.EOF then
		response.write("<table cellpadding=""1"" cellspacing=""1"" border=""0"">")
		if intDaysToKeep > 0 then
			response.write "<tr><td align=""center""><h3>Drafs are kept for " & intDaysToKeep & " days.</h3></td></tr>"
		end if
		rs.MoveFirst
		Do While NOT rs.EOF
			response.write("<tr>")
			response.write("<td colspan=""2""><div id=""" & rs("DRAFT_ID") & """>" & rs("DRAFT_TEXT") & "</div></td>")
			response.write("</tr>")
			response.write("<tr><td align=""center""><input type=""button"" id=""insert"" name=""insert"" value=""{$lang_insert}"" onClick=""javascript:tinyMCE.execCommand('mceSetContent', false, document.getElementById('" & rs("DRAFT_ID") & "').innerHTML);tinyMCEPopup.close();"" /><br /><hr /></td></tr>")
			rs.MoveNext
		Loop
		%>
		</table>
<%
	else	
%>
		<span class="fSubTitle">No drafts found.</span>
<%
	end if
	%> </div> <%
set rs = nothing
end function

sub config_drafts() 
'nothing here for now :)
end sub

incDraftsFp = true
%>