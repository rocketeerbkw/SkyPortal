<%
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

'/**
' * SkyPortal Roster Module
' *
' * This file handles uploads for the roster module
' *
' * LICENSE: You may copy, modify and redistribute this work,
' *          provided that you do not remove this copyright notice
' *
' * @copyright  2008 Brandon Williams. Some Rights Reserved.
' * @license    http://creativecommons.org/licenses/BSD/   BSD License
' */

Server.ScriptTimeout = 3600
uploadPg = true
cid = 0
sid = 0
%>
<!--#include file="config.asp" -->
<!--#INCLUDE file="includes/inc_clsUpload.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" --> 
<!--#include file="modules/roster/roster_functions.asp"-->
<!-- #include file="includes/core_module_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
response.clear

setAppPerms "roster","iName"

if not bAppWrite then
    showMsg "warn","You don't have permission to do that."
else
    if not isObject(objUpload) then
        strOops = "Your session has expired.<br />You will need to refresh the submission page<br />to get the session back."
        showMsg "err",strOops
        response.end
    else        
        if objUpload.Fields("cid").Value <> "" or  objUpload.Fields("cid").Value <> " " then
        	if IsNumeric(objUpload.Fields("cid").Value) = True then
        		c_id = cLng(objUpload.Fields("cid").Value)
        	else
        		closeAndGo("stop")
        	end if
        end if
        if objUpload.Fields("sid").Value <> "" or  objUpload.Fields("sid").Value <> " " then
        	if IsNumeric(objUpload.Fields("sid").Value) = True then
        		s_id = cLng(objUpload.Fields("sid").Value)
        	else
        		closeAndGo("stop")
        	end if
        end if
        rstrFolder = chkString(objUpload.Fields("folder").Value,"sqlstring")
        'rstrFolder = request.querystring("folder")
		
		'**sigh**
        'Hardcoding this because I'm too lazy to sanitize a "return" querystring parameter
        select case rstrFolder
            case "team"
                rstrReturnUrl = "pop_roster.asp?mode=team&cmd=edit&upPhoto=upped&cid=" & c_id & "&sid=" & s_id & "&photourl="
            case "player"
                rstrReturnUrl = "admin_roster.asp?v=pl&c=2&i=" & s_id & "&upPhoto=upped&photourl="
            case "sponsor"
                rstrReturnUrl = "admin_roster.asp?v=s&c=2&i=" & s_id & "&upPhoto=upped&photourl="
            case "volunteer"
                rstrReturnUrl = "admin_roster.asp?v=v&c=2&i=" & s_id & "&upPhoto=upped&photourl="
            case else
                closeAndGo("stop")
        end select

        if sString <> "" then
            Session.Contents("rosterErr") = sString
            Response.Redirect rstrReturnUrl & "&err=true"
        end if
		
		'Rename file to sid
		fileExt = Right(filename,Len(filename)-InstrRev(filename,"."))
		newFileName = s_id & "." & lcase(fileExt)
        		
		basePath = remotePath
		tmpFilePath = basePath & filename
		folderPath = basePath & rstrFolder
		if c_id > 0 then
			extFolderPath = folderPath & "/" & c_id
			fullFilePath = extFolderPath & "/" & newFileName
		else
			extFolderPath = ""
			fullFilePath = folderPath & "/" & newFileName
		end if
		
        set fso = Server.CreateObject("Scripting.FileSystemObject")
			'Skip any FSO errors
			on error resume next
			'Check to see if our folders exist
			if fso.FolderExists(server.MapPath(folderPath)) = false then
				fso.CreateFolder(server.MapPath(folderPath))
			end if
			if extFolderPath <> "" then
				if fso.FolderExists(server.MapPath(extFolderPath)) = false then
					fso.CreateFolder(server.MapPath(extFolderPath))
				end if
			end if
			'Check to see if our file exists
			if fso.FileExists(server.MapPath(fullFilePath)) = true then
				deleteFile(server.MapPath(fullFilePath))
			end if
            'If the upload was successful, it is waitin in the tmpFilePath
            'Let's move to final destination
    		if fso.FileExists(server.MapPath(tmpFilePath)) = true then
    			fso.MoveFile server.MapPath(tmpFilePath), server.MapPath(fullFilePath)
    		end if
			'Resume errors
			on error goto 0
        set fso = nothing
        
        'we've got this far, we're done
        'send back to form page
        Response.Redirect rstrReturnUrl & server.urlencode(fullFilePath)
    end if
end if

%>
<!--#INCLUDE FILE="inc_footer.asp" -->