<!--#include file="config.asp" --><%
bDebugTest = false
bDeleteOrphans = false
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

'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'<><>
'<><>	This script will sync your database with the uploaded files
'<><>	on your server. Unzip to your site root and run it.
'<><>
'<><>	It will report LOST files in the database that are not in the
'<><>	filesystem. If a file is LOST, it will attempt to find it by
'<><>	recursing the module upload folder. If found, it will move the
'<><>	file to the correct directory.
'<><>
'<><>	It will compare the file category against the category in the
'<><>	URL path. If they are different, it will make the necessary
'<><>	adjustments.
'<><>
'<><>	The script will also locate ORPHANED files and will either
'<><>	DELETE them or move them to a special folder named 'orph'.
'<><>	You set this variable bDeleteOrphans to 'true' or 'false'.
'<><>
'<><>	For your first run, leave the variable bDebugTest set to 'true'
'<><>	and run the script. This will allow you to see what is going
'<><>	to happen without actually making any changes. Change the
'<><>	variable bDebugTest to 'false' and run the script to actually
'<><>	do the work and sync your database with your filesystem.
'<><>
'<><>	At the bottom of the page you will see a list of the files
'<><>	that are LOST. Just in case you have copies that you can 
'<><>	FTP to the correct directory.
'<><>
'<><>	Run this script anytime you feel like it to keep your
'<><>	database in order.
'<><>
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
%>
<!--#include file="inc_functions.asp" -->
<!--INCLUDE file="includes/inc_DBFunctions.asp" -->
<!--#include file="inc_top.asp" -->
<br /><br /><br /><br />
<p align="left">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
<%
iLost = 0
iFound = 0
iMoved = 0
iBadPath = 0
iBadPathFixed = 0
iOrphans = 0
sOrphans = ""
sLostFiles = ""
sDirPath = "files/downloads/"
sSql = "SELECT * FROM DL"
sUpdSql = "UPDATE DL SET URL='[%URL%]' WHERE DL_ID=[%ID%]"
  
  on error resume next
  set oFs = new clsSFSO
  sDirFullPath = server.MapPath(sDirPath)
  ssSql = ""
  
  ':: check db for invalid characters
  set rsA = my_Conn.execute(sSql)
  if rsA.eof then
	Response.Write sSql
  else
	do until rsA.eof
	  call checkDots(rsA)
	  rsA.movenext
	loop
  end if
  set rsA = nothing
  
  if ssSql <> "" then
    updateDots(ssSql)
  end if
  ssSql = ""
  
  ':: check files for invalid characters
  fileDots()

  ':: Sync db with files
	set rsA = my_Conn.execute(sSql)
	if rsA.eof then
	  Response.Write sSql
	else
	  do until rsA.eof
		call checkSync(rsA)
	    rsA.movenext
	  loop
	end if
	set rsA = nothing
	
  ':: find Orphaned files
  findOrphans()
	
  if iLost = 0 then
    'Response.Write "<br /><br /><br />None Lost!"
  else
  end if
    Response.Write "<br /><br /><br>Bad Path items: " & iBadPath
    Response.Write "<br />Lost items: " & iLost
    Response.Write "<br />Lost items Found: " & iFound
    Response.Write "<br />Items moved: " & iMoved
    Response.Write "<br />Orphaned Items: " & iOrphans
	if sLostFiles <> "" then
	Response.Write "<br><br><br /><h4>Lost Files</h4>" & sLostFiles
	end if
	if sOrphans <> "" then
	  Response.Write "<br><br><br /><h4>Orphaned Files</h4>" & sOrphans
	end if
	
  set oFs = nothing
  on error goto 0
%></td>
  </tr>
</table></p>

<!--#include file="inc_footer.asp" -->

<%
sub findOrphans()
  sOrphPath = "files/downloads/orph/"
  sOrphFullPath = server.MapPath(sOrphPath)
  oFs.CreateFolder(sOrphFullPath)  
  
  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  'fname = fs.GetFileName(server.MapPath(surl))
  Set fo = fs.GetFolder(sDirFullPath)
  for each x in fo.SubFolders
    Set fu = fs.GetFolder(x.Path)
	if lcase(fu.Name) <> "orph" then
     for each g in fu.files
	  tmpUrl =  sDirPath & fu.Name & "/" & g.Name
	  tmpFullUrl =  sDirFullPath & "\" & fu.Name & "\" & g.Name
	  tmpOrphPath = sOrphFullPath & "\" & g.Name
	  sSql = "SELECT * FROM DL WHERE URL='" & tmpUrl & "'"
	  set rsA = my_Conn.execute(sSql)
	  if rsA.eof then
	    ':: found orphan
		iOrphans = iOrphans + 1
		sOrphans = sOrphans & fu.Name & "/" & g.Name & "<br />"
    	Response.Write "<ul><li>Orphan found: " & fu.Name & "/" & g.Name
		if bDeleteOrphans then
    	  Response.Write "<ul><li>File Deleted! "
		  Response.Write "</li></ul>"
		  deleteFile(tmpFullUrl)
		else
		  call moveSyncfile(tmpFullUrl,tmpOrphPath)
		end if
		Response.Write "</li></ul>"
	  end if
	  set rsA = nothing
     next
	end if
    Set fu = nothing
  next  
  Set fo = nothing
  Set fs = nothing
  if bDeleteOrphans then
    deleteFolder(sOrphFullPath)
  end if
end sub

sub fileDots()
  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  'fname = fs.GetFileName(server.MapPath(surl))
  Set fo = fs.GetFolder(sDirFullPath)
  for each x in fo.SubFolders
    Set fu = fs.GetFolder(x.Path)
    for each g in fu.files
	  if instr(lcase(g.Name),"..") > 0 then
		g.Name = replace(g.Name,"..",".")
	  end if
	  if instr(lcase(g.Name),"~") > 0 then
		g.Name = replace(g.Name,"~","-")
	  end if
    next
    Set fu = nothing
  next
  Set fo = nothing
  Set fs = nothing
end sub

sub checkDots(o)
  if instr(o("URL"),"..") > 0 then
    ssSql = ssSql & "UPDATE DL SET URL='" & replace(o("URL"),"..",".") & "'"
	ssSql = ssSql & " WHERE DL_ID=" & o("DL_ID") & "|"
  end if
  if instr(o("URL"),"~") > 0 then
    ssSql = ssSql & "UPDATE DL SET URL='" & replace(o("URL"),"~","-") & "'"
	ssSql = ssSql & " WHERE DL_ID=" & o("DL_ID") & "|"
  end if
end sub

sub updateDots(s)
  s = left(s,len(s)-1)
  if instr(s,"|") > 0 then
    aSt = split(s,"|")
	for a = 0 to ubound(aSt)
	  if not bDebugTest then
	    executeThis(aSt(a))
	  end if
	next
  else
	if not bDebugTest then
      executeThis(s)
	end if
  end if
end sub

sub checkSync(o)
  scat = o("CATEGORY")
  surl = o("URL")
  sname = o("NAME")
  if oFs.FileExist(server.MapPath(trim(o("URL")))) then
    checkBadPath(o)
  else
	iLost = iLost + 1
    Response.Write "<ul><li>Lost: " & scat & " - " & sname & " - " & surl & ""
	searchForFile(o)
	Response.Write "</li></ul>"
  end if
end sub

sub checkBadPath(ox)
   if left(ox("URL"),4) <> "http" then
    if left(ox("URL"),len(sDirPath & ox("CATEGORY") & "/")) <> sDirPath & ox("CATEGORY") & "/" then
	  iBadPath = iBadPath + 1
      Response.Write "<ul><li>"
	  Response.Write "Bad Path: " & ox("CATEGORY") & " - " & ox("NAME") & " - " & ox("URL")
	  fixBadPath(ox)
	  Response.Write "</li>"
	  Response.Write "</ul>"
	end if
   end if
end sub

sub fixBadPath(ou)
  iBadPathFixed = iBadPathFixed + 1
  aTmp = split(ou("URL"),"/")
  nPath = trim(sDirPath & ou("CATEGORY") & "/" & aTmp(3))
  Response.Write "<ul><li>"
  Response.Write "FixBadPath()"
  Response.Write "<br />From Path: " & server.MapPath(trim(ou("URL")))
  Response.Write "<br />To Path: " & server.MapPath(nPath)
  Response.Write "</li></ul>"
  if oFs.FileExist(server.MapPath(trim(ou("URL")))) then
    call moveSyncfile(server.MapPath(trim(ou("URL"))),server.MapPath(nPath))
  end if
  call updateDbPath(ou("DL_ID"),nPath)
end sub

sub updateDbPath(i,np)
  sSql = replace(sUpdSql,"[%URL%]",np)
  sSql = replace(sSql,"[%ID%]",i)
  executeThis(sSql)
  Response.Write "<ul><li>"
  Response.Write "updateDbPath()"
  Response.Write "<br />Url: " & np
  Response.Write "<br />DL ID: " & i
  Response.Write "<br />" & sSql
  Response.Write "</li></ul>"
end sub

function searchForFile(oo)
  bt = false
  scat = oo("CATEGORY")
  surl = trim(oo("URL"))
  sname = oo("NAME")
  Set fs = Server.CreateObject("Scripting.FileSystemObject")
  fname = fs.GetFileName(server.MapPath(surl))
  Set fo = fs.GetFolder(sDirFullPath)
  for each x in fo.SubFolders
    Set fu = fs.GetFolder(x.Path)
    for each g in fu.files
	  if lcase(fname) = lcase(g.Name) then
	    iFound = iFound + 1
	    Response.Write "<ul><li>"
        Response.write "Found it: " & g.Path & ""
		Response.Write "</li></ul>"
		call doFoundLost(oo,g.Path)
		bt = true
	  end if
    next
    Set fu = nothing
  next
  for each fg in fo.files
	  if lcase(fname) = lcase(fg.Name) then
	    iFound = iFound + 1
	    Response.Write "<ul><li>"
        Response.write "Found it: " & fg.Path & ""
		Response.Write "</li></ul>"
		call doFoundLost(oo,fg.Path)
		bt = true
	  end if
  next
  Set fo = nothing
  Set fs = nothing
  if not bt then
	sLostFiles = sLostFiles & split(surl,"/")(2) & "/" & split(surl,"/")(3) & "<br />"
	checkBadPath(oo)
  end if
end function

sub doFoundLost(ob,px)
  call moveSyncfile(px,server.MapPath(ob("URL")))
  call checkBadPath(ob)
end sub

sub moveSyncfile(f,t)
  Response.Write "<ul><li>"
  Response.Write "moveSyncfile()"
  if lcase(f) <> lcase(t) then
    Response.Write "<br />From Path: " & f
    Response.Write "<br />To Path: " & t
	if not bDebugTest then
      oFs.MoveFile f,t
      iMoved = iMoved + 1
	  if oFs.FileExist(t) then
        Response.Write "<br />File moved!"
	  else
        Response.Write "<br />File could not be moved!"
	  end if
	else
      Response.Write "<br />File would be moved!"
	end if
  else
    Response.Write "<br />File not moved - same path!"
  end if
  Response.Write "</li></ul>"
end sub

%>
