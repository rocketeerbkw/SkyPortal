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
'::::: START DB_FUNCTIONS :::::::::::::::::::::::::::::::::::::::

txtDBSQLExecSucc = "SQL executed successfully"
txtDBSQLNotSucc = "SQL was not executed successfully"
txtDBConstNotCreated = "Constraint not created"
txtDBConstExists = "Constraint already exists"
txtDBConstCreated = "Constraint created succesfully"
txtDBTableAltSucc = "Table altered succesfully"
txtDBTblNoCreated = "Table not created"
txtDBTblExists = "Table already exists"
txtDBTblCreated = "Table created succesfully"
txtDBTblDropped = "Table dropped succesfully"
txtDBTblNoDropped = "Table not dropped"
txtDBTable = "Table"
txtDBIndxDrop = "Index dropped succesfully"
txtDBIndxNotDrop = "Index was not dropped"
txtDBIndxCreated = "Index created succesfully"
txtDBIndxNotCreated = "Index was not created"
txtDBTblPopulated = "Table was populated successfully with default data"
txtDBTblNotPopulated = "Error, table was not populated"
txtDBTblFldNoExist = "Table or Field does not exist"

'::::: END DB_FUNCTIONS :::::::::::::::::::::::::::::::::::::::::

Dim tableExists
Dim tableNotExist
Dim relationexists
Dim fieldExists
Dim ErrorCount
Dim primaryExists
tableExists   = -2147217900
tableNotExist = -2147217865 
relationexists = -2147217900
fieldExists   = -2147217887
ErrorCount = 0
primaryExists = -2147467259


function checkIt(str)
		select case strDBType
			case "sqlserver"
				select case sqlVer
					case "7"
'					  str = replace(str,"DATE","DATETIME")
				end select
				
				select case strUnicode
					Case "YES"
					  str = replace(str,"TEXT(","NVARCHAR(")
					  str = replace(str,"MEMO","NTEXT")
					Case else
					  str = replace(str,"TEXT(","VARCHAR(")
					  str = replace(str,"MEMO","TEXT")
				end select
				str = replace(str,"LONG","INT")
				str = replace(str,"BYTE","SMALLINT")
			case "mysql"
				str = replace(str,"BYTE","TINYINT")
				str = replace(str,"LONG","INT")
				str = replace(str,"TEXT","VARCHAR")
				str = replace(str,"MEMO","TEXT")
				str = replace(str,"[","")
				str = replace(str,"]","")
				str = replace(str,"IDENTITY (1, 1) PRIMARY KEY NOT NULL","UNSIGNED AUTO_INCREMENT PRIMARY KEY")
		end select
		checkIt = str
end function

function doSQL(ix)
	sMg = ""
	on error resume next	
	Err.Clear
	my_Conn.Execute (ix)
	dbHits = dbHits + 1
	if err.number = 0 THEN
	  if strDebug then
		Response.Write("<b>" & ix & "<br />" & txtDBSQLExecSucc & "</b><br /><br />" & vbNewLine)
	  end if
	elseif err.number <> -2147217887 then
		ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBSQLNotSucc & "</b><br />" & vbNewLine)
		response.Write("SQL: " & ix & "<br />" & err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	end if
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function doSQL2(ix,typ)
	on error resume next	
	Err.Clear
	my_Conn.Execute(ix)
	dbHits = dbHits + 1
	fldExists = -2147217887
  if typ = 0 then
	if err.number = 0 THEN
	  'if strDebug then
		Response.Write("<b>" & ix & "<br />" & txtDBSQLExecSucc & "</b><br /><br />" & vbNewLine)
	  'end if
	elseif err.number <> fldExists then
		ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBSQLNotSucc & "</b><br />" & vbNewLine)
		response.Write("SQL: " & ix & "<br />" & err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	end if
  end if
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function CreateConstraints(ConstraintSQL,iRelationship)
	on error resume next
	Err.Clear

	my_Conn.execute ConstraintSQL
	    dbHits = dbHits + 1
		if err.number <> 0 and err.number <> relationshipExists then
			ErrorCount = ErrorCount + 1
'			response.Write "    " & ConstraintSQL & "<br />" & vbNewLine
			Response.Write("    <font color=""#FF0000""><b>" & txtDBConstNotCreated & ": " & iRelationship & "</b><br />" & ContrraintSQL & "</font><br /><br />" & vbNewLine)
			response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		else
			if err.number = -2147217900 then 
	  		  if strDebug then
				Response.Write("    <font color=""#FF0000""><b>" & txtDBConstExists & ": " & iRelationship & "</b></font><br /><br />" & vbNewLine)
	  		  end if
			else
	  		  if strDebug then
				Response.Write("    <b>" & txtDBConstCreated & ": " & iRelationship & "</b><br /><br />" & vbNewLine)
	  		  end if
			end if
		end if	
		Err.Clear
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")	
end function

function alterTable(tbl)
	on error resume next
	Err.Clear
'	response.write "    " & tbl & "<br />" & vbNewLine
	my_Conn.Execute(tbl),,adCmdText + adExecuteNoRecords
	dbHits = dbHits + 1
		
	if err.number <> 0 then
		Response.Write("<font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
	  'if strDebug then
		Response.Write("    <b>" & txtDBTableAltSucc & "</b><br /><br />" & vbNewLine)
	  'end if
	end if
	Err.Clear
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")	
end function

function alterTable2(tbl)
	on error resume next	
	Err.Clear
'	response.write "    " & tbl & "<br />" & vbNewLine
	altString = split(tbl,",")
	ssSQL = "ALTER TABLE " & altString(0) & " ADD "
	for str = 1 to ubound(altString)
		altSQL = ssSQL & altString(str)
		my_Conn.Execute(altSQL),,adCmdText + adExecuteNoRecords
		dbHits = dbHits + 1
		if err.number <> 0 and err.number <> fieldExists then
			Response.Write("<font color=""#FF0000"">" & altString(str) & ":<br />" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
			ErrorCount = ErrorCount + 1
		else
	  	  'if strDebug then
			Response.Write("<b>" & txtDBTableAltSucc & ": """ & altString(str) & """</b><br /><br />" & vbNewLine)
	  	  'end if
		end if
		Err.Clear
	next
	Err.Clear
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")	
end function

function createTable(tbl)
	on error resume next	
	Err.Clear
  
'	response.write "    " & tbl & "<br />" & vbNewLine
	my_Conn.Execute(tbl),,adCmdText + adExecuteNoRecords
	dbHits = dbHits + 1
	
	if err.number <> 0 and err.number <> 13 and err.number <> tableExists then
'		response.Write "    " & tbl & "<br />" & vbNewLine
		Response.Write("    <font color=""#FF0000""><b>" & txtDBTblNoCreated & "</b></font><br /><br />" & vbNewLine)
		response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableExists then 
			Response.Write("    <font color=""#FF0000""><b>" & txtDBTblExists & "</b></font><br /><br />" & vbNewLine)
		else
	  	  'if strDebug then
			Response.Write("    <b>" & txtDBTblCreated & "</b><br /><br />" & vbNewLine)
	  	  'end if
		end if
	end if	
	Err.Clear
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")	
end function

function droptable(table)
	on error resume next	
	Err.Clear
  'if strDebug then
	response.write("    <hr size=""1"" width=""400"" align=""center"" color=""green"">" & vbNewLine)
	response.write("    <hr size=""1"" width=""400"" align=""center"" color=""green"">" & vbNewLine)
	Response.Write("    <b>" & txtDBTable & ": <font color=""#0000FF"">" & table & "</font></b><br /><br />" & vbNewLine)
  'end if
		stSQL = "DROP TABLE " & table
	my_Conn.Execute (stSQL),,adCmdText + adExecuteNoRecords
	dbHits = dbHits + 1
	if err.number = 0 then ' and err.number <> 13 and err.number <> tableExists
	  'if strDebug then
			Response.Write("    <b>" & txtDBTblDropped & "</b><br /><br />" & vbNewLine)
	  'end if
	else
		'ErrorCount = ErrorCount + 1
		if err.number = tableNotExist then 
			'Response.Write("    <font color=""#FF0000""><b>Table does not exist</b></font><br /><br />" & vbNewLine)	
			Err.Clear
		else
			Response.Write("    <font color=""#FF0000""><b>" & txtDBTblNoDropped & "</b></font><br />" & vbNewLine)
		response.Write("    <font color=""#FF0000""> '" & table & "': " & err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)	
		Err.Clear
		end if
	end if		
	Err.Clear
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function doIndex(idnx,typ)
	on error resume next	
	Err.Clear
	my_Conn.Execute (idnx),,adCmdText + adExecuteNoRecords
	dbHits = dbHits + 1
	if err.number = 0 THEN
	  if typ = "drop" then
	   'if strDebug then
		Response.Write("<b>" & txtDBIndxDrop & "</b><br /><br />" & vbNewLine)
	   'end if
	  else
	   'if strDebug then
		Response.Write("<b>" & txtDBIndxCreated & "</b><br /><br />" & vbNewLine)
	   'end if
	  end if
	else
	  if typ = "drop" then
	    if err.number <> -2147217887 then
	     'if strDebug then
		  Response.Write("<font color=""#FF0000""><b>" & txtDBIndxNotDrop & "</b><br /><br />" & vbNewLine)
		  Response.Write(err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	     'end if
		end if
	  else
		'ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBIndxNotCreated & "</b><br />" & vbNewLine)
		Response.Write(err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	  end if
	end if
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function createIndex(idnx)
	on error resume next	
	Err.Clear
	 select case strDBType
	   case "mysql"
		my_Conn.Execute(checkit(idnx))
	   case else
		my_Conn.Execute(idnx),,adCmdText + adExecuteNoRecords
	 end select
	dbHits = dbHits + 1
	if err.number = 0 THEN
	  'if strDebug then
		Response.Write("<b>" & txtDBIndxCreated & "</b><br /><br />" & vbNewLine)
	  'end if
	else
	  if err.number <> -2147217887 and err.number <> primaryExists then
		'ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBIndxNotCreated & "</b><br />" & vbNewLine)
		response.Write(err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	  end if
	end if
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function createIndx(idnx)
		on error resume next
	for i = 0 to uBound(idnx)
		Err.Clear
	 select case strDBType
	   case "mysql"
		my_Conn.Execute(checkit(idnx(i)))
	   case else
		my_Conn.Execute(idnx(i)),,adCmdText + adExecuteNoRecords
	 end select
		dbHits = dbHits + 1
	 if err.number = 0 THEN
	   'if strDebug then
		Response.Write("<b>" & txtDBIndxCreated & "</b><br /><br />" & vbNewLine)
	   'end if
	 else
	  if err.number <> -2147217887 and err.number <> primaryExists then
		'ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBIndxNotCreated & "</b><br />" & vbNewLine)
		response.Write(err.number & " | " & err.description & "</font><br /><br />" & vbNewLine)
	  end if
	 end if
	next
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function ParseMySQLdata(s) 
'	' Single quote  
'	" Double quote  
'	\ Blackslash 
'	% Percent  
'	_ Underscore 
  if strDBType = "mysql" then
    s = replace(s,"\","\\")
    s = replace(s,"''","'")
    s = replace(s,"'","\'")
    s = replace(s,"""","\""")
    s = replace(s,"%","\%")
    's = replace(s,"_","\_")
    's = replace(s,"","")
  else
  end if
  ParseMySQLdata = s
end function

function populateA(str)
	on error resume next
	Err.Clear
	my_Conn.Execute(str)
	dbHits = dbHits + 1
	if err.number = 0 THEN
	   'if strDebug then
		Response.Write("    <b>" & txtDBTblPopulated & "</b><br />" & vbNewLine)
	   'end if
	else
		ErrorCount = ErrorCount + 1
		Response.Write("<font color=""#FF0000""><b>" & txtDBTblNotPopulated & "</b></font><br />" & vbNewLine)
'		Response.Write("    " & str & "<br />" & vbNewLine)
		if err.count = 1 and err.number <> 438 then
			response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		else
			if err.number = 438 then
				Response.Write("    <font color=""#FF0000"">" & txtDBTblFldNoExist & "</font><br />" & vbNewLine)
			else
				intErrors = err.count
				for er = 0 to intErrors - 1
				response.Write("    <font color=""#FF0000"">" & err(er).number & " | " & err.description & "</font><br />" & vbNewLine)
				next
			end if
		end if
		Err.Clear
	end if
	on error goto 0
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function populateB(str)
	dim iECnt
	iECnt = 0
	for i = 2 to uBound(str)
		on error resume next
		Err.Clear
		strSql = "INSERT INTO " & checkit(str(0)) & " (" & checkit(str(1)) &  ") VALUES (" & str(i) & ")"
		'response.Write strSql & " X<br />"
		my_Conn.Execute(strSql)
		dbHits = dbHits + 1
	if err.number = 0 THEN
		'Response.Write("    <b>" & txtDBTblPopulated & "</b><br />" & vbNewLine)
	else
		ErrorCount = ErrorCount + 1
		iECnt = iECnt + 1
		Response.Write("    <font color=""#FF0000""><b><br />" & txtDBTblNotPopulated & "</b></font><br />" & vbNewLine)
			if err.count = 1 then
				if err.number = 438 then
					Response.Write("    <font color=""#FF0000"">" & txtDBTblFldNoExist & "</font><br />" & vbNewLine)
				else
					response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
				end if
			else
				intErrors = err.count
				for er = 0 to intErrors - 1
				response.Write("    <font color=""#FF0000"">" & err(er).number & " | " & err.description & "</font><br />" & vbNewLine)
				next
			end if
			'response.Write(strSql & "<br />")
		Err.Clear
	end if
	next
	if iECnt = 0 then
	   'if strDebug then
	  Response.Write("<b>" & txtDBTblPopulated & "</b><br />")
	   'end if
	else
	end if
	'response.Write("Error Count: " & ErrorCount & "<br />")
end function

function UpdateErrorCheck()

dim intErrorNumber
dim counter

	intErrorNumber = 0
	for counter = 0 to my_Conn.Errors.Count -1
		intErrorNumber = my_Conn.Errors(counter).Number
		if intErrorNumber <> 0 then  
			select case intErrorNumber
				case -2147217900
					UpdateErrorCheck = 1
					counter = my_Conn.Errors.Count -1
				case -2147467259
					UpdateErrorCheck = 2
					counter = my_Conn.Errors.Count -1	
				case else
					UpdateErrorCheck = intErrorNumber
			end select
		end if
	next
end function

':::::: MISC FUNCTIONS

sub delFPusers(num)
  dim mmSQL, tmp_Col
  tmp_Col = ""
  mmSQL = "select fp_leftcol, fp_rightcol, fp_maincol, fp_uid from PORTAL_FP_USERS where fp_leftcol like '%" & num & "%' or fp_maincol like '%" & num & "%' or fp_rightcol like '%" & num & "%'"
  set rsFPd = my_Conn.execute(mmSQL)
  if not rsFPd.eof then
    do until rsFPd.eof
      Col1 = rsFPd("fp_leftcol")
	  Col2 = rsFPd("fp_rightcol")
	  Col3 = rsFPd("fp_maincol")
	  membID = rsFPd("fp_uid")
	  if Col1 <> "" then
		tmp_Col = ""
	    if instr(Col1,",") then
		  alCol = split(Col1,",")
		  cnt = ubound(alCol)
		  for al = 0 to cnt
		    if instr(trim(alCol(al)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & alCol(al) & ","
		    end if
		  next
		else
		  if instr(Col1,num) = 0 then
		    tmp_Col = Col1
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col1 = tmp_Col
	  end if
	  if Col2 <> "" then
		tmp_Col = ""
	    if instr(Col2,",") then
		  arCol = split(Col2,",")
		  cnt = ubound(arCol)
		  for ar = 0 to ubound(arCol)
		    if instr(trim(arCol(ar)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & arCol(ar) & ","
		    end if
		  next
		else
		  if instr(Col2,num) = 0 then
		    tmp_Col = Col2
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col2 = tmp_Col
	  end if
	  if Col3 <> "" then
		tmp_Col = ""
	    if instr(Col3,",") then
		  amCol = split(Col3,",")
		  cnt = ubound(amCol)
		  for ar = 0 to ubound(amCol)
		    if instr(trim(amCol(ar)),num) = 0 then
		    'if num <> trim(alCol(al)) then
		      tmp_Col = tmp_Col & amCol(ar) & ","
		    end if
		  next
		else
		  if instr(Col3,num) = 0 then
		    tmp_Col = Col3
		  end if
		end if
		if right(tmp_Col,1) = "," then
		   tmp_Col = left(tmp_Col,len(tmp_Col)-1)
		end if
		Col3 = tmp_Col
	    end if
		sSQL = "UPDATE PORTAL_FP_USERS SET "
		sSQL = sSQL & "fp_leftcol = '" & Col1 & "'"
		sSQL = sSQL & ",fp_maincol = '" & Col3 & "'"
		sSQL = sSQL & ",fp_rightcol = '" & Col2 & "'"
		sSQL = sSQL & " WHERE fp_uid = " & membID
		'response.Write(sSQL & "<br />")
		executeThis(sSQL)
		rsFPd.movenext
	loop
  end if
  set rsFPd = nothing
end sub

sub updateVersion(v,i)
  ':: update the app module version
  strSql = "UPDATE " & strTablePrefix & "APPS SET APP_VERSION = '"& v &"'"
  strSql = strSql & ",APP_DATE = '"& DateToStr(now()) &"'"
  strSql = strSql & " WHERE APP_INAME = '" & i & "';"
  executeThis(strSql)
end sub

':: this creates a new group and returns the group ID
function createGroup(n,i,d)
  ng = 0
  sSql = "INSERT INTO " & strTablePrefix & "GROUPS ("
  sSql = sSql & "G_NAME"
  sSql = sSql & ", G_INAME"
  sSql = sSql & ", G_DESC"
  sSql = sSql & ", G_CREATE"
  sSql = sSql & ", G_ACTIVE"
  sSql = sSql & ", G_ADDMEM"
  sSql = sSql & ") VALUES ("
  sSql = sSql & "'" & n & "'"
  sSql = sSql & ", '" & i & "'"
  sSql = sSql & ", '" & d & "'"
  sSql = sSql & ", '" & DateToStr(now()) & "'"
  sSql = sSql & ", 1"
  sSql = sSql & ", 1"
  sSql = sSql & ")"
  executeThis(sSql)
  
  sSql = "SELECT G_ID FROM " & strTablePrefix & "GROUPS WHERE G_INAME='" & i & "'"
  set rsA = my_Conn.execute(sSql)
    ng = rsA("G_ID")
  set rsA = nothing
  createGroup = ng
end function

function addMember2Group(m,g)
  if left(g,1) = "," then
    g = right(g,len(g)-1)
  end if
  sSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS ("
  sSql = sSql & "G_MEMBER_ID"
  sSql = sSql & ", G_GROUP_ID"
  sSql = sSql & ", G_GROUP_LEADER"
  sSql = sSql & ") VALUES ("
  sSql = sSql & m
  sSql = sSql & ", " & g
  sSql = sSql & ", 0"
  sSql = sSql & ")"
  executeThis(sSql)
end function
%>