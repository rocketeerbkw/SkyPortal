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
dim bCanSubmit
bCanSubmit = false

function modGrpEdit(pg,m,c,s,al,ih)
  if len(al) > 3 then
    algn = " align=""" & al & """"
  end if
  select case ih
    case 0
	  gimg = "icon_access_off.gif"
	case 1
	  gimg = "icon_access_on.gif"
	case 2
	  gimg = "icon_access.gif"
  end select
  'popUpWind('calendar_pop.asp?cmd=1&event_id=" & rsEvents("event_id") & "','event','400','400','yes','yes');
  if isnumeric(m) and isnumeric(c) then
    modGrpEdit = "<a href=""javascript:;"" onclick=""JavaScript:popUpWind('" & pg & "?mode=editAccess&amp;cmd=1&amp;cid=" & c & "&amp;sid=" & s & "','groups','430','580','yes','yes');""><img src=""themes/" & strTheme & "/icons/" & gimg & """ title=""" & txtEdGrpAccess & """ alt=""" & txtGrpAccess & """ border=""0"" style=""display:inline;"" hspace=""2""" & algn & "></a>"
  else
    modGrpEdit = "<a href=""" & pg & """><img src=""themes/" & strTheme & "/icons/" & gimg & """ title=""" & txtManage & """ alt=""" & txtManage & """ border=""0"" hspace=""2""" & algn & "></a>"
  end if
end function

function grpCompare(cur,old)
  tmpLst = ""
  if cur <> "" and cur <> "0" then
    if instr(cur,",") > 0 then
	  arrCur = split(cur,",")
	else
	  redim arrCur(0)
	  arrCur(0) = cur
	end if
    if instr(old,",") > 0 then
	  arrOld = split(old,",")
	else
	  redim arrOld(0)
	  arrOld(0) = old
	end if
	for c = 0 to ubound(arrCur)
	  gotOne = false
	  for o = 0 to ubound(arrOld)
	    if trim(arrCur(c)) = trim(arrOld(o)) then
		  gotOne = true
		end if
	  next
	  if not gotOne and arrCur(c) <> 0 then
	    tmpLst = tmpLst & "," & arrCur(c)
	  end if
	next
  else
  end if
  if left(tmpLst,1) = "," then
    tmpLst = right(tmpLst,len(tmpLst)-1)
  end if
  grpCompare = tmpLst
end function

'grpCatUpdate("PIC_CATEGORIES","CG_",g_read,grpRDel,g_write,grpWDel,g_full,grpFDel)
sub grpCatUpdate(tbl,gPfx,gr,grD,gw,gwD,gf,gfD)
  select case gPfx
    case "CG_"
	  gid = "CAT_ID"
	  sAnd = ""
    case "SG_"
	  gid = "SUBCAT_ID"
	  if cid > 0 then
	    sAnd = " AND CAT_ID=" & cid
	  end if
  end select
	  sSql = "SELECT * FROM " & tbl & " WHERE " & gPfx & "INHERIT <> 1" & sAnd
	  set rsG = my_Conn.execute(sSql)
	  if not rsG.eof then
	    do until rsG.eof
	      g_R = rsG(gPfx & "READ")
		  g_W = rsG(gPfx & "WRITE")
		  g_F = rsG(gPfx & "FULL")
		  g_id = rsG(gid)

		  rsG.movenext
		
	      g_Rn = grpCompare(g_R,grD)
	      g_Wn = grpCompare(g_W,gwD)
	      g_Fn = grpCompare(g_F,gfD)
		  'response.Write("<br />with inherit=false<br />")
		  'response.Write("g_db: " & g_R & "<br />")
		  'response.Write("g_db: " & g_W & "<br />")
		  'response.Write("g_db: " & g_F & "<br /><br />")
		  'response.Write("g_Rn: " & g_Rn & "<br />")
		  'response.Write("g_Wn: " & g_Wn & "<br />")
		  'response.Write("g_Fn: " & g_Fn & "<br /><br />")
	  	  sSql = "UPDATE " & tbl & " SET "
	  	  sSql = sSql & gPfx & "READ='" & g_Rn & "',"
	  	  sSql = sSql & gPfx & "WRITE='" & g_Wn & "',"
	  	  sSql = sSql & gPfx & "FULL='" & g_Fn & "' "
	  	  sSql = sSql & "WHERE " & gid & " = " & g_id & sAnd
		  executeThis(sSql)
		loop
	  end if
		
		  'response.Write("<br />with inherit=true<br />")
		  'response.Write("g_Rn: " & gr & "<br />")
		  'response.Write("g_Wn: " & gw & "<br />")
		  'response.Write("g_Fn: " & gf & "<br /><br />")
	  sSql = "UPDATE " & tbl & " SET "
	  sSql = sSql & gPfx & "READ='" & gr & "',"
	  sSql = sSql & gPfx & "WRITE='" & gw & "',"
	  sSql = sSql & gPfx & "FULL='" & gf & "' "
	  sSql = sSql & "WHERE " & gPfx & "INHERIT = 1" & sAnd
	  executeThis(sSql)
end sub

sub shoGroupAccess(frm,gRead,gWrite,gFull,gLst) %><fieldset style='margin:10px;'>
		<!-- <legend><b>Group Access</b></legend> --><br />
      <table border="0" cellpadding="0" cellspacing="0">
	  <tr><td colspan="2" align="center" class="fSubTitle"><b><%= txtGrpsRead %></b></td></tr>
  	  <tr><td align="right" valign="middle" width="50%" class="fNorm" nowrap>
  		<a href="JavaScript:allowgroups('<%= frm %>','g_read','<%= gLst %>');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:removeGroup('<%= frm %>','g_read');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:eGroup('<%= frm %>','g_read');" title="<%= txtCM10 %>"><b><%= txtEditGrp %></b></a>&nbsp;&nbsp;
		<!-- <a href="JavaScript:eGroup('<%= frm %>','g_read','');" title="<%= txtCM10 %>"><b>Edit Group</b></a>&nbsp;&nbsp;<br /> -->
		</td>
          <td align="left"><p>
            <select size="5" name="g_read" style="width:120;" multiple>
			  <% if gRead <> "" then
			  		getOptGroups(gRead)
				 end if %>
			  <option value="0"></option>
            </select><br />&nbsp;</p>
          </td>
        </tr>
  
	  <tr><td colspan="2" align="center" class="fSubTitle"><b><%= txtGrpsWrite %></b></td></tr>
  <tr><td align="right" valign="middle" class="fNorm" nowrap><a href="JavaScript:moveGroup('Add','<%= frm %>','g_read','g_write');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
				<a href="JavaScript:removeGroup('<%= frm %>','g_write');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:eGroup('<%= frm %>','g_write');" title=""><b><%= txtEditGrp %></b></a>&nbsp;&nbsp;</td>
          <td align="left"><p>
            <select size="5" name="g_write" style="width:120;" multiple>
			  <% if gWrite <> "" then
			  		getOptGroups(gWrite)
				 end if %>
			  <option value="0"></option>
            </select><br />&nbsp;</p>
          </td>
        </tr>
  
	  <tr><td colspan="2" align="center" class="fSubTitle"><b><%= txtGrpsFull %></b></td></tr>
  <tr><td align="right" valign="middle" class="fNorm" nowrap><a href="JavaScript:moveGroup('Add','<%= frm %>','g_write','g_full');" title="<%= txtCM10 %>"><b><%= txtCM09 %></b></a>&nbsp;&nbsp;<br />
				<a href="JavaScript:removeGroup('<%= frm %>','g_full');" title="<%= txtCM12 %>"><b><%= txtCM11 %></b></a>&nbsp;&nbsp;<br />
		<a href="JavaScript:eGroup('<%= frm %>','g_full');" title=""><b><%= txtEditGrp %></b></a>&nbsp;&nbsp;</td>
          <td align="left"><p>
            <select size="5" name="g_full" style="width:120;" multiple>
			  <% if gFull <> "" then
			  		getOptGroups(gFull)
				 end if %>
			  <option value="0"></option>
            </select><br />&nbsp;</p>
          </td>
        </tr></table>
		</fieldset>
<%
end sub

sub getOptGroups(gid)
  if gid <> "" then
	gid = gid & ",0"
	arrTemp = split(gid,",")
	for xp = 0 to ubound(arrTemp)-1
	  sSQL = "select G_ID, G_NAME from " & strTablePrefix & "GROUPS where G_ID = " & arrTemp(xp) & ";"
	  'response.Write(sSQL & "<br />")
	  set rsGrp = my_Conn.execute(sSQL)
	  if not rsGrp.eof then
	    response.Write("<option value=""" & rsGrp("G_ID") & """>" & rsGrp("G_NAME") & "</option>" & vbCrLf)
	  end if
	next
	set rsGrp = nothing
  end if
end sub

function chkGrpAdmin(strUsrs)
	sAdmin = false
	tstr = ""
	tmpUsrs = ""
	if len(strUsrs) > 0 and strUsrs <> "0" then
	  tmpArr = split(strUsrs,",")
	  for sp = 0 to ubound(tmpArr)
		if tmpArr(sp) = 1 then
		  sAdmin = true
		end if
	  next
	  if not sAdmin then
		tmpUsrs = "1," & strUsrs
	  else
		tmpUsrs = strUsrs
	  end if
	  'if instr(tmpUsrs,",0") > 0 then
	  tArr = split(tmpUsrs,",")
	  for sp = 0 to ubound(tArr)
		if trim(tArr(sp)) <> "0" then
		  tstr = tstr & "," & tArr(sp)
		end if
	  next
	  tstr = right(tstr,len(tstr)-1)
	  tmpUsrs = tstr
	  'end if
	else
	  tmpUsrs = "1"
	end if
	chkGrpAdmin = tmpUsrs
end function

function hasAccess(mstr)
  tmpAccess = false
 if len(mstr) > 0 and isArray(arrGroups) then
  if not isArray(mstr) and instr(mstr,",") = 0 then
  	    for xg = 0 to ubound(arrGroups)
	      if cLng(trim(arrGroups(xg,0))) = 1 then 'is admin group
		    tmpAccess = true
		  else
		    if cLng(trim(arrGroups(xg,0))) = cLng(trim(mstr)) then
		      tmpAccess = true
			end if
		  end if
	    next
  else 'it is an array or is comma delimited
    if isArray(mstr) then
	  tmpMstr = join(mstr,",")
	else
	  tmpMstr = mstr
	end if
	arrChk = split(tmpMstr,",")
	for xb = 0 to ubound(arrChk)
	  if tmpAccess = false then
		  for xg = 0 to ubound(arrGroups)
			if tmpAccess = false then
	      	  if cLng(trim(arrGroups(xg,0))) = 1 then 'is admin group
		    	  tmpAccess = true
		  	  else
		    	if cLng(trim(arrGroups(xg,0))) = cLng(trim(arrChk(xb))) then
		      	  tmpAccess = true
				end if
		  	  end if
		    end if
		  next
	  end if
	next
  end if 'is array or comma delimited
 else 
   ':: parameter was empty
 end if
  hasAccess = tmpAccess
end function

Function chkApp(app,fld)
  dim appActive, bHasAccess, ckA, strTmpA, agb, agx
  appActive = false
  bHasAccess = false
  for ckA = 0 to ubound(arrAppPerms)
    if arrAppPerms(ckA,1) = app then
	  if arrAppPerms(ckA,2) = 0 then 'app not active
	    if intIsSuperAdmin then
		  appActive = true
		end if
	  else
	    appActive = true
	  end if
      if appActive then
        select case fld
	      case "USERS"
			  strTmpA = arrAppPerms(ckA,3) ':: app 'read' access
			  arTmpA = split(strTmpA,",")
			  for agb = 0 to ubound(arTmpA)
				if intIsSuperAdmin then
				  bHasAccess = true
				else
	              for agx = 0 to ubound(arrGroups)
				    'response.Write(arTmpA(agx) & "<br />")
	          		if trim(arrGroups(agx,0)) = trim(arTmpA(agb)) then
		        		bHasAccess = true
		      		end if
		      	  next
				end if
			  next
			  if app = "PM" and PMaccess = 0 then
			    bHasAccess = false
			  end if
	      case else
		    bHasAccess = false
	    end select
      end if
	  exit for
	end if
  next
  set arTmpA = nothing
  chkApp = bHasAccess
end function

function bldArrUserGroup()
  if strUserMemberID > 0 then ':: they are a member
	strSql = "SELECT G_GROUP_ID, G_GROUP_LEADER FROM " & strTablePrefix & "GROUP_MEMBERS WHERE G_MEMBER_ID = " & strUserMemberID
	set rsApp = my_Conn.execute(strSql)
	if not rsApp.eof then
		tmpArr1 = "2," 'add member group by default
		tmpArr2 = "0,"
		do until rsApp.eof
			tmpArr1 = tmpArr1 & rsApp("G_GROUP_ID") & ","
			tmpArr2 = tmpArr2 & rsApp("G_GROUP_LEADER") & ","
			rsApp.movenext
		loop
		if tmpArr1 <> "" then
			tmpArr3 = split(tmpArr1,",")
			tmpArr4 = split(tmpArr2,",")
			acnt = ubound(tmpArr3)-1
			redim arrGroups(acnt,1)
			for ag = 0 to ubound(tmpArr3)-1
				arrGroups(ag,0) = tmpArr3(ag)
				arrGroups(ag,1) = tmpArr4(ag)
			next
		end if
	else
		redim arrGroups(0,1)
	 	arrGroups(0,0) = "2" 'members group
	 	arrGroups(0,1) = "0" 'not group leader
	end if
	set rsApp = nothing
	
  else '::they are a guest
	redim arrGroups(0,1)
	arrGroups(0,0) = "3" 'GUEST group
	arrGroups(0,1) = "0" 'not group leader
  end if
end function

Function setAppPerms(app,fld)
  dim appActive, ckG, chk
  appActive = false
  'intAppID = 0
  sAppRead = "1"
  sAppWrite = "1"
  sAppFull = "1"
  bAppRead = false
  bAppWrite = false
  bAppFull = false
  chk = ""
  for ckG = 0 to ubound(arrAppPerms)
    if fld = "id" then
	  chk = arrAppPerms(ckG,0)
	else
	  chk = trim(arrAppPerms(ckG,1))
	end if
    if chk = trim(app) then
	  if arrAppPerms(ckG,2) = 0 then 'app not active
	    if intIsSuperAdmin then
	      intAppID = arrAppPerms(ckG,0)
		  intAppActive = 0
  		  sAppRead = "1"
  		  sAppWrite = "1"
  		  sAppFull = "1"
  		  bAppRead = true
  		  bAppWrite = true
  		  bAppFull = true
	      if intSubscriptions = 1 then
	        intSubscriptions = arrAppPerms(ckG,6)
	      end if
	      if intBookmarks = 1 then
	        intBookmarks = arrAppPerms(ckG,7)
	      end if
	      intSecCode = arrAppPerms(ckG,8)
		  iData1 = arrAppPerms(ckG,9)
		  iData2 = arrAppPerms(ckG,10)
		  iData3 = arrAppPerms(ckG,11)
		  iData4 = arrAppPerms(ckG,12)
		  iData5 = arrAppPerms(ckG,13)
		  iData6 = arrAppPerms(ckG,14)
		  iData7 = arrAppPerms(ckG,15)
		  iData8 = arrAppPerms(ckG,16)
		  iData9 = arrAppPerms(ckG,17)
		  iData10 = arrAppPerms(ckG,18)
		  tData1 = arrAppPerms(ckG,19)
		  tData2 = arrAppPerms(ckG,20)
		  tData3 = arrAppPerms(ckG,21)
		  tData4 = arrAppPerms(ckG,22)
		  tData5 = arrAppPerms(ckG,23)
		  exit for
		end if
	  else
	    appActive = true
	  end if
      if appActive then
	    intAppID = arrAppPerms(ckG,0)
		intAppActive = arrAppPerms(ckG,2)
	    sAppRead = arrAppPerms(ckG,3) ':: app 'Read' access
	   	sAppWrite = arrAppPerms(ckG,4) ':: app 'Write' access
	    sAppFull = arrAppPerms(ckG,5) ':: app 'Full' access
  		bAppFull = hasAccess(sAppFull)
		if bAppFull then
  		  bAppWrite = true
  		  bAppRead = true
		else
  		  bAppWrite = hasAccess(sAppWrite)
  		  bAppRead = hasAccess(sAppRead)
		end if
	    if intSubscriptions = 1 then
	      intSubscriptions = arrAppPerms(ckG,6)
	    end if
	    if intBookmarks = 1 then
	      intBookmarks = arrAppPerms(ckG,7)
	    end if
	    intSecCode = arrAppPerms(ckG,8)
		iData1 = arrAppPerms(ckG,9)
		iData2 = arrAppPerms(ckG,10)
		iData3 = arrAppPerms(ckG,11)
		iData4 = arrAppPerms(ckG,12)
		iData5 = arrAppPerms(ckG,13)
		iData6 = arrAppPerms(ckG,14)
		iData7 = arrAppPerms(ckG,15)
		iData8 = arrAppPerms(ckG,16)
		iData9 = arrAppPerms(ckG,17)
		iData10 = arrAppPerms(ckG,18)
		tData1 = arrAppPerms(ckG,19)
		tData2 = arrAppPerms(ckG,20)
		tData3 = arrAppPerms(ckG,21)
		tData4 = arrAppPerms(ckG,22)
		tData5 = arrAppPerms(ckG,23)
      end if
	  exit for
	end if
  next
end function

Function getAppPerms(app,fld,typ)
  dim appActive, ckG, chk, tGr
  appActive = false
  tGr = ""
  chk = ""
  for ckG = 0 to ubound(arrAppPerms)-1
    if typ = "id" then
	  chk = arrAppPerms(ckG,0)
	else
	  chk = trim(arrAppPerms(ckG,1))
	end if
    if chk = trim(app) then
	  if arrAppPerms(ckG,2) = 0 then 'app not active
	    if intIsSuperAdmin then
	      appActive = true
		  tGr = "1"
		  exit for
		end if
	  else
	    appActive = true
	  end if
      if appActive then
	    select case fld
		  case "read"
		    tGr = arrAppPerms(ckG,3) ':: app 'Read' access
		  case "write"
		    tGr = arrAppPerms(ckG,4) ':: app 'Write' access
		  case "full"
		    tGr = arrAppPerms(ckG,5) ':: app 'Full' access
		  case else
		    tGr = "1"
		end select
      end if
	  exit for
	end if
  next
  getAppPerms = tGr
end function

function setPermVars(sq,typ)
  'sAppRead, sAppWrite, sAppFull
  'bAppRead, bAppWrite, bAppFull
  bCatRead = false
  bCatWrite = false
  bCatFull = false
  bSCatRead = false
  bSCatWrite = false
  bSCatFull = false
  if typ = 1 or typ = 2 then
      sCatFull = sq("CG_FULL")
      sCatRead = sq("CG_READ")
      sCatWrite = sq("CG_WRITE")
	  bCatFull = hasAccess(sCatFull)
  end if
  if typ = 2 then
	  sSCatFull = sq("SG_FULL")
	  sSCatRead = sq("SG_READ")
	  sSCatWrite = sq("SG_WRITE")
	bSCatFull = hasAccess(sSCatFull)
  end if
  if bAppFull then
	bCatRead = true
	bCatWrite = true
	bCatFull = true
	bSCatRead = true
	bSCatWrite = true
	bSCatFull = true
  else
    if bCatFull then	  
	  bCatRead = true
	  bCatWrite = true
	  bSCatRead = true
	  bSCatWrite = true
	  bSCatFull = true
	else
	  bCatRead = hasAccess(sCatRead)
	  bCatWrite = hasAccess(sCatWrite)
  	  if typ = 2 then
	    if bSCatFull then
	      bSCatRead = true
	      bSCatWrite = true
	    else
	      bSCatRead = hasAccess(sSCatRead)
	      bSCatWrite = hasAccess(sSCatWrite)
	    end if
	  end if
  	  if typ = 1 or typ = 2 then
	    if bCatWrite then
	      'bSCatWrite = bCatWrite
	    end if
	  end if
	end if
  end if
end function

Function chkAppActive(app)
  dim appActive
  appActive = true
  sSQL = "Select APP_ACTIVE from " & strTablePrefix & "APPS WHERE APP_INAME = '" & app & "'"
  set appChk = my_Conn.execute(sSQL)
  if not appChk.eof then
    if appChk("APP_ACTIVE") = 0 then
      appActive = false
	end if
  end if
  set appChk = nothing
  chkAppActive = appActive
end function

sub debugPermVars()
		response.Write("Cat: " & cid & "<br />")
		response.Write("inherit: " & inherit & "<br /><br />")
		response.Write("app_R: " & sAppRead & "<br />")
		response.Write("app_W: " & sAppWrite & "<br />")
		response.Write("app_F: " & sAppFull & "<br /><br />")
		response.Write("frm_R: " & g_read & "<br />")
		response.Write("frm_W: " & g_write & "<br />")
		response.Write("frm_F: " & g_full & "<br /><br />")
		response.Write("db_R: " & sCatRead & "<br />")
		response.Write("db_W: " & sCatWrite & "<br />")
		response.Write("db_F: " & sCatFull & "<br /><br />")
		response.Write("grpRDel: " & grpRDel & "<br />")
		response.Write("grpWDel: " & grpWDel & "<br />")
		response.Write("grpFDel: " & grpFDel & "<br /><br />")
end sub

sub updateAccess()
  tCat = strTablePrefix & "M_CATEGORIES"
  tSub = strTablePrefix & "M_SUBCATEGORIES"
  pag = app_pop
  
  g_full = chkGrpAdmin(request.Form("g_full"))
  g_write = chkGrpAdmin(request.Form("g_write"))
  g_read = chkGrpAdmin(request.Form("g_read"))
  inherit = request.Form("inherit")
  if inherit <> 1 then inherit = 0
  'cid
  'sid
  select case request.Form("cmd")
    case 1
	  ':: update module groups
	  if not bAppFull then
	    ':: show no access
	  else
		grpRDel = grpCompare(sAppRead,g_read)
		grpWDel = grpCompare(sAppWrite,g_write)
		grpFDel = grpCompare(sAppFull,g_full)
		'response.Write("App: " & intAppID & "<br />")
		'response.Write("grpRDel: " & grpRDel & "<br />")
		'response.Write("grpWDel: " & grpWDel & "<br />")
		'response.Write("grpFDel: " & grpFDel & "<br /><br />")
		grpCatUpdate tCat,"CG_",g_read,grpRDel,g_write,grpWDel,g_full,grpFDel
		grpCatUpdate tSub,"SG_",g_read,grpRDel,g_write,grpWDel,g_full,grpFDel

			strSql = "UPDATE " & strTablePrefix & "APPS SET "
			strSql = strSql & "APP_GROUPS_USERS = '" & g_read & "'"
			strSql = strSql & ", APP_GROUPS_WRITE = '" & g_write & "'"
			strSql = strSql & ", APP_GROUPS_FULL = '" & g_full & "'"
			strSql = strSql & " WHERE APP_ID = " & intAppID
			executeThis(strSql)
	  end if
	case 2
	  ':: update category groups
	  if sid = 0 and cid > 0 then
		sSQL = "select CAT_NAME, CG_READ, CG_WRITE, CG_FULL, CG_INHERIT, CG_PROPAGATE from " & tCat & " WHERE CAT_ID=" & cid
  		set rsT = my_Conn.execute(sSQL)
  		cat_name = rsT("CAT_NAME")
		cg_inherit = rsT("CG_INHERIT")
		cg_propagate = rsT("CG_PROPAGATE")
		sub_name = ""
  '		sub_name = rsT("SUBCAT_NAME")
  		call setPermVars(rsT,1)
  		set rsT = nothing
		
	    if not bCatFull then
	    ':: show no access
		  Response.Write("<b>No Access!</b>")
	    else
  		  c_name = chkstring(replace(request.Form("c_name"),"'",""),"sqlstring")
		  if len(c_name) = 0 then 
		    c_name = cat_name
		  end if
		  if inherit = 1 then
		    g_read = sAppRead
			g_write = sAppWrite
			g_full = sAppFull
		  else
		    ':: Keep form variables
		  end if
		  grpRDel = grpCompare(sCatRead,g_read)
		  grpWDel = grpCompare(sCatWrite,g_write)
		  grpFDel = grpCompare(sCatFull,g_full)
		  'grpRAdd = grpCompare(g_read,sCatRead)
		  'grpWAdd = grpCompare(g_write,sCatWrite)
		  'grpFAdd = grpCompare(g_full,sCatFull)
		  
		  'debugPermVars()

		  strSql = "UPDATE " & tCat & " SET "
		  strSql = strSql & "CG_READ = '" & g_read & "'"
		  strSql = strSql & ", CG_WRITE = '" & g_write & "'"
		  strSql = strSql & ", CG_FULL = '" & g_full & "'"
		  strSql = strSql & ", CG_INHERIT = " & inherit & ""
		  strSql = strSql & ", CAT_NAME = '" & c_name & "'"
		  strSql = strSql & " WHERE CAT_ID = " & cid
		  'Response.Write(strSql & "<br />")
		  executeThis(strSql)
		  
		  grpCatUpdate tSub,"SG_",g_read,grpRDel,g_write,grpWDel,g_full,grpFDel
		  
		end if
	  else
	    ':: show no access
	  end if
	case 3
	  ':: update subcategory groups
	  if sid > 0 and cid > 0 then
	    'response.Write("This is a subcategory group edit")
		sSQL = "SELECT " & tCat & ".CAT_ID, " & tCat & ".CAT_NAME, " & tCat & ".CG_READ, " & tCat & ".CG_WRITE, " & tCat & ".CG_FULL, " & tCat & ".CG_INHERIT, " & tCat & ".CG_PROPAGATE, " & tSub & ".SUBCAT_ID, " & tSub & ".SUBCAT_NAME, " & tSub & ".SG_READ, " & tSub & ".SG_WRITE, " & tSub & ".SG_FULL, " & tSub & ".SG_INHERIT "
		sSQL = sSQL & "FROM " & tCat & " INNER JOIN " & tSub & " ON " & tCat & ".CAT_ID = " & tSub & ".CAT_ID "
		sSQL = sSQL & "WHERE (((" & tCat & ".CAT_ID)=" & cid & ") AND ((" & tSub & ".SUBCAT_ID)=" & sid & "));"
  		set rsT = my_Conn.execute(sSQL)
  		cat_name = rsT("CAT_NAME")
		sub_name = rsT("SUBCAT_NAME")
		'inherit = rsT("SG_INHERIT")
  		call setPermVars(rsT,2)
  		set rsT = nothing
	
	    if not bSCatFull then
	    ':: show no access
	    else
  		  cat = clng(request.Form("cat"))
  		  s_name = chkstring(replace(request.Form("s_name"),"'",""),"sqlstring")
		  if len(s_name) = 0 then 
		    s_name = sub_name
		  end if
		  if inherit = 1 then
		    g_read = sCatRead
			g_write = sCatWrite
			g_full = sCatFull
		  else
		    ':: Keep form variables
		  end if
		  
		  ':: check for category change
		  if cat <> cid then
		    sSql = "SELECT CG_READ,CG_WRITE,CG_FULL FROM " & tCat & " WHERE CAT_ID=" & cat
			set rsT = my_Conn.execute(sSql)
			if rst.eof then
			  cat = cid
			else
		      g_read = rsT("CG_READ")
			  g_write = rsT("CG_WRITE")
			  g_full = rsT("CG_FULL")
			  inherit = 1
			  cid = cat
			end if
			set rsT = nothing
		  else
			cat = cid
		  end if
		  
		    strSql = "UPDATE " & tSub & " SET "
		    strSql = strSql & "SG_READ = '" & g_read & "'"
		    strSql = strSql & ", SG_WRITE = '" & g_write & "'"
		    strSql = strSql & ", SG_FULL = '" & g_full & "'"
		    strSql = strSql & ", SG_INHERIT = " & inherit & ""
		    strSql = strSql & ", CAT_ID = " & cat & ""
		    strSql = strSql & ", SUBCAT_NAME = '" & s_name & "'"
		    strSql = strSql & " WHERE SUBCAT_ID = " & sid
		  'Response.Write(strSql & "<br />")
		  executeThis(strSql)
	    end if
	  else
	    ':: show no access
	  end if
	    ':: show no access
	case else
  end select
  resetCoreConfig()
  %>
  <script type="text/javascript"> 
	opener.document.location.reload();
  </script>
<%  
  editAccessForm()
end sub

sub editAccessForm()
  tC = strTablePrefix & "M_CATEGORIES"
  tS = strTablePrefix & "M_SUBCATEGORIES"
  pg = app_pop
  		  iLoc = 0
  		  bAccess = false
  		  iApp = 0
  		  iCat = 2
  		  iSub = 3
  response.Write("<br />")
	  grpRead = ""
	  grpWrite = ""
	  grpFull = ""
	  p_read = ""
	if sid = 0 and cid = 0 and bAppFull then
	    'response.Write("This is a module group edit")
  		iLoc = 1
  		bAccess = true
		g_full = sAppFull
		g_write = sAppWrite
		g_read = sAppRead
		p_read = ""
  '		cat_name = txtPics
		sub_name = ""
		stMsg = stMsg & "&nbsp;" & txtModGrpEdit & "<hr />"
	elseif sid = 0 and cid > 0 then
	    'response.Write("This is a category group edit")
		sSQL = "select CAT_NAME, CG_READ, CG_WRITE, CG_FULL, CG_INHERIT, CG_PROPAGATE from " & strTablePrefix & "M_CATEGORIES WHERE CAT_ID=" & cid & " AND APP_ID=" & intAppID
  		set rsT = my_Conn.execute(sSQL)
  		cat_name = rsT("CAT_NAME")
		cg_inherit = rsT("CG_INHERIT")
		cg_propagate = rsT("CG_PROPAGATE")
		sub_name = ""
  '		sub_name = rsT("SUBCAT_NAME")
  		call setPermVars(rsT,1)
  		set rsT = nothing
		
	  if bCatFull then
  		bAccess = true
  		iLoc = 2
		g_full = sCatFull
		g_write = sCatWrite
		g_read = sCatRead
		p_read = sAppRead
	    stMsg = replace(txtEditCatGrp,"[%cat_name%]",cat_name) & "<br /><br />"
		if cg_inherit = 1 then
		  stMsg = stMsg & txtGrpModInherit & "&nbsp;"
		else
		  stMsg = stMsg & txtGrpModNoInherit & "&nbsp;"
		end if
		if cg_propagate = 1 then
		stMsg = stMsg & txtGrpPropSub
		end if
		stMsg = stMsg & "<br /><hr />"
	  end if
	elseif sid > 0 and cid > 0 then
	    'response.Write("This is a subcategory group edit")
		sSQL = "SELECT " & tC & ".CAT_ID, " & tC & ".CAT_NAME, " & tC & ".CG_READ, " & tC & ".CG_WRITE, " & tC & ".CG_FULL, " & tC & ".CG_INHERIT, " & tC & ".CG_PROPAGATE, " & tS & ".SUBCAT_ID, " & tS & ".SUBCAT_NAME, " & tS & ".SG_READ, " & tS & ".SG_WRITE, " & tS & ".SG_FULL, " & tS & ".SG_INHERIT "
		sSQL = sSQL & "FROM " & tC & " INNER JOIN " & tS & " ON " & tC & ".CAT_ID = " & tS & ".CAT_ID "
		sSQL = sSQL & "WHERE (((" & tC & ".CAT_ID)=" & cid & ") AND ((" & tS & ".SUBCAT_ID)=" & sid & "));"
		
  		set rsT = my_Conn.execute(sSQL)
  		cat_name = rsT("CAT_NAME")
		sub_name = rsT("SUBCAT_NAME")
		'cg_inherit = rsT("CG_INHERIT")
		'cg_propagate = rsT("CG_PROPAGATE")
		cg_inherit = rsT("SG_INHERIT")
  		call setPermVars(rsT,2)
  		set rsT = nothing
	
	  if bSCatFull then
  		iLoc = 3
  		bAccess = true
		g_full = sSCatFull
		g_write = sSCatWrite
		g_read = sSCatRead
		p_read = sCatRead
	    stMsg = "" & replace(replace(txtEditSubGrp,"[%cat_name%]",cat_name),"[%sub_name%]",sub_name)
		if cg_inherit = 1 then
		stMsg = stMsg & "<br /><br />" & replace(txtGrpCatInherit,"[%cat_name%]",cat_name)
		else
		stMsg = stMsg & "<br /><br />" & replace(txtGrpCatNoInherit,"[%cat_name%]",cat_name)
		end if
		stMsg = stMsg & "<br /></p><hr /><p>"
	  end if
	end if
	if bAccess then
  	  spThemeTitle = txtEdGrpAccess
	  spThemeBlock1_open(intSkin)
	  %>
      <form action="<%= pg %>" method="post" id="GrpAccess" name="GrpAccess" onSubmit="selectUsers('GrpAccess')"><p>
	<input type="hidden" name="mode" value="updAccess">
	<input type="hidden" name="cmd" value="<%= iLoc %>">
	<input type="hidden" name="cid" value="<%= cid %>">
	<input type="hidden" name="sid" value="<%= sid %>">
	<%= stMsg %>
	<%
	  if iLoc = 2 or iLoc = 3 then
	  response.Write("<br /><span class=""fNorm"">" & txtCat & ": </span>")
	  end if
	  if iLoc = 2 then
	    response.Write("<input name=""c_name"" class=""textbox"" type=""text"" value=""" & cat_name & """ maxlength=""40"" style=""margin:0px""><br />")
	  end if
	  if iLoc = 3 then
	    'getSelCats(tC)
		Call mod_selectCats(cid,"FULL")
	    response.Write("<br /><br /><span class=""fNorm"">" & txtSubCatNam & ": </span><input name=""s_name"" class=""textbox"" type=""text"" value=""" & sub_name & """ maxlength=""40"" style=""margin:0px""><br />")
	  end if
	  'end if
	  'p_read
	  if iLoc > 1 then %>
	   <input id="inherit" name="inherit" type="checkbox" value="1"<%= chkRadio(cg_inherit,1) %>> <span class="fNorm"><%= txtInhPerms %></span>
   <% End If
   
	  Call shoGroupAccess("GrpAccess",g_read,g_write,g_full,p_read) %>
   	   <br />
	   <input name="submit" type="submit" class="button" value="<%= txtSubmit %>">
	  </p></form><br />
	  <%
	  spThemeBlock1_close(intSkin)
	else 
	  ':: no access
	end if
  'response.Write("<br /><br />")
end sub

sub getSelCats(c)
  response.Write("<select name=""cat"">" & vbCRLF)
  sSql = "SELECT CAT_ID, CAT_NAME, CG_READ FROM " & c & " WHERE APP_ID=" & intAppID & " ORDER BY C_ORDER, CAT_NAME"
  set oRs = my_Conn.execute(sSql)
  if oRs.eof then
    response.Write("<option value=""0""> [ NO CATEGORIES ] </option>" & vbCRLF)
  else
    do until oRs.eof
	  if hasAccess(oRs("CG_READ")) then
        response.Write("<option value=""" & oRs("CAT_ID") & """" & chkSelect(cid,oRs("CAT_ID")) & "> " & oRs("CAT_NAME") & "</option>" & vbCRLF)
	  end if
	  oRs.movenext
	loop
  end if
  set oRs = nothing
  response.Write("</select>" & vbCRLF)
end sub

sub updateGroup(g_id)
  if not isGrpLeader(g_id) then
    ':: show no permission
  else
	Err_Msg = ""
	sMsg = ""
	'g_name = Request.Form("g_name")
	'g_desc = Request.Form("g_desc")
	g_members = Request.Form("g_write")
	g_leaders = Request.Form("g_full")
	g_modify = strCurDateString
	g_create = g_modify
	if g_members = "" then
	  Err_Msg = txtGrpNeedMbrs
	end if

	if Err_Msg = "" then
		'g_id = cInt(Request.Form("g_id"))
			
		'response.Write("Name: " & g_name & "<br />")
		'response.Write("Desc: " & g_desc & "<br />")
		'response.Write("Members: " & g_members & "<br />")
		'response.Write("Leaders: " & g_leaders & "<br />")

	  strSql = "UPDATE " & strTablePrefix & "GROUPS"
	  strSql = strSql & " SET G_MODIFIED = '" & g_desc & "'"
	  'strSql = strSql & ", G_NAME = '" & g_name & "'"
	  'strSql = strSql & ", G_DESC = '" & g_modify & "'"
	  strSql = strSql & " WHERE G_ID = " & g_id
	  executeThis(strSql)

	  strSql = "DELETE FROM " & strTablePrefix & "GROUP_MEMBERS"
	  strSql = strSql & " WHERE G_GROUP_ID = " & g_id
	  executeThis(strSql)
			
	  if g_members <> "" then
		if inStr(g_members,",") > 0 then
		  arrMembers = split(g_members,",")
		  for g = 0 to ubound(arrMembers)
			strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
			strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
			strSql = strSql & " (" & trim(arrMembers(g)) & "," & g_id & ",0)"
			executeThis(strSql)				  
		  next
		else
		  strSql = "INSERT INTO " & strTablePrefix & "GROUP_MEMBERS"
		  strSql = strSql & " (G_MEMBER_ID, G_GROUP_ID, G_GROUP_LEADER) VALUES"
		  strSql = strSql & " (" & trim(g_members) & "," & g_id & ",0)"
		  executeThis(strSql)			  
		end if
			  
		' check and insert group leaders	
		if inStr(g_leaders,",") > 0 then
		  arrLeaders = split(g_leaders,",")
		  for h = 0 to ubound(arrLeaders)
			strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
			strSql = strSql & " SET G_GROUP_LEADER = 1"
			strSql = strSql & " WHERE G_MEMBER_ID = " & arrLeaders(h)
			strSql = strSql & " AND G_GROUP_ID = " & g_id
			executeThis(strSql)				  
		  next
		elseif len(g_leaders) > 0 then
		  strSql = "UPDATE " & strTablePrefix & "GROUP_MEMBERS"
		  strSql = strSql & " SET G_GROUP_LEADER = 1"
		  strSql = strSql & " WHERE G_MEMBER_ID = " & g_leaders
		  strSql = strSql & " AND G_GROUP_ID = " & g_id
		  executeThis(strSql)
		end if		
	  else 'g_members = ""
			
	  end if 'g_members check
	  'closeandgo("admin_config_groups.asp")
	  sMsg = "<span class=""fAlert""><b>" & txtCG01 & "</b></span>"
	else 
	  ':: there is an error message
	  sMsg = "<span class=""fSubTitle"">" & txtThereIsProb & "</span>"
	  sMsg = sMsg & "<ul>" & Err_Msg & "</ul>"
	end if
	  'response.Write(sMsg)
	  editGroupForm(g_id)
  end if
end sub

function isGrpLeader(gp)
  bAcc = false
  if hasAccess(1) then
    bAcc = true
  else
    sSql = "SELECT PORTAL_GROUPS.G_ID, PORTAL_GROUPS.G_NAME, PORTAL_GROUPS.G_DESC, PORTAL_GROUPS.G_ADDMEM, PORTAL_GROUP_MEMBERS.G_MEMBER_ID, PORTAL_GROUP_MEMBERS.G_GROUP_ID, PORTAL_GROUP_MEMBERS.G_GROUP_LEADER "
    sSql = sSql & "FROM PORTAL_GROUPS INNER JOIN PORTAL_GROUP_MEMBERS ON PORTAL_GROUPS.G_ID = PORTAL_GROUP_MEMBERS.G_GROUP_ID "
    sSql = sSql & "WHERE (((PORTAL_GROUP_MEMBERS.G_MEMBER_ID)=" & strUserMemberID & "));"
    set rsT = my_Conn.execute(sSql)
    if not rsT.eof then
      if rsT("G_GROUP_LEADER") = 1 then
        bAcc = true
	  end if
    end if
    set rsT = nothing
  end if
  isGrpLeader = bAcc
end function

sub editGroupForm(g)
  if not isGrpLeader(g) then
  ':: show no access
    Response.Write(txtNoAccess)
  else
   'bGrpLeader = isGrpLeader(g)
	g_field = ""
	sSQL = "select * from " & strTablePrefix & "GROUPS where G_ID = " & g
	set rsGrp = my_Conn.execute(sSQL)
	if not rsGrp.eof then
	  gid = rsGrp("G_ID")
	  gname = rsGrp("G_NAME")
	  gdesc = rsGrp("G_DESC")
	  gcreate = strtodate(rsGrp("G_CREATE"))
	  if trim(rsGrp("G_MODIFIED")) <> "" then
		gmodify = strtodate(rsGrp("G_MODIFIED"))
	  end if
	  gactive = rsGrp("G_ACTIVE")
	  btn = txtSknUpdate
	  g_field = ""
	end if
	Set rsGrp = nothing %>
<script type="text/javascript">
function selectPeeps(){
  //alert(fm);
  selectAll('Grp','g_write');
  selectAll('Grp','g_full');
}
</script>
	<form action="pop_portal.asp" method="post" id="Grp" name="Grp">
	<input type="hidden" name="cmd" value="12">
	<input type="hidden" name="cid" value="<%= g %>">
	<% spThemeBlock1_open(intSkin) %>
<table class="tCellAlt2" border="0" cellspacing="0" cellpadding="0" width="100%">
    <tr> 
      <td> 
        <table border="0" cellspacing="1" cellpadding="3" class="tCellAlt1" width="100%">
		<tr><td align="center" class="tSubTitle">
			<%= txtEdit %>&nbsp;<%= txtGroup %></td>
          </tr>
		  <% If sMsg <> "" Then %>
          <tr> 
            <td align="center"><%= sMsg %></td>
          </tr>
		  <% End If %>
		<tr><td align="left">
		<table border="0" cellpadding="0" cellspacing="5" width="100%">
			<tr><td width="100" align="right"><b><%= txtGrpNam %>:</b> </td><td>&nbsp;<b><%= gname %></b></td></tr>
			<tr><td align="right"><b><%= txtDesc %>:</b> </td><td>&nbsp;<%= gdesc %></td></tr>
		</table>
		
      <br />
      <fieldset style='margin:10px;'>
		<legend><b><%= txtMbrshp %></b></legend>
      <table border=0 cellpadding=0 cellspacing=0 width="100%">
        <tr><td align="center" valign="middle"><a href="JavaScript:allmemberList('Grp','g_write');" title="<%= txtCG14 %>"><b><%= txtAddMem %></b></a><br />
				<a href="JavaScript:removeGroup('Grp','g_write');" title="<%= txtCG15 %>"><b><%= txtRemMember %></b></a></td>
          <td align="center"><b><%= txtCG13 %></b><br />
            <select size="9" name="g_write" style="width:150;" multiple>
			  <% getGroupMembers(g) %>
			  <option value="0"></option>
            </select>
          </td>
        </tr>
        <tr> <td align="center" valign="middle"><br />
				<a href="JavaScript:moveGroup('Add','Grp','g_write','g_full');" title="<%= txtCG16 %>">
				<b><%= txtCG18 %></b></a><br />
				<a href="JavaScript:removeGroup('Grp','g_full');" title="<%= txtCG17 %>">
				<b><%= txtCG19 %></b></a></td>
          <td align=center> <br />
            <br />
            <b><%= txtCG20 %></b><br />
            <select size="9" name="g_full" style="width:150;" multiple>
			  <% getGroupLeaders(g) %>
			  <option value="0"></option>
            </select><br /><br />
          </td>
        </tr>
      </table>
		</fieldset>		
		<br /><div align=center><input name="Submit" type="submit" value="<%= txtSubmit %>" onclick="selectPeeps()" class="button">
		</div><br />
		</td></tr>
	</table>
		</td></tr>
	</table>
	<% spThemeBlock1_close(intSkin) %>
</form><br />
<% 
  end if
End Sub 

sub getGroupMembers(gid)
	sSQL = "select * from " & strTablePrefix & "GROUP_MEMBERS where G_GROUP_ID = " & gid
	'response.Write(sSQL & "<br />")
	set rsGrp = my_Conn.execute(sSQL)
	if not rsGrp.eof then
	  do until rsGrp.eof
	    response.Write("<option value=""" & rsGrp("G_MEMBER_ID") & """>" & getmembername(rsGrp("G_MEMBER_ID")) & "</option>" & vbnewline)
	    rsGrp.movenext
	  loop
	end if
	set rsGrp = nothing
end sub

sub getGroupLeaders(gid)
	sSQL = "select * from " & strTablePrefix & "GROUP_MEMBERS where G_GROUP_ID = " & gid & " AND G_GROUP_LEADER = 1"
	response.Write(sSQL & "<br />")
	set rsGrp = my_Conn.execute(sSQL)
	if not rsGrp.eof then
	  do until rsGrp.eof
	    response.Write("<option value=""" & rsGrp("G_MEMBER_ID") & """>" & getmembername(rsGrp("G_MEMBER_ID")) & "</option>" & vbnewline)
	    rsGrp.movenext
	  loop
	end if
	set rsGrp = nothing
end sub
%>