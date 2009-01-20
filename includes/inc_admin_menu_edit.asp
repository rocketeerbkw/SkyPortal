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

If Request.Form("cmd") <> "" Then
  Ctitle = Trim(Request.Form("mnuTitle"))
  CINAME = Trim(Request.Form("CINAME"))
  if CINAME = "" then CINAME = Ctitle
  mnuErr = ""
  
  Select Case Request.Form("cmd")
    case "editchild"
    		Cid = Trim(Request.Form("id"))
    		Cname = replace(Trim(Request.Form("Cname")),"'","''")
    		Clink = Trim(Request.Form("Clink"))
    		Cimage = Trim(Request.Form("CImage"))
    		Conclick = replace(Trim(Request.Form("Conclick")),"'","''")
    		Ctarget = Trim(Request.Form("Ctarget"))
    		Cfunct = Trim(Request.Form("Cfunct"))
    		CMenu = Trim(Request.Form("CMenu"))
    		CgrpAccess = Trim(Request.Form("g_read"))
			
			if Cname = "" then
			  mnuErr = mnuErr & "<br /><br />" & txtLink & "&nbsp;" & xx & ":&nbsp;" & txtNamNoBlank
			else
			  chkMnu(Cname)
			end if
			
			if mnuErr = "" then
			  if Clink = "" then
'			    mnuErr = mnuErr & "<br />Link " & xx & ": Link cannot be blank"
			  else
			    chkMnu(Clink)
			  end if
			end if
			
			if mnuErr = "" then
			  if Cimage <> "" then
			    chkMnu(Cimage)
			  end if
			end if
			
			if mnuErr = "" then
			  if CgrpAccess = "0" then CgrpAccess = ""
			  if len(CgrpAccess) > 0 then
			    CgrpAccess = chkGrpAdmin(CgrpAccess)
			  end if
			end if
			
			if mnuErr = "" then
			
			sSql = "SELECT NAME, INAME from MENU WHERE ID=" & Cid
			set rsEd = my_Conn.execute(sSQL)
			  pName = rsEd("NAME")
			  pTitle = rsEd("INAME")
			
			  strSQL = "update Menu set Parent='"& Cname &"' where Parent='"& pName &"' and INAME='"& pTitle &"'"
			  executeThis(strSql)
			  
			  'response.Write("Cname: " & Cname & "<br />")
			  'response.Write("pName: " & pName & "<br />")
			  'response.Write("pTitle: " & pTitle & "<br />")
			  'response.End()
			set rsEd = nothing
			
			strSQL = "update Menu set Name = '" & Cname & "', "
			strSQL = strSQL & "Link = '" & Clink & "', "
			strSQL = strSQL & "mnuImage = '" & Cimage & "', "
			strSQL = strSQL & "onclick = '" & Conclick & "', "
			strSQL = strSQL & "Target = '" & Ctarget & "', "
			strSQL = strSQL & "mnuAccess = '" & CgrpAccess & "', "
			strSQL = strSQL & "mnuFunction = '" & Cfunct & "' "
			strSQL = strSQL & "where id = " & Cid
			executeThis(strSql)
			
			end if
			
			CINAME = CMenu
		
	case "addchild"
    		CparentID = Trim(Request.Form("CparentID"))
    		Cparent = Trim(Request.Form("Cparent"))
			if Cparent = "" then
			Cparent = Trim(Request.Form("mnuTitle"))
			end if
    		Cname = replace(Trim(Request.Form("Cname")),"'","''")
    		Clink = Trim(Request.Form("Clink"))
    		Cimage = Trim(Request.Form("CImage"))
    		Conclick = replace(Trim(Request.Form("Conclick")),"'","''")
    		Ctarget = Trim(Request.Form("Ctarget"))
    		Corder = Trim(Request.Form("Corder"))
    		Ctitle = Trim(Request.Form("mnuTitle"))
    		CNINAME = Trim(Request.Form("CINAME"))
    		Cfunct = Trim(Request.Form("Cfunct"))
    		CaddMenu = Trim(Request.Form("CaddMenu"))
    		CMenu = Trim(Request.Form("CMenu"))
    		CgrpAccess = Trim(Request.Form("g_read"))
			
			if CaddMenu <> "" then
			  sSql = "SELECT * FROM MENU WHERE NAME='" & Cparent & "' AND INAME = '" & CaddMenu & "'"
			  set rsT = my_Conn.execute(sSql)
			  if not rsT.eof then
			    CparentID = rsT("ID")
			    Ctitle = rsT("mnuTitle")
			    CNINAME = rsT("INAME")
			    CaddMenu = ""
			  end if
			  set rsT = nothing
			end if
			
			if Cname = "" then
			  mnuErr = "<br /><br />" & txtNamNoBlank
			else
			  chkMnu(Cname)
			end if
			
			if mnuErr = "" then
			  if Clink <> "" then
			    chkMnu(Clink)
			  end if
			end if
			
			if mnuErr = "" then
			  if Cimage <> "" then
			    chkMnu(Cimage)
			  end if
			end if
			
			if mnuErr = "" then
			  if CgrpAccess = "0" then CgrpAccess = ""
			  if len(CgrpAccess) > 0 then
			    CgrpAccess = chkGrpAdmin(CgrpAccess)
			  end if
			end if
			
			if mnuErr = "" then
						
			strSQL = "INSERT into Menu (" _
							&"Name, "_
							&"Parent, "_
							&"ParentID, "_
							&"Link, "_
							&"mnuImage, "_
							&"onclick, "_
							&"Target, "_
							&"mnuTitle, "_
							&"INAME, "_
							&"mnuFunction, "_
							&"mnuAccess, "_
							&"mnuOrder) "_
							&"VALUES ('" & Cname & "'"_
							&",'" & Cparent & "'"_
							&",'" & CparentID & "'"_
							&",'" & Clink & "'"_
							&",'" & Cimage & "'"_
							&",'" & Conclick & "'"_
							&",'" & Ctarget & "'"_
							&",'" & Ctitle & "'"_
							&",'" & CNINAME & "'"_
							&",'" & Cfunct & "'"_
							&",'" & CgrpAccess & "'"_
							&",'" & Corder & "')"
'							response.write("<br />" & strSQL & "<br />")
'							response.End()
			my_Conn.Execute (strSql)
			
			CINAME = CMenu
			end if
		
	case "Cdelete"
		strSQL = "delete from Menu where id = " & Request.Form("Cid")				
		executeThis(strSql)
		
	case "Pdelete"
	 	sSql = "SELECT * FROM MENU WHERE ID = " & Request.Form("Pid")
		set rsT = my_Conn.execute(sSql)
		if not rsT.eof then
		  if rsT("mnuAdd") <> "" then
			set rsT = nothing
			strSQL = "delete from Menu where id = " & Request.Form("Pid")				
			executeThis(strSql)
		  else
			set rsT = nothing
			strSQL = "delete from Menu where id = " & Request.Form("Pid")				
			executeThis(strSql)
		
			strSQL = "delete from Menu where Parent = '" & Request.Form("Cid") & "' AND INAME='" & CINAME & "'"				
			executeThis(strSql)
		  end if
		end if
		
	case "Mdelete"
		strSQL = "delete from Menu where INAME = '" & CINAME & "' or mnuAdd='" & CINAME & "'"				
		executeThis(strSql)
		CINAME = ""
			
	case "editparent"
    		Pid = Trim(Request.Form("id"))
    		Pname = replace(Trim(Request.Form("Pname")),"'","''")
    		PNname = replace(Trim(Request.Form("PNname")),"'","''")
    		Plink = Trim(Request.Form("Plink"))
    		Pimage = Trim(Request.Form("PImage"))
    		Ponclick = replace(Trim(Request.Form("Ponclick")),"'","''")
    		Ptarget = Trim(Request.Form("Ptarget"))
    		Porder = Trim(Request.Form("Porder"))
    		Ctitle = Trim(Request.Form("mnuTitle"))
    		Cfunct = Trim(Request.Form("Cfunct"))
    		CaddMenu = Trim(Request.Form("CaddMenu"))
    		CgrpAccess = Trim(Request.Form("g_read"))
			
			if PNname = "" then
			  PNname = Pname
			else
			  chkMnu(Pname)
			end if
			
			if mnuErr = "" then
			  if Plink <> "" then
			    chkMnu(Plink)
			  end if
			end if
			
			if mnuErr = "" then
			  if Pimage <> "" then
			    chkMnu(Pimage)
			  end if
			end if
			
			if mnuErr = "" then
			  if CgrpAccess = "0" then CgrpAccess = ""
			  if len(CgrpAccess) > 0 then
			    CgrpAccess = chkGrpAdmin(CgrpAccess)
			  end if
			end if
			
		if mnuErr = "" then
			strSQL = "update Menu set Name = '" & PNname & "', "
			strSQL = strSQL & "Link = '" & Plink & "', "
			strSQL = strSQL & "mnuImage = '" & Pimage & "', "
			strSQL = strSQL & "onclick = '" & Ponclick & "', "
			strSQL = strSQL & "Target = '" & Ptarget & "', "
			'strSQL = strSQL & "mnuOrder = " & Porder & ", "
			strSQL = strSQL & "mnuAccess = '" & CgrpAccess & "', "
			strSQL = strSQL & "mnuFunction = '" & Cfunct & "' "
			strSQL = strSQL & "where id = " & Pid
'			response.Write(strSQL & "<br />")				
			executeThis(strSql)
		  if CaddMenu <> "" then
			  sSql = "SELECT * FROM MENU WHERE PARENT='" & CaddMenu & "' AND NAME='" & Pname & "' AND INAME='" & CaddMenu & "'"
			  set rsAM = my_Conn.execute(sSql)
			  if rsAM.eof then
			    'mnuErr = mnuErr & "<br /><br />Menu was not found"
			  else
			    Pid = rsAM("ID")
			  end if
			  set rsAM = nothing
			strSQL = "update Menu set Name = '" & PNname & "', "
			strSQL = strSQL & "Link = '" & Plink & "', "
			strSQL = strSQL & "mnuImage = '" & Pimage & "', "
			strSQL = strSQL & "onclick = '" & Ponclick & "', "
			strSQL = strSQL & "Target = '" & Ptarget & "', "
			'strSQL = strSQL & "mnuOrder = " & Porder & ", "
			strSQL = strSQL & "mnuAccess = '" & CgrpAccess & "', "
			strSQL = strSQL & "mnuFunction = '" & Cfunct & "' "
			strSQL = strSQL & "where id = " & Pid
'			response.Write(strSQL & "<br />")				
			executeThis(strSql)
			
		  else
		  end if
			strSQL = "update Menu set Parent = '" & PNname & "' where Parent = '" & Pname & "' and mnuTitle = '" & Ctitle & "'"
			executeThis(strSql)
		end if
		
	case "addparent"
			Cparent = Trim(Request.Form("Cparent"))
    		Cname = replace(Trim(Request.Form("Cname")),"'","''")
    		Clink = Trim(Request.Form("Clink"))
    		Cimage = Trim(Request.Form("CImage"))
    		Conclick = replace(Trim(Request.Form("Conclick")),"'","''")
    		Ctarget = Trim(Request.Form("Ctarget"))
    		Corder = Trim(Request.Form("Corder"))
    		Ctitle = Trim(Request.Form("mnuTitle"))
    		'CINAME = Trim(Request.Form("CINAME"))
    		Cfunct = Trim(Request.Form("Cfunct"))
    		CaddMenu = Trim(Request.Form("AddMenu"))
    		CappID = Trim(Request.Form("app_id"))
    		CgrpAccess = Trim(Request.Form("g_read"))
			
			if CgrpAccess = "0" then CgrpAccess = ""
			if len(CgrpAccess) > 0 then
			  CgrpAccess = chkGrpAdmin(CgrpAccess)
			end if
			
			if CappID = "" then CappID = 0
			
		  if CaddMenu <> "" then
			  sSql = "SELECT * FROM MENU WHERE PARENT='" & CaddMenu & "' AND INAME='" & CaddMenu & "'"
			  set rsAM = my_Conn.execute(sSql)
			  if rsAM.eof then
			    'mnuErr = mnuErr & "<br /><br />Menu was not found"
			  else
			    'do until rsAM.eof
    			  'Cname = rsAM("Name")
      			  'Clink = rsAM("Link")
    			  'Cimage = rsAM("mnuImage")
    			  'Conclick = rsAM("onclick")
    			  'Ctarget = rsAM("Target")
				  'CappID = rsAM("app_id")
     			  'Cfunct = rsAM("mnuFunction")
    		      'CaddMenu = rsAM("mnuAdd")
    			  Cname = rsAM("mnuTitle") & "&nbsp;" & txtMenu
      			  Clink = ""
    			  Cimage = ""
    			  Conclick = ""
    			  Ctarget = ""
				  CappID = rsAM("app_id")
     			  Cfunct = ""
    		      'CaddMenu = rsAM("INAME")
				  
				  if rsAM("mnuAdd") <> "" then
				   'CaddMenu = rsAM("mnuAdd")
				  else
				   'CaddMenu = rsAM("INAME")
				  end if
				  
				  if CappID = "" then CappID = 0			
				  strSQL = "INSERT into Menu (" _
							&"Name, "_
							&"Parent, "_
							&"Link, "_
							&"mnuImage, "_
							&"onclick, "_
							&"Target, "_
							&"mnuTitle, "_
							&"INAME, "_
							&"mnuFunction, "_
							&"mnuAccess, "_
							&"mnuAdd, "_
							&"app_id, "_
							&"mnuOrder) "_
							&"VALUES ('" & Cname & "'"_
							&",'" & Cparent & "'"_
							&",'" & Clink & "'"_
							&",'" & Cimage & "'"_
							&",'" & Conclick & "'"_
							&",'" & Ctarget & "'"_
							&",'" & Ctitle & "'"_
							&",'" & CINAME & "'"_
							&",'" & Cfunct & "'"_
							&",'" & CgrpAccess & "'"_
							&",'" & CaddMenu & "'"_
							&"," & CappID & ""_
							&"," & Corder & ")"
				  executeThis(strSql)
				  
				  'rsAM.movenext
				'loop
			  end if
			  set rsAM = nothing
		  else
			
			if mnuErr = "" then
			  if Cname = "" then
			    mnuErr = mnuErr & "<br /><br />" & txtNamNoBlank
			  else
			    chkMnu(Cname)
			  end if
			end if
			
			if mnuErr = "" then
			  if Clink <> "" then
			    chkMnu(Clink)
			  end if
			end if
			
			if mnuErr = "" then
			  if Cimage <> "" then
			    chkMnu(Cimage)
			  end if
			end if
			
			if mnuErr = "" then			
			strSQL = "INSERT into Menu (" _
							&"Name, "_
							&"Parent, "_
							&"Link, "_
							&"mnuImage, "_
							&"onclick, "_
							&"Target, "_
							&"mnuTitle, "_
							&"INAME, "_
							&"mnuFunction, "_
							&"mnuAccess, "_
							&"mnuAdd, "_
							&"app_id, "_
							&"mnuOrder) "_
							&"VALUES ('" & Cname & "'"_
							&",'" & Cparent & "'"_
							&",'" & Clink & "'"_
							&",'" & Cimage & "'"_
							&",'" & Conclick & "'"_
							&",'" & Ctarget & "'"_
							&",'" & Ctitle & "'"_
							&",'" & CINAME & "'"_
							&",'" & Cfunct & "'"_
							&",'" & CgrpAccess & "'"_
							&",'" & CaddMenu & "'"_
							&"," & CappID & ""_
							&"," & Corder & ")"
			executeThis(strSql)
			end if
		  end if
			
		
	case "addmenu"
    		Cname = replace(Trim(Request.Form("Cname")),"'","''")
    		Clink = Trim(Request.Form("Clink"))
    		Cimage = Trim(Request.Form("CImage"))
    		Conclick = replace(Trim(Request.Form("Conclick")),"'","''")
    		Ctarget = Trim(Request.Form("Ctarget"))
    		Ctitle = "** " & Trim(replace(Request.Form("mnuTitle"),"'","''"))
    		Caddmenu = Trim(Request.Form("AddMenu"))
    		CINAME = "m_" & lcase(replace(replace(replace(Trim(Request.Form("mnuTitle")),"'","")," ",""),"&nbsp:",""))
			
			if Caddmenu <> "" then
			  CNname = Caddmenu
			else
			  CNname = CINAME
			end if
			
			if Ctitle = "" then
			  mnuErr = "<br /><br />" & txtMnuHvTitle
			else
			  chkMnu(Ctitle)
			end if
			
			if mnuErr = "" then
			  if Cname = "" then
			    mnuErr = mnuErr & "<br /><br />" & txtNamNoBlank
			  else
			    chkMnu(Cname)
			  end if
			end if
			
			if mnuErr = "" then
			  if Clink <> "" then
			    chkMnu(Clink)
			  end if
			end if
			
			if mnuErr = "" then
			  if Cimage <> "" then
			    chkMnu(Cimage)
			  end if
			end if
			
			if mnuErr = "" then
			  if Conclick <> "" then
			    chkMnu(Conclick)
			  end if
			end if
			
			if mnuErr = "" then			
			strSQL = "INSERT into Menu (" _
							&"Name, "_
							&"Parent, "_
							&"INAME, "_
							&"Link, "_
							&"mnuImage, "_
							&"onclick, "_
							&"Target, "_
							&"mnuTitle, "_
							&"mnuFunction, "_
							&"mnuOrder) "_
							&"VALUES ('" & Cname & "'"_
							&",'" & CNname & "'"_
							&",'" & CNname & "'"_
							&",'" & Clink & "'"_
							&",'" & Cimage & "'"_
							&",'" & Conclick & "'"_
							&",'" & Ctarget & "'"_
							&",'" & Ctitle & "'"_
							&",'" & Cfunct & "'"_
							&",'1')"
			executeThis(strSql)
			end if
		
	case "updateOrder"
    	Ucount = Request("count")
    	Utitle = Request("mnuTitle")
  		For ux = 1 to Ucount
    		Uid = Request("id" & ux)
    		Uorder = Request("mnuOrder" & ux)
			strSQL = "update Menu set mnuOrder = '" & Uorder & "' where id = " & Uid & " and mnuTitle = '" & Utitle & "'"
'			response.Write(Ctitle & "<br />")
'			response.Write(strSQL & "<br />")				
			executeThis(strSql)
  		next
'		response.End()
						
  end select
  
  if mnuErr <> "" then
	mnuErr = mnuErr & "<br /><br />" & txtMnuChgsAbort
	response.Write("<center>" & mnuErr & "</center>")
	response.Write("<meta http-equiv=""Refresh"" content=""2; URL=admin_menu.asp"">")
	response.End()
  else
	if bFso then
	  mnu.DelMenuFiles(CINAME)
	end if
	'response.end()
  end if
  response.Redirect("admin_menu.asp?menu=" & CINAME)
else

  if request("mode") = "resetmenu" then
	if bFso then
	  mnu.DelMenuFiles("")
	  resetCoreConfig() 
	  Call setSession("sMsg","Menu files rebuilt")
	  closeAndGo("admin_menu.asp")
	end if
  end if
  
  if request("menu") <> "" then
	Menu = request("menu")
  else
    Menu = "def_main"
  end if

  i = 0
  ed = 0
end if

Function chkMnu(str)
	uIP = request.ServerVariables("REMOTE_ADDR")
	muEr = "<br /><br /><span class=""fAlert""><b>" & txtMalScrDet & "</b></span>"
	muEr = muEr & "<span class=""fAlert"">" & uIP & "</span>"
	muEr = muEr & "<br /><b>" & txtIPLogged & "</b>"
	if inStr(str,"<") > 0 or inStr(str,">") > 0 then 
	  mnuErr = mnuErr & muEr
	  exit function
	end if
End Function
%>
