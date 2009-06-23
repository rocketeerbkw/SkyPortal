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
%><!-- JUMP TO <span id="formEle"> --> 
    <form name="formJmpTo" id="formEle" action="" method="post">
    <select name="SelectMenu" size="1" onchange="jumpTo(this)" style="font-size:10px;">
      <option value="./">Jump To:</option>
      <option value="default.asp">Home</option>
      <option value="fhome.asp">Forums</option>
      <option value="forum_active_topics.asp">Active Topics</option>
      <option value="active_users.asp">Active Users</option>
      <option value="cp_main.asp">Control Panel</option>
      <option value="">&nbsp;--------------------</option>
<%
	'## Get all Forum Categories From DB
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_NAME, "
	strSql = strSql & strTablePrefix & "CATEGORY.CAT_ORDER "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY"
	strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_ORDER ASC"
	strSql = strSql &  ", " & strTablePrefix & "CATEGORY.CAT_NAME ASC;"

	set rsCat = my_conn.Execute (strSql)

	do until rsCat.eof '## Grab the Categories.

		'##  Build SQL to get forums via category
		'
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_TYPE, " & strTablePrefix & "FORUM.F_SUBJECT, "
		strSql = strSql & strTablePrefix & "FORUM.F_URL, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ORDER "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rsCat("CAT_ID")
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.FORUM_ORDER ASC"
		strSql = strSql &  ", " & strTablePrefix & "FORUM.F_SUBJECT ASC;"
		set rsForum =  my_Conn.Execute (StrSql)

		if rsForum.eof or rsForum.bof then
			'nothing
		else
			iNewCat = rsForum("CAT_ID")
			iOldCat = 0
			do until rsForum.Eof
				if chkForumAccess(strUserMemberID,rsForum("FORUM_ID")) then
					if iNewCat <> iOldCat Then
						Response.Write "      <option value='fhome.asp'>" & rsCat("CAT_NAME") & "</option>" & vbCrLf
						iOldCat = iNewCat
					end if
					if rsForum("F_TYPE") = 0 then
						Response.Write "      <option value='forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&amp;CAT_ID=" & rsForum("CAT_ID") & "&amp;Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & "'"
					else
						if rsForum("F_TYPE") = 1 then
							Response.Write "      <option value='" & rsForum("F_URL") & "'"
						end if
					end if
					if rsForum("FORUM_ID") = Request.Querystring("Forum_ID") then 
						Response.Write(" selected=""selected""") 
					end if
					Response.Write ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & rsForum("F_SUBJECT")& "</option>" & vbCrLf
				end if
				rsForum.MoveNext
			loop
		end if
		rsCat.MoveNext
	loop
%>
    </select>
    </form>
<!-- END JUMP TO -->
