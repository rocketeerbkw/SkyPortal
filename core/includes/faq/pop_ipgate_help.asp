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
'#################################################################################
'## NET IPGATE v2.0.0 RC3 Orig Idea by alex042@aol.com(c)Aug 2002, 
'## inc_ipgate.asp rewritten by www.gpctexas.net admin@gpctexas.net
'##
'## MOD re-Written by Hawk92 hawk@SkyPortal.com to be compatible with SkyPortal 1.31
'#################################################################################

Response.Write("<table class=""tPlain"">")
Response.Write		"        <tr>" & vbNewLine & _
		"          <td>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgateban""></a><b>" & txtIPFAQBanning & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQBanningDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatelck""></a><b>" & txtIPFAQLockdown & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQLockdownDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatecok""></a><b>" & txtIPFAQCookies & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQCookiesDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatelog""></a><b>" & txtIPFAQLoggingUsers & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQLoggingUsersDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatetyp""></a><b>" & txtIPFAQLoggingAll & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQLoggingAllDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgateexp""></a><b>" & txtIPFAQLogExp & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQLogExpDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatestartip""></a><b>" & txtIPFAQIPandHost & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _				
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQIPandHostDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatestartdate""></a><b>" & txtIPFAQStartDate & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _				
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQStartDateDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatepagekey""></a><b>" & txtIPFAQPageKey & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _				
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQPageKeyDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _								
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgatestatus""></a><b>" & txtIPFAQStatus & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _				
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPFAQStatusDesc & "<br /><br />" & txtIPStatusNoCSS & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _	
				"              <tr>" & vbNewLine & _
				"                <td class=""tSubTitle"" valign=""middle""><a name=""ipgateredir""></a><b>" & txtIPRedirection & "</b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _				
				"              <tr>" & vbNewLine & _
				"                <td>" & vbNewLine & _
				"                " & txtIPRedirectionDesc & vbNewLine & _
				"                <a href=""#top""><img src="""& strHomeUrl &"themes/" & strTheme & "/icons/icon_go_up.gif"" border=""0"" align=""right""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine	
				
Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine
Response.Write("</table>")
%>