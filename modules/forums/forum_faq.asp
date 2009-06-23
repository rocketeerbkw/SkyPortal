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
pagtype = chkString(Request.QueryString("page"),"sqlstring")
select case pagtype
case "home"
CurPageType="home"
case "forums"
CurPageType="forums"
case else
CurPageType="home"
end select
%>
<!--#INCLUDE file="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<br />
<%
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = "Site FAQs"
	PageAction = "Viewing<br />" 
	PageLocation = "forum_faq.asp"
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function

select case pagtype

case "forums"

	intSkin = getSkin(intSubSkin,2)
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tCellAlt1"">")
%>
  <tr class="tSubTitle">
    <td>Forums FAQ Table of Contents<% if (hasAccess(1) or mlev = 3) then%> <a href="forum_faq_moderator.asp">Moderators Click here</a><% end if %></td>
  </tr>
  <tr>
    <td>
    <p>
    <ul class="fNorm">
    <li><a href="#register">Do I have to register?</a></li>
<% if (strIcons = "1") then %>
    <li><a href="#smilies">How can I use smilies and images?</a></li>
<% end if %>
    <li><a href="#hyperlink">Can I add a hyperlink to my messages?</a></li>
    <li><a href="#format">Can I change the format of my text?</a></li>
    <li><a href="#mods">What are Moderators?</a></li>
    <li><a href="#profile">How can I change my registration profile?</a></li>
    <li><a href="#cookies">Are cookies used?</a></li>
    <li><a href="#activetopics">What are active topics?</a></li>
    <li><a href="#avatar">Can I upload my own avatars?</a></li>
    <li><a href="#edit">Can I edit my own posts?</a></li>
    <li><a href="#attach">Can I attach files?</a></li>
    <li><a href="#search">Can I search?</a></li>
    <li><a href="#EditProfile">Can I edit my profile?</a></li>
    <li><a href="#signature">Can I attach my own signature to my posts?</a></li>
<% if strBadWordFilter = 1 then %>
    <li><a href="#censor">Are there any censor features?</a></li>
<% end if %>
<% if strEmail = 1 then %>
    <li><a href="#pw">What do I do if I forget my UserName and/or Password?</a></li>
    <li><a href="#notify">Can I be notified by email if someone responds to my topic?</a></li>
<% end if %>
    <li><a href="#COPPA">What is COPPA?</a></li>
    <li><a href="#getforum">Where can I get my own copy of this Forum?</a></li>
    </ul>
    </p>
    </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="register"></a>Registering</td>
  </tr>
  <tr>
      <td>
      <p>
      Registration is not required to view current topics in the Forums; however, 
      if you wish to post a new topic or reply to an existing topic, registration is 
      required.&nbsp; Registration is free and only takes a few moments.&nbsp; The only 
      required fields are your UserName, which may be your real name or a nickname, a password, and a 
      valid e-mail address.&nbsp; The information you provide during registration is not 
      outsourced or used for any advertising by <% =strSiteTitle %>.&nbsp; If you believe someone 
      is sending you advertisements as a result of the information you provided through 
      your registration, please notify us immediately.</p>
      
      </td>
  </tr>
<% If (strIcons = "1") then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="smilies"></a>Smilies</td>
  </tr>
  <tr>
      <td>
      <p>
      You've probably seen others use smilies before in email messages or other Forum posts. Smilies are keyboard characters used to convey an emotion, such as a smile 
      <img border="0" src="<%= strHomeUrl %>images/Smilies/smile.gif"> or a frown 
      <img border="0" src="<%= strHomeUrl %>images/Smilies/sad.gif">. This Forum 
      automatically converts certain text to a graphical representation when it is 
      inserted between brackets [].&nbsp; Here are the smilies that are currently 
      supported by <% =strSiteTitle %>:<br />
      <table border="0" align="center" cellpadding="5">
        <tr valign="top">
          <td>
          <table border="0" align="center">
              <tr>
                      <td><img border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/smile.gif"></td>
                      <td>smile</td>
                      <td>[:)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/big.gif"></td>
                      <td>big smile</td>
                      <td>[:D]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/cool.gif"></td>
                      <td>cool</td>
                      <td>[8D]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/blush.gif"></td>
                      <td>blush</td>
                      <td>[:I]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/tongue.gif"></td>
                      <td>tongue</td>
                      <td>[:P]</td>
             </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/evil.gif"></td>
                      <td>evil</td>
                      <td>[}:)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/wink.gif"></td>
                      <td>wink</td>
                      <td>[;)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/clown.gif"></td>
                      <td>clown</td>
                      <td>[:o)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/blackeye.gif"></td>
                      <td>black eye</td>
                      <td>[B)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/8ball.gif"></td>
                      <td>eightball</td>
                      <td>[8]</td>
              </tr>
      </table>
      </td>
      <td>
      <table border="0" align="center">
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/sad.gif"></td>
                      <td>frown</td>
                      <td>[:(]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/shy.gif"></td>
                      <td>shy</td>
                      <td>[8)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/shock.gif"></td>
                      <td>shocked</td>
                      <td>[:O]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/angry.gif"></td>
                      <td>angry</td>
                      <td>[:(!]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/dead.gif"></td>
                      <td>dead</td>
                      <td>[xx(]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/sleepy.gif"></td>
                      <td>sleepy</td>
                      <td>[|)]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/kisses.gif"></td>
                      <td>kisses</td>
                      <td>[:X]</td>
              </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/approve.gif"></td>
                      <td>approve</td>
                      <td>[^]</td>
             </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/dissapprove.gif"></td>
                      <td>disapprove</td>
                      <td>[V]</td>
             </tr>
              <tr>
                      <td><img alt border="0" hspace="10" src="<%= strHomeUrl %>images/Smilies/question.gif"></td>
                      <td>question</td>
                      <td>[?]</td>
              </tr>
      </table>
          </td>
        </tr>
      </table>
      </p>
      </td>
  </tr>
<% end if %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="hyperlink"></a>Creating a Hyperlink in your message</td>
  </tr>
  <tr>
      <td>
      <p>You can easily add a hyperlink to your message.</p>

      <p>All that you need to do is type the URL (<% =strHomeURL %>), and it will automatically be converted to a URL (<a href="<% =strHomeURL %>" target="_blank"><% =strHomeURL %></a>)!</p>
      
      <p>The trick here is to make sure you prefix your URL with the <b>http://</b>, <b>https://</b> or <b>file://</b></p>

      <p>You can also add a mailto link to your message by typing in your email address.<br />
      <blockquote class="fNorm">
			<i>This Example:</i><br />
			<b><%= replace(strSender,"@","[no-spam]@") %></b><br />
			<i>Outputs this:</i><br />
			<a href="mailto:<%= replace(strSender,"@","[no-spam]@") %>"><%= replace(strSender,"@","[no-spam]@") %></a>.</p>
      </blockquote>
      
      <p>Another way to add hyperlinks is to use the <b>[url]</b>linkto<b>[/url]</b> tags</p>
	  <blockquote class="fNorm">
              <i>This Example:</i><br />
              <b>[url]</b><% =strHomeURL %><b>[/url]</b> takes you home!<br />
              <i>Outputs This:</i><br />
              <a href="<% =strHomeURL %>"><% =strHomeURL %></a> takes you home!
      </blockquote></p>
	  <p> 
      <p>If you use this tag: <b>[url=&quot;</b>linkto<b>&quot;]</b>description[/url]</b> you can add a description to the link.</p>
      <blockquote class="fNorm">
              <i>This Example:</i><br />
              Take me to <b>[url=&quot;<% =strHomeURL %>&quot;]</b><% =strSiteTitle %><b>[/url]</b><br />
              <i>Outputs This:</i><br />
              Take me to <a href="<% =strHomeURL %>"><% =strSiteTitle %></a>
      </blockquote>
      <blockquote class="fNorm">
              <i>This Example:</i><br />
              If you have a question <b>[url=&quot;<% =strSender %>&quot;]</b>Mail Me<b>[/url]</b><br />
              <i>Outputs This:</i><br />
              If you have a question <a href="mailto:<% =strSender %>">Mail Me</a>
      </blockquote>
      <p>You can make clickable images by combining the <b>[url="</b>linkto<b>"]</b>desc<b>[/url]</b> and <b>[img]</b>image_url<b>[/img]</b> tags</p><br />
      <blockquote class="fNorm">
              <i>This Example:</i><br />
              <b>[url=&quot;<% =strHomeURL %>&quot;][img]</b><% =strHomeURL %>images/site_logo.jpg <b>[/img][/url]</b><br />
              <i>Outputs This:</i><br />
              <a href="<% =strHomeURL %>"><img src="<% if right(strHomeURL,1) <> "/" then Response.Write(strHomeURL & "/") else Response.Write(strHomeURL) end if%>themes/<%= strTheme %>/site_logo.jpg" target="_new" border="0"></a>
      </blockquote>
      </p>
      </td>
  </tr>
<% if strAllowForumCode = "1" then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="format"></a>How to format text with Bold, Italic, Quote, etc...</td>
  </tr>
  <tr>
    <td>
    <p>
    There are several Forum Codes you may use to change the appearance 
    of your text.&nbsp; Following is the list of codes currently available:</p>
    <blockquote>
      <p><b>Bold:</b> enclose your text with [b] and [/b] .&nbsp; <i>Example:</i> This is <b>[b]</b>bold<b>[/b]</b> text. = This is <b>bold</b> text.</p>

      <p><i>Italic:</i> enclose your text with [i] and [/i] .&nbsp; <i>Example:</i> This is <b>[i]</b>italic<b>[/i]</b> text. = This is <i>italic</i> text.</p>

      <p><u>Underline:</u> enclose your text with [u] and [/u]. <i>Example:</i> This is <b>[u]</b>underline<b>[/u]</b> text. =  This is <u>underline</u> text.</p>

      <p>Aligning Text Left:<br />
        Enclose your text with [left] and [/left]
      </p>

      <p>Aligning Text Center:<br />
        Enclose your text with [center] and [/center]
      </p>

      <p>Aligning Text Right:<br />
        Enclose your text with [right] and [/right]
      </p>

      <p>Pre Text:<br />
        Enclose your text with [pre] and [/pre]
      </p>

      <p>Striking Text:<br />
        Enclose your text with [s] and [/s]. <i>Example:</i> <b>[s]</b>mistake<b>[/s]</b> = <s>mistake</s>
      </p>

      <p>Marquee:<br />
        Enclose your text with [marquee] and [/marquee]. <br /><i>Example:</i> <b>[marquee]</b>moving text<b>[/marquee]</b> = <br /><marquee>moving text</marquee>
      </p>
      
      <p>Superscript Text:<br />
        Enclose your text with [sup] and [/sup]. <i>Example:</i> <b>[sup]</b>Superscript<b>[/sup]</b> = <sup>Superscript</sup>
      </p>
      
      <p>Subscript Text:<br />
        Enclose your text with [sub] and [/sub]. <i>Example:</i> <b>[sub]</b>Subscript<b>[/sub]</b> = <sub>Subscript</sub>
      </p>
      
      <p>Teletype Text:<br />
        Enclose your text with [tt] and [/tt]. <i>Example:</i> <b>[tt]</b>Teletype<b>[/tt]</b> = <tt>Teletype</tt>
      </p>
      
      <p>Highlight Text:<br />
        Enclose your text with [hl] and [/hl]. <i>Example:</i> <b>[hl]</b>Highlight<b>[/hl]</b> = <span style="background-color: #FFFF00">Highlight</span>
      </p>
      
      <p>Horizontal Line:<br />
        Insert [hr] to where you want to add the horizontal line. <i>Example:</i> <b>[hr]</b> = <hr />
      </p>
      
      <p>&nbsp; </p>

      <p><b>Font Colors:</b><br />
        Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>] <br />
        <i>Example:</i> <b>[red]</b>Text<b>[/red]</b> = <font color="red">Text</font id=red><br />
        <i>Example:</i> <b>[blue]</b>Text<b>[/blue]</b> = <font color="blue">Text</font id=blue><br />
        <i>Example:</i> <b>[pink]</b>Text<b>[/pink]</b> = <font color="pink">Text</font id=pink><br />
        <i>Example:</i> <b>[brown]</b>Text<b>[/brown]</b> = <font color="brown">Text</font id=brown><br />
        <i>Example:</i> <b>[black]</b>Text<b>[/black]</b> = <font color="black">Text</font id=black><br />
        <i>Example:</i> <b>[orange]</b>Text<b>[/orange]</b> = <font color="orange">Text</font id=orange><br />
        <i>Example:</i> <b>[violet]</b>Text<b>[/violet]</b> = <font color="violet">Text</font id=violet><br />
        <i>Example:</i> <b>[yellow]</b>Text<b>[/yellow]</b> = <font color="yellow">Text</font id=yellow><br />
        <i>Example:</i> <b>[green]</b>Text<b>[/green]</b> = <font color="green">Text</font id=green><br />
        <i>Example:</i> <b>[gold]</b>Text<b>[/gold]</b> = <font color="gold">Text</font id=gold><br />
        <i>Example:</i> <b>[white]</b>Text<b>[/white]</b> = <font color="white">Text</font id=white><br />
        <i>Example:</i> <b>[purple]</b>Text<b>[/purple]</b> = <font color="purple">Text</font id=purple>
      </p>

      <p>&nbsp; </p>

      <span class="fSubTitle"><b>Headings:</b></span><br />
        <p>Enclose your text with [h<i>number</i>] and [/h<i>n</i>]<br />
        <table border=0>
          <tr>
            <td>
            <i>Example:</i> <b>[h1]</b>Text<b>[/h1]</b> =
            </td>
            <td>
            <h1>Text</h1>
            </td>
          </tr>
          <tr>
            <td>
            <i>Example:</i> <b>[h2]</b>Text<b>[/h2]</b> =
            </td>
            <td>
            <h2>Text</h2>
            </td>
          <tr>
            <td>
            <i>Example:</i> <b>[h3]</b>Text<b>[/h3]</b> =
            </td>
            <td>
            <h3>Text</h3>
            </td>
          </tr>
          <tr>
            <td>
            <i>Example:</i> <b>[h4]</b>Text<b>[/h4]</b> =
            </td>
            <td>
            <h4>Text</h4>
            </td>
          </tr>
          <tr>
            <td>
            <i>Example:</i> <b>[h5]</b>Text<b>[/h5]</b> =
            </td>
            <td>
            <h5>Text</h5>
            </td>
          </tr>
          <tr>
            <td>
            <i>Example:</i> <b>[h6]</b>Text<b>[/h6]</b> =
            </td>
            <td>
            <h6>Text</h6>
            </td>
          </tr>
        </table>
      </p>

      <p>&nbsp; </p>

      <span class="fSubTitle"><b>Font Sizes:</b></span><br />
        <p><i>Example:</i> <b>[size=1]</b>text<b>[/size=1]</b> = <font size=1>Text</font id=size1><br />
        <i>Example:</i> <b>[size=2]</b>text<b>[/size=2]</b> = <font size=2>Text</font id=size2><br />
        <i>Example:</i> <b>[size=3]</b>text<b>[/size=3]</b> = <font size=3>Text</font id=size3><br />
        <i>Example:</i> <b>[size=4]</b>text<b>[/size=4]</b> = <font size=4>Text</font id=size4><br />
        <i>Example:</i> <b>[size=5]</b>text<b>[/size=5]</b> = <font size=5>Text</font id=size5><br />
        <i>Example:</i> <b>[size=6]</b>text<b>[/size=6]</b> = <font size=6>Text</font id=size6>
      </p>

      <p>&nbsp; </p>

      <p>Bulleted List: <b>[list]</b> and <b>[/list]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Ordered Alpha List: <b>[list=a]</b> and <b>[/list=a]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Ordered Number List: <b>[list=1]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Code: enclose your text with <b>[code]</b> and <b>[/code]</b>.</p>

      <p>Quote: enclose your text with <b>[quote]</b> and <b>[/quote]</b>.</p>
      <p>Images: enclose the address with <b>[img]</b> and <b>[/img]</b>. You can make clickable images by combining the [url=""][img][/img][/url]</p>
    </blockquote></td>
  </tr>
    
  
<% end if %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="mods"></a>Moderators</td>
  </tr>
  <tr>
      <td>
      <p>
      Moderators control individual forums. 
      They may edit, delete, or prune any posts in their forums. 
<%	if (strShowModerators = "1") then %>
      If you have a question about a particular forum, you should direct it to your forum moderator.
<%	end if %>
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="profile"></a>Changing Your Profile</td>
  </tr>
  <tr>
      <td>
      <p>
      You may easily change any 
      information stored in your registration profile by using the &quot;Control Panel&quot; link located near 
      the top of each page. Simply identify yourself by typing your UserName and 
      Password and all of your profile information will appear on screen. You may 
      edit any information (except your UserName).
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="cookies"></a>Cookies</td>
  </tr>
  <tr>
      <td>
      <p>
      These Forums use cookies to store the following information: the last time you logged in, your UserName and 
      your Password, if you set it in preferences. These cookies are stored on your hard drive. Cookies are not used 
      to track your movement or perform any function other than to enhance your use of these forums. 
<% if (strNoCookies = "0") then %>
      If you have not enabled cookies in your browser, many of these time-saving features will not work properly. <b>Also, you 
      need to have cookies enabled if you want to enter a private forum or post a topic/reply.</b>
<% end if %>
      </p>
      <p>You may delete all cookies set by these forums in selecting the &quot;logout&quot; button at the top of any page.
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="activetopics"></a>Active Topics</td>
  </tr>
  <tr>
      <td>
      <p>Active Topics are tracked by cookies. When you click on the &quot;active topics&quot; a page is generated listing all topics that have been posted since your last visit to these forums (or approximately 20 minutes).</p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="Avatar"></a>Uploading Avatars</td>
  </tr>
  <tr>
      <td>
      <p>
      You may upload 1 personal avatar to our server. However, when you upload another Avatar, it overwrites your old Avatar. Please use good taste when uploading Avatars. Admins and Moderators can and will delete Avatars that are deemed unfit.
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="Edit"></a>Editing Your Posts</td>
  </tr>
  <tr>
      <td>
      <p>
      You may edit or delete your own posts at any time. Just go to the topic where the 
      post to be edited or deleted is located 
      and you will see a edit or delete icon (<img border="0" src="<%= strHomeUrl %>images/icons/icon_edit_topic.gif" hspace="6"><img border="0" src="<%= strHomeUrl %>images/icons/icon_delete_reply.gif" hspace="6">) 
      on the line that begins &quot;posted on...&quot; Click on this icon to edit or 
      delete the post. No one else can edit your post, except for the forum Moderator 
      or the site administrator. 
<% if (strEditedByDate = "1") then %>
      A note is generated at the bottom of each edited post displaying when and by whom the post was edited.
<% end if %>
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="Attach"></a>Attaching Files</td>
  </tr>
  <tr>
      <td>
      <p>
      For security reasons, you may 
      not attach files to any posts. However, you may cut and paste text into your post.
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="Search"></a>Searching For Specific Posts</td>
  </tr>
  <tr>
      <td>
      <p>
      You may search for 
      specific posts based on a word or words found in the posts, user name, date, and 
      particular forum(s). Simply click on the &quot;search&quot; link at the top of most pages. 
<!--      Note: announcements are not included in the search returns. -->
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="EditProfile"></a>Editing Your Profile</td>
  </tr>
  <tr>
      <td>
      <P>You may easily change any information stored in your registration profile by using the "Control Panel" link located near the top of each page. Simply identify yourself by typing your UserName and Password and all of your profile information will appear on screen. You may edit any information (except your UserName).</P>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="Signature"></a>Signatures</td>
  </tr>
  <tr>
      <td>
      <p>You may attach signatures to the end of your posts when you post either a New Topic or Reply. Your signature is editable by clicking on &quot;profile&quot; at the top of any forum page and entering your UserName and Password.</p>
      <p>NOTE: HTML can't be used in Signatures.</p>
      </td>
  </tr>
<% if (strBadWordFilter = "1") then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="censor"></a>Censoring Posts</td>
  </tr>
  <tr>
      <td>
      <p>
      The Forum does censor certain words that may be posted; however, this 
      censoring is not an exact science, and is being done based on the words that are being 
      screened, so certain words may be censored out of context. Words that are censored are replaced with asterisks.
      </p>
      </td>
  </tr>
<% end if %>
<% if (strEmail = "1") then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="pw"></a>Lost User Name and/or Password</td>
  </tr>
  <tr>
      <td>
      <p>
      Retrieving your 
      UserName and Password is simple, assuming that email features are turned on for 
      this portal. All of the pages that require you to identify yourself with 
      your UserName and Password carry a &quot;lost Password&quot; link that you can use to have 
      your UserName and Password mailed instantly to your email address of 
      record.
      </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="notify"></a>Email Notification</td>
  </tr>
  <tr>
      <td>
      <p>
      When you create a new topic, 
      you have the option of receiving an email notification every time someone posts 
      a reply to your topic. If you wish to use this feature, simply check the email notification box on the &quot;New Topic&quot; 
      page when you create your new topic.
      </p>
      </td>
  </tr>
<% end if %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="COPPA"></a>What is COPPA</td>
  </tr>
  <tr>
      <td>
      <p>The Children's Online Privacy Protection Act and Rule apply to individually identifiable
      information about a child that is collected online, such as full name, home address, email address,
      telephone number or any other information that would allow someone to identify or contact the
      child. The Act and Rule also cover other types of information -- for example, hobbies, interests
      and information collected through cookies or other types of tracking mechanisms -- when they
      are tied to individually identifiable information. More information can be found 
      <a href="http://www.ftc.gov/bcp/conline/pubs/buspubs/coppa.htm" title="What is COPPA?">here</a>.</p>
</td></tr>
<tr><td align="center" class="tSubTitle"><a href="forum_faq.asp?page=forums">Back to FAQ Table of Contents</a></td></tr>

<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
case "links"
spThemeTableCustomCode = "align=""center"" width=""99%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
  <tr><td class="tSubTitle"><span class="fSubTitle"><b>How to add a link to our website?</b></span></td></tr>
  <tr><td>

<ol><li>Go to a subcategory that matches with the link you're supplying.</li>
<li>You will see &quot;Add Link&quot; link towards the bottom, click it.</li>
<li>Fill in the provided form.</li>
<li>Wait 2-3 days to give us some time to review your supplied link.</li>
<li>The link will then be added to our database.</li></ol>
</td></tr>
<tr><td align="center" class="tSubTitle"><a href="forum_faq.asp?page=forums">Back to FAQ Table of Contents</a></td></tr>

<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
case "downloads"
spThemeTableCustomCode = "align=""center"" width=""99%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
  <tr><td class="tSubTitle"><span class="fSubTitle"><b>How to add a URL to our website?</b></span></td></tr>
  <tr><td>

<ol><li>Go to a subcategory that matches with the File you're supplying.</li>
<li>You will see &quot;Add a file&quot; link towards the bottom, click it.</li>
<li>Fill in the provided form.</li>
<li>Wait 2-3 days to give us some time to review your supplied URL.</li>
<li>The URL will then be added to our database.</li></ol>
</td></tr>
<tr><td align="center" class="tSubTitle"><a href="forum_faq.asp?page=forums">Back to FAQ Table of Contents</a></td></tr>

<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
case "articles"
spThemeTableCustomCode = "align=""center"" width=""99%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")%>
  <tr><td class="tSubTitle"><span class="fSubTitle"><b>How to add an article to our website?</b></span></td></tr>
  <tr><td>

<ol><li>Go to a subcategory that matches with the article you're supplying.</li>
<li>You will see &quot;Add an Article&quot; link towards the bottom, click it.</li>
<li>Fill in the provided form.</li>
<li>Wait 2-3 days to give us some time to review your supplied article.</li>
<li>The article will then be added to our database.</li></ol>
</td></tr>
<tr><td align="center" class="tSubTitle"><a href="forum_faq.asp?page=forums">Back to FAQ Table of Contents</a></td></tr>

<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
case "pictures"
spThemeTableCustomCode = "align=""center"" width=""99%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
  <tr><td class="tSubTitle"><span class="fSubTitle"><b>How to add a picture on our website?</b></span></td></tr>
  <tr><td>

<b>IMPORTANT</b><br /><br />
All pictures contained on pages herein were collected freely from the internet and are believed to be public domain. If you own the copyright to any images that appear on this site, please send an email to the webmaster.<br /><br />
<ol><li>Go to a subcategory that matches with the picture you're supplying.</li>
<li>You will see &quot;Add a Picture&quot; link towards the bottom, click it.</li>
<li>Fill in the provided form.</li>
<li>Wait 2-3 days to give us some time to review your picture.</li>
<li>The picture will then be added to our database.</li></ol>
</td></tr>
<tr><td align="center" class="tSubTitle"><a href="forum_faq.asp?page=forums">Back to FAQ Table of Contents</a></td></tr>

<%
case else
spThemeTableCustomCode = "align=""center"" width=""99%"""
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tPlain"">")
%>
  <tr>
    <td class="tSubTitle"><span class="fSubTitle"><b>FAQ Table of Contents</b></span></td>
  </tr>
  <tr>
    <td>
    <p>
    <ul type="square">
    <li><a href="forum_faq.asp#general">General Questions</a></li>
    <li><a href="forum_faq.asp?page=forums">Forums Questions</a></li>
<!--     <li><a href="forum_faq.asp?page=links">Links Manager Questions</a></li>
    <li><a href="forum_faq.asp?page=downloads">Downloads Manager Questions</a></li>
    <li><a href="forum_faq.asp?page=articles">Articles Manager Questions</a></li>
    <li><a href="forum_faq.asp?page=games">Games Questions</a></li> -->
    <li><a href="#getportal">Where can I get my own copy of this Portal?</a></li>
    <li><a href="mailto:<% =strSender %>">Can't find your answer here? Send us an e-mail.</a></li>
    </ul>
    </p>
    </td>
  </tr>
<tr>
    <td class="tSubTitle"><a name="general"></a><span class="fSubTitle"><b>General Questions</b></span></td>
  </tr>
  <tr>
  <td>
    <p>
    <ul>
    <li><a href="forum_faq.asp#register">Do I have to register to use this site?</a></li>
    </ul>
    </p>
</td></tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="register"></a><span class="fSubTitle"><b>Do I have to register to use this site?</b></span></td>
  </tr>
  <tr>
      <td>
      <p>
      Registration is not required to browse this site; however, 
      if you wish to customize the theme or post a new topic or reply a topic in our forums, registration is 
      required.&nbsp; Registration is FREE and only takes a few moments.&nbsp; The only 
      required fields are your UserName, which may be your real name or a nickname, a Password, and a 
      valid e-mail address.&nbsp; The information you provide during registration is not 
      outsourced or used for any advertising by <% =strSiteTitle %>.&nbsp; If you believe someone 
      is sending you advertisements as a result of the information you provided through 
      your registration, please notify us immediately.</p>
      </td>
  </tr>
<tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="GetPortal"></a><span class="fSubTitle"><b>Getting Your Own Portal</b></span></td>
  </tr>
  <tr>
      <td>
      <p>The most recent version of SkyPortal can be downloaded at <a href="http://www.SkyPortal.net/" title="Link to SkyPortal.net Website!">this website</a>.</p>
      </td>
</tr>

<%
Response.Write("</table>")
spThemeBlock1_close(intSkin)
end select%>
<!--#INCLUDE file="inc_footer.asp" -->
