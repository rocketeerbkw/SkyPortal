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
CurPageType="forums" %>
<!--#INCLUDE file="config.asp" -->
<!-- #include file="lang/en/forum_core.asp" -->
<%
CurPageInfoChk = "1"
function CurPageInfo ()
	strOnlineQueryString = ChkActUsrUrl(Request.QueryString)
	PageName = "Moderator FAQs"
	PageAction = "Viewing<br />" 
	PageLocation = "faq_moderator.asp"
	CurPageInfo = PageAction & " " & "<a href=" & PageLocation & ">" & PageName & "</a>"

end function
%>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<%
intSkin = getSkin(intSubSkin,2)
spThemeBlock1_open(intSkin)
Response.Write("<table class=""tCellAlt1"">")%>
  <tr>
    <td class="tTitle">Moderator FAQ Table of Contents</td>
  </tr>
  <tr>
    <td>
    <p>
    <ul class="fNorm">
    <li><a href="#general">Ok, I'm a Moderator, now what?</a></li>
    <li><a href="#edittopic">How do I Edit a Topic/Post?</a></li>
    <li><a href="#locktopic">What is Locking/Un-Locking a Topic and how do I do that?</a></li>
    <li><a href="#delete">How and when should I Delete a Topic/Post</a></li>
	<li><a href="#mainforum">I see some icons on the Main Forum page, what do those do?</a></li>
    
<% if (strBadWordFilter = "1") then %>
    <li><a href="#badwords">Some words aren't being censored, can I change this?</a></li>
<% end if %>
    <li><a href="mailto:<% =strSender %>">Can't find your answer here? Send us an e-mail.</a></li>
    </ul>
    </p>
    </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="general"></a>Moderator Generalities</td>
  </tr>
  <tr>
      <td>
      <p>
      A Moderator is in charge of making sure that all posters within his/her forum(s) follow
	  all the rules and guidelines that have been set forth for that/those forum(s).&nbsp;
	  Any Posts/Topics found in violation of these rules is subject to immediate edition or deletion.
	  This is accomplished by using the privileges granted to a Moderator, such as editing
	  a topics name or content, editing individual posts, locking topics, and completely
	  deleting topics.&nbsp; All of these abilities are discussed in-depth below.
	  </p>
      </td>
  </tr>
<% If (strIcons = "1") then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="edittopic"></a>Editing A Topic</td>
  </tr>
  <tr>
      <td>
      <p>
      In the main page of each forum you will see a listing off all topics for that forum, and
	  next to each topic you will see an icon that looks like this <img src="<%= strHomeUrl %>images/icons/icon_pencil.gif" alt="Edit Message" border="0" hspace="0"> and
	  within every Topic you will see this <img src="<%= strHomeUrl %>images/icons/icon_folder_pencil.gif" alt="Edit Message" border="0" hspace="0"> icon at the
	  top of the Topic.&nbsp; Both of these icons can be used to Edit the Topics information such as the Topic Name, the Forum the Topic is/should be posted in,
	  and any information in the initial post.</p>
		
	  <p>At the top of each Reply within a Topic you will see a number of icons, one of which being <img src="<%= strHomeUrl %>images/icons/icon_edit_topic.gif" alt="Edit Message" border="0" hspace="0"> 
	  which is used to edit any of the information within that reply.  
      </p>
      </td>
  </tr>
<% end if %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="locktopic"></a>Locking/Un-Locking A Topic</td>
  </tr>
  <tr>
      <td>
      <p>Locking a topic is normally done when the topic has finished discussion and further 
      discussion 
	  of the subject, in this particular thread, is unnecessary.&nbsp; A post would be made reporting the bug, discussion would go on, and upon the
	  bug being fixed, a final reply would be made mentioning the fix and closing the topic.&nbsp; Further
	  discussion would take place in another Topic if necessary.</p>
	  
	  <p>To Lock a Topic simply click the <img src="<%= strHomeUrl %>images/icons/icon_lock.gif" alt="Lock Topic" border="0" hspace="0"> icon
	  from the Topics listing, click the <img src="<%= strHomeUrl %>images/icons/icon_folder_locked.gif" alt="Lock Topic" border="0" hspace="0"> icon
	  from within the specific Topic, or when replying to the Topic click on the checkbox 
      labelled 
	  "Check here to lock the topic after this post." and submit your reply as normal.&nbsp; This will lock the
	  topic and prevent anyone else from posting.&nbsp; When choosing either of the 2 icons you will be prompted
	  with a small pop-up window confirming your decision to lock the topic.&nbsp; Press the button, refresh the 
	  browser containing the Topic and it will now be locked.</p>

	  <p>At times it will become necessary to Un-Lock a Topic after it has been locked.&nbsp; This can happen
	  for several reasons ( accidental locking, the topic/discussion is found to still be open, etc ).&nbsp; To
	  Un-Lock a Topic simply click the <img src="<%= strHomeUrl %>images/icons/icon_unlock.gif" alt="Lock Topic" border="0" hspace="0"> icon
	  from the Topics listing or click the <img src="<%= strHomeUrl %>images/icons/icon_folder_unlocked.gif" alt="Lock Topic" border="0" hspace="0"> icon
	  from within the specific Topic and you will be prompted with a small pop-up window confirming your decision to Un-Lock 
	  the Topic.  Click "Send", and refresh the browser to see the Topic is now open for posting again.
	  </p>
      </td>
  </tr>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="delete"></a>Deleting Topics</td>
  </tr>
  <tr>
    <td>
    <p>
	When it comes to the deletion of Topics/Replies, first know that this is completely FINAL and there is no
	way to undo this action.</p>
	
	<p>
	There are 2 levels of deletion.&nbsp; You can delete a single Reply using the <img src="<%= strHomeUrl %>images/icons/icon_delete_reply.gif" alt="Lock Topic" border="0" hspace="0"> icon
	or you delete a whole Topic using the <img src="<%= strHomeUrl %>images/icons/icon_trashcan.gif" alt="Lock Topic" border="0" hspace="0"> icon found on the Topics listing
	page or the <img src="<%= strHomeUrl %>images/icons/icon_folder_delete.gif" alt="Lock Topic" border="0" hspace="0"> icon found at the top of every Topic.
	Deleting a whole Topic deletes the initial post along with all Replies that it contains.</p>

	<p>In either case, upon clicking one of those icons you will be prompted with a small pop-up window confirming your decision
	to delete this Topic/Reply.&nbsp; Press the button and refresh your browser to see that the Topic/Reply is no longer there.
    </td>
  </tr>

<% if (strBadWordFilter = "1") then %>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="badwords"></a>Censoring</td>
  </tr>
  <tr>
      <td>
      <p>
      The censoring feature is only modifiable by the Admin(s) of the boards.&nbsp; Please send <a href="mailto:<% =strSender %>">the Admin</a> an e-mail
	  and express your concern over the violating words.
      </p>
      </td>
  </tr>
<%end if%>
  <tr>
    <td class="tSubTitle"><a href="#top"><img src="<%= strHomeUrl %>themes/<%= strTheme %>/icons/icon_go_up.gif" align="right" border="0"></a><a name="mainforum"></a>Main Forum Listings Options</td>
  </tr>
  <tr>
      <td>
      <p>
      For each Forum that you are in charge of, you will see 2 options next to the name. <img src="<%= strHomeUrl %>images/icons/icon_lock.gif" alt="Lock Forum" border="0" hspace="0"> and 
	  <img src="<%= strHomeUrl %>images/icons/icon_pencil.gif" alt="Edit Forum Properties" border="0" hspace="0">.</p>

	  <p>
	  The <img src="<%= strHomeUrl %>images/icons/icon_pencil.gif" alt="Edit Message" border="0" hspace="0"> is for Editing the Forums properties.
	  Such properties are Name, Description, Category, Authorization Type, Password, and Selected Members Allowed to 
	  enter that Forum.</p>

	  <p>
	  As seen previously, the <img src="<%= strHomeUrl %>images/icons/icon_lock.gif" alt="Lock Forum" border="0" hspace="0"> icon is used for locking, but
	  in this case it is used for locking an entire Forum. This prevents any members from posting in this Forum, though
	  viewing of the forum is available.	  
	  </p>
      </td>
  </tr>
<tr><td align="center" class="tSubTitle"><a href="faq.asp">Back to FAQ Table of Contents</a></td></tr>
<%Response.Write("</table>")
spThemeBlock1_close(intSkin)%>

<!--#INCLUDE file="inc_footer.asp" -->
