::::::::::::::::: Drafts Module ::::::::::::::::::::

Info
----------------------------------------------------
Author:
	Brandon Williams (Battousai)
	
Copyright:
	Some code Copyright (C) 2005-2006 Dogg Software All Rights Reserved
		Please see the SkyPortal End-User License Agreement at http://www.skyportal.net
	Everything else is 100% free, no strings attached.  If you like it, just send a thx my way, that will make my day :)
	
Compatibility:
	TinyMCE v2.1.0
		Should theoretically be OK with versions after this, but has not been tested
	SkyPortal v1
		Should theoretically be OK with versions after this, but has not been tested

Description:
	This SkyPortal Module lets you save and open drafts directly from the tinyMCE editor. After installation you will see 2 new
	options added to the left part of the middle section of buttons.  These buttons will save the current contents or open drafts
	previously saved. It also provides a user control panel where each member can see, add and delete drafts. It works like Private
	Messages, users can only see drafts they authored.

Credits
----------------------------------------------------
Idea came from here http://www.skyportal.net/link.asp?TOPIC_ID=4002
Original save plugin from moxicode but is so heavily modified that it doesn't even perform the same function

Files
----------------------------------------------------
createDrafts.asp
drafts.asp
pop_drafts.asp
README_drafts.txt
includes/inc_editor.asp
modules/drafts/drafts_functions.asp
tiny_mce/plugins/save/editor_plugin.js
tiny_mce/plugins/save/editor_plugin_src.js
tiny_mce/plugins/save/images/cancel.gif
tiny_mce/plugins/save/images/open.gif
tiny_mce/plugins/save/images/save.gif
tiny_mce/plugins/save/langs/en.js

Upgrading
----------------------------------------------------
NOTE: This install assumes you have not changed any of the files mentioned above.  It also asuumes you have not manually upgraded your version of TinyMCE to that other than what is officially released by SkyPortal. If you did any of the those things, the install cannot be a guaranteed success and may overwrite changes you have made.

1) Upload all files in the zip
2) Run createDrafts.asp
3) Enjoy :)

Installation
----------------------------------------------------
NOTE: This install assumes you have not changed any of the files mentioned above.  It also asuumes you have not manually upgraded your version of TinyMCE to that other than what is officially released by SkyPortal. If you did any of the those things, the install cannot be a guaranteed success and may overwrite changes you have made.

1) Upload all files in the zip
2) Add <!--#INCLUDE FILE="modules/drafts/drafts_functions.asp" --> to fp_custom.asp (HINT: next to all the other ones)
3) Run createDrafts.asp
4) Add delDraftsDaemon() to fp_custom.asp at end of file right before %>
5) Enjoy :)

Support
----------------------------------------------------
If you need help, or something broke, you can get support at http://www.skyportal.net
NOTE: This is NOT an official SkyPortal module.  You may not be able to get support from anyone other than the author so please be respectful to everyone who is trying to help you.

To Do
----------------------------------------------------
Got any suggestions? Drop a line at http://www.skyportal.net

CHANGELOG
----------------------------------------------------
1.00
Fixed - Hover text for open button on TinyMCE
Feature/Fixed - Support for Skyportal v1
Feature - Added Group Access support (Only people in Read can use)
0.95
Feature/Fixed - Support for RC7
0.9
Fixed - createDrafts.asp Add App SQL issue
Fixed - createDrafts.asp Add menu to nav bar issue
Fixed - config_drafts() missing
Feature - AJAXified the saving of drafts using tinyMCE
Feature - The save/open buttons on tinyMCE obey the active settings in the module manager
MISC - Spruced up the open drafts popup
MISC - Updated this readme to include more descriptive information
0.8
Feature - AJAXified the deletion of drafts from the users control panel
MISC - Renamed some files to better match the skyportal nomenclature
Fixed - Tweaked createDrafts.asp to behave nicely
MISC - Added raiseHackAttempt()
MISC - Security double check
0.5
Initial creation
Drafts are successfully saved and opened directly from TinyMCE
Created users Control Panel