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

'admin configuration help file

%>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="tCellAlt2" valign="top">
                
      <table border="0" width="100%" cellspacing="1" cellpadding="4">
        <tr> 
          <td class="tSubTitle"><a name="FORUMtitle"></a><b> What's Site Title? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Site Title is the title 
            that shows up in the upper right hand corner of the portal. It is 
            also used in email's to show where the email came from when posting 
            replies are sent and when new users register. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="copyright"></a><b> What's Copyright? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> This copyright statements 
            location is basically saying that any topics or replies that are posted 
            are copyrighted material of your organization. This copyright location 
            also helps to copyright the images of your logo and any other material 
            that may be posted on portal pages. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="homeurl"></a><b> What's the Home URL? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> The Home URL is the base 
            address for your website. An example would be:<br />
            <b>http://www.SkyPortal.net</b><br />
            <br />
            NOTE: Include the full path of the URL whether that begin with <b>http://</b>. 
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="AuthType"></a><b> Authorization Type? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> You can select DataBase, 
            NT Domain or Active Directory (AD) authorization.<br>
            <br>
            If you are planning to use this for an INTRANET, then select either 
            NT Domain or AD (Active Directory), otherwise, select DataBase. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="PMtype"></a><b> Private Messaging 
            Type </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ul>
              <li><b>Graphic</b><br />
                This is the blinking icon that appears next to your name in the 
                nav bar when you recieve a Private Message.</li>
              <li><b>Toast</b><br />
                This is the pop-up notification that appears in the bottom right 
                corner of your browser when you recieve a Private Message.</li>
              <li><b>Both</b><br />
                With this option selected, both of the above will be active when 
                you get a new Private message.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="allowuploads"></a><b> Allow Uploads 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Do you want to allow people 
            to be able to upload. This covers all upload areas.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="headtype"></a><b> Header Type </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ul>
              <li><b>None</b><br />
                This will display nothing in the header area except your logo.</li>
              <li><b>Icons</b><br />
                This will display ICON links for the different areas of the website.<br />
                IE: Classifieds, Articles, Downloads, etc</li>
              <li><b>Rotating Banner</b><br />
                This will change the banner every 10 seconds. This option will 
                not count the 'impressions', but will count the hits.</li>
              <li><b>Random Banner</b><br />
                This will display a random banner with each page. This option 
                counts impressions as well as hits on the banner</li>
              <li><b>Other</b><br />
                This option calls a function in inc_header.asp named <i>showOther()</i>. 
                You can populate this function for your own code to display in 
                the header.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="lockdown"></a><b> Site Lock Down </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ul>
              <li><b>Yes</b><br />
                This setting will effectively close the site to all non-logged 
                in visitors. No part of the site will be visable to them. They 
                will simply see a login box and a link to the registration page 
                (if applicable) 
              <li><b>No</b><br />
                This setting will allow guest visitors to view all areas of the 
                site except those areas marked as 'private'.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> </li> </ul> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="loginloc"></a><b> Login Box Location 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ul>
              <li><b>Header</b><br />
                This places the login box in the header on all pages.</li>
              <li><b>NavBar</b><br />
                This places the login box in the nav bar on all pages except the 
                home page.</li>
              <li><b>Other</b><br />
                This setting will not display a login box nor a logout button. 
                This setting assumes you have coded the login/logout buttons/links 
                to display elsewhere.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="allowreg"></a><b> New Registrations 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ul>
              <li><b>Users</b><br />
                Users can register themselves. Approval is based on the <i>Member 
                Validation</i> selection below. </li>
              <li><b>Admin</b><br />
                Users cannot register themselves. They must be sent an email invitation 
                when the admin registers them on the 'register' page.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="UniqueEmail"></a><b> Require Unique 
            Email </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Do you want to require each 
            user to have an email address that is different from every other member?<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="valtype"></a><b> Registration Validation 
            and Notifications </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall">The <b>Validation</b> box 
            determines the contents of the <b>Notification</b> box. 
            <ul>
              <b>Notifications</b> 
              <li>Who will recieve <i>email notification</i> when a new member 
                registers or is <i>accepted/rejected</i> for membership.</li>
            </ul>
            <ul>
              <b>Validation </b><br />
              <li><b>None</b><br />
                Guests register and are automatically members.<br />
                No email validation</li>
              <li><b>Member</b><br />
                Each user is required to validate their e-mail address when they 
                first Register and anytime they change their e-mail address The 
                user will receive an e-mail with a link in it that will validate 
                that the e-mail address they entered is a valid e-mail address.</li>
              <li><b>Admin</b><br />
                User does not validate their email, they become 'Pending Members' 
                upon registration. An email is then sent to the site admin, notifying 
                them of a new member. When visiting the site, the pending members 
                can be confirmed or denied from within the admin panel. Email 
                is then sent to the user notifying them of <i>acceptance</i> or 
                <i>denial</i> 
              <li><b>Member &amp; Admin</b><br />
                After users validate their email during Registration, they become 
                'Pending Members'. An email is then sent to the site admin, notifying 
                them of a new member. When visiting the site, the pending members 
                can be confirmed or denied from within the admin panel. Email 
                is then sent to the user notifying them of <i>acceptance</i> or 
                <i>denial</i>.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> </li> </ol> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="sectype"></a><b> Security Image Protection 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> The Security Image is a 
            graphic code (Gimpy Captcha) displayed and requiring the users to 
            input. This is to prevent automated login and registration to your 
            site. It is a good anti-spam measure. 
            <ul>
              <b>Security Image Settings </b><br />
              <li><b>Off</b><br />
                No Security Image will be displayed.</li>
              <li><b>Registration</b><br />
                Each user is required to enter the displayed security code when 
                Registering for the site</li>
              <li><b>Users</b><br />
                Security Code will be required when registering, editing profile 
                and on each site login</i> 
              <li><b>Users &amp; Admin</b><br />
                Security Code will be required when registering, editing profile,on 
                each site login, and on Admin area logins.</li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> </li> </ol> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="imgComp"></a><b> Image Component</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> This is a list of any available 
            Image manipulating components that are installed on your server. SkyPortal 
            currently only supports Asp.Net, AspJpeg, and AspImage..<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="uallow"></a><b> Allow Uploads </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Do you want to allow people 
            to be able to upload. This covers all upload areas.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="upComp"></a><b> Upload Component</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> This is a list of any available 
            Upload components that are installed on your server. SkyPortal currently 
            only supports Asp for uploading..<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="usize"></a><b> Upload File Size Limit 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Maximum file size you want 
            to allow to be uploaded in this area. Number must be entered in kilobytes. 
            ie: a 1 megabyte (mb) file is 1000 kilobytes (kb)<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="uextentions"></a><b> Upload Extentions 
            Allowed </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> List the file extentions 
            you want to allow to be uploaded in this area. Separate each type 
            with a comma.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ulogfile"></a><b> Upload Log File 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Input the name of the file 
            that you want to log the uploads to.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ulog"></a><b> Log Uploads </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Turned <b>ON</b>, this will 
            log all upload attempts, whether successfull or not, to the file named 
            in the <b>Log File</b> textbox.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="uwho"></a><b> Who can Upload </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Determines which group of 
            members you will allow access to the upload functions on the site. 
            <ul>
              <li><b>Everyone</b><br />
                This will allow guests the ability to use the upload functionality</li>
              <li><b>All Members</b><br />
                All Members can upload. Guests canno upload.</li>
              <li><b>Mods and Admin</b><br />
                Only Moderators and Administrators can use the upload functions.</li>
              <li><b>Admin</b><br />
                Only Administrators can use the upload functionality</li>
              <li><b>Super Admin</b><br />
                Only the Super Administrators of the site can use the upload functionality. 
              </li>
            </ul>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a></span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="timetype"></a><b> Time Display? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Choose 24Hr to display all 
            times in military (24 hour) format or 12Hr to display all times in 
            12 hour format appended with an AM or PM depending on whether it's 
            before or after midday. Default is 24 hour. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="TimeAdjust"></a><b> Time Adjustment? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Enter either a positive 
            or negative integer value between +12 and 0 and -12. This may come 
            in handy if your located in one part of the world, and your server 
            is in another, and you need the time displayed in the portal to be 
            converted to a local time for you! (Default value is 0, meaning no 
            adjustment) <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="StrIcsLocation"></a><b> .ICS File 
            Location? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Enter a directory where 
            the .ics file for sending iCal event information will be stored on 
            your server.&nbsp; Anonymous users must have write permissions to 
            the directory as well as read permissions.&nbsp; By default this is 
            set to \database\somefile.ics&nbsp; As a security measure you should 
            at least change the filename, as an example to your site name: \database\mysite.ics 
            Even if you do not allow iCal event information features on your site, 
            this variable must be filled out.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="StriCalNew"></a><b> Allow iCal for 
            New Events? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Toggles whether or not you 
            will allow members to send themselves an iCal for an event they are 
            entering.&nbsp; Creation of an iCal during event entry requires FileSystemObject 
            (FSO) to be available on your webserver.&nbsp; Some hosting providers 
            do not allow FSO, if FSO is not present you must check No for this 
            item.&nbsp; If it is present, checking Yes will allow users to send 
            themselves the iCal during event entry.&nbsp; If iCals are also allowed 
            for existing events then members will have the option of receiving 
            an iCal during event entry if this option is checked or from already 
            existing events if the other option for existing iCals is checked.&nbsp; 
            The Existing iCal option does not require FSO.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="email"></a><b> What does Email do? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Disabling the Email function 
            will turn off any features that involve sending mail. If you don't 
            have an SMTP server of any type, you will want to turn this feature 
            off. If you do have an SMTP (mail) server, however, then also select 
            the type of server you have from the dropdown menu. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="mailserver"></a><b> What is a Mail 
            Server Address? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> The mail server address 
            is the actual domain name that resolves your mail server. This could 
            be something like:<br />
            <b>mail.SkyPortal.net</b><br />
            or it could be the same address as the web server:<br />
            <b>www.SkyPortal.net</b><br />
            Either way, don't put the <b>http://</b> on it.<br />
            <br />
            <b>NOTE:</b> If you are using CDONTS as a mail server type, <br />
            and your server supports it, you do not need to fill in this field. 
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="mailserverusername"></a><b> Email 
            Server Username</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> If your Mailserver requires 
            a Username and/or Password, enter the Username here.<br>
            <br>
            This is an OPTIONAL field. Only use it if your Mailserver requires 
            it. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="mailserverpassword"></a><b> Email 
            Server Password</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> If your Mailserver requires 
            a Username and/or Password, enter the Password here.<br>
            <br>
            This is an OPTIONAL field. Only use it if your Mailserver requires 
            it. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="mailserverport"></a><b> Email 
            Server Port</b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> If your Mailserver requires 
            you to use a specific PORT, enter the port here.<br>
            <br>
            This is an OPTIONAL field. Only use it if your Mailserver requires 
            it. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="sender"></a><b> Portal Email Address? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> This address is referenced 
            by the portal in a couple ways.<br />
            <ol>
              <li>When mail is sent, it is sent from this user email account.</li>
              <li>This Email is also the point of contact given if there is a 
                problem with the portal.</li>
            </ol>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="LogonForMail"></a><b> Require Logon 
            for sending Mail? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Do you require a user to 
            be logged on before being able to use the <i>Send Topic To a Friend</i> 
            or <i>Email Poster</i> options ? <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="privateforums"></a><b> What are Private 
            Forums? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Private Forums enable you 
            to only allow certain members to see that the forum exists. If it's 
            only password protected, everyone can see that it exists, however, 
            they are prompted for a password to get in. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ShowRank"></a><b> Showing Titles? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> 
            <ol>
              <li>Don't Show Any</li>
              <li>Show Titles Only</li>
              <li>Show Icons Only</li>
              <li>Show Both Titles and Icons</li>
            </ol>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="IPLogging"></a><b> What is IP Logging? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> IP Logging will record in 
            the database the IP address of the person who posted a new Topic or 
            Reply. A moderator or administrator then could view the IP by clicking 
            on an icon above the post in the topic. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ShowModerator"></a><b> What does Show 
            Moderators do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Basically, if this function 
            is on, it shows the name of the moderator beside the forum that they 
            moderate on the main default page. If it is off, however, visitors 
            won't see who is moderating the forum they are posting in. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="AllowForumCode"></a><b> Enable Forum 
            Code? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> By turning on Forum Code, 
            you can allow users to mark up their posts with safe codes. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="AllowHTML"></a><b> Why would I allow 
            HTML? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Enabling HTML will place 
            a HTML editor in places where members input messages. If their browser 
            does not allow this editor, then they will see the &quot;forum code&quot; 
            editor in its place.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="FeaturedPoll"></a><b> What is featured 
            poll? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Featured poll is the poll 
            that shows up on the front page. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a>Turning 
            this 'on' will allow them on the front page to be voted on. You can 
            only have one featured poll at a time.</span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="AllowMemPoll"></a><b> What is Allow 
            Member Poll? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> This allow members to create 
            their own polls. Only admins and moderator can create new polls if 
            this feature is disabled. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="hottopics"></a><b> What are Hot Topics? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Hot Topics change the topic 
            folder icon in the Forum view from a normal folder to a flaming folder 
            to let people know that your minimum number of posts has been met 
            to categorize this topic as one that is seeing a lot of action. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="imginposts"></a><b> Why enable Images 
            in Posts? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows users to place images 
            into their Posts. However, you should be aware that this feature would 
            allow anyone to post ANY image in your forums. This may lead to broken 
            links and potentially objectionable material being displayed. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="homepages"></a><b> What is Homepages 
            For? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allow your users to display 
            their homepage link by their name on each post. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="icq"></a><b> What is the ICQ Option 
            For? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Turns On/Off features that 
            allow users to enter their ICQ number... then for other users to send 
            them ICQ messages and/or see if they are online. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="yahoo"></a><b> What is the YAHOO Option 
            For? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Turns On/Off features that 
            allow users to enter their YAHOO username... then for other users 
            to send them messages and/or add them to their buddy list. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="aim"></a><b> What is the AIM Option 
            For? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Turns On/Off features that 
            allow users to enter their AIM username... then for other users to 
            send them messages and/or add them to their buddy list. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="icons"></a><b> What do Icons do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allow users to post smiley 
            faces and other icons allowed by the Forums within the body of their 
            posts! <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ShowPaging"></a><b> What does Show 
            Quick Paging do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Shows the page numbers of 
            any given Topic whenever there are more than one page for that Topic 
            (there will be more than one page whenever there are more Topic Replies 
            posted in that Topic than the number specified by the <a href="#ItemsPerPage"><b>Items 
            per Page</b></a> Setting. This will be displayed as <br />
            <CENTER>
              <span class="fAlert"><b>Previous 1 2 3 4 5 6 7 8 9 NEXT</b></span>, 
            </CENTER>
            <br />
            for instance, at the bottom of each of the Topic's pages. For this 
            example, the Topic has 9 pages, and the user will be able to navigate 
            to any page by clicking on the page number, or can click on the Previous 
            Page or Next Page links as well to navigate in an ordinal way through 
            all the pages in ascending or descending order. If turned OFF, the 
            only links that will appear will be <br />
            <center>
              <span class="fAlert"><b>Previous &nbsp;Next</b></span> 
            </center>
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="QuickReply"></a><b> What is the Show 
            Quick Reply Box setting? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> If turned ON, allows the 
            Quick Reply field (textarea) to be present at the bottom of all Forum 
            Topic pages. Allows the user to post a Reply to that Topic in question, 
            with formatted text. If turned OFF, the user is forced to click on 
            the <img src="<%= strHomeURL %>images/icons/icon_reply_topic.gif"> 
            icon in order to post a Reply Message, taking him to another page 
            to post his or her Reply. The Quick Reply Box is a very convenient 
            feature and it's recommended that you leave it ON. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="ItemsPerPage"></a><b> What does the 
            Items per Page setting do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows you to configure 
            the number (how many) Topic Replies (single posts) are to be displayed 
            on each Topic Page on any given Topic (there will be more than one 
            page whenever there are more Topic Replies posted in that Topic than 
            the number specified by this setting. For example, if the number of 
            items set is 15, then the number of Replies displayed on a single 
            Topic page will be 15, accordingly. Once the count reaches 16 Replies 
            for that Topic, a new page will be generated and the Paging function 
            will build the corresponding Paging links at the bottom of each page 
            of that Topic.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="PageNumbersPerRow"></a><b> What does 
            the Pagenumbers per Row setting do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows you to configure 
            the number of page numbers to display whenever there is more than 
            one page to be displayed on each Topic Page on any given Topic (there 
            will be more than one page whenever there are more Topic Replies posted 
            in that Topic than the number specified by this setting. (see the 
            <a href="#ShowPaging">Quick Paging reference</a>). Once the number 
            of pages has reached the number set by this feature, then a new row 
            of page numbers will be created underneath. For example, if the number 
            (Page numbers) is set to "10", then the number of page numbers displayed 
            on a single row will be 10, and if there are additional pages, the 
            Paging Links shown starting from Page 11 and up will be displayed 
            on a row just under the one showing Pages 1 through 10, like this: 
            <br />
            <CENTER>
              <span class="fAlert"><b> Previous 1 2 3 4 5 6 7 8 9 10 NEXT <br />
              Previous 11 12 13 14 NEXT</b></span>, 
            </CENTER>
            <br />
            at the bottom of each page of that Topic.<a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="secureadminmode"></a><b> Secure Admin 
            Mode? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> <b><span class="fAlert">WARNING:</span> 
            Only turn Secure Admin off if you absolutely need to. If this option 
            is turned off, anyone can change your portal's configuration!</b> 
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="allownoncookies"></a><b> Why would 
            I want Non-Cookie Mode on? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> If your user base does not 
            use cookies, then you would want to turn this function "ON". WARNING: 
            all your admin functions will be visible to all users if this function 
            is "ON". <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="RankColor"></a><b> Color of Icons? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> You can change the color 
            of Icons that show up for each rank of member. (only when the Icons 
            function is turned on) <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="editedbydate"></a><b> What would Edited 
            By on Date do? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> When a post is edited, there 
            is an appending to the end of the post that says when and by whom 
            the post was edited. Turning this function off would make it so that 
            the footer would not be placed on the end of the post. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="badwordfilter"></a><b> Bad Word Filter? 
            </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Screen out words you and 
            your guests would find offensive. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="FloodCheck"></a><b> What is Flood 
            Check? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> With Flood Check enabled, 
            normal users will have to wait the specified amount of time between 
            posts before they can post again. <br />
            <br />
            Admins and Moderators are not affected by this limitation. <br />
            <br />
            You can choose 30 seconds, 60 seconds, 90 seconds or 120 seconds. 
            <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="Picture"></a><b> What is the Picture 
            setting for? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows each member to add 
            his or her own personal Photo to their Member Profile. If set to OFF, 
            they will not see that option on their profile page. This is a site-wide 
            setting (affects ALL users). <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="RecentTopics"></a><b> What is the 
            Recent Topics setting for? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows each member to have 
            a small list of the most Recent Forum Topics on which they have participated 
            (either created or replied) on their Personal Profile page. This list 
            will appear at the end of the left-hand column of such page. If set 
            to OFF, they will not see that option (the list) on their profile 
            page. This is a site-wide setting (affects ALL users). <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="Occupation"></a><b> What is the Occupation 
            setting for? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows each member to specify 
            their current Occupation, Profession or line of work on their Personal 
            Profile page. If set to OFF, they will not see that option on their 
            profile page. This is a site-wide setting (affects ALL users). <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="FavLinks"></a><b> What is the Favorite 
            Links setting for? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows members to specify 
            their favorite links to other websites (either their own or not) on 
            their Personal Profile page. If set to OFF, they will not see that 
            option on their profile page. This is a site-wide setting (affects 
            ALL users). <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
        <tr> 
          <td class="tSubTitle"><a name="Vars1to4"></a><b> What are the Var 
            1 to Var 4 fields used for? </b></td>
        </tr>
        <tr> 
          <td class="tCellAlt1"><span class="fSmall"> Allows you as the administrator 
            to configure and specify the LABEL for any extra fields you might 
            need to add to your members' Profiles. Feel free to name these four 
            variables any way you want or need (examples: My Car is a, My Favorite 
            Food, How many Children I Have, Color of My Eyes, etc, etc, etc.). 
            <br />
            If left empty, the fields and their corresponding labels will not 
            appear on their Profile page. The Labels you can see already there 
            are only suggestions. <a href="#top"><img src="<%= strHomeURL %>themes/<%= strTheme %>/icons/icon_go_up.gif" border="0" align="right"></a> 
            </span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
