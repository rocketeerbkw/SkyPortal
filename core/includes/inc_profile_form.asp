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
 'on error resume next 
 %> 
<script type="text/javascript">
var selectedtablink="";

function handlelink(aele,tab){
//selectedtablink=aobject.href;
if (document.getElementById){
  
var tabobj=document.getElementById("tablist");
var tabobjlinks=tabobj.getElementsByTagName("A");
for (i=0; i<tabobjlinks.length; i++)
tabobjlinks[i].className=""
document.getElementById("" + tab + "").className="current";
  $('pg_html', 'ed_pass', 'ed_basics', 'ed_misc', 'ed_sig', 'ed_contact').invoke('hide');
  $(aele).show();
return false;
}
else
return true;
}
</script>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td align=left valign="top">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td align="center" colspan="2">
		<p><b><%= txtReg1a %>&nbsp;<span class="fAlert"><b>*</b></span>&nbsp;<%= txtReg1b %></b></p>
		<!-- S k y D o g g - S k y P o r t a l - is here - december 2007-->
	    <hr /></td>
	  </tr>
  <tr>
	<td colspan="2" valign="top"><br>
	  <ul id="tablist">
      <li><a id="tab1" class="current" href="javascript:;" onclick="handlelink('ed_basics','tab1');">Basics</a></li>
      <li><a id="tab2" class="" href="javascript:;" onclick="handlelink('ed_pass','tab2');">Edit Password</a></li>
	  <li><a id="tab3" class="" href="javascript:;" onclick="handlelink('ed_contact','tab3');">Contact Info</a></li>
      <li><a id="tab4" class="" href="javascript:;" onclick="handlelink('ed_misc','tab4');">Misc</a></li>
      <li><a id="tab5" class="" href="javascript:;" onclick="handlelink('ed_sig','tab5');">Signature</a></li>
    </ul>
		  <div class="tabframe" style="width:90%;height:400;overflow:scroll;">
		  <%
		  call editBasics(rs,"edit",1)
		  call editPassword(rs,"edit",0)
		  call editContact(rs,"edit",0)
		  call editMisc(rs,"edit",0)
		  call editSig(rs,"edit",0)
		  %>
		  </div>
	</td>
	</tr>
	</table>
	</td>
  </tr> 
	    <%
        'if dtyp = "edit" then %>
        <tr><td colspan="2" class="fNorm" align="center"><br /></td></tr>
        <tr><td colspan="2" class="fNorm" align="center">
		  <b><%= txtRefFrndUrl %>: </b></td></tr>
        <tr><td colspan="2" class="fNorm" align="center">  
            <%= strHomeURL %>policy.asp?rname=<%'= rs("M_NAME")%><br />&nbsp;</td></tr>
        <%
		'end if %>
</table>
<%
'<!-- ::::::::::::::::::::::::: start BASICS info ::::::::::::::::::::::::::::: --> %>