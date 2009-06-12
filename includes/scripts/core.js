<!-- //
//':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
//':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
//'<> Copyright (C) 2005-2008 Dogg Software All Rights Reserved
//'<>
//'<> By using this program, you are agreeing to the terms of the
//'<> SkyPortal End-User License Agreement.
//'<>
//'<> All copyright notices regarding SkyPortal must remain 
//'<> intact in the scripts and in the outputted HTML.
//'<> The "powered by" text/logo with a link back to 
//'<> http://www.SkyPortal.net in the footer of the pages MUST
//'<> remain visible when the pages are viewed on the internet or intranet.
//'<>
//'<> Support can be obtained from support forums at:
//'<> http://www.SkyPortal.net
//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

window.status=js_welcome;
/* ********** some hide/show script ********************************** */

  for (i=0; i < mmImages.length; i++) {
    var preload = new Image();
    preload.src = mmImages[i];
  }
  
function swap(imgID,img) {
 document[''+imgID+'Img'].src = mmImages[img];
 if(document[''+imgID+'Img'].alt != js_collapse ){
	document[''+imgID+'Img'].title = js_collapse;
	document[''+imgID+'Img'].alt = js_collapse;
 }else{
	document[''+imgID+'Img'].title = js_expand;
	document[''+imgID+'Img'].alt = js_expand;
 }
}
// forum min/max function
function mwpHS(obj,idd,tagg){
	if(document.getElementById){
	var ar = document.getElementById(obj).getElementsByTagName(tagg);
	var cook = jsUniqueID + "hide";
	var clsNam = obj + idd;
		for (var i=0; i<ar.length; i++){
		  if(ar[i].id==clsNam){
			if(ar[i].style.display != "none"){
				swap(clsNam,0);
				ar[i].style.display = "none";
				setCookieSubKey(cook,clsNam,"1");
			}else{
				swap(clsNam,1);
				ar[i].style.display = "";
				setCookieSubKey(cook,clsNam,"0");
			}
		  }
		} 
	}
}
//themebox min/max
function mwpHSx(obj){
  if(document.getElementById){
	var ele = document.getElementById(obj);
	var cook = jsUniqueID + "hide";
	if(ele.style.display != "none"){
		swap(obj,0);
		ele.style.display = "none";
		setCookieSubKey(cook,obj,"1");
	}else{
		swap(obj,1);
		ele.style.display = "block";
		setCookieSubKey(cook,obj,"0");
	}
	//alert(cook + ' : ' + obj );
  }
}

var ns4 = (document.layers);
var ie4 = (document.all && !document.getElementById);
var ie5 = (document.all && document.getElementById);
var ns6 = (!document.all && document.getElementById);

// alternate hide/show
function mwpHSa(obj,typ){
	// Netscape 4
	if(ns4){
		if (document.layers[obj]){
			if (document.layers[obj].visibility != "hide"){
				document.layers[obj].visibility = "hide";
			}else{
				document.layers[obj].visibility = "show";
			}
		}
	}
	// Explorer 4
	else if(ie4){
	  if (document.all[obj]){
		if (document.all[obj].style.visibility != "hidden"){
		  document.all[obj].style.visibility = "hidden";
		}else{
		  document.all[obj].style.visibility = "visible";
		}
	  }
	}
	// W3C - Explorer 5+ and Netscape 6+
	else if(ie5 || ns6){
		if (document.getElementById(obj)){
			var ela = document.getElementById(obj);
			//for (var i=0; i<el.length; i++){
				if (ela.style.display != "none"){		
				  if(typ != 3){
					swap(obj,2);
				  }
				  ela.style.display = "none";
				}else{
				  if(typ != 3){
				    swap(obj,3);
				  }
				  ela.style.display = "block";
				}
			//}
		}
	}
}

//simple hide/show
function mwpHSs(obj,typ){
	if (document.getElementById(obj)){
		var ela = document.getElementById(obj);
		//for (var i=0; i<el.length; i++){
		if (ela.style.display != "none"){		
			if(typ != 1){
				swap(obj,0);
			}
			ela.style.display = "none";
		}else{
			if(typ != 1){
				 swap(obj,1);
			}
			ela.style.display = "block";
		}
		//}
	}
}

function checkBrowser(){
	this.ver=navigator.appVersion
	this.dom=document.getElementById?1:0
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.ns5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.ns4=(document.layers && !this.dom)?1:0;
	this.bw=(this.ie5 || this.ie4 || this.ns4 || this.ns5)
	return this
}
bw=new checkBrowser()

function showhide(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	if(obj.display=='block' || obj.display=='block') obj.display='none'
	else obj.display='block'
}

function show(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	obj.display='block'
}

function hide(div,nest){
	obj=bw.dom?document.getElementById(div).style:bw.ie4?document.all[div].style:bw.ns4?nest?document[nest].document[div]:document[div]:0; 
	obj.display='none'
}

function stripHTML(){
var re= /<\S[^><]*>/g
for (i=0; i<arguments.length; i++)
arguments[i].value=arguments[i].value.replace(re, "")
}

// ++++++++++++++++++++++++ Cookie code +++++++++++++++++++++++++
    function setCookie(cname,value) {
	var timeout=60*60*24;
	var today = new Date();
	var the_date = new Date();
	the_date.setTime(today.getTime() + 365000 * timeout);
	//alert(the_date);
	var the_cookie_date = the_date.toGMTString();
	var the_cookie = cname +"="+value;
	var the_cookie = the_cookie + ";expires=" + the_cookie_date;
    document.cookie= the_cookie; 
    //E.g. setCookie("name1","dogg")
	}

    function getCookie(name) {
    	//alert(getCookie("name1"));
    	var result = ""; 
    	var myCookie = " " + document.cookie + ";";
    	var searchName = " " + name + "=";
    	var startOfCookie = myCookie.indexOf(searchName); 	
    	var endOfCookie; 
		if (startOfCookie != -1) {
        		startOfCookie += searchName.length; 
        		endOfCookie = myCookie.indexOf(";", startOfCookie); 
        		result = unescape(myCookie.substring(startOfCookie, endOfCookie)); 
        }
        	return result; 
    }
    //get multi value cookie value e.g. 
    //     Person=name=dogg&age=25;
    function getCookieSubKey(cookiename,cookiekey) {
        var cookievalue=getCookie(cookiename);
        if ( cookievalue == "")  return "";
        cookievaluesep=cookievalue.split("&");
        	for (c=0;c<cookievaluesep.length;c++)	{
            	cookienamevalue=cookievaluesep[c].split("=");
            	if (cookienamevalue.length > 1) {  //it has multi valued cookie
					if ( cookienamevalue[0] == cookiekey )
						return cookienamevalue[1].toString();			
                }
                else		
                	return "";		
            }	
    	return "";
	}
    //set multi value cookie value e.g. 
    //     Person=name=dogg&age=25;
	function setCookieSubKey(cookiename,cookiekey,cookiekeyvalue){
		var cookievalue=getCookie(cookiename);
        if ( cookievalue.trim() == "" ){
        	setCookie(cookiename,cookiekey+"="+cookiekeyvalue);
            return;
        }		
        //check if cookie already exist
        getcookiekeyvalue=getCookieSubKey(cookiename,cookiekey);
        newCookieValue=cookievalue.trim();
        if ( getcookiekeyvalue == "")	//key cookie never exist		
        	newCookieValue += "&" + cookiekey + "=" + cookiekeyvalue;
        else
		{
        	if ( newCookieValue.substr(0,cookiekey.length+1) == (cookiekey + "=") ) {  //Check if at first location . no beginning with &
		  	//pick rest keys = keylength+equalsign+cookiekeyvalue+nextampesand
             totalcookiekeylength=cookiekey.length+1+getCookieSubKey(cookiename,cookiekey).length+1;
             newCookieValue = newCookieValue.substr(totalcookiekeylength);
             if (newCookieValue.trim() == "")			
                newCookieValue = cookiekey + "=" + cookiekeyvalue;
             else
                newCookieValue += "&" + cookiekey + "=" + cookiekeyvalue;
           }
           else 
		   {
          	  fullcookiekey="&"+cookiekey+"="+getcookiekeyvalue;
              if ( newCookieValue.indexOf(fullcookiekey) != -1 ) //cookie key inside the cookie value
			  {
              	  newCookieValue = ReplaceAll(newCookieValue, fullcookiekey, "");
                  if (newCookieValue.trim() == "")			
                      newCookieValue = cookiekey + "=" + cookiekeyvalue;
                  else
                      newCookieValue += "&" + cookiekey + "=" + cookiekeyvalue;
               }
            }
		}
        setCookie(cookiename,newCookieValue);
	}
	//Replace all given string from a string
	//
	function ReplaceAll(varb, replaceThis, replaceBy){	
    	newvarbarray=varb.split(replaceThis);
        newvarb=newvarbarray.join(replaceBy);	
        return newvarb;
	}
	
	String.prototype.trim = function(){
    // Use a regular expression to replace
    //      leading and trailing 
    // spaces with the empty string
    return this.replace(/(^\s*)|(\s*$)/g, "");
    }
// +++++++++ End Cookie code +++++++++++++++++++++++++++++++++++++++++++++++++

// ------------functions for codebox mod

function docodebox(el){
var id1='thecode'+el;
//alert('X' + id1 + 'X');
//var codbox=eval('document.selectcode'+el+'.thecode'+el);
//var id1='thecode'+el;
var codbox=document.getElementById(id1);
codbox.focus();
codbox.select();
}

function expand(el){
var id1='thecode'+el;
var codebox=document.getElementById(id1);
var scheight = codebox.scrollHeight +10;
if (txttype=='opera') {
    codebox.style.height='100%';}
else if (txttype=='ie') {
    codebox.style.height=scheight+'px';
    codebox.style.overflowX='auto';
    codebox.style.overflowY='auto';}
else {
codebox.style.height=scheight+'px';
codebox.style.overflow='visible';}
}

function contract(el){
//alert('X' + el + 'X');
var id1='thecode'+el;
var codebox=document.getElementById(id1);
codebox.style.height=45+'px';
codebox.style.overflow='auto';
}
function dohelp(){
// Help Code Popup
var doPopUpHelpCodeX = (screen.width/2)-110;
var doPopUpHelpCodeY = (screen.height/2)-150;
var pos = "left="+doPopUpHelpCodeX+",top="+doPopUpHelpCodeY;
doPopUpHelpCodeWindow = window.open("includes/code_help.asp","HelpCode","width=220,height=325,"+pos);
}
// ------------------ end codebox code
var arrItems1 = new Array();
var arrItemsGrp1 = new Array();
arrItems1[0] = js_none;
arrItemsGrp1[0] = 1;
arrItems1[1] = js_member;
arrItemsGrp1[1] = 1;
arrItems1[2] = js_admin;
arrItemsGrp1[2] = 1;
arrItems1[3] = js_member + " & " + js_admin;
arrItemsGrp1[3] = 1;
arrItems1[4] = js_member;
arrItemsGrp1[4] = 2;
arrItems1[5] = js_member + " & " + js_admin;
arrItemsGrp1[5] = 2;
arrItems1[6] = js_member;
arrItemsGrp1[6] = 3;
arrItems1[7] = js_member;
arrItemsGrp1[7] = 4;

function selectChange(control, controlToPopulate, ItemArray, GroupArray)
{
  var x;
  var cntTotalOptions = 0;
  // Empty the second drop down box of any choices
  controlToPopulate.options.length = 0;
  // ADD Default Choice - in case there are no values
  controlToPopulate.options[cntTotalOptions] = new Option('[SELECT]',0);
  
  for ( x = 0 ; x < ItemArray.length  ; x++ )
    {
      if ( GroupArray[x] == control.value )
        {
          cntTotalOptions++;
          controlToPopulate.options[cntTotalOptions] = new Option(ItemArray[x],x+1);
        }
    }
}

//var maxWidth=100;
//var maxHeight=100;
var maxWidth;
var maxHeight;
var fileTypes=["bmp","gif","png","jpg","jpeg"];
var outImage="previewField";
var defaultPic="images/spacer.gif";

var globalPic;
function preview(what,mwid,mhgt){
  maxWidth=mwid;
  maxHeight=mhgt;
  var source=what.value;
  var ext=source.substring(source.lastIndexOf(".")+1,source.length).toLowerCase();
  for (var i=0; i<fileTypes.length; i++) if (fileTypes[i]==ext) break;
  globalPic=new Image();
  if (i<fileTypes.length) {
    globalPic.src=source;
	//globalPic.width=maxWidth;
    //globalPic.height=maxHeight;
  } else {
    globalPic.src=defaultPic;
    alert("THAT IS NOT A VALID IMAGE\nPlease load an image with a valid extention");
  }
  setTimeout("applyChanges()",200);
}
function applyChanges(){
  var field=document.getElementById(outImage);
  var x=parseInt(globalPic.width);
  var y=parseInt(globalPic.height);
  if (x>maxWidth) {
    y*=maxWidth/x;
    x=maxWidth;
  }
  if (y>maxHeight) {
    x*=maxHeight/y;
    y=maxHeight;
  }
  //alert(globalPic.src);
  field.style.display=(x<1 || y<1)?"none":"";
  field.src=globalPic.src;
  field.width=x;
  field.height=y;
}

// Group Access functions

function selectUsers(fm){
  //alert(fm);
  selectAll(fm,'g_read');
  selectAll(fm,'g_write');
  selectAll(fm,'g_full');
}

function selectAll(fom,ob){
  //alert(fm + ' : ' + ob);
	var oFrm = document[fom];
	for (x = 0;x < oFrm[ob].length ;x++)
	  if(oFrm[ob].options[x].value != '0'){
		oFrm[ob].options[x].selected = true;
	  } else {
		oFrm[ob].options[x].selected = false;
	  }
}

function moveGroup(strAction,sForm,sFrom,sTo){
	var pos,user,mText;
	var count,finished,ir;
	var oFrm = document[sForm]
	if (strAction == "Add")
	{
		pos = oFrm[sTo].length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (oFrm[sFrom].options[count].text == "")
			{
				//alert("You must select a Group\nfrom the 'Allowed Users'");
				finished = true;
				continue;
			}
			if (oFrm[sFrom].options[count].selected)
			{
			  for (ir=0; ir<oFrm[sTo].length; ir++) {
			    if (oFrm[sTo].options[ir].value==oFrm[sFrom].options[count].value) {
			    // group already added
			    return;
			    }
			  }
				oFrm[sTo].length +=1;
				oFrm[sTo].options[pos].value = oFrm[sTo].options[pos-1].value;	
				oFrm[sTo].options[pos].text = oFrm[sTo].options[pos-1].text;
				oFrm[sTo].options[pos-1].value = oFrm[sFrom].options[count].value;	
				oFrm[sTo].options[pos-1].text = oFrm[sFrom].options[count].text;
				oFrm[sTo].options[pos-1].selected = true;
			}
			pos = oFrm[sTo].length;
			count += 1;
		}while (!finished); //finished adding
	}

	if (strAction == "Del")
	{
		pos = document.PostTopic.AuthUsersCombo.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected)
			{
				document.PostTopic.AuthUsersCombo.length +=1;
				document.PostTopic.AuthUsersCombo.options[pos].value = document.PostTopic.AuthUsersCombo.options[pos-1].value;	
				document.PostTopic.AuthUsersCombo.options[pos].text = document.PostTopic.AuthUsersCombo.options[pos-1].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].value = document.PostTopic.AuthUsers.options[count].value;	
				document.PostTopic.AuthUsersCombo.options[pos-1].text = document.PostTopic.AuthUsers.options[count].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].selected = true;
			}
			count += 1;
			pos = document.PostTopic.AuthUsersCombo.length;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsers.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}		
}

function remGroup(fm,ob,it){
	var count,finished;
	var oFrm = document[fm]
	finished = false;
	count = 0;
	count = oFrm[ob].length - 1;
	if (count<1) {
		return;
	}
	do //remove from source
	{	
		if (oFrm[ob].options[count].text == ""){
			--count;
			continue;
		}
		if (oFrm[ob].options[count].value == it ){
		  for ( z = count ; z < oFrm[ob].length-1;z++){	
			oFrm[ob].options[z].value = oFrm[ob].options[z+1].value;	
			oFrm[ob].options[z].text = oFrm[ob].options[z+1].text;
		  }
		  oFrm[ob].length -= 1;
		}
		--count;
		if (count < 0)
			finished = true;
	}while(!finished) //finished removing
}

function removeGroup(fm,ob){
	var user,mID;
	var count,finished;
	var oFrm = document[fm]
	finished = false;
	count = 0;
	count = oFrm[ob].length - 1;
	if (count<1) {
		return;
	}
	do //remove from source
	{	
		if (oFrm[ob].options[count].text == ""){
			--count;
			continue;
		}
		if (oFrm[ob].options[count].selected ){
		  mID = oFrm[ob].options[count].value
		  for ( z = count ; z < oFrm[ob].length-1;z++){	
			oFrm[ob].options[z].value = oFrm[ob].options[z+1].value;	
			oFrm[ob].options[z].text = oFrm[ob].options[z+1].text;
		  }
		  oFrm[ob].length -= 1;
		}
		--count;
		if (count < 0)
			finished = true;
	}while(!finished) //finished removing
	
	if (ob=="g_read") {
		remGroup(fm,"g_write",mID);
		remGroup(fm,"g_full",mID);
	}
	if (ob=="g_write") {
		remGroup(fm,'g_full',mID);
	}
		//return;
}

function eGroup(fm,ob){
	var user,mID;
	var count,finished;
	var oFrm = document[fm]
	finished = false;
	mID = ""
	count = 0;
	count = oFrm[ob].length - 1;
	if (count<1) {
		return;
	}
	do
	{	
		if (oFrm[ob].options[count].text == ""){
			--count;
			continue;
		}
		if (oFrm[ob].options[count].selected ){
		  mID = oFrm[ob].options[count].value
		  user = oFrm[ob].options[count].text
		}
		--count;
		if (count < 0)
			finished = true;
	}while(!finished) //finished removing
	
		//alert(user + " : " + mID);
	if (mID != 0){
	if (mID != 2){
	if (mID != 3){
    var whereto = "pop_portal.asp?cmd=11&cid=" + mID;
	popUpWind(whereto,'egroups','430','580','yes','yes');
	}}}
}

function validate(){
    document.PostTopic.Message.focus()
	if (document.PostTopic.Subject) {
		if (trim(document.PostTopic.Subject.value)=="") {
			alert("You must enter a Subject");
			return false;
		}
	}
	if (setTimeout((document.PostTopic.Message), 1000)) {
		if (trim(document.PostTopic.Message.value)=="") {
			alert("You must enter a Message");
			return false;
		}
	}
	return true
}

// Banner functions
	  
  function imageItem(image_location) {
	this.image_item = new Image();
	this.image_item.src = image_location;
	}
	
  function get_ImageItemLocation(imageObj) {
    return(imageObj.image_item.src);
  }
  
  function generate(x, y) {
  var range = y - x + 1;
  return Math.floor(Math.random() * range) + x;
  }

function allmemberList(frm,obj) { 
  var whereto = "pop_memberlist.asp?pageMode=shoall&frm=" + frm + "&sel=" + obj;
  var MainWindow = window.open(whereto, "memLst","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=500,top=100,left=100,status=no"); }

function allowgroups(fm,ob,grp) { 
  var whereto = "pop_portal.asp?cmd=5&mode=1&frm=" + fm + "&sel=" + ob + "&grps=" + grp;
  var MainWindow = window.open (whereto, "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=300,height=330,top=100,left=100,status=no"); }

function openWindow(url) {
  LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-400)/2 : 0;
  popupWin = window.open(url,'new_page','width=400,height=400,top='+TopPosition+',left='+LeftPosition+'');
}
function openWindow2(url) {
  LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-480)/2 : 0;
  popupWin = window.open(url,'new_page','width=400,height=480,top='+TopPosition+',left='+LeftPosition+'');
}
function openWindow3(url) {
  LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-450)/2 : 0;
  popupWin = window.open(url,'new_page','width=400,height=450,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes');
}
function openWindow4(url) {
  LeftPosition = (screen.width) ? (screen.width-400)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-525)/2 : 0;
  popupWin = window.open(url,'new_page','width=400,height=525,top='+TopPosition+',left='+LeftPosition+'');
}
function openWindow5(url) {
  LeftPosition = (screen.width) ? (screen.width-450)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-525)/2 : 0;
  popupWin = window.open(url,'new_page','width=450,height=525,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes,toolbar=yes,menubar=yes,resizable=yes');
}
function openWindow6(url) {
  LeftPosition = (screen.width) ? (screen.width-550)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-525)/2 : 0;
  popupWin = window.open(url,'new_page','width=550,height=525,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes,resizable=yes');
}
function openWindowCT(url) {
  popupWin = window.open(url,'new_page','width=450,height=480');
}
function openWindowPM(url) {
  LeftPosition = (screen.width) ? (screen.width-635)/2 : 0;
  TopPosition = (screen.height) ? (screen.height-550)/2 : 0;
  popupWin = window.open(url,'pm_pop_send','resizable=yes,width=635,height=550,top='+TopPosition+',left='+LeftPosition+',scrollbars=yes')
}
function openWindowPager(url) {
  popupWin = window.open(url,'pager','resizable,width=210,height=310,left=10,top=75,scrollbars=auto')
}
var popwin = null;
function popUpWind(mypage,myname,w,h,scr,resiz){
LeftPosition = (screen.width) ? (screen.width-w)/2 : 0;
TopPosition = (screen.height) ? (screen.height-h)/2 : 0;
settings =
'height='+h+',width='+w+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scr+',toolbar=no,resizable='+resiz+',menubar=no'
popwin = window.open(mypage,myname,settings)
}
// onclick="popUpWind('default.asp','name','400','400','yes','yes');return false"

	function openJsLayer(ob,w,h) {
	    Dialog.alert($(ob).innerHTML, {windowParameters: {width:w, height:h}, 
        okLabel: "cancel"});
	}

	function openDialog2(n,ob,t,w,h) {
      if(document.getElementById){
		  var ele = document.getElementById(ob);
		  ihtml = ele.innerHTML;
		  var win = new Window(n, {className: "dialog", title: t, width:w, height:h, zIndex:150, opacity:1, resizable: true, maximizable: false})
		  win.getContent().innerHTML = ihtml;
		  win.toFront();
		  win.setDestroyOnClose();
		  win.showCenter();	
	}}

function SwitchMenu(div,obj){
if(document.getElementById){
var el = document.getElementById(obj);
var ar = document.getElementById(div).getElementsByTagName("span");
if(el.style.display != "block"){
for (var i=0; i<ar.length; i++){
if (ar[i].className=="submenu")
ar[i].style.display = "none";
}
el.style.display = "block";
}else{
el.style.display = "none";
}
}
}
//  End -->