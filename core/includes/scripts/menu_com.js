//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
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
//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
//'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

function hidFm(id){var Lbx=document.getElementById('dogin');if(Lbx.style.display!="none"){if(ns4){if(document.layers[id]){document.layers[id].visibility="hide";}}else if(ie4){if(document.all[id]){document.all[id].style.visibility="hidden";}}else if(ie5||ns6){if(document.getElementById(id)){var el=document.getElementById(id);var ar=document.getElementById(id).getElementsByTagName("select");for(var i=0;i<ar.length;i++){ar[i].style.visibility="hidden";}}}}else{if(ns4){if(document.layers[id]){if(document.layers[id].visibility!="hide"){document.layers[id].visibility="hide";}else{document.layers[id].visibility="show";}}}else if(ie4){if(document.all[id]){if(document.all[id].style.visibility!="hidden"){document.all[id].style.visibility="hidden";}else{document.all[id].style.visibility="visible";}}}else if(ie5||ns6){if(document.getElementById(id)){var el=document.getElementById(id);var ar=document.getElementById(id).getElementsByTagName("select");for(var i=0;i<ar.length;i++){if(ar[i].style.visibility!="hidden"){ar[i].style.visibility="hidden";}else{ar[i].style.visibility="visible";}}}}}}function parentrefresh(){window.opener.refreshpage();window.close();}function add_block(section_name,add_name){section_select=document.getElementById(section_name);add_select=document.getElementById(add_name);if((section_select)&&(add_select)){if(add_select.selectedIndex==0)return false;add_option=add_select.options[add_select.selectedIndex];section_select.options[section_select.length]=new Option(add_option.text,add_option.value);}}function shwFm(id){if(ns4){if(document.layers[id]){document.layers[id].visibility="show";}}else if(ie4){if(document.all[id]){document.all[id].style.visibility="visible";}}else if(ie5||ns6){if(document.getElementById(id)){var el=document.getElementById(id);var ar=document.getElementById(id).getElementsByTagName("select");for(var i=0;i<ar.length;i++){ar[i].style.visibility="visible";}}}}function Browser(){var ua,s,i;this.isIE=false;this.isOP=false;this.isNS=false;this.version=null;ua=navigator.userAgent;s="Opera";if((i=ua.indexOf(s))>=0){this.isOP=true;this.version=parseFloat(ua.substr(i+s.length));return;}s="Netscape6/";if((i=ua.indexOf(s))>=0){this.isNS=true;this.version=parseFloat(ua.substr(i+s.length));return;}s="Gecko";if((i=ua.indexOf(s))>=0){this.isNS=true;this.version=6.1;return;}s="MSIE";if((i=ua.indexOf(s))){this.isIE=true;this.version=parseFloat(ua.substr(i+s.length));return;}}var browser=new Browser();function remove_block(section_name){section_select=document.getElementById(section_name);if(section_select){if(section_select.selectedIndex==-1)return false;section_select.options[section_select.selectedIndex]=null;}}function move_up_block(section_name){section_select=document.getElementById(section_name);if(section_select){if(section_select.selectedIndex<=0)return false;index=section_select.selectedIndex;temp=new Option(section_select.options[index-1].text,section_select.options[index-1].value);section_select.options[index-1]=new Option(section_select.options[index].text,section_select.options[index].value);section_select.options[index]=temp;section_select.selectedIndex=index-1;}}function move_down_block(section_name){section_select=document.getElementById(section_name);if(section_select){if(section_select.selectedIndex<0)return false;if(section_select.selectedIndex>=section_select.length-1)return false;index=section_select.selectedIndex;temp=new Option(section_select.options[index+1].text,section_select.options[index+1].value);section_select.options[index+1]=new Option(section_select.options[index].text,section_select.options[index].value);section_select.options[index]=temp;section_select.selectedIndex=index+1;}}function move_left_right_block(add_to_column,remove_from_column){section_select=document.getElementById(remove_from_column);add_select=document.getElementById(add_to_column);if((section_select)&&(add_select)){add_option=add_select.options[add_select.selectedIndex];section_select.options[section_select.length]=new Option(add_option.text,add_option.value);add_select.options[add_select.selectedIndex]=null;}}var activeButton=null;function buttonClick(event,menuId){var button;if(browser.isIE)button=window.event.srcElement;else button=event.currentTarget;button.blur();if(button.menu==null){button.menu=document.getElementById(menuId);if(button.menu.isInitialized==null)menuInit(button.menu);}if(button.onmouseout==null)button.onmouseout=buttonOrMenuMouseout;if(button==activeButton)return false;if(activeButton!=null)resetButton(activeButton);if(button!=activeButton){depressButton(button);activeButton=button;}else activeButton=null;return false;}function buttonmouseover(event,menuId){var button;m_sub="nav_menu";if(activeButton==null){buttonClick(event,menuId);return;}if(browser.isIE)button=window.event.srcElement;else button=event.currentTarget;if(activeButton!=null&&activeButton!=button)buttonClick(event,menuId);}function depressButton(button){var x,y;button.className+=" menuButtonActive";if(button.onmouseout==null)button.onmouseout=buttonOrMenuMouseout;if(button.menu.onmouseout==null)button.menu.onmouseout=buttonOrMenuMouseout;x=getPageOffsetLeft(button);y=getPageOffsetTop(button)+button.offsetHeight;if(browser.isIE){x+=button.offsetParent.clientLeft;y+=button.offsetParent.clientTop;}button.menu.style.left=x+"px";button.menu.style.top=y+"px";button.menu.style.visibility="visible";if(button.menu.iframeEl!=null){button.menu.iframeEl.style.left=button.menu.style.left;button.menu.iframeEl.style.top=button.menu.style.top;button.menu.iframeEl.style.width=button.menu.offsetWidth+"px";button.menu.iframeEl.style.height=button.menu.offsetHeight+"px";button.menu.iframeEl.style.display="";}}function resetButton(button){removeClassName(button,"menuButtonActive");if(button.menu!=null){closeSubMenu(button.menu);button.menu.style.visibility="hidden";if(button.menu.iframeEl!=null)button.menu.iframeEl.style.display="none";}}function menuMouseover(event){var menu;if(browser.isIE)menu=getContainerWith(window.event.srcElement,"DIV",m_sub);else menu=event.currentTarget;if(menu.activeItem!=null)closeSubMenu(menu);}function select_options(){section_select=document.getElementById('left_sticky');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('main_sticky');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('right_sticky');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('left_select');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('main_select');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('right_select');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}
section_select=document.getElementById('maintop_select');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}section_select=document.getElementById('mainbottom_select');if(section_select){for(i=0;i<section_select.length;i++){section_select.options[i].selected=true;}}
return true;}function show_description(add_name){add_select=document.getElementById(add_name);instruct=document.getElementById('instructions');if(add_select&&instruct){if(add_select.selectedIndex==0)instruct.innerHTML="";else instruct.innerHTML=block_descr[add_select.options[add_select.selectedIndex].value];}}function menuItemMouseover(event,menuId){var item,menu,x,y;if(browser.isIE)item=getContainerWith(window.event.srcElement,"A",m_sub+"Item");else item=event.currentTarget;menu=getContainerWith(item,"DIV",m_sub);if(menu.activeItem!=null)closeSubMenu(menu);menu.activeItem=item;item.className+=" menuItemHighlight";if(item.subMenu==null){item.subMenu=document.getElementById(menuId);if(item.subMenu.isInitialized==null)menuInit(item.subMenu);}if(item.subMenu.onmouseout==null)item.subMenu.onmouseout=buttonOrMenuMouseout;x=getPageOffsetLeft(item)+item.offsetWidth;y=getPageOffsetTop(item);var maxX,maxY;if(browser.isIE){maxX=Math.max(document.documentElement.scrollLeft,document.body.scrollLeft)+(document.documentElement.clientWidth!=0?document.documentElement.clientWidth:document.body.clientWidth);maxY=Math.max(document.documentElement.scrollTop,document.body.scrollTop)+(document.documentElement.clientHeight!=0?document.documentElement.clientHeight:document.body.clientHeight);}if(browser.isOP){maxX=document.documentElement.scrollLeft+window.innerWidth;maxY=document.documentElement.scrollTop+window.innerHeight;}if(browser.isNS){maxX=window.scrollX+window.innerWidth;maxY=window.scrollY+window.innerHeight;}maxX-=item.subMenu.offsetWidth;maxY-=item.subMenu.offsetHeight;if(x>maxX)x=Math.max(0,x-item.offsetWidth-item.subMenu.offsetWidth+(menu.offsetWidth-item.offsetWidth));y=Math.max(0,Math.min(y,maxY));item.subMenu.style.left=x+"px";item.subMenu.style.top=y+"px";item.subMenu.style.visibility="visible";if(item.subMenu.iframeEl!=null){item.subMenu.iframeEl.style.left=item.subMenu.style.left;item.subMenu.iframeEl.style.top=item.subMenu.style.top;item.subMenu.iframeEl.style.width=item.subMenu.offsetWidth+"px";item.subMenu.iframeEl.style.height=item.subMenu.offsetHeight+"px";item.subMenu.iframeEl.style.display="";}if(browser.isIE)window.event.cancelBubble=true;else event.stopPropagation();}function closeSubMenu(menu){if(menu==null||menu.activeItem==null)return;if(menu.activeItem.subMenu!=null){closeSubMenu(menu.activeItem.subMenu);menu.activeItem.subMenu.style.visibility="hidden";if(menu.activeItem.subMenu.iframeEl!=null)menu.activeItem.subMenu.iframeEl.style.display="none";menu.activeItem.subMenu=null;}removeClassName(menu.activeItem,"menuItemHighlight");menu.activeItem=null;}function buttonOrMenuMouseout(event){var el;if(activeButton==null)return;if(browser.isIE)el=window.event.toElement;else if(event.relatedTarget!=null)el=(event.relatedTarget.tagName?event.relatedTarget:event.relatedTarget.parentNode);if(getContainerWith(el,"DIV",m_sub)==null){resetButton(activeButton);activeButton=null;}}function menuInit(menu){var itemList,spanList;var textEl,arrowEl;var itemWidth;var w,dw;var i,j;if(browser.isIE){menu.style.lineHeight="2.5ex";spanList=menu.getElementsByTagName("SPAN");for(i=0;i<spanList.length;i++)if(hasClassName(spanList[i],"menuItemArrow")){spanList[i].style.fontFamily="Webdings";spanList[i].firstChild.nodeValue="4";}}itemList=menu.getElementsByTagName("A");if(itemList.length>0)itemWidth=itemList[0].offsetWidth;else return;for(i=0;i<itemList.length;i++){spanList=itemList[i].getElementsByTagName("SPAN");textEl=null;arrowEl=null;for(j=0;j<spanList.length;j++){if(hasClassName(spanList[j],"menuItemText"))textEl=spanList[j];if(hasClassName(spanList[j],"menuItemArrow"))arrowEl=spanList[j];}if(textEl!=null&&arrowEl!=null){textEl.style.paddingRight=(itemWidth-(textEl.offsetWidth+arrowEl.offsetWidth))+"px";if(browser.isOP)arrowEl.style.marginRight="0px";}}if(browser.isIE){w=itemList[0].offsetWidth;itemList[0].style.width=w+"px";dw=itemList[0].offsetWidth-w;w-=dw;itemList[0].style.width=w+"px";}if(browser.isIE){var iframeEl=document.createElement("IFRAME");iframeEl.frameBorder=0;iframeEl.src="javascript:;";iframeEl.style.display="none";iframeEl.style.position="absolute";iframeEl.style.filter="progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)";menu.iframeEl=menu.parentNode.insertBefore(iframeEl,menu);}menu.isInitialized=true;}function getContainerWith(node,tagName,className){while(node!=null){if(node.tagName!=null&&node.tagName==tagName&&hasClassName(node,className))return node;node=node.parentNode;}return node;}function hasClassName(el,name){var i,list;list=el.className.split(" ");for(i=0;i<list.length;i++)if(list[i]==name)return true;return false;}function removeClassName(el,name){var i,curList,newList;if(el.className==null)return;newList=new Array();curList=el.className.split(" ");for(i=0;i<curList.length;i++)if(curList[i]!=name)newList.push(curList[i]);el.className=newList.join(" ");}

function getPageOffsetLeft2(el){var x;x=el.offsetLeft;if(el.offsetParent!=null)x+=getPageOffsetLeft(el.offsetParent);return x;}
function getPageOffsetTop2(el){var y;y=el.offsetTop;if(el.offsetParent!=null)y+=getPageOffsetTop(el.offsetParent);return y;}

function getPageOffsetLeft(el){
  var x;
  x=el.offsetLeft;
  if(el.offsetParent!=null)x+=getPageOffsetLeft(el.offsetParent);
  //x+=20;
  return x;
}
function getPageOffsetTop(el){
  var y;
  y=el.offsetTop;
  if(el.offsetParent!=null)y+=getPageOffsetTop(el.offsetParent);
  return y;
}

var m_sub

function buttonmouseover2(event,menuId,sSub){var button;m_sub=sSub;if(activeButton==null){buttonClick2(event,menuId);return;}if(browser.isIE)button=window.event.srcElement;else button=event.currentTarget;if(activeButton!=null&&activeButton!=button)buttonClick2(event,menuId);}

function buttonClick2(event,menuId){var button;if(browser.isIE)button=window.event.srcElement;else button=event.currentTarget;button.blur();if(button.menu==null){button.menu=document.getElementById(menuId);if(button.menu.isInitialized==null)menuInit(button.menu);}if(button.onmouseout==null)button.onmouseout=buttonOrMenuMouseout;if(button==activeButton)return false;if(activeButton!=null)resetButton(activeButton);if(button!=activeButton){depressButton2(button);activeButton=button;}else activeButton=null;return false;}

function depressButton2(button){var x,y;button.className+=" menuButtonActive";if(button.onmouseout==null)button.onmouseout=buttonOrMenuMouseout;if(button.menu.onmouseout==null)button.menu.onmouseout=buttonOrMenuMouseout;x=getPageOffsetLeft(button)+button.offsetWidth-15;y=getPageOffsetTop(button);if(browser.isIE){x+=button.offsetParent.clientLeft;y+=button.offsetParent.clientTop;}button.menu.style.left=x+"px";button.menu.style.top=y+"px";button.menu.style.visibility="visible";if(button.menu.iframeEl!=null){button.menu.iframeEl.style.left=button.menu.style.left;button.menu.iframeEl.style.top=button.menu.style.top;button.menu.iframeEl.style.width=button.menu.offsetWidth+"px";button.menu.iframeEl.style.height=button.menu.offsetHeight+"px";button.menu.iframeEl.style.display="";}}

function getObject(obj) {
  var theObj;
  if(document.all) {
    if(typeof obj=="string") {
      return document.all(obj);
    } else {
      return obj.style;
    }
  }
  if(document.getElementById) {
    if(typeof obj=="string") {
      return document.getElementById(obj);
    } else {
      return obj.style;
    }
  }
  return null;
}

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
if( top.parent.frames.length > 0){
top.parent.location.href=self.location.href;
}
//Count Characters
function cntChar(fInput,fCntr,strTxt,maxlen) {
  var objInput=getObject(fInput);
  var objCnt=getObject(fCntr);
  var curLen=maxlen - objInput.value.length;
  if(curLen <= 0) {
    curLen=0;
    strTxt='<span class="fAlert"> '+strTxt+' </span>';
    objInput.value=objInput.value.substr(0,maxlen);
  }
  objCnt.innerHTML = strTxt.replace("{CHAR}",curLen);
}