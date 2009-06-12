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

//::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::
//::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::
function ajax_UpdateBlock(t,d,m,c,s,sh,co){
  new Ajax.Updater({ success: d, failure: d }, t, {
  method:'get',
  parameters: {
    mode: m,
	cid: c,
	sid: s,
	show: sh,
	col: co
    }
  });
}

function openJsLayer(ob,w,h) {
  Dialog.alert($(ob).innerHTML, {windowParameters: {width:w, height:h}, okLabel: "cancel"});
}
//::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::<>::

function ajaxSpDynTpl(t,d){
  var p = "files/templates/"+t;
  new Ajax.Request('rss_ajax.asp',{
  method:'get',
  onSuccess: function(transport){
    var response = transport.responseText || "no response text";
	
    $(d).innerHTML=response;
    },
  onFailure: function(){ 
    $(d).innerHTML='Something went wrong...'
    }
  });
}

// -->