tinyMCE.importPluginLanguagePack('emotions');var TinyMCE_EmotionsPlugin={getInfo:function(){return{longname:'Emotions',author:'Moxiecode Systems AB',authorurl:'http://tinymce.moxiecode.com',infourl:'http://wiki.moxiecode.com/index.php/TinyMCE:Plugins/emotions',version:tinyMCE.majorVersion+"."+tinyMCE.minorVersion}},getControlHTML:function(cn){switch(cn){case"emotions":return tinyMCE.getButtonHTML(cn,'lang_emotions_desc','{$pluginurl}/images/emotions.gif','mceEmotion')}return""},execCommand:function(editor_id,element,command,user_interface,value){switch(command){case"mceEmotion":var template=new Array();template['file']='../../plugins/emotions/emotions.htm';template['width']=225;template['height']=290;template['width']+=tinyMCE.getLang('lang_emotions_delta_width',0);template['height']+=tinyMCE.getLang('lang_emotions_delta_height',0);tinyMCE.openWindow(template,{editor_id:editor_id,inline:"yes"});return true}return false}};tinyMCE.addPlugin('emotions',TinyMCE_EmotionsPlugin);