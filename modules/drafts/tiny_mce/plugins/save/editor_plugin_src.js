tinyMCE.importPluginLanguagePack('save');

var TinyMCE_SavePlugin = {
    getInfo:function() {
        return{
            longname:'Save',
            author:'Brandon Williams',
            authorurl:'http://www.skyportal.net',
            infourl:'http://www.skyportal.net',
            version:'1.0'
        }
    },
    
    initInstance:function(inst) {
        inst.addShortcut('ctrl','s','lang_save_desc','mceSave')
    },
    
    getControlHTML:function(cn) {
        switch(cn) {
            case"save":
                return tinyMCE.getButtonHTML(cn,'lang_save_desc','{$pluginurl}/images/save.gif','mceSave');
                
            case"cancel":
                return tinyMCE.getButtonHTML(cn,'lang_cancel_desc','{$pluginurl}/images/cancel.gif','mceCancel');
                
            case"open":
                return tinyMCE.getButtonHTML(cn, 'lang_open_desc', '{$pluginurl}/images/open.gif', 'mceOpen');
        }
        
        return""
    },
    
    execCommand:function(editor_id,element,command,user_interface,value) {
        switch(command) {
            case"mceSave":
                return this._save(editor_id,element,command,user_interface,value);
                
            case"mceCancel":
                return this._cancel(editor_id,element,command,user_interface,value);
                
            case"mceOpen":
                return this._open(editor_id, element, command, user_interface, value);
        }
        
        return false
    },
    
    _save:function(editor_id, element, command, user_interface, value) {
        var inst=tinyMCE.getInstanceById(editor_id), os, h=tinyMCE.trim(inst.startContent);
        new Ajax.Request(
            'drafts.asp?cmd=3&newDraft='+tinyMCE.getContent(editor_id),
            {
                method:'post',
                onSuccess:function(transport) {
                    if(transport.responseText.match(/draftmodxxxdraftsaveddraftmodxxx/)) {
                        alert('Draft Saved!')
                    } else if(transport.responseText.match(/draftmodxxxtextboxemptydraftmodxxx/)) {
                        alert('Text Box Empty!')
                    };
                }
            }
        );
        
        tinyMCE.triggerNodeChange(false, true);
        
        return true;
    },
    
    _cancel:function(editor_id,element,command,user_interface,value) {
        var inst=tinyMCE.getInstanceById(editor_id),os,h=tinyMCE.trim(inst.startContent);
        if((os=tinyMCE.getParam("save_oncancelcallback"))) {
            if(eval(os+'(inst);'))
                return true
        }
        
        inst.setHTML(h);
        inst.undoRedo.undoLevels=[];
        inst.undoRedo.add({content:h});
        inst.undoRedo.undoIndex=0;
        inst.undoRedo.typingUndoIndex=-1;
        tinyMCE.triggerNodeChange(false,true);
        
        return true
    },
    
    _open:function(editor_id, element, command, user_interface, value) {
        var template=new Array();
        template['file']='../../../pop_drafts.asp?mode=1';
        template['width']=600;
        template['height']=600;
        tinyMCE.openWindow(template, {editor_id : editor_id, inline : "yes"});
        tinyMCE.triggerNodeChange(false, true);
        
        return true;
    },
    
    handleNodeChange:function(editor_id,node,undo_index,undo_levels,visual_aid,any_selection) {
        var inst;
        if(tinyMCE.getParam("fullscreen_is_enabled")) {
            tinyMCE.switchClass(editor_id+'_save','mceButtonDisabled');
            return true
        }
        if(tinyMCE.getParam("save_enablewhendirty")) {
            inst=tinyMCE.getInstanceById(editor_id);
            if(inst.isDirty()) {
                tinyMCE.switchClass(editor_id+'_save','mceButtonNormal');
                return true
            }
            tinyMCE.switchClass(editor_id+'_save','mceButtonDisabled')
        }
        
        return true
    }
};

tinyMCE.addPlugin("save",TinyMCE_SavePlugin);
