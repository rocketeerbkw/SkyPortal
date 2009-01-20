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
helpstat = false;
stprompt = true;
basic = false;
text = ""

function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 2) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 0) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}


function getActiveText(selectedtext) { 
	text = (document.all) ? document.selection.createRange().text : document.getSelection();
		if (selectedtext.createTextRange) {	
   			selectedtext.caretPos = document.selection.createRange().duplicate();	
  		}
		return true;
}

function AddText(NewCode) {
if (document.PostTopic.Message.createTextRange && document.PostTopic.Message.caretPos) {
var caretPos = document.PostTopic.Message.caretPos;
caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? NewCode + ' ' : NewCode;
}
else {
document.PostTopic.Message.value+=NewCode
}
setfocus();
}
function setfocus() {
document.PostTopic.Message.focus();
}

function email() {
	if (helpstat) {
		alert("Email Tag Turns an email address into a mailto hyperlink.\n\nUSE #1: [url]someone\@anywhere.com[/url] \nUSE #2: [url=\"someone\@anywhere.com\"]link text[/url]");
		}
	else if (basic) {
		AddTxt="[url]"+text+"[/url]";
		AddText(AddTxt);
		}
	else { 
		txt2=prompt("Text to be shown for the link. Leave blank if you want the url to be shown for the link.",""); 
		if (txt2!=null) {
			txt=prompt("URL for the link.","mailto:");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function showsize(size) {
	if (helpstat) {
		alert("Size Tag Sets the text size. Possible values are 1 to 6.\n1 being the smallest and 6 the largest.\n\nUSE: [size="+size+"]This is size "+size+" text[/size="+size+"]");
	} else if (basic) {
		AddTxt="[size="+size+"]"+text+"[/size="+size+"]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("Text to be size "+size,"Text"); 
		if (txt!=null) {             
			AddTxt="[size="+size+"]"+txt+"[/size="+size+"]";
			AddText(AddTxt);
		}        
	}
}

function bold() {
	if (helpstat) {
		alert("Bold Tag Makes the enlosed text bold.\n\nUSE: [b]This is some bold text[/b]");
	} else if (basic) {
		AddTxt="[b]"+text+"[/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BOLD.","Text");     
		if (txt!=null) {           
			AddTxt="[b]"+txt+"[/b]";
			AddText(AddTxt);
		}       
	}
}

function italicize() {
	if (helpstat) {
		alert("Italicize Tag Makes the enlosed text italicized.\n\nUSE: [i]This is some italicized text[/i]");
	} else if (basic) {
		AddTxt="[i]"+text+"[/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Text to be italicized","Text");     
		if (txt!=null) {           
			AddTxt="[i]"+txt+"[/i]";
			AddText(AddTxt);
		}	        
	}
}

function quote() {
	if (helpstat){
		alert("Quote tag Quotes the enclosed text to reference something specific that someone has posted.\n\nUSE: [quote]This is a quote[/quote]");
	} else if (basic) {
		AddTxt="[quote]"+text+"[/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Text to be quoted","Text");     
		if(txt!=null) {          
			AddTxt="[quote]"+txt+"[/quote]";
			AddText(AddTxt);
		}	        
	}
}

function showcolor(color) {
	if (helpstat) {
		alert("Color Tag Sets the text color. Any named color can be used.\n\nUSE: ["+color+"]This is some "+color+" text[/"+color+"]");
	} else if (basic) {
		AddTxt="["+color+"]"+text+"[/"+color+"]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("Text to be "+color,"Text");
		if(txt!=null) {
			AddTxt="["+color+"]"+txt+"[/"+color+"]";
			AddText(AddTxt);        
		} 
	}
}

function center() {
 	if (helpstat) {
		alert("Centered tag Centers the enclosed text.\n\nUSE: [center]This text is centered[/center]");
	} else if (basic) {
		AddTxt="[center]"+text+"[/center]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be centered","Text");     
		if (txt!=null) {          
			AddTxt="[center]"+txt+"[/center]";
			AddText(AddTxt);
		}	       
	}
}

function hyperlink() {
	if (helpstat) {
		alert("Hyperlink Tag \nTurns an url into a hyperlink.\n\nUSE: [url]http://www.anywhere.com[/url]\n\nUSE: [url=http://www.anywhere.com]link text[/url]");
	} else if (basic) {
		AddTxt="[url]"+text+"[/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("Text to be shown for the link.\nLeave blank if you want the url to be shown for the link.",""); 
		if (txt2!=null) {
			txt=prompt("URL for the link.","http://");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt+"[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url=\""+txt+"\"]"+txt2+"[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function image() {
	if (helpstat){
		alert("Image Tag Inserts an image into the post.\n\nUSE: [img]http://www.anywhere.com/image.gif[/img]");
	} else if (basic) {
		AddTxt="[img]"+text+"[/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("URL for graphic","http://");    
		if(txt!=null) {            
			AddTxt="[img]"+txt+"[/img]";
			AddText(AddTxt);
		}	
	}
}

function showcode() {
	if (helpstat) {
		alert("Code Tag Blockquotes the text you reference and preserves the formatting.\nUsefull for posting code.\n\nUSE: [code]This is formated text[/code]");
	} else if (basic) {
		AddTxt="[code]"+text+"[/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Enter code","");     
		if (txt!=null) {          
			AddTxt="[code]"+txt+"[/code]";
			AddText(AddTxt);
		}	       
	}
}

function list() {
	if (helpstat) {
		alert("List Tag Builds a bulleted, numbered, or alphabetical list.\n\nUSE: [list] [*]item1[/*] [*]item2[/*] [*]item3[/*] [/list]");
	} else if (basic) {
		AddTxt="[list]"+text+"[*]  [/*]"+text+"[*]  [/*]"+text+"[*]  [/*]"+text+"[/list]";
		AddText(AddTxt);
	} else {  
		type=prompt("Type of list Enter \'A\' for alphabetical, \'1\' for numbered, Leave blank for bulleted.","");               
		while ((type!="") && (type!="A") && (type!="a") && (type!="1") && (type!=null)) {
			type=prompt("ERROR! The only possible values for type of list are blank 'A' and '1'.","");               
		}
		if (type!=null) {
			if (type=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+type+"]";
			} 
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("List item Leave blank to end list",""); 
				if (txt!="") {             
					AddTxt+="[*]"+txt+"[/*]"; 
				}                   
			} 
			if (type=="") {
				AddTxt+="[/list] ";
			} else {
				AddTxt+="[/list="+type+"]";
			} 
			AddText(AddTxt); 
		}
	}
}

function underline() {
  	if (helpstat) {
		alert("Underline Tag Underlines the enclosed text.\n\nUSE: [u]This text is underlined[/u]");
	} else if (basic) {
		AddTxt="[u]"+text+"[/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be Underlined.","Text");     
		if (txt!=null) {           
			AddTxt="[u]"+txt+"[/u]";
			AddText(AddTxt);
		}	        
	}
}

function showfont(font) {
 	if (helpstat){
		alert("Font Tag Sets the font face for the enclosed text.\n\nUSE: [font="+font+"]The font of this text is "+font+"[/font]");
	} else if (basic) {
		AddTxt="[font="+font+"]"+text+"[/font="+font+"]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("Text to be in "+font,"Text");
		if (txt!=null) {             
			AddTxt="[font="+font+"]"+txt+"[/font="+font+"]";
			AddText(AddTxt);
		}        
	}  
}

function red() {
	if (helpstat) {
		alert("Red Tag Makes the enlosed text Red.\n\nUSE: [red]This is some red text[/red]");
	} else if (basic) {
		AddTxt="[red]"+text+"[/red]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made RED.","Text");     
		if (txt!=null) {           
			AddTxt="[red]"+txt+"[/red]";
			AddText(AddTxt);
		}       
	}
}

function blue() {
	if (helpstat) {
		alert("Blue Tag Makes the enlosed text Blue.\n\nUSE: [blue]This is some blue text[/blue]");
	} else if (basic) {
		AddTxt="[blue]"+text+"[/blue]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BLUE.","Text");     
		if (txt!=null) {           
			AddTxt="[blue]"+txt+"[/blue]";
			AddText(AddTxt);
		}       
	}
}

function pink() {
	if (helpstat) {
		alert("Pink Tag Makes the enlosed text Pink.\n\nUSE: [pink]This is some pink text[/pink]");
	} else if (basic) {
		AddTxt="[pink]"+text+"[/pink]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made PINK.","Text");     
		if (txt!=null) {           
			AddTxt="[pink]"+txt+"[/pink]";
			AddText(AddTxt);
		}       
	}
}

function brown() {
	if (helpstat) {
		alert("Brown Tag Makes the enlosed text Brown.\n\nUSE: [brown]This is some brown text[/brown]");
	} else if (basic) {
		AddTxt="[brown]"+text+"[/brown]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BROWN.","Text");     
		if (txt!=null) {           
			AddTxt="[brown]"+txt+"[/brown]";
			AddText(AddTxt);
		}       
	}
}

function black() {
	if (helpstat) {
		alert("Black Tag Makes the enlosed text Black.\n\nUSE: [black]This is some black text[/black]");
	} else if (basic) {
		AddTxt="[black]"+text+"[/black]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BLACK.","Text");     
		if (txt!=null) {           
			AddTxt="[black]"+txt+"[/black]";
			AddText(AddTxt);
		}       
	}
}

function orange() {
	if (helpstat) {
		alert("Orange Tag Makes the enlosed text Orange.\n\nUSE: [orange]This is some orange text[/orange]");
	} else if (basic) {
		AddTxt="[orange]"+text+"[/orange]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made ORANGE.","Text");     
		if (txt!=null) {           
			AddTxt="[orange]"+txt+"[/orange]";
			AddText(AddTxt);
		}       
	}
}

function violet() {
	if (helpstat) {
		alert("Violet Tag Makes the enlosed text Violet.\n\nUSE: [violet]This is some violet text[/violet]");
	} else if (basic) {
		AddTxt="[violet]"+text+"[/violet]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made VIOLET.","Text");     
		if (txt!=null) {           
			AddTxt="[violet]"+txt+"[/violet]";
			AddText(AddTxt);
		}       
	}
}

function yellow() {
	if (helpstat) {
		alert("Yellow Tag Makes the enlosed text Yellow.\n\nUSE: [yellow]This is some yellow text[/yellow]");
	} else if (basic) {
		AddTxt="[yellow]"+text+"[/yellow]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made YELLOW.","Text");     
		if (txt!=null) {           
			AddTxt="[yellow]"+txt+"[/yellow]";
			AddText(AddTxt);
		}       
	}
}

function green() {
	if (helpstat) {
		alert("Green Tag Makes the enlosed text Green.\n\nUSE: [green]This is some green text[/green]");
	} else if (basic) {
		AddTxt="[green]"+text+"[/green]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made GREEN.","Text");     
		if (txt!=null) {           
			AddTxt="[green]"+txt+"[/green]";
			AddText(AddTxt);
		}       
	}
}

function gold() {
	if (helpstat) {
		alert("Gold Tag Makes the enlosed text Gold.\n\nUSE: [gold]This is some gold text[/gold]");
	} else if (basic) {
		AddTxt="[gold]"+text+"[/gold]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made GOLD.","Text");     
		if (txt!=null) {           
			AddTxt="[gold]"+txt+"[/gold]";
			AddText(AddTxt);
		}       
	}
}

function white() {
	if (helpstat) {
		alert("White Tag Makes the enlosed text White.\n\nUSE: [white]This is some white text[/white]");
	} else if (basic) {
		AddTxt="[white]"+text+"[/white]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made WHITE.","Text");     
		if (txt!=null) {           
			AddTxt="[white]"+txt+"[/white]";
			AddText(AddTxt);
		}       
	}
}

function purple() {
	if (helpstat) {
		alert("Purple Tag Makes the enlosed text Purple.\n\nUSE: [purple]This is some purple text[/purple]");
	} else if (basic) {
		AddTxt="[purple]"+text+"[/purple]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made PURPLE.","Text");     
		if (txt!=null) {           
			AddTxt="[purple]"+txt+"[/purple]";
			AddText(AddTxt);
		}       
	}
}

function beige() {
	if (helpstat) {
		alert("Beige Tag Makes the enlosed text Beige.\n\nUSE: [beige]This is some Beige text[/beige]");
	} else if (basic) {
		AddTxt="[beige]"+text+"[/beige]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made BEIGE.","Text");     
		if (txt!=null) {           
			AddTxt="[beige]"+txt+"[/beige]";
			AddText(AddTxt);
		}       
	}
}

function teal() {
	if (helpstat) {
		alert("Teal Tag Makes the enlosed text Teal.\n\nUSE: [teal]This is some teal text[/teal]");
	} else if (basic) {
		AddTxt="[teal]"+text+"[/teal]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made TEAL.","Text");     
		if (txt!=null) {           
			AddTxt="[teal]"+txt+"[/teal]";
			AddText(AddTxt);
		}       
	}
}

function navy() {
	if (helpstat) {
		alert("Navy Tag Makes the enlosed text Navy.\n\nUSE: [navy]This is some navy text[/navy]");
	} else if (basic) {
		AddTxt="[navy]"+text+"[/navy]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made NAVY.","Text");     
		if (txt!=null) {           
			AddTxt="[navy]"+txt+"[/navy]";
			AddText(AddTxt);
		}       
	}
}

function maroon() {
	if (helpstat) {
		alert("Maroon Tag Makes the enlosed text Maroon.\n\nUSE: [maroon]This is some maroon text[/maroon]");
	} else if (basic) {
		AddTxt="[maroon]"+text+"[/maroon]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made MAROON.","Text");     
		if (txt!=null) {           
			AddTxt="[maroon]"+txt+"[/maroon]";
			AddText(AddTxt);
		}       
	}
}

function limegreen() {
	if (helpstat) {
		alert("Limegreen Tag Makes the enlosed text Limegreen.\n\nUSE: [limegreen]This is some limegreen text[/limegreen]");
	} else if (basic) {
		AddTxt="[limegreen]"+text+"[/limegreen]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be made LIMEGREEN.","Text");     
		if (txt!=null) {           
			AddTxt="[limegreen]"+txt+"[/limegreen]";
			AddText(AddTxt);
		}       
	}
}

function strike() {
	if (helpstat) {
		alert("strike Tag Makes the enlosed text striked.\n\nUSE: [s]This is some striked text[/]");
	} else if (basic) {
		AddTxt="[s]"+text+"[/s]";
		AddText(AddTxt);
	} else {   
		txt=prompt("Text to be striked","Text");     
		if (txt!=null) {           
			AddTxt="[s]"+txt+"[/s]";
			AddText(AddTxt);
		}	        
	}
}

function aleft() {
 	if (helpstat) {
		alert("Align left tag aligns the enclosed text left.\n\nUSE: [left]This text is aligned left[/left]");
	} else if (basic) {
		AddTxt="[left]"+text+"[/left]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be aligned left","Text");     
		if (txt!=null) {          
			AddTxt="[left]"+txt+"[/left]";
			AddText(AddTxt);
		}	       
	}
}

function aright() {
 	if (helpstat) {
		alert("Align right tag aligns the enclosed text right.\n\nUSE: [right]This text is aligned right[/right]");
	} else if (basic) {
		AddTxt="[right]"+text+"[/right]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be aligned right","Text");     
		if (txt!=null) {          
			AddTxt="[right]"+txt+"[/right]";
			AddText(AddTxt);
		}	       
	}
}

function pre() {
 	if (helpstat) {
		alert("Pre tag allows you to write freely.\n\nUSE: [pre]This   text   is   written   freely[/pre]");
	} else if (basic) {
		AddTxt="[pre]"+text+"[/pre]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be pre","Text");     
		if (txt!=null) {          
			AddTxt="[pre]"+txt+"[/pre]";
			AddText(AddTxt);
		}	       
	}
}

function marquee() {
 	if (helpstat) {
		alert("Marquee tag moves the enclosed text.\n\nUSE: [marquee]This text is moving[/marquee]");
	} else if (basic) {
		AddTxt="[marquee]"+text+"[/marquee]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be moved","Text");     
		if (txt!=null) {          
			AddTxt="[marquee]"+txt+"[/marquee]";
			AddText(AddTxt);
		}	       
	}
}

function sup() {
 	if (helpstat) {
		alert("Sup tag writes the enclosed text as sup.\n\nUSE: [sup]This text is sup[/sup]");
	} else if (basic) {
		AddTxt="[sup]"+text+"[/sup]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be written sup","Text");     
		if (txt!=null) {          
			AddTxt="[sup]"+txt+"[/sup]";
			AddText(AddTxt);
		}	       
	}
}

function sub() {
 	if (helpstat) {
		alert("Sub tag writes the enclosed text as sub.\n\nUSE: [sub]This text is sub[/sub]");
	} else if (basic) {
		AddTxt="[sub]"+text+"[/sub]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be written sub","Text");     
		if (txt!=null) {          
			AddTxt="[sub]"+txt+"[/sub]";
			AddText(AddTxt);
		}	       
	}
}

function tt() {
 	if (helpstat) {
		alert("Teletype tag writes the enclosed text as TeleType.\n\nUSE: [tt]This text is TeleType[/tt]");
	} else if (basic) {
		AddTxt="[tt]"+text+"[/tt]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be written as TeleType","Text");     
		if (txt!=null) {          
			AddTxt="[tt]"+txt+"[/tt]";
			AddText(AddTxt);
		}	       
	}
}





function hl() {
  	if (helpstat) {
		alert("Highlight Tag highlights the enclosed text yellow.\n\nUSE: [hl]This text is highlighted with yellow[/hl]");
	} else if (basic) {
		AddTxt="[hl]"+text+"[/hl]";
		AddText(AddTxt);
	} else {  
		txt=prompt("Text to be highlighted yellow.","Text");     
		if (txt!=null) {           
			AddTxt="[hl]"+txt+"[/hl]";
			AddText(AddTxt);
		}	        
	}
}


function hr() {
  	if (helpstat) {
		alert("Horizontal Rule Tag adds a horizontal rule.\n\nUSE: Horizontal Rule[hr]");
	} else if (basic) {
		AddTxt="[hr]";
		AddText(AddTxt);
	} else {  
		AddTxt="[hr]";
		AddText(AddTxt);
	}
}

function OpenPreview(){
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_portal.asp?cmd=6', 'preview_page', 'scrollbars=yes,width=750,height=450')	
}

function ltrim(s) {
	return s.replace( /^\s*/, "" );
}
function rtrim(s) {
	return s.replace( /\s*$/, "" );
}
function trim ( s ) {
	return rtrim(ltrim(s));
}