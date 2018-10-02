const Writable = require("stream").Writable;
const {
	RTFObj, 
	RTFDoc, 
	RTFGroup, 
	ParameterGroup, 
	DocTable, 
	ColourTable, 
	FontTable, 
	Font, 
	ListTable, 
	List, 
	ListLevel, 
	ListOverrideTable, 
	ListOverride, 
	Field, 
	Fldrslt, 
	Picture
} = require("./RTFGroups.js");

const win_1252 = ` !"#$%&'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_\`abcdefghijklmnopqrstuvqxyz{|}~ €�‚ƒ„…†‡ˆ‰Š‹Œ�Ž��‘’“”•–—˜™š›œ�žŸ ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ`

class LargeRTFSubunit extends Writable{
	constructor() {
		super({
			write(chunk, encoding, callback) {
				this.curInstruction = JSON.parse(chunk);
				this.followInstruction(this.curInstruction);	
				callback();
			}
		});
		this.curInstruction = {};
		this.output = {};
		this.curIndex = 0;
		this.defState = {
			font:0,
			fontsize:22,
			bold:false,
			italics:false,
			underline:false,
			strikethrough:false,
			smallcaps:false,
			subscript:false,
			superscript:false,
			foreground:false,
			background:false
		};
		this.doc = new RTFDoc;
		this.curGroup = this.doc;
		this.paraTypes = ["paragraph", "listitem"];
		this.textTypes = ["text", "listtext", "field", "fragment"];
	}
	followInstruction(instruction) {
		switch(instruction.type) {
			case "control":
				this.parseControl(instruction.value);
				break;
			case "text":
				if (this.curGroup.type !== "paragraph") {
					this.curGroup.contents.push(instruction.value);
				} else {
					this.newGroup("fragment");
					this.curGroup.contents.push(instruction.value);
				}		
				break;
			case "groupStart":
				this.newGroup("span");
				break;
			case "groupEnd":
				this.endGroup();
				break;
			case "ignorable":
				this.curGroup.attributes.ignorable = true;
				break;
			case "break":
				if (this.curGroup.type === "fragment") {this.endGroup();}
				break;
			case "documentEnd":
				while (this.curGroup !== this.doc) {this.endGroup();}
				this.output = this.doc.dumpContents();
				break;
		}
	}
	parseControl(instruction) {
		const numPos = instruction.search(/\d|\-/);
		let val = null;
		if (numPos !== -1) {
			val = parseFloat(instruction.substr(numPos).replace(/,/g,""));
			instruction = instruction.substr(0,numPos);
		}
		const command = "cmd$" + instruction;
		if (this[command]) {
			this[command](val);
		}
	}
	newGroup(type) {
		this.curGroup = new RTFGroup(this.curGroup, type);
		this.curGroup.style = this.curGroup.parent.style ? this.curGroup.parent.curstyle : this.defState;
	}
	endGroup() {
		this.curGroup.dumpContents();
		if (this.curGroup.parent) {
			this.curGroup = this.curGroup.parent;
		} else {
			this.curGroup = this.doc;
		}
	}

	/* Header */
	cmd$rtf(val) {
		this.doc.attributes.rtfversion = val;
	}
	cmd$ansi() {
		this.doc.attributes.charset = "ansi";
	}
	cmd$mac() {
		this.doc.attributes.charset = "mac";
	}
	cmd$pc() {
		this.doc.attributes.charset = "pc";
	}
	cmd$pca() {
		this.doc.attributes.charset = "pca";
	}
	cmd$ansicpg(val) {
		this.doc.attributes.ansipg = val;
	}
	cmd$fbidis() {
		this.doc.attributes.fbidis = true;
	}

	/* Default Fonts and Languages */
	cmd$fromtext() {
		this.doc.attributes.fromtext = true;
	}
	cmd$fromhtml(val) {
		this.doc.attributes.fromhtml = val;
	}
	cmd$deff(val) {
		this.doc.attributes.defaultfont = val;
	}
	cmd$adeff(val) {
		this.doc.attributes.defaultbidifont = val;
	}
	cmd$stshfdbch(val) {
		this.doc.attributes.defaulteastasian = val;
	}
	cmd$stshfloch(val) {
		this.doc.attributes.defaultascii = val;
	}
	cmd$stshfhich(val) {
		this.doc.attributes.defaulthighansi = val;
	}
	cmd$stshfbi(val) {
		this.doc.attributes.defaultbidi = val;
	}
	cmd$deflang(val) {
		this.doc.attributes.defaultlanguage = val;
	}
	cmd$deflangfe(val) {
		this.doc.attributes.defaultlanguageeastasia = val;
	}
	cmd$adeflang(val) {
		this.doc.attributes.defaultlanguagesouthasia = val;
	}

	/*Themes */
	cmd$themedata() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "themedata");
	}
	cmd$colorschememapping() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "colorschememapping");
	}
	cmd$flomajor() {
		this.doc.attributes.fmajor = "ascii";
	}
	cmd$fhimajor() {
		this.doc.attributes.fmajor = "default";
	}
	cmd$fdbmajor() {
		this.doc.attributes.fmajor = "eastasian";
	}
	cmd$fbimajor() {
		this.doc.attributes.fmajor = "complexscripts";
	}
	cmd$flominor() {
		this.doc.attributes.fminor = "ascii";
	}
	cmd$fhiminor() {
		this.doc.attributes.fminor = "default";
	}
	cmd$fdbminor() {
		this.doc.attributes.fminor = "eastasian";
	}
	cmd$fbiminor() {
		this.doc.attributes.fminor = "complexscripts";
	}

	/* Code Page */
	cmd$cpg(val) {
		this.curGroup.attributes.codepage = val;
	}

	/* File Table */
	cmd$filetbl() {
		this.curGroup = new FileTable(this.doc);
	}
	cmd$file() {
		this.curGroup = new File(this.doc);
	}
	cmd$fid(val) {
		this.curGroup.attributes.id = val;
	}
	cmd$frelative(val) {
		this.curGroup.attributes.relative = val;
	}
	cmd$fosnum(val) {
		this.curGroup.attributes.osnumber = val;
	}
	cmd$fvalidmac() {
		this.curGroup.attributes.filesystem = "mac";
	}
	cmd$fvaliddos() {
		this.curGroup.attributes.filesystem = "ms-dos";
	}
	cmd$fvalidntfs() {
		this.curGroup.attributes.filesystem = "ntfs";
	}
	cmd$fvalidhpfs() {
		this.curGroup.attributes.filesystem = "hpfs";
	}
	cmd$fnetwork() {
		this.curGroup.attributes.networkfilesystem = true;
	}
	cmd$fnonfilesys() {
		this.curGroup.attributes.nonfilesys = true;
	}

	/* Paragraphs */
	cmd$par() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			const prevStyle = this.curGroup.curstyle;
			this.endGroup()
			this.newGroup("paragraph");
			this.curGroup.style = prevStyle;
		} else {
			this.newGroup("paragraph");
		}	
	}
	cmd$pard() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.style = JSON.parse(JSON.stringify(this.defState));
		} else {
			this.newGroup("paragraph");
			this.curGroup.style = JSON.parse(JSON.stringify(this.defState));
		}
	}
	cmd$plain() {
		Object.keys(this.defState).forEach(key => {
			if (this.curGroup.style[key]) {
				this.curGroup.style[key] = this.defState[key];
			}	
		});
	}

	/* Alignment */
	cmd$qc() {
		this.curGroup.style.alignment = "center";
	}
	cmd$qj() {
		this.curGroup.style.alignment = "justified";
	}
	cmd$qr() {
		this.curGroup.style.alignment = "right";
	}
	cmd$ql() {
		this.curGroup.style.alignment = "left";
	}

	/* Text Direction */
	cmd$rtlch() {
		this.curGroup.style.direction = "rtl";
	}
	cmd$ltrch() {
		this.curGroup.style.direction = "ltr";
	}

	/* Character Stylings */
	cmd$i(val) {
		this.curGroup.style.italics = val !== 0;
	}
	cmd$b(val) {
		this.curGroup.style.bold = val !== 0;
	}
	cmd$strike(val) {
		this.curGroup.style.strikethrough = val !== 0;
	}
	cmd$scaps(val) {
		this.curGroup.style.smallcaps = val !== 0;
	}
	cmd$ul(val) {
		this.curGroup.style.underline = val !== 0;
	}
	cmd$ulnone(val) {
		this.curGroup.style.underline = false;
	}
	cmd$sub() {
		this.curGroup.style.subscript = true;
	}
	cmd$super() {
		this.curGroup.style.superscript = true;
	}
	cmd$nosupersub() {
		this.curGroup.style.subscript = false;
		this.curGroup.style.superscript = false;
	}
	cmd$cf(val) {
		this.curGroup.style.foreground = this.doc.colourTable[val - 1];
	}
	cmd$cb(val) {
		this.curGroup.style.background = this.doc.colourTable[val - 1];
	}

	/* Lists */
	cmd$ilvl(val) {
		this.curGroup.style.ilvl = val;
		this.curGroup.type = "listitem";
	}
	cmd$listtext(val) {
		this.curGroup.type = "listtext";
	}

	/* Special Characters */
	cmd$emdash() {
		this.curGroup.contents.push("—");
	}
	cmd$endash() {
		this.curGroup.contents.push("–");
	}
	cmd$tab() {
		this.curGroup.contents.push("\t");
	}
	cmd$line() {
		this.curGroup.contents.push("\n");
	}
	cmd$hrule() {
		this.curGroup.contents.push({type:"hr"});
	}

	/* Unicode Characters */
	cmd$uc(val) {
		if (this.curGroup.type !== "span") {
			this.curGroup.uc = val
		} else {
			this.curGroup.parent.uc = val
		}
	}
	cmd$u(val) {
		if (!this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.contents.push(String.fromCharCode(parseInt(val)));			
		} else {
			this.newGroup("fragment");
			this.curGroup.contents.push(String.fromCharCode(parseInt(val)));
			this.endGroup();
		}
		if(this.curGroup.uc) {
			this.curIndex += this.curGroup.uc;
		} else if (this.curGroup.parent.uc) {
			this.curIndex += this.curGroup.parent.uc;
		} else {
			this.curIndex += 1;
		}
	}

	/* Ascii Extended Characters (Windows 1252) */
	cmd$hex(val) {
		if (!this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.contents.push(win_1252.charAt(parseInt(val, 16) - 32));		
		} else {
			this.newGroup("fragment");
			this.curGroup.contents.push(win_1252.charAt(parseInt(val, 16) - 32));
			this.endGroup();
		}
	}

	/* Fonts */
	cmd$f(val) {
		if (this.curGroup.parent instanceof RTFObj) {
			this.curGroup.style.font = val;
		} else if (this.curGroup.parent instanceof FontTable) {
			this.curGroup = new Font(this.curGroup.parent);
			this.curGroup.attributes.font = val;
		}	
	}
	cmd$fs(val) {
		this.curGroup.style.fontsize = val;
	}

	/* Fields */
	cmd$field() {
		this.curGroup = new Field(this.curGroup.parent);
	}
	cmd$fldinst() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "fieldInst");
	}
	cmd$fldrslt() {
		this.curGroup = new Fldrslt(this.curGroup.parent);
	}

	/* Pictures */
	cmd$shppict() {
		this.curGroup.type = "shppict";
	}
	cmd$pict() {
		this.curGroup = new Picture(this.curGroup.parent);
	}
	cmd$nisusfilename() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "nisusfilename");
	}
	cmd$nonshppict() {
		this.curGroup.attributes.nonshppict = false;
	}
	cmd$emfblip() {
		this.curGroup.attributes.source = "EMF";
	}
	cmd$pngblip() {
		this.curGroup.attributes.source = "PNG";
	}
	cmd$jpegblip() {
		this.curGroup.attributes.source = "JPEG";
	}
	cmd$macpict() {
		this.curGroup.attributes.source = "QUICKDRAW";
	}
	cmd$pmmetafile(val) {
		this.curGroup.attributes.source = "OS/2 METAFILE";
		this.curGroup.attributes.sourcetype = val;
	}
	cmd$wmetafile(val) {
		this.curGroup.attributes.source = "WINDOWS METAFILE";
		this.curGroup.attributes.mappingmode = val;
	}
	cmd$dibitmap(val) {
		this.curGroup.attributes.source = "WINDOWS DI BITMAP";
		this.curGroup.attributes.sourcetype = val;
	}
	cmd$wbitmap(val) {
		this.curGroup.attributes.source = "WINDOWS DD BITMAP";
		this.curGroup.attributes.sourcetype = val;
	}
	cmd$wbmbitspixel(val) {
		this.curGroup.attributes.bitspixel = val;
	}
	cmd$wbmplanes(val) {
		this.curGroup.attributes.planes = val;
	}
	cmd$wbmwidthbytes(val) {
		this.curGroup.attributes.widthbytes = val;
	}
	cmd$picw(val) {
		this.curGroup.style.width = val;
	}
	cmd$pich(val) {
		this.curGroup.style.height = val;
	}
	cmd$picwgoal(val) {
		this.curGroup.style.widthgoal = val;
	}
	cmd$pichgoal(val) {
		this.curGroup.style.heightgoal = val;
	}
	cmd$picscalex(val) {
		this.curGroup.style.scalex = val;
	}
	cmd$picscaley(val) {
		this.curGroup.style.scaley = val;
	}
	cmd$picscaled() {
		this.curGroup.style.scaled = true;
	}
	cmd$piccropt(val) {
		this.curGroup.style.croptop = val;
	}
	cmd$piccropb(val) {
		this.curGroup.style.cropbottom = val;
	}
	cmd$piccropl(val) {
		this.curGroup.style.cropleft = val;
	}
	cmd$piccropr(val) {
		this.curGroup.style.cropright = val;
	}
	cmd$picprop(val) {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "prop");
	}
	cmd$defshp() {
		this.curGroup.style.shape = true;
	}
	cmd$picbmp() {
		this.curGroup.attributes.bitmap = true;
	}
	cmd$picbpp(val) {
		this.curGroup.attributes.bpp = val;
	}
	cmd$bin(val) {
		this.curGroup.attributes.binary = val;
	}
	cmd$blipupi(val) {
		this.curGroup.attributes.upi = val;
	}
	cmd$blipuid() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "uid");
	}
	cmd$bliptag(val) {
		this.curGroup.attributes.tag = val;
	}


	/* Font Table */
	cmd$fonttbl() {
		this.curGroup = new FontTable(this.doc);
	}
	cmd$fcharset(val) {
		this.curGroup.attributes.charset = val;
	}
	cmd$fprq(val) {
		this.curGroup.attributes.pitch = val;
	}
	cmd$fbias(val) {
		this.curGroup.attributes.bias = val;
	}
	cmd$falt() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "alternate");
	}
	cmd$panose() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "panose");
	}
	cmd$fname() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "taggedname");
	}
	cmd$fnil() {
		this.curGroup.attributes.family = "nil";
	}
	cmd$froman() {
		this.curGroup.attributes.family = "roman";
	}
	cmd$fswiss() {
		this.curGroup.attributes.family = "swiss";
	}
	cmd$fmodern() {
		this.curGroup.attributes.family = "modern";
	}
	cmd$fscript() {
		this.curGroup.attributes.family = "script";
	}
	cmd$fdecor() {
		this.curGroup.attributes.family = "decor";
	}
	cmd$ftech() {
		this.curGroup.attributes.family = "tech";
	}
	cmd$fbidi() {
		this.curGroup.attributes.family = "bidi";
	}
	cmd$ftnil() {
		this.curGroup.attributes.type = "nil";
	}
	cmd$fttruetype() {
		this.curGroup.attributes.type = "truetype";
	}

	/* Colour Table */
	cmd$colortbl() {
		this.curGroup = new ColourTable(this.doc);
	}
	cmd$red(val) {
		if (this.curGroup instanceof ColourTable) {
			this.curGroup.addColour("red", val);
		}
	}
	cmd$blue(val) {
		if (this.curGroup instanceof ColourTable) {
			this.curGroup.addColour("blue", val);
		}
	}
	cmd$green(val) {
		if (this.curGroup instanceof ColourTable) {
			this.curGroup.addColour("green", val);
		}
	}

	/* List Table */
	cmd$listtable() {
		this.curGroup = new ListTable(this.doc);
	}

	cmd$list() {
		this.curGroup = new List(this.curGroup.parent);
	}
	cmd$listid(val) {
		this.curGroup.id = val;
	}
	cmd$listtemplateid(val) {
		this.curGroup.templateid = val;
	}
	cmd$listsimple(val) {
		this.curGroup.attributes.simple = val;
	}
	cmd$listhybrid(val) {
		this.curGroup.attributes.hybrid = true;
	}
	cmd$listname() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "listname");
	}
	cmd$liststyleid(val) {
		this.curGroup.attributes.styleid = val;
	}
	cmd$liststylename(val) {
		this.curGroup.attributes.stylename = val;
	}
	cmd$liststartat(val) {
		this.curGroup.attributes.startat = val;
	}
	cmd$lvltentative() {
		this.curGroup.attributes.lvltentative = true;
	}

	cmd$listlevel() {
		this.curGroup = new ListLevel(this.curGroup.parent);
	}
	cmd$levelstartat(val) {
		this.curGroup.attributes.startat = val;
	}
	cmd$levelnfc(val) {
		this.curGroup.attributes.nfc = val;
	}
	cmd$levelnfcn(val) {
		this.curGroup.attributes.nfcn = val;
	}
	cmd$leveljc(val) {
		this.curGroup.attributes.jc = val;
	}
	cmd$leveljcn(val) {
		this.curGroup.attributes.jcn = val;
	}
	cmd$leveltext() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "leveltext");
	}
	cmd$levelnumbers(val) {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "levelnumbers");
	}
	cmd$levelfollow(val) {
		this.curGroup.attributes.follow = val;
	}
	cmd$levellegal(val) {
		this.curGroup.attributes.legal = val;
	}
	cmd$levelnorestart(val) {
		this.curGroup.attributes.norestart = val;
	}
	cmd$levelold(val) {
		this.curGroup.attributes.old = val;
	}
	cmd$levelprev(val) {
		this.curGroup.attributes.prev = val;
	}
	cmd$levelprevspace(val) {
		this.curGroup.attributes.prevspace = val;
	}
	cmd$levelindent(val) {
		this.curGroup.attributes.indent = val;
	}
	cmd$levelspace(val) {
		this.curGroup.attributes.space = val;
	}

	/* List Override Table */
	cmd$listoverridetable() {
		this.curGroup = new ListOverrideTable(this.doc);
	}
	cmd$listoverride() {
		this.curGroup = new ListOverride(this.curGroup.parent);
	}
	cmd$ls(val) {
		if (this.curGroup instanceof ListOverride) {
	      	this.curGroup.ls = val;
	    } else {
	      	this.curGroup.style.ls = val;
	    }
	}
	cmd$listoverridecount(val) {
		this.curGroup.attributes.overridecount = val;
	}
	cmd$listoverridestartat() {
		this.curGroup.attributes.overridestartat = true;
	}
	cmd$listoverrideformat(val) {
		this.curGroup.attributes.overrideformat = val;
	}
}

module.exports = LargeRTFSubunit;