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
	FileTable,
	File,
	Default,
	Stylesheet,
	Style,
	StyleRestrictions,
	ListTable, 
	List, 
	ListLevel, 
	ListOverrideTable, 
	ListOverride, 
	ParagraphGroupTable,
	ParagraphGroup, 
	RevisionTable,
	RSIDTable,
	ProtectedUsersTable,
	UserProperty,
	XMLNamespaceTable,
	XMLNamespace,
	Field, 
	Fldrslt, 
	Picture,
	DateGroup,
	NonGroup
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
		this.skip = 0;
		this.defCharState = {
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
		this.defParState = {
			alignment:"left",
			direction: 'ltr'
		}
		this.doc = new RTFDoc;
		this.curGroup = this.doc;
		this.paraTypes = ["paragraph", "listitem"];
		this.textTypes = ["text", "listtext", "field", "fragment"];
	}
	followInstruction(instruction) {
		if (this.skip > 0) {
			this.skip--;
			return;
		}
		switch(instruction.type) {
			case "control":
				this.parseControl(instruction.value);
				break;
			case "text":
				if (!this.paraTypes.includes(this.curGroup.type)) {
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
			case "listBreak":
				if (this.curGroup.listType) {this.curGroup.flush();}
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
		if (this.curGroup.parent instanceof Stylesheet && !(this.curGroup instanceof Style)) {
			this.curGroup = new Style(this.curGroup.parent, instruction);
		} else {
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
	}
	newGroup(type) {
		this.curGroup = new RTFGroup(this.curGroup, type);
		this.curGroup.style = this.curGroup.parent.style ? this.curGroup.parent.curstyle : this.defCharState;
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

	/* Colour Table */
	cmd$colortbl() {
		this.curGroup = new ColourTable(this.doc);
	}
	cmd$red(val) {
		this.curGroup.red = val
	}
	cmd$blue(val) {
		this.curGroup.blue = val
	}
	cmd$green(val) {
		this.curGroup.green = val	
	}
	cmd$ctint(val) {
		this.curGroup.attributes.tint = val;
	}
	cmd$cshade(val) {
		this.curGroup.attributes.shade = val;
	}
	cmd$cmaindarkone() {
		this.curGroup.attributes.themecolour = "maindarkone";
	}
	cmd$cmaindarktwo() {
		this.curGroup.attributes.themecolour = "maindarktwo";
	}
	cmd$cmainlightone() {
		this.curGroup.attributes.themecolour = "mainlightone";
	}
	cmd$cmainlighttwo() {
		this.curGroup.attributes.themecolour = "mainlighttwo";
	}
	cmd$caccentone() {
		this.curGroup.attributes.themecolour = "accentone";
	}
	cmd$caccenttwo() {
		this.curGroup.attributes.themecolour = "accenttwo";
	}
	cmd$caccentthree() {
		this.curGroup.attributes.themecolour = "accentthree";
	}
	cmd$caccentfour() {
		this.curGroup.attributes.themecolour = "accentfour";
	}
	cmd$caccentfive() {
		this.curGroup.attributes.themecolour = "accentfive";
	}
	cmd$caccentsix() {
		this.curGroup.attributes.themecolour = "accentsix";
	}
	cmd$chyperlink() {
		this.curGroup.attributes.themecolour = "hyperlink";
	}
	cmd$cfollowedhyperlink() {
		this.curGroup.attributes.themecolour = "followedhyperlink";
	}
	cmd$cbackgroundone() {
		this.curGroup.attributes.themecolour = "backgroundone";
	}
	cmd$cbackgroundtwo() {
		this.curGroup.attributes.themecolour = "backgroundtwo";
	}
	cmd$ctextone() {
		this.curGroup.attributes.themecolour = "textone";
	}
	cmd$ctexttwo() {
		this.curGroup.attributes.themecolour = "texttwo";
	}

	/* Defaults */
	cmd$defchp() {
		this.curGroup = new Default(this.doc, this.defCharStyle, "character");
	}
	cmd$defpap() {
		this.curGroup = new Default(this.doc, this.defParStyle, "paragraph");
	}

	/* Stylesheet */
	cmd$stylesheet() {
		this.curGroup = new Stylesheet(this.doc);
	}
	cmd$tsrowd() {
		this.curGroup.attributes.tsrowd = true;
	}
	cmd$additive() {
		this.curGroup.attributes.additive = true;
	}
	cmd$sbasedon(val) {
		this.curGroup.attributes.basedon = val;
	}
	cmd$snext(val) {
		this.curGroup.attributes.next = val;
	}
	cmd$sautoupd() {
		this.curGroup.attributes.autoupdate = true;
	}
	cmd$shidden() {
		this.curGroup.attributes.hidden = true;
	}
	cmd$slink(val) {
		this.curGroup.attributes.link = true;
	}
	cmd$slocked() {
		this.curGroup.attributes.locked = true;
	}
	cmd$spersonal() {
		this.curGroup.attributes.emailstyle = "personal";
	}
	cmd$scompose() {
		this.curGroup.attributes.emailstyle = "compose";
	}
	cmd$reply() {
		this.curGroup.attributes.emailstyle = "reply";
	}
	cmd$styrsid(val) {
		this.curGroup.attributes.rsid = val;
	}
	cmd$ssemihidden(val) {
		if (val === null) {val = 0}
		this.curGroup.attributes.semihidden = val;
	}
	cmd$keycode() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "keycode");
	}
	cmd$alt() {
		this.curGroup.contents.push("ALT ");
	}
	cmd$shift() {
		this.curGroup.contents.push("SHIFT ");
	}
	cmd$ctrl() {
		this.curGroup.contents.push("CTRL ");
	}
	cmd$fn(val) {
		this.curGroup.contents.push("FN" + val + " ");
	}
	cmd$sqformat() {
		this.curGroup.attributes.primary = true;
	}
	cmd$spriority(val) {
		this.curGroup.attributes.priority = val;
	}
	cmd$sunhideused(val) {
		this.curGroup.attributes.unhideused = val;
	}

	cmd$s(val) {
		this.curGroup.attributes.styledesignation = "s" + val;
	}
	cmd$cs(val) {
		this.curGroup.attributes.styledesignation = "cs" + val;
	}
	cmd$ds(val) {
		this.curGroup.attributes.styledesignation = "ds" + val;
	}
	cmd$ts(val) {
		this.curGroup.attributes.styledesignation = "ts" + val;
	}

	cmd$noqfpromote() {
		this.doc.attributes.noqfpromote = true;
	}

	/* Table Styles */
	cmd$tscellwidth(val) {
		this.curGroup.style.cellwidth = val;
	}
	cmd$tscellwidthfts(val) {
		this.curGroup.style.cellwidthfts = val;
	}
	cmd$tscellpaddt(val) {
		this.curGroup.style.toppadding = val;
	}
	cmd$tscellpaddl(val) {
		this.curGroup.style.leftpadding = val;
	}
	cmd$tscellpaddr(val) {
		this.curGroup.style.rightpadding = val;
	}
	cmd$tscellpaddb(val) {
		this.curGroup.style.bottompadding = val;
	}
	cmd$tscellpaddft(val) {
		this.curGroup.style.toppaddingunits = val;
	}
	cmd$tscellpaddfl(val) {
		this.curGroup.style.leftpaddingunits = val;
	}
	cmd$tscellpaddfr(val) {
		this.curGroup.style.rightpaddingunits = val;
	}
	cmd$tscellpaddfb(val) {
		this.curGroup.style.bottompaddingunits = val;
	}
	cmd$tsvertalt() {
		this.curGroup.style.cellalignment = "top";
	}
	cmd$tsvertalc() {
		this.curGroup.style.cellalignment = "center";
	}
	cmd$tsvertalb() {
		this.curGroup.style.cellalignment = "bottom";
	}
	cmd$tsnowrap() {
		this.curGroup.style.nowrap = true;
	}
	cmd$tscellcfpat(val) {
		this.curGroup.style.foregroundshading = val;
	}
	cmd$tscellcbpat(val) {
		this.curGroup.style.backgroundshading = val;
	}
	cmd$tscellpct(val) {
		this.curGroup.style.shadingpercentage = val;
	}
	cmd$tsbgbdiag() {
		this.curGroup.style.shadingpattern = "backwardsdiagonal";
	}
	cmd$tsbgfdiag() {
		this.curGroup.style.shadingpattern = "forwardsdiagonal";
	}
	cmd$tsbgdkbdiag() {
		this.curGroup.style.shadingpattern = "darkbackwardsdiagonal";
	}
	cmd$tsbgdkfdiag() {
		this.curGroup.style.shadingpattern = "darkforwardsdiagonal";
	}
	cmd$tsbgcross() {
		this.curGroup.style.shadingpattern = "cross";
	}
	cmd$tsbgdcross() {
		this.curGroup.style.shadingpattern = "diagonalcross";
	}
	cmd$tsbgdkcross() {
		this.curGroup.style.shadingpattern = "darkcross";
	}
	cmd$tsbgdkdcross() {
		this.curGroup.style.shadingpattern = "darkdiagonalcross";
	}
	cmd$tsbghoriz() {
		this.curGroup.style.shadingpattern = "horizontal";
	}
	cmd$tsbgvert() {
		this.curGroup.style.shadingpattern = "vertical";
	}
	cmd$tsbgdkhor() {
		this.curGroup.style.shadingpattern = "darkhorizontal";
	}
	cmd$tsbgdkvert() {
		this.curGroup.style.shadingpattern = "darkvertical";
	}
	cmd$tsbrdrt() {
		theis.curGroup.style.cellborder = "top";
	}
	cmd$tsbrdrb() {
		theis.curGroup.style.cellborder = "bottom";
	}
	cmd$tsbrdrl() {
		theis.curGroup.style.cellborder = "left";
	}
	cmd$tsbrdrr() {
		theis.curGroup.style.cellborder = "right";
	}
	cmd$tsbrdrh() {
		theis.curGroup.style.cellborder = "horizontal";
	}
	cmd$tsbrdrv() {
		theis.curGroup.style.cellborder = "vertical";
	}
	cmd$tsbrdrdgl() {
		theis.curGroup.style.cellborder = "diagonalullr";
	}
	cmd$tsbrdrdgr() {
		theis.curGroup.style.cellborder = "diagonalllur";
	}
	cmd$tscbandsh(val) {
		theis.curGroup.style.rowbandcount = val;
	}
	cmd$tscbandsv(val) {
		theis.curGroup.style.cellbandcount = val;
	}

	/* Style Restrictions */
	cmd$latentstyles() {
		this.curGroup = new StyleRestrictions(this.doc);
	}
	cmd$lsdstimax(val) {
		this.curGroup.attributes.dstimax = val;
	}
	cmd$lsdlockeddef(val) {
		this.curGroup.attributes.lockeddef = val;
	}
	cmd$lsdlockedexcept() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "lockedexceptions");
	}
	cmd$lsdsemihiddendef(val) {
		this.curGroup.attributes.ssemihiddendefault = val;
	}
	cmd$lsdunhideuseddef(val) {
		this.curGroup.attributes.sunhideuseddefault = val;
	}
	cmd$lsdqformatdef(val) {
		this.curGroup.attributes.sqformatdefault= val;
	}
	cmd$lsdprioritydef(val) {
		this.curGroup.attributes.sprioritydefault = val;
	}
	cmd$lsdpriority(val) {
		this.curGroup.attributes.sprioritylatentdefault = val;
	}
	cmd$lsdunhideused(val) {
		this.curGroup.attributes.sunhideusedlatentdefault = val;
	}
	cmd$lsdsemihidden(val) {
		this.curGroup.attributes.ssemihiddenlatentdefault = val;
	}
	cmd$lsdqformat(val) {
		this.curGroup.attributes.sqformatlatentdefault = val;
	}
	cmd$lsdlocked(val) {
		this.curGroup.attributes.slockedlatentdefault = val;
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
	cmd$lfolevel() {
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

	/* Paragraph Group Properties */
	cmd$pgptbl() {
		this.curGroup = new ParagraphGroupTable(this.doc);
	}
	cmd$pgp() {
		this.curGroup = new ParagraphGroup(this.curGroup.parent);
	}
	cmd$ipgp(val) {
		this.curGroup.attributes.id = val;
	}

	/* Revision Marks */
	cmd$revtbl() {
		this.curGroup = new RevisionTable(this.doc);
	}

	/* RSID */
	cmd$rsidtbl() {
		this.curGroup = new RSIDTable(this.doc);
	}
	cmd$rsid(val) {
		this.curGroup.table.push(val);
	}

	cmd$insrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "insert";
	}
	cmd$rsidroot(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "root";
	}
	cmd$delrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "delete";
	}
	cmd$charrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "characterformat";
	}
	cmd$sectrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "sectionformat";
	}
	cmd$pararsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "paragraphformat";
	}
	cmd$tblrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidtype = "tableformat";
	}

	/* Old Properties */
	cmd$oldcProps() {
		this.curGroup.type = "oldcprop";
	}
	cmd$oldpProps() {
		this.curGroup.type = "oldpprop";
	}
	cmd$oldtProps() {
		this.curGroup.type = "oldtprop";
	}
	cmd$oldsProps() {
		this.curGroup.type = "oldsprop";
	}

	/* User Protection Information */
	cmd$protusertbl() {
		this.curGroup = new ProtectedUsersTable(this.doc);
	}

	/* Generator */
	cmd$generator() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "generator");
	}

	/* Information */
	cmd$info() {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}
	cmd$title() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "title");
	}
	cmd$subject() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "subject");
	}
	cmd$author() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "author");
	}
	cmd$manager() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "manager");
	}
	cmd$company() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "company");
	}
	cmd$operator() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "operator");
	}
	cmd$category() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "category");
	}
	cmd$keywords() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "keywords");
	}
	cmd$comment() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "comment");
	}
	cmd$version(val) {
		this.doc.attributes.version = val;
	}
	cmd$title() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "doccomment");
	}
	cmd$hlinkbase() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "hlinkbase");
	}

	cmd$userprops() {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}
	cmd$propname(val) {
		this.curGroup.parent = new UserProperty(this.doc.attributes);
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertyname");
	}
	cmd$proptype(val) {
		this.curGroup.propertyType = val;
	}
	cmd$staticval() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertyvalue");
	}
	cmd$linkval() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertylink");
	}

	cmd$vern(val) {
		this.doc.attributes.initialversion = val;
	}
	cmd$creatim() {
		this.curGroup = new DateGroup(this.doc.attributes, "createtime");
	}
	cmd$revtim() {
		this.curGroup = new DateGroup(this.doc.attributes, "revisiontime");
	}
	cmd$printtim() {
		this.curGroup = new DateGroup(this.doc.attributes, "lastprinttim");
	}
	cmd$buptim() {
		this.curGroup = new DateGroup(this.doc.attributes, "backuptime");
	}
	cmd$edmins(val) {
		this.doc.attributes.editingminutes = val;
	}
	cmd$nofpages(val) {
		this.doc.attributes.pages = val;
	}
	cmd$nofwords(val) {
		this.doc.attributes.words = val;
	}
	cmd$nofchars(val) {
		this.doc.attributes.chars = val;
	}
	cmd$nofcharsws(val) {
		this.doc.attributes.charsnospaces = val;
	}
	cmd$id(val) {
		this.doc.attributes.id = val;
	}

	cmd$yr(val) {
		this.curGroup.year = val;
	}
	cmd$mo(val) {
		this.curGroup.month = val;
	}
	cmd$dy(val) {
		this.curGroup.day = val;
	}
	cmd$hr(val) {
		this.curGroup.hour = val;
	}
	cmd$min(val) {
		this.curGroup.minute = val;
	}
	cmd$sec(val) {
		this.curGroup.second = val;
	}

	/* Read-Only Password Protection */
	cmd$passwordhash() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "passwordhash");
	}

	/* XML Namespace Table */
	cmd$xmlnstbl() {
		this.curGroup = new XMLNamespaceTable(this.doc);
	}
	cmd$xmlns(val) {
		this.curGroup = new XMLNamespace(this.curGroup.parent, val);
	}

	/* Document Formatting Properties */
	cmd$deftab(val) {
		this.doc.style.defaulttab = val;
	}
	cmd$hyphhotz(val) {
		this.doc.style.hyphenhotzone = val;
	}
	cmd$hyphconsec(val) {
		this.doc.style.hyphenconsecutive = val;
	}
	cmd$hyphcaps(val) {
		this.doc.style.hyphencaps = val !== 0;
	}
	cmd$hyphauto(val) {
		this.doc.style.hyphenauto = val !== 0;
	}
	cmd$linestart(val) {
		this.doc.style.linestart = val;
	}
	cmd$fracwidth(val) {
		this.doc.style.fractionalwidths = true;
	}
	cmd$nextfile() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "nextfile");
	}
	cmd$template() {
		this.curGroup = new ParameterGroup(this.doc.style, "template");
	}
	cmd$makebackup() {
		this.doc.attributes.makebackup = true;
	}
	cmd$muser() {
		this.doc.attributes.compatability = true;
	}
	cmd$defformat() {
		this.doc.attributes.defformat = true;
	}
	cmd$psover() {
		this.doc.attributes.psover = true;
	}
	cmd$doctemp() {
		this.doc.attributes.boilerplate = true;
	}
	cmd$windowcaption() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "caption");
	}
	cmd$doctype(val) {
		this.doc.attributes.doctype = val;
	}
	cmd$ilfomacatclnup(val) {
		this.doc.attributes.ilfomacatclnup = val;
	}
	cmd$horzdoc() {
		this.doc.style.rendering = "horizontal";
	}
	cmd$horzdoc() {
		this.doc.style.rendering = "vertical";
	}
	cmd$jcompress() {
		this.doc.style.justification = "compressing";
	}
	cmd$jexpand() {
		this.doc.style.justification = "expanding";
	}
	cmd$lnongrid() {
		this.doc.style.lineongrid = true;
	}
	cmd$grfdocevents(val) {
		this.doc.attributes.grfdocevents = val;
	}
	cmd$themelang(val) {
		this.doc.style.themelanguage = val;
	}
	cmd$themelangfe(val) {
		this.doc.style.themelanguagefe = val;
	}
	cmd$themelangcs(val) {
		this.doc.style.themelanguagecs = val;
	}
	cmd$relyonvml(val) {
		this.doc.attributes.vml = val;
	}
	cmd$validatexml(val) {
		this.doc.attributes.validatexml = val;
	}
	cmd$xform() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "xform");
	}
	cmd$donotembedsysfont(val) {
		this.doc.attributes.donotembedsysfont = val;
	}
	cmd$donotembedlingdata(val) {
		this.doc.attributes.donotembedlingdata = val;
	}
	cmd$showplaceholdtext(val) {
		this.doc.attributes.showplaceholdertext = val;
	}
	cmd$trackmoves(val) {
		this.doc.attributes.trackmoves = val;
	}
	cmd$trackformatting(val) {
		this.doc.attributes.trackformatting = val;
	}
	cmd$ignoreremixedcontent(val) {
		this.doc.attributes.ignoreremixedcontent = val;
	}
	cmd$saveinvalidxml(val) {
		this.doc.attributes.saveinvalidxml = val;
	}
	cmd$showxmlerrors(val) {
		this.doc.attributes.showxmlerrors = val;
	}
	cmd$stylelocktheme(val) {
		this.doc.style.locktheme = true;
	}
	cmd$stylelockqfset(val) {
		this.doc.style.lockqfset = true;
	}
	cmd$usenormstyforlist(val) {
		this.doc.style.usenormstyforlist = true;
	}
	cmd$wgrffmtfilter(val) {
		this.doc.style.wgrffmtfilter = val;
	}
	cmd$readonlyrecommended() {
		this.doc.style.readonlyrecommended = true;
	}
	cmd$stylesortmethod(val) {
		this.doc.style.stylesortmethod = val;
	}
	cmd$writereservhash() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "reservehash");
	}
	cmd$writereservation() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "reservation");
	}
	cmd$saveprevpict() {
		this.doc.attributes.saveprevpict = true;
	}
	cmd$viewkind(val) {
		this.doc.style.viewkind = val;
	}
	cmd$viewscale(val) {
		this.doc.style.viewscale = val;
	}
	cmd$viewzk(val) {
		this.doc.style.viewzk = val;
	}
	cmd$viewbksp(val) {
		this.doc.style.viewbksp = val;
	}

	/* Footnotes and Endnotes */
	cmd$fet(val) {
		this.doc.style.fet = val;
	}
	cmd$ftnsep() {
		this.curGroup = new ParameterGroup(this.doc.style, "footnotesep");
	}
	cmd$ftnsepc() {
		this.curGroup = new ParameterGroup(this.doc.style, "footnotesepc");
	}
	cmd$ftncn() {
		this.curGroup = new ParameterGroup(this.doc.style, "footnotenotice");
	}
	cmd$aftnsep() {
		this.curGroup = new ParameterGroup(this.doc.style, "endnotesep");
	}
	cmd$aftnsepc() {
		this.curGroup = new ParameterGroup(this.doc.style, "endnotesepc");
	}
	cmd$aftncn() {
		this.curGroup = new ParameterGroup(this.doc.style, "endnotenotice");
	}
	cmd$endnotes() {
		this.doc.attributes.footnoteposition = "endnotes";
	}
	cmd$enddoc() {
		this.doc.attributes.footnoteposition = "enddoc";
	}
	cmd$ftntj() {
		this.doc.attributes.footnoteposition = "beneathtexttj";
	}
	cmd$ftnbj() {
		this.doc.attributes.footnoteposition = "bottomofpagebj";
	}
	cmd$aendnotes() {
		this.doc.attributes.endnoteposition = "endnotes";
	}
	cmd$aenddoc() {
		this.doc.attributes.endnoteposition = "enddoc";
	}
	cmd$aftntj() {
		this.doc.attributes.endnoteposition = "beneathtexttj";
	}
	cmd$aftnbj() {
		this.doc.attributes.endnoteposition = "bottomofpagebj";
	}
	cmd$ftnstart(val) {
		this.doc.style.footnotestart = val;
	}
	cmd$aftnstart(val) {
		this.doc.style.endnotestart = val;
	}
	cmd$ftnrstpg() {
		this.doc.style.footnoterestart = "page";
	}
	cmd$ftnrestart() {
		this.doc.style.footnoterestart = "section";
	}
	cmd$ftnrstcont() {
		this.doc.style.footnoterestart = "none";
	}
	cmd$aftnrestart() {
		this.doc.style.endnoterestart = "section";
	}
	cmd$aftnrstcont() {
		this.doc.style.endnoterestart = "none";
	}
	cmd$ftnnar() {
		this.doc.style.footnotenumbering = "1";
	}
	cmd$ftnnalc() {
		this.doc.style.footnotenumbering = "a";
	}
	cmd$ftnnauc() {
		this.doc.style.footnotenumbering = "A";
	}
	cmd$ftnnrlc() {
		this.doc.style.footnotenumbering = "i";
	}
	cmd$ftnnruc() {
		this.doc.style.footnotenumbering = "I";
	}
	cmd$ftnnchi() {
		this.doc.style.footnotenumbering = "*";
	}
	cmd$ftnnchosung() {
		this.doc.style.footnotenumbering = "CHOSUNG";
	}
	cmd$ftnncnum() {
		this.doc.style.footnotenumbering = "CIRCLENUM";
	}
	cmd$ftnndbnum() {
		this.doc.style.footnotenumbering = "DBNUM1";
	}
	cmd$ftnndbnumd() {
		this.doc.style.footnotenumbering = "DBNUM2";
	}
	cmd$ftnndbnumt() {
		this.doc.style.footnotenumbering = "DBNUM3";
	}
	cmd$ftnndbnumk() {
		this.doc.style.footnotenumbering = "DBNUM4";
	}
	cmd$ftnndbar() {
		this.doc.style.footnotenumbering = "DBCHAR";
	}
	cmd$ftnnganada() {
		this.doc.style.footnotenumbering = "GANADA";
	}
	cmd$ftnngbnum() {
		this.doc.style.footnotenumbering = "GB1";
	}
	cmd$ftnngbnumd() {
		this.doc.style.footnotenumbering = "GB2";
	}
	cmd$ftnngbnuml() {
		this.doc.style.footnotenumbering = "GB3";
	}
	cmd$ftnngbnumk() {
		this.doc.style.footnotenumbering = "GB4";
	}
	cmd$ftnnzodiac() {
		this.doc.style.footnotenumbering = "ZODIAC1";
	}
	cmd$ftnnzodiacd() {
		this.doc.style.footnotenumbering = "ZODIAC2";
	}
	cmd$ftnnzodiacl() {
		this.doc.style.footnotenumbering = "ZODIAC3";
	}
	cmd$aftnnar() {
		this.doc.style.endnotenumbering = "1";
	}
	cmd$aftnnalc() {
		this.doc.style.endnotenumbering = "a";
	}
	cmd$aftnnauc() {
		this.doc.style.endnotenumbering = "A";
	}
	cmd$aftnnrlc() {
		this.doc.style.endnotenumbering = "i";
	}
	cmd$aftnnruc() {
		this.doc.style.endnotenumbering = "I";
	}
	cmd$aftnnchi() {
		this.doc.style.endnotenumbering = "*";
	}
	cmd$aftnnchosung() {
		this.doc.style.endnotenumbering = "CHOSUNG";
	}
	cmd$aftnncnum() {
		this.doc.style.endnotenumbering = "CIRCLENUM";
	}
	cmd$aftnndbnum() {
		this.doc.style.endnotenumbering = "DBNUM1";
	}
	cmd$aftnndbnumd() {
		this.doc.style.endnotenumbering = "DBNUM2";
	}
	cmd$aftnndbnumt() {
		this.doc.style.endnotenumbering = "DBNUM3";
	}
	cmd$aftnndbnumk() {
		this.doc.style.endnotenumbering = "DBNUM4";
	}
	cmd$aftnndbar() {
		this.doc.style.endnotenumbering = "DBCHAR";
	}
	cmd$aftnnganada() {
		this.doc.style.endnotenumbering = "GANADA";
	}
	cmd$aftnngbnum() {
		this.doc.style.endnotenumbering = "GB1";
	}
	cmd$aftnngbnumd() {
		this.doc.style.endnotenumbering = "GB2";
	}
	cmd$aftnngbnuml() {
		this.doc.style.endnotenumbering = "GB3";
	}
	cmd$aftnngbnumk() {
		this.doc.style.endnotenumbering = "GB4";
	}
	cmd$aftnnzodiac() {
		this.doc.style.endnotenumbering = "ZODIAC1";
	}
	cmd$aftnnzodiacd() {
		this.doc.style.endnotenumbering = "ZODIAC2";
	}
	cmd$aftnnzodiacl() {
		this.doc.style.endnotenumbering = "ZODIAC3";
	}

	/* Page Information */
	cmd$paperw(val) {
		this.doc.style.paperwidth = val;
	}
	cmd$paperh(val) {
		this.doc.style.paperheight = val;
	}
	cmd$margl(val) {
		this.doc.style.marginleft = val;
	}
	cmd$margr(val) {
		this.doc.style.marginright = val;
	}
	cmd$margt(val) {
		this.doc.style.margintop = val;
	}
	cmd$margb(val) {
		this.doc.style.marginbottom = val;
	}
	cmd$facingp() {
		this.doc.style.facingpages = true;
	}
	cmd$gutter(val) {
		this.doc.style.gutterwidth = val;
	}
	cmd$ogutter(val) {
		this.doc.style.outsidegutterwidth = val;
	}
	cmd$rtlgutter() {
		this.doc.style.gutterright = true;
	}
	cmd$gutterprl() {
		this.doc.style.gutterparallel = true;
	}
	cmd$margmirror() {
		this.doc.style.mirroredmargins = true;
	}
	cmd$landscape() {
		this.doc.style.landscape = true;
	}
	cmd$pgnstart(val) {
		this.doc.style.pagenumberstart = val;
	}
	cmd$widowctrl() {
		this.doc.style.widowcontrol = true;
	}
	cmd$twoonone() {
		this.doc.style.twoonone = true;
	}
	cmd$bookfold() {
		this.doc.style.bookfold = true;
	}
	cmd$bookfoldrev() {
		this.doc.style.bookfoldrev = true;
	}
	cmd$bookfoldsheets(val) {
		this.doc.style.bookfoldsheets = val;
	}

	/* Linked Styles */
	cmd$linkstyles() {
		this.doc.attributes.linkstyles = true;
	}







	/* Mail Merge */

	/* Section Formatting Properties */

	/* Headers, Footers */

	/* Paragraphs */
	cmd$par() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			const prevStyle = this.curGroup.curstyle;
			this.endGroup();
			this.newGroup("paragraph");
			this.curGroup.style = prevStyle;
		} else {
			this.newGroup("paragraph");
		}	
	}
	cmd$pard() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.style = Object.assign(JSON.parse(JSON.stringify(this.defCharState)),JSON.parse(JSON.stringify(this.defParState)));
		} else {
			this.newGroup("paragraph");
			this.curGroup.style = Object.assign(JSON.parse(JSON.stringify(this.defCharState)),JSON.parse(JSON.stringify(this.defParState)));
		}
	}
	cmd$plain() {
		Object.keys(this.defCharState).forEach(key => {
			if (this.curGroup.style[key]) {
				this.curGroup.style[key] = this.defCharState[key];
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
		this.curGroup.style.foreground = this.doc.tables.colourTable[val - 1];
	}
	cmd$cb(val) {
		this.curGroup.style.background = this.doc.tables.colourTable[val - 1];
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
			this.skip += this.curGroup.uc;
		} else if (this.curGroup.parent.uc) {
			this.skip += this.curGroup.parent.uc;
		} else {
			this.skip += 1;
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
}

module.exports = LargeRTFSubunit;