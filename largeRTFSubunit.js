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
	MailMergeTable,
	Odso,
	FieldMap,
	OdsoRecip,
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
		this.curGroup.style.defaulttab = val;
	}
	cmd$hyphhotz(val) {
		this.curGroup.style.hyphenhotzone = val;
	}
	cmd$hyphconsec(val) {
		this.curGroup.style.hyphenconsecutive = val;
	}
	cmd$hyphcaps(val) {
		this.curGroup.style.hyphencaps = val !== 0;
	}
	cmd$hyphauto(val) {
		this.curGroup.style.hyphenauto = val !== 0;
	}
	cmd$linestart(val) {
		this.curGroup.style.linestart = val;
	}
	cmd$fracwidth(val) {
		this.curGroup.style.fractionalwidths = true;
	}
	cmd$nextfile() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "nextfile");
	}
	cmd$template() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "template");
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
		this.curGroup.style.rendering = "horizontal";
	}
	cmd$horzdoc() {
		this.curGroup.style.rendering = "vertical";
	}
	cmd$jcompress() {
		this.curGroup.style.justification = "compressing";
	}
	cmd$jexpand() {
		this.curGroup.style.justification = "expanding";
	}
	cmd$lnongrid() {
		this.curGroup.style.lineongrid = true;
	}
	cmd$grfdocevents(val) {
		this.doc.attributes.grfdocevents = val;
	}
	cmd$themelang(val) {
		this.curGroup.style.themelanguage = val;
	}
	cmd$themelangfe(val) {
		this.curGroup.style.themelanguagefe = val;
	}
	cmd$themelangcs(val) {
		this.curGroup.style.themelanguagecs = val;
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
		this.curGroup.style.locktheme = true;
	}
	cmd$stylelockqfset(val) {
		this.curGroup.style.lockqfset = true;
	}
	cmd$usenormstyforlist(val) {
		this.curGroup.style.usenormstyforlist = true;
	}
	cmd$wgrffmtfilter(val) {
		this.curGroup.style.wgrffmtfilter = val;
	}
	cmd$readonlyrecommended() {
		this.curGroup.style.readonlyrecommended = true;
	}
	cmd$stylesortmethod(val) {
		this.curGroup.style.stylesortmethod = val;
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
		this.curGroup.style.viewkind = val;
	}
	cmd$viewscale(val) {
		this.curGroup.style.viewscale = val;
	}
	cmd$viewzk(val) {
		this.curGroup.style.viewzk = val;
	}
	cmd$viewbksp(val) {
		this.curGroup.style.viewbksp = val;
	}

	/*-- Footnotes and Endnotes --*/
	cmd$fet(val) {
		this.curGroup.style.fet = val;
	}
	cmd$ftnsep() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnotesep");
	}
	cmd$ftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnotesepc");
	}
	cmd$ftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnotenotice");
	}
	cmd$aftnsep() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pagesep");
	}
	cmd$aftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pagesepc");
	}
	cmd$aftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pagenotice");
	}
	cmd$pages() {
		this.doc.attributes.footnoteposition = "pages";
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
	cmd$apages() {
		this.doc.attributes.pageposition = "pages";
	}
	cmd$aenddoc() {
		this.doc.attributes.pageposition = "enddoc";
	}
	cmd$aftntj() {
		this.doc.attributes.pageposition = "beneathtexttj";
	}
	cmd$aftnbj() {
		this.doc.attributes.pageposition = "bottomofpagebj";
	}
	cmd$ftnstart(val) {
		this.curGroup.style.footnotestart = val;
	}
	cmd$aftnstart(val) {
		this.curGroup.style.pagestart = val;
	}
	cmd$ftnrstpg() {
		this.curGroup.style.footnoterestart = "page";
	}
	cmd$ftnrestart() {
		this.curGroup.style.footnoterestart = "section";
	}
	cmd$ftnrstcont() {
		this.curGroup.style.footnoterestart = "none";
	}
	cmd$aftnrestart() {
		this.curGroup.style.pagerestart = "section";
	}
	cmd$aftnrstcont() {
		this.curGroup.style.pagerestart = "none";
	}
	cmd$ftnnar() {
		this.curGroup.style.footnotenumbering = "1";
	}
	cmd$ftnnalc() {
		this.curGroup.style.footnotenumbering = "a";
	}
	cmd$ftnnauc() {
		this.curGroup.style.footnotenumbering = "A";
	}
	cmd$ftnnrlc() {
		this.curGroup.style.footnotenumbering = "i";
	}
	cmd$ftnnruc() {
		this.curGroup.style.footnotenumbering = "I";
	}
	cmd$ftnnchi() {
		this.curGroup.style.footnotenumbering = "*";
	}
	cmd$ftnnchosung() {
		this.curGroup.style.footnotenumbering = "CHOSUNG";
	}
	cmd$ftnncnum() {
		this.curGroup.style.footnotenumbering = "CIRCLENUM";
	}
	cmd$ftnndbnum() {
		this.curGroup.style.footnotenumbering = "DBNUM1";
	}
	cmd$ftnndbnumd() {
		this.curGroup.style.footnotenumbering = "DBNUM2";
	}
	cmd$ftnndbnumt() {
		this.curGroup.style.footnotenumbering = "DBNUM3";
	}
	cmd$ftnndbnumk() {
		this.curGroup.style.footnotenumbering = "DBNUM4";
	}
	cmd$ftnndbar() {
		this.curGroup.style.footnotenumbering = "DBCHAR";
	}
	cmd$ftnnganada() {
		this.curGroup.style.footnotenumbering = "GANADA";
	}
	cmd$ftnngbnum() {
		this.curGroup.style.footnotenumbering = "GB1";
	}
	cmd$ftnngbnumd() {
		this.curGroup.style.footnotenumbering = "GB2";
	}
	cmd$ftnngbnuml() {
		this.curGroup.style.footnotenumbering = "GB3";
	}
	cmd$ftnngbnumk() {
		this.curGroup.style.footnotenumbering = "GB4";
	}
	cmd$ftnnzodiac() {
		this.curGroup.style.footnotenumbering = "ZODIAC1";
	}
	cmd$ftnnzodiacd() {
		this.curGroup.style.footnotenumbering = "ZODIAC2";
	}
	cmd$ftnnzodiacl() {
		this.curGroup.style.footnotenumbering = "ZODIAC3";
	}
	cmd$aftnnar() {
		this.curGroup.style.pagenumbering = "1";
	}
	cmd$aftnnalc() {
		this.curGroup.style.pagenumbering = "a";
	}
	cmd$aftnnauc() {
		this.curGroup.style.pagenumbering = "A";
	}
	cmd$aftnnrlc() {
		this.curGroup.style.pagenumbering = "i";
	}
	cmd$aftnnruc() {
		this.curGroup.style.pagenumbering = "I";
	}
	cmd$aftnnchi() {
		this.curGroup.style.pagenumbering = "*";
	}
	cmd$aftnnchosung() {
		this.curGroup.style.pagenumbering = "CHOSUNG";
	}
	cmd$aftnncnum() {
		this.curGroup.style.pagenumbering = "CIRCLENUM";
	}
	cmd$aftnndbnum() {
		this.curGroup.style.pagenumbering = "DBNUM1";
	}
	cmd$aftnndbnumd() {
		this.curGroup.style.pagenumbering = "DBNUM2";
	}
	cmd$aftnndbnumt() {
		this.curGroup.style.pagenumbering = "DBNUM3";
	}
	cmd$aftnndbnumk() {
		this.curGroup.style.pagenumbering = "DBNUM4";
	}
	cmd$aftnndbar() {
		this.curGroup.style.pagenumbering = "DBCHAR";
	}
	cmd$aftnnganada() {
		this.curGroup.style.pagenumbering = "GANADA";
	}
	cmd$aftnngbnum() {
		this.curGroup.style.pagenumbering = "GB1";
	}
	cmd$aftnngbnumd() {
		this.curGroup.style.pagenumbering = "GB2";
	}
	cmd$aftnngbnuml() {
		this.curGroup.style.pagenumbering = "GB3";
	}
	cmd$aftnngbnumk() {
		this.curGroup.style.pagenumbering = "GB4";
	}
	cmd$aftnnzodiac() {
		this.curGroup.style.pagenumbering = "ZODIAC1";
	}
	cmd$aftnnzodiacd() {
		this.curGroup.style.pagenumbering = "ZODIAC2";
	}
	cmd$aftnnzodiacl() {
		this.curGroup.style.pagenumbering = "ZODIAC3";
	}

	/*-- Page Information --*/
	cmd$paperw(val) {
		this.curGroup.style.paperwidth = val;
	}
	cmd$paperh(val) {
		this.curGroup.style.paperheight = val;
	}
	cmd$margl(val) {
		this.curGroup.style.marginleft = val;
	}
	cmd$margr(val) {
		this.curGroup.style.marginright = val;
	}
	cmd$margt(val) {
		this.curGroup.style.margintop = val;
	}
	cmd$margb(val) {
		this.curGroup.style.marginbottom = val;
	}
	cmd$facingp() {
		this.curGroup.style.facingpages = true;
	}
	cmd$gutter(val) {
		this.curGroup.style.gutterwidth = val;
	}
	cmd$ogutter(val) {
		this.curGroup.style.outsidegutterwidth = val;
	}
	cmd$rtlgutter() {
		this.curGroup.style.gutterright = true;
	}
	cmd$gutterprl() {
		this.curGroup.style.gutterparallel = true;
	}
	cmd$margmirror() {
		this.curGroup.style.mirroredmargins = true;
	}
	cmd$landscape() {
		this.curGroup.style.landscape = true;
	}
	cmd$pgnstart(val) {
		this.curGroup.style.pagenumberstart = val;
	}
	cmd$widowctrl() {
		this.curGroup.style.widowcontrol = true;
	}
	cmd$twoonone() {
		this.curGroup.style.twoonone = true;
	}
	cmd$bookfold() {
		this.curGroup.style.bookfold = true;
	}
	cmd$bookfoldrev() {
		this.curGroup.style.bookfoldrev = true;
	}
	cmd$bookfoldsheets(val) {
		this.curGroup.style.bookfoldsheets = val;
	}

	/*-- Linked Styles --*/
	cmd$linkstyles() {
		this.doc.attributes.linkstyles = true;
	}

	/*-- Compatability Options --*/
	cmd$notabind() {
		this.curGroup.style.notabindent = true;
	}
	cmd$wraptrsp() {
		this.curGroup.style.wraptrailingwhitespace = true;
	}
	cmd$prcolbl() {
		this.curGroup.style.printcolourblack = true;
	}
	cmd$noextrasprl() {
		this.curGroup.style.noextraspacerl = true;
	}
	cmd$nocolbal() {
		this.curGroup.style.nocolumnbalance = true;
	}
	cmd$cvmme() {
		this.curGroup.style.cvmailmergeescape = true;
	}
	cmd$sprstsp() {
		this.curGroup.style.surpressextraline = true;
	}
	cmd$sprsspbf() {
		this.curGroup.style.surpressspacebefore = true;
	}
	cmd$otblrul() {
		this.curGroup.style.combinetableborders = true;
	}
	cmd$transmf() {
		this.curGroup.style.transparentmetafile = true;
	}
	cmd$swpbdr() {
		this.curGroup.style.swapborder = true;
	}
	cmd$brkfrm() {
		this.curGroup.style.hardbreaks = true;
	}
	cmd$sprslnsp() {
		this.curGroup.style.surpresslinespace = true;
	}
	cmd$subfontbysize() {
		this.curGroup.style.subfontbysize = true;
	}
	cmd$truncatefontheight() {
		this.curGroup.style.truncatefontheight = true;
	}
	cmd$truncex() {
		this.curGroup.style.noleadingspace = true;
	}
	cmd$bdbfhdr() {
		this.curGroup.style.bodybeforehf = true;
	}
	cmd$dntblnsbdb() {
		this.curGroup.style.nobalancesbdb = true;
	}
	cmd$expshrtn() {
		this.curGroup.style.expandcharspace = true;
	}
	cmd$lytexcttp() {
		this.curGroup.style.nocenterlhlines = true;
	}
	cmd$lytprtmet() {
		this.curGroup.style.useprintermetrics = true;
	}
	cmd$msmcap() {
		this.curGroup.style.macsmallcaps = true;
	}
	cmd$nolead() {
		this.curGroup.style.noleading = true;
	}
	cmd$nospaceforul() {
		this.curGroup.style.nounderlinespace = true;
	}
	cmd$noultrlspc() {
		this.curGroup.style.nounderlinetrailing = true;
	}
	cmd$noxlattoyen() {
		this.curGroup.style.noyentranslate = true;
	}
	cmd$oldlinewrap() {
		this.curGroup.style.oldlinewrap = true;
	}
	cmd$sprsbsp() {
		this.curGroup.style.surpressextraspaceb = true;
	}
	cmd$sprstsm() {
		//Does literally nothing. Why is it here? We don't know.
	}
	cmd$wpjst() {
		this.curGroup.style.wpjustify = true;
	}
	cmd$wpsp() {
		this.curGroup.style.wpspacewidth = true;
	}
	cmd$wptb() {
		this.curGroup.style.wptabadvance = true;
	}
	cmd$splytwnine() {
		this.curGroup.style.nolegacyautoshape = true;
	}
	cmd$ftnlytwnine() {
		this.curGroup.style.nolegacyfootnote = true;
	}
	cmd$htmautsp() {
		this.curGroup.style.htmlautospace = true;
	}
	cmd$useltbaln() {
		this.curGroup.style.noforgetlattab = true;
	}
	cmd$alntblind() {
		this.curGroup.style.noindependentrowalign = true;
	}
	cmd$lytcalctblwd() {
		this.curGroup.style.norawtablewidth = true;
	}
	cmd$lyttblrtgr() {
		this.curGroup.style.notablerowapart = true;
	}
	cmd$oldas() {
		this.curGroup.style.ninetyfiveautospace = true;
	}
	cmd$lnbrkrule() {
		this.curGroup.style.nolinebreakrule = true;
	}
	cmd$bdrrlswsix() {
		this.curGroup.style.uselegacyborderrules = true;
	}
	cmd$nolnhtadjtbl() {
		this.curGroup.style.noadjusttablelineheight = true;
	}
	cmd$ApplyBrkRules() {
		this.curGroup.style.applybreakrules = true;
	}
	cmd$rempersonalinfo() {
		this.curGroup.style.removepersonalinfo = true;
	}
	cmd$remdttm() {
		this.curGroup.style.removedatetime = true;
	}
	cmd$snaptogridincell() {
		this.curGroup.style.snaptexttogrid = true;
	}
	cmd$wrppunct() {
		this.curGroup.style.hangingpunctuation = true;
	}
	cmd$asianbrkrule() {
		this.curGroup.style.asianbreakrules = true;
	}
	cmd$nobrkwrptbl() {
		this.curGroup.style.nobreakwrappedtable = true;
	}
	cmd$toplinepunct() {
		this.curGroup.style.toplinepunct = true;
	}
	cmd$viewnobound() {
		this.curGroup.style.hidepagebetweenspace = true;
	}
	cmd$donotshowmarkup() {
		this.curGroup.style.noshowmarkup = true;
	}
	cmd$donotshowcomments() {
		this.curGroup.style.noshowcomments = true;
	}
	cmd$donotshowinsdel() {
		this.curGroup.style.noshowinsdel = true;
	}
	cmd$donotshowprops() {
		this.curGroup.style.noshowformatting = true;
	}
	cmd$allowfieldendsel() {
		this.curGroup.style.fieldendselect = true;
	}
	cmd$nocompatoptions() {
		this.curGroup.style.compatabilitydefaults = true;
	}
	cmd$nogrowautofit() {
		this.curGroup.style.notableautofit = true;
	}
	cmd$newtblstyruls() {
		this.curGroup.style.newtablestylerules = true;
	}
	cmd$background() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "docbackground");
	}
	cmd$nouicompat() {
		this.curGroup.style.nouicompatability = true;
	}
	cmd$nofeaturethrottle(val) {
		this.curGroup.style.nofeaturethrottle = val;
	}
	cmd$forceupgrade() {
		this.curGroup.style.mayupgrade = true;
	}
	cmd$noafcnsttbl() {
		this.curGroup.style.notableautowrap = true;
	}
	cmd$noindnmbrts() {
		this.curGroup.style.bullethangindent = true;
	}
	cmd$felnbrelev() {
		this.curGroup.style.alternatebeginend = true;
	}
	cmd$indrlsweleven() {
		this.curGroup.style.ignorefloatingobjectvector = true;
	}
	cmd$nocxsptable() {
		this.curGroup.style.notableparspace = true;
	}
	cmd$notcvasp() {
		this.curGroup.style.notablevectorvertical = true;
	}
	cmd$notvatxbx() {
		this.curGroup.style.notextboxvertical = true;
	}
	cmd$spltpgpar() {
		this.curGroup.style.splitpageparagraph = true;
	}
	cmd$hwelev() {
		this.curGroup.style.hangulfixedwidth = true;
	}
	cmd$afelev() {
		this.curGroup.style.tableautofitmimic = true;
	}
	cmd$cachedcolbal() {
		this.curGroup.style.cachedcolumnbalance = true;
	}
	cmd$utinl() {
		this.curGroup.style.underlinenumberedpar = true;
	}
	cmd$notbrkcnstfrctbl () {
		this.curGroup.style.notablerowsplit = true;
	}
	cmd$krnprsnet() {
		this.curGroup.style.ansikerning = true;
	}
	cmd$usexform() {
		this.curGroup.style.noxsltransform = true;
	}

	/*-- Forms --*/
	cmd$formprot() {
		this.doc.attributes.protectedforms = true;
	}
	cmd$allprot() {
		this.doc.attributes.protectedall = true;
	}
	cmd$formshade() {
		this.curGroup.style.formfieldshading = true;
	}
	cmd$formdisp() {
		this.doc.attributes.formboxselected = true;
	}
	cmd$formprot() {
		this.doc.attributes.printformdata = true;
	}

	/*-- Revision Marks --*/
	cmd$revprot() {
		this.doc.attributes.protectedrevisions = true;
	}
	cmd$revisions() {
		this.curGroup.style.revisions = true;
	}
	cmd$revprop(val) {
		this.curGroup.style.revisiontextdisplay = val;
	}
	cmd$revbar(val) {
		this.curGroup.style.revisionlinemarking = val;
	}

	/*-- Write Protection --*/
	cmd$readprot() {
		this.doc.attributes.protectedread = true;
	}

	/*-- Comment Protection --*/
	cmd$annotprot() {
		this.doc.attributes.protectedcomments = true;
	}

	/*-- Style Protection --*/
	cmd$stylelock() {
		this.doc.attributes.stylelock = true;
	}
	cmd$stylelockenforced() {
		this.doc.attributes.stylelockenforced = true;
	}
	cmd$stylelockbackcomp() {
		this.doc.attributes.stylelockcompatability = true;
	}
	cmd$autofmtoverride() {
		this.doc.attributes.autoformatlockoverride = true;
	}

	/*-- Style and Formatting Properties --*/
	cmd$enforceprot(val) {
		this.doc.attributes.enforceprotection = val;
	}
	cmd$protlevel(val) {
		this.doc.attributes.protectionlevel = val;
	}

	/*-- Tables --*/
	cmd$tsd(val) {
		this.curGroup.style.defaulttablestyle = val;
	}
	
	/*-- Bidirectional Controls --*/
	cmd$rtldoc() {
		this.doc.attributes.paginationdirection = "rtl";
	}
	cmd$ltrdoc() {
		this.doc.attributes.paginationdirection = "ltr";
	}

	/*-- Click and Type --*/
	cmd$cts(val) {
		this.doc.attributes.clickandtype = val;
	}

	/*-- Kinsoku Characters --*/
	cmd$jsksu() {
		this.doc.attributes.japanesekinsoku = true;
	}
	cmd$ksulang(val) {
		this.doc.attributes.kinsokulang = val;
	}
	cmd$fchars() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "followingkinsoku");
	}
	cmd$lchars() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "leadingkinsoku");
	}
	cmd$nojkernpunct() {
		this.doc.attributes.latinkerningonly = true;
	}

	/*-- Drawing Grid --*/
	cmd$dghspace(val) {
		this.curGroup.style.drawgridhorizontalspace = val;
	}
	cmd$dgvspace(val) {
		this.curGroup.style.drawgridverticalspace = val;
	}
	cmd$dghorigin(val) {
		this.curGroup.style.drawgridhorizontalorigin = val;
	}
	cmd$dgvorigin(val) {
		this.curGroup.style.drawgridverticalorigin = val;
	}
	cmd$dghshow(val) {
		this.curGroup.style.drawgridhorizontalshow = val;
	}
	cmd$dgvshow(val) {
		this.curGroup.style.drawgridverticalshow = val;
	}
	cmd$dgsnap() {
		this.curGroup.style.drawgridsnap = true;
	}
	cmd$dgmargin() {
		this.curGroup.style.drawgridmargin = true;
	}

	/*-- Page Borders --*/
	cmd$pgbrdrhead() {
		this.curGroup.style.pagebordersurroundsheader = true;
	}
	cmd$pgbrdrfoot() {
		this.curGroup.style.pagebordersurroundsfooter = true;
	}
	cmd$pgbrdrt() {
		this.curGroup.style.pagebordertop = true;
	}
	cmd$pgbrdrb() {
		this.curGroup.style.pageborderbottom = true;
	}
	cmd$pgbrdrl() {
		this.curGroup.style.pageborderleft = true;
	}
	cmd$pgbrdrr() {
		this.curGroup.style.pageborderright = true;
	}
	cmd$brdrart(val) {
		this.curGroup.style.pageborderart = val;
	}
	cmd$pgbrdropt(val) {
		this.curGroup.style.pageborderoptions = val;
	}
	cmd$pgbrdrsnap() {
		this.curGroup.style.pagebordersnap = true;
	}

	/* Mail Merge */
	cmd$mailmerge() {
		this.curGroup = new MailMergeTable(this.doc);
	}
	cmd$mmlinktoquery() {
		this.curGroup.attributes.linktoquery = true;
	}
	cmd$mmdefaultsql() {
		this.curGroup.attributes.defaultsql = true;
	}
	cmd$mmconnectstrdata() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectstringdata");
	}
	cmd$mmconnectstr() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectstring");
	}
	cmd$mmquery() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectstringdata");
	}
	cmd$mmdatasource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "datasource");
	}
	cmd$mmheadersource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "headersource");
	}
	cmd$mmblanklinks() {
		this.curGroup.attributes.blanklinks = true;
	}
	cmd$mmaddfieldname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "fieldname");
	}
	cmd$mmmailsubject() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "subject");
	}
	cmd$mmattatch() {
		this.curGroup.attributes.attatch = true;
	}
	cmd$mmshowdata() {
		this.curGroup.attributes.showdata = true;
	}
	cmd$mmreccur(val) {
		this.curGroup.attributes.reccur = val;
	}
	cmd$mmerrors(val) {
		this.curGroup.attributes.errorreporting = val;
	}
	cmd$mmodso() {
		this.curGroup = new Odso(this.curGroup.parent);
	}
	cmd$mmodsoudldata() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "udldata");
	}
	cmd$mmodsoudl() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "udl");
	}
	cmd$mmodsotable() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "table");
	}
	cmd$mmodsosrc() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "source");
	}
	cmd$mmodsofilter() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "filter");
	}
	cmd$mmodsosort() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "sort");
	}
	cmd$mmodsofldmpdata() {
		this.curGroup = new FieldMap(this.curGroup.parent);
	}
	cmd$mmodsoname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "name");
	}
	cmd$mmodsomappedname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "mappedname");
	}
	cmd$mmodsofmcolumn(val) {
		this.curGroup.attributes.columnindex = val;
	}
	cmd$mmodsodynaddr(val) {
		this.curGroup.attributes.addressorder = val;
	}
	cmd$mmodsosolid(val) {
		this.curGroup.attributes.language = val;
	}
	cmd$mmodsocoldelim(val) {
		this.curGroup.attributes.columndelimiter = val;
	}
	cmd$mmjdsotype(val) {
		this.curGroup.attributes.datasourcetype = val;
	}
	cmd$mmodsofhdr(val) {
		this.curGroup.attributes.firstrowheader = val;
	}
	cmd$mmodsorecipdata() {
		this.curGroup = new OdsoRecip(this.curGroup.parent);
	}
	cmd$mmodsoactive(val) {
		this.curGroup.attributes.active = val;
	}
	cmd$mmodsohash(val) {
		this.curGroup.attributes.hash = val;
	}
	cmd$mmodsocolumn(val) {
		this.curGroup.attributes.column = val;
	}
	cmd$mmodsouniquetag() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "uniquetag");
	}

	cmd$mmfttypenull() {
		this.curGroup.attributes.datatype = "null";
	}
	cmd$mmfttypedbcolumn() {
		this.curGroup.attributes.datatype = "databasecolumn";
	}
	cmd$mmfttypeaddress() {
		this.curGroup.attributes.datatype = "address";
	}
	cmd$mmfttypesalutation() {
		this.curGroup.attributes.datatype = "salutation";
	}
	cmd$mmfttypemapped() {
		this.curGroup.attributes.datatype = "mapped";
	}
	cmd$mmfttypebarcode() {
		this.curGroup.attributes.datatype = "barcode";
	}

	cmd$mmdestnewdoc() {
		this.curGroup.attributes.destination = "newdocument";
	}
	cmd$mmdestprinter() {
		this.curGroup.attributes.destination = "printer";
	}
	cmd$mmdestemail() {
		this.curGroup.attributes.destination = "email";
	}
	cmd$mmdestfax() {
		this.curGroup.attributes.destination = "fax";
	}

	cmd$mmmaintypecataloq() {
		this.curGroup.attributes.sourcetype = "cataloq";
	}
	cmd$mmmaintypeenvelopes() {
		this.curGroup.attributes.sourcetype = "envelope";
	}
	cmd$mmmaintypelabels() {
		this.curGroup.attributes.sourcetype = "label";
	}
	cmd$mmmaintypeletters() {
		this.curGroup.attributes.sourcetype = "letter";
	}
	cmd$mmmaintypeemail() {
		this.curGroup.attributes.sourcetype = "email";
	}
	cmd$mmmaintypefax() {
		this.curGroup.attributes.sourcetype = "fax";
	}

	cmd$mmdatatypeaccess() {
		this.curGroup.attributes.connectiontype = "dde-access";
	}
	cmd$mmdatatypeexcel() {
		this.curGroup.attributes.connectiontype = "dde-excel";
	}
	cmd$mmdatatypeqt() {
		this.curGroup.attributes.connectiontype = "externalquery";
	}
	cmd$mmdatatypeodbc() {
		this.curGroup.attributes.connectiontype = "odbc";
	}
	cmd$mmdatatypeodso() {
		this.curGroup.attributes.connectiontype = "odso";
	}
	cmd$mmdatatypefile() {
		this.curGroup.attributes.connectiontype = "dde-textfile";
	}

	/* Sections */
	cmd$sect() {
		if (this.curGroup.type === "section") {
			const prevStyle = this.curGroup.curstyle;
			this.endGroup();
			this.newGroup("section");
			this.curGroup.style = prevStyle;
		} else {
			this.newGroup("section");
		}	
	}
	cmd$sectd() {
		let defSectStyle = {};
		if (this.doc.style.paperwidth) {defSectStyle.paperwidth = this.doc.style.paperwidth}
		if (this.doc.style.paperheight) {defSectStyle.paperheight = this.doc.style.paperheight}
		if (this.doc.style.marginleft) {defSectStyle.marginleft = this.doc.style.marginleft}
		if (this.doc.style.marginright) {defSectStyle.marginright = this.doc.style.marginright}
		if (this.doc.style.margintop) {defSectStyle.margintop = this.doc.style.margintop}
		if (this.doc.style.marginbottom) {defSectStyle.marginbottom = this.doc.style.marginbottom}
		if (this.doc.style.gutterwidth) {defSectStyle.gutterwidth = this.doc.style.gutterwidth}
		if (this.curGroup.type === "section") {
			this.curGroup.style = defSectStyle;
		} else {
			this.newGroup("section");
			this.curGroup.style = defSectStyle;
		}
	}
	cmd$endnhere() {
		this.curGroup.attributes.pagesincluded = true;
	}
	cmd$binfsxn(val) {
		this.curGroup.attributes.firstprinterbin = val;
	}
	cmd$binsxn(val) {
		this.curGroup.attributes.printerbin = val;
	}
	cmd$pnseclvl(val) {
		this.curGroup.style.liststyle = val;
	}
	cmd$sectunlocked() {
		this.curGroup.attributes.unlocked = true;
	}

	cmd$sbknone() {
		this.curGroup.style.sectionbreak = "none";
	}
	cmd$sbkcol() {
		this.curGroup.style.sectionbreak = "column";
	}
	cmd$sbkpage() {
		this.curGroup.style.sectionbreak = "page";
	}
	cmd$sbkeven() {
		this.curGroup.style.sectionbreak = "even";
	}
	cmd$sbkodd() {
		this.curGroup.style.sectionbreak = "odd";
	}

	cmd$cols(val) {
		this.curGroup.style.columns = val;
	}
	cmd$colsx(val) {
		this.curGroup.style.columnspace = val;
	}
	cmd$colno(val) {
		this.curGroup.style.columnnumber = val;
	}
	cmd$colsr(val) {
		this.curGroup.style.columnrightspace = val;
	}
	cmd$colw(val) {
		this.curGroup.style.columnwidth = val;
	}
	cmd$linebetcol() {
		this.curGroup.style.linebetweencolumns = true;
	}

	cmd$sftntj() {
		this.curGroup.style.footnotepos = "beneath";
	}
	cmd$sftnbj() {
		this.curGroup.style.footnotepos = "bottom";
	}
	cmd$sftnstart(val) {
		this.curGroup.style.footnotestart = val;
	}
	cmd$pgnnstart(val) {
		this.curGroup.style.pagestart = val;
	}
	cmd$sftnrstpg() {
		this.curGroup.style.footnoterestart = "page";
	}
	cmd$sftnrestart() {
		this.curGroup.style.footnoterestart = "section";
	}
	cmd$sftnrstcont() {
		this.curGroup.style.footnoterestart = "none";
	}
	cmd$pgnnrstpg() {
		this.curGroup.style.pagerestart = "page";
	}
	cmd$pgnnrestart() {
		this.curGroup.style.pagerestart = "section";
	}
	cmd$pgnnrstcont() {
		this.curGroup.style.pagerestart = "none";
	}
	cmd$sftnnar() {
		this.curGroup.style.footnotenumbering = "1";
	}
	cmd$sftnnalc() {
		this.curGroup.style.footnotenumbering = "a";
	}
	cmd$sftnnauc() {
		this.curGroup.style.footnotenumbering = "A";
	}
	cmd$sftnnrlc() {
		this.curGroup.style.footnotenumbering = "i";
	}
	cmd$sftnnruc() {
		this.curGroup.style.footnotenumbering = "I";
	}
	cmd$sftnnchi() {
		this.curGroup.style.footnotenumbering = "*";
	}
	cmd$sftnnchosung() {
		this.curGroup.style.footnotenumbering = "CHOSUNG";
	}
	cmd$sftnncnum() {
		this.curGroup.style.footnotenumbering = "CIRCLENUM";
	}
	cmd$sftnndbnum() {
		this.curGroup.style.footnotenumbering = "DBNUM1";
	}
	cmd$sftnndbnumd() {
		this.curGroup.style.footnotenumbering = "DBNUM2";
	}
	cmd$sftnndbnumt() {
		this.curGroup.style.footnotenumbering = "DBNUM3";
	}
	cmd$sftnndbnumk() {
		this.curGroup.style.footnotenumbering = "DBNUM4";
	}
	cmd$sftnndbar() {
		this.curGroup.style.footnotenumbering = "DBCHAR";
	}
	cmd$sftnnganada() {
		this.curGroup.style.footnotenumbering = "GANADA";
	}
	cmd$sftnngbnum() {
		this.curGroup.style.footnotenumbering = "GB1";
	}
	cmd$sftnngbnumd() {
		this.curGroup.style.footnotenumbering = "GB2";
	}
	cmd$sftnngbnuml() {
		this.curGroup.style.footnotenumbering = "GB3";
	}
	cmd$sftnngbnumk() {
		this.curGroup.style.footnotenumbering = "GB4";
	}
	cmd$sftnnzodiac() {
		this.curGroup.style.footnotenumbering = "ZODIAC1";
	}
	cmd$sftnnzodiacd() {
		this.curGroup.style.footnotenumbering = "ZODIAC2";
	}
	cmd$sftnnzodiacl() {
		this.curGroup.style.footnotenumbering = "ZODIAC3";
	}
	cmd$pgnnnar() {
		this.curGroup.style.pagenumbering = "1";
	}
	cmd$pgnnnalc() {
		this.curGroup.style.pagenumbering = "a";
	}
	cmd$pgnnnauc() {
		this.curGroup.style.pagenumbering = "A";
	}
	cmd$pgnnnrlc() {
		this.curGroup.style.pagenumbering = "i";
	}
	cmd$pgnnnruc() {
		this.curGroup.style.pagenumbering = "I";
	}
	cmd$pgnnnchi() {
		this.curGroup.style.pagenumbering = "*";
	}
	cmd$pgnnnchosung() {
		this.curGroup.style.pagenumbering = "CHOSUNG";
	}
	cmd$pgnnncnum() {
		this.curGroup.style.pagenumbering = "CIRCLENUM";
	}
	cmd$pgnnndbnum() {
		this.curGroup.style.pagenumbering = "DBNUM1";
	}
	cmd$pgnnndbnumd() {
		this.curGroup.style.pagenumbering = "DBNUM2";
	}
	cmd$pgnnndbnumt() {
		this.curGroup.style.pagenumbering = "DBNUM3";
	}
	cmd$pgnnndbnumk() {
		this.curGroup.style.pagenumbering = "DBNUM4";
	}
	cmd$pgnnndbar() {
		this.curGroup.style.pagenumbering = "DBCHAR";
	}
	cmd$pgnnnganada() {
		this.curGroup.style.pagenumbering = "GANADA";
	}
	cmd$pgnnngbnum() {
		this.curGroup.style.pagenumbering = "GB1";
	}
	cmd$pgnnngbnumd() {
		this.curGroup.style.pagenumbering = "GB2";
	}
	cmd$pgnnngbnuml() {
		this.curGroup.style.pagenumbering = "GB3";
	}
	cmd$pgnnngbnumk() {
		this.curGroup.style.pagenumbering = "GB4";
	}
	cmd$pgnnnzodiac() {
		this.curGroup.style.pagenumbering = "ZODIAC1";
	}
	cmd$pgnnnzodiacd() {
		this.curGroup.style.pagenumbering = "ZODIAC2";
	}
	cmd$pgnnnzodiacl() {
		this.curGroup.style.pagenumbering = "ZODIAC3";
	}

	cmd$linemod(val) {
		this.curGroup.style.linemodulus = val;
	}
	cmd$linex(val) {
		this.curGroup.style.linedistance = val;
	}
	cmd$linestarts(val) {
		this.curGroup.style.linestarts = val;
	}
	cmd$linerestart() {
		this.curGroup.style.linerestart = "onlinestarts";
	}
	cmd$lineppage() {
		this.curGroup.style.linerestarts = "page";
	}
	cmd$lineppage() {
		this.curGroup.style.linerestarts = "none";
	}

	cmd$pgwsxn(val) {
		this.curGroup.style.pagewidth = val;
	}
	cmd$pghsxn(val) {
		this.curGroup.style.pageheight = val;
	}
	cmd$marglsxn(val) {
		this.curGroup.style.marginleft = val;
	}
	cmd$margrsxn(val) {
		this.curGroup.style.marginright = val;
	}
	cmd$margtsxn(val) {
		this.curGroup.style.margintop = val;
	}
	cmd$margbsxn(val) {
		this.curGroup.style.marginbottom = val;
	}
	cmd$guttersxn(val) {
		this.curGroup.style.gutterwidth = val;
	}
	cmd$margmirsxn() {
		this.curGroup.style.marginswap = true;
	}
	cmd$lndscpsxn() {
		this.curGroup.style.landscape = true;
	}
	cmd$titlepg() {
		this.curGroup.style.titlepage = true;
	}
	cmd$headery(val) {
		this.curGroup.style.headertop = val;
	}
	cmd$footery(val) {
		this.curGroup.style.footerbottom = val;
	}

	cmd$pgncont() {
		this.curGroup.style.pagenumberrestart = "none";
	}
	cmd$pgnrestart() {
		this.curGroup.style.pagenumberrestart = "onpagenumber";
	}
	cmd$pgnx(val) {
		this.curGroup.style.pagenumberright = val;
	}
	cmd$pgny(val) {
		this.curGroup.style.pagenumbertop = val;
	}
	cmd$pgndec() {
		this.curGroup.style.pagenumbering = "DECIMAL";
	}
	cmd$pgnucrm() {
		this.curGroup.style.pagenumbering = "I";
	}
	cmd$pgnlcrm() {
		this.curGroup.style.pagenumbering = "i";
	}
	cmd$pgnucltr() {
		this.curGroup.style.pagenumbering = "A";
	}
	cmd$pgnlcltr() {
		this.curGroup.style.pagenumbering = "a";
	}
	cmd$pgnbidia() {
		this.curGroup.style.pagenumbering = "BIDIA";
	}
	cmd$pgnbidib() {
		this.curGroup.style.pagenumbering = "BIDIB";
	}
	cmd$pgnchosung() {
		this.curGroup.style.pagenumbering = "CHOSUNG";
	}
	cmd$pgncnum() {
		this.curGroup.style.pagenumbering = "CIRCLENUM";
	}
	cmd$pgndbnum() {
		this.curGroup.style.pagenumbering = "KANJINODIGIT";
	}
	cmd$pgndbnumd() {
		this.curGroup.style.pagenumbering = "KANJIDIGIT";
	}
	cmd$pgndbnumt() {
		this.curGroup.style.pagenumbering = "DBNUM3";
	}
	cmd$pgndbnumk() {
		this.curGroup.style.pagenumbering = "DBNUM4";
	}
	cmd$pgndecd() {
		this.curGroup.style.pagenumbering = "DOUBLEBYTE";
	}
	cmd$pgnganada() {
		this.curGroup.style.pagenumbering = "GANADA";
	}
	cmd$pgngbnum() {
		this.curGroup.style.pagenumbering = "GB1";
	}
	cmd$pgngbnumd() {
		this.curGroup.style.pagenumbering = "GB2";
	}
	cmd$pgngbnuml() {
		this.curGroup.style.pagenumbering = "GB3";
	}
	cmd$pgngbnumk() {
		this.curGroup.style.pagenumbering = "GB4";
	}
	cmd$pgnzodiac() {
		this.curGroup.style.pagenumbering = "ZODIAC1";
	}
	cmd$pgnzodiacd() {
		this.curGroup.style.pagenumbering = "ZODIAC2";
	}
	cmd$pgnzodiacl() {
		this.curGroup.style.pagenumbering = "ZODIAC3";
	}
	cmd$pgnhindia() {
		this.curGroup.style.pagenumbering = "HINDIVOWEL";
	}	
	cmd$pgnhindib() {
		this.curGroup.style.pagenumbering = "HINDICONSONANT";
	}
	cmd$pgnhindic() {
		this.curGroup.style.pagenumbering = "HINDIDIGIT";
	}
	cmd$pgnhindid() {
		this.curGroup.style.pagenumbering = "HINDIDESCRIPTIVE";
	}
	cmd$pgnthaia() {
		this.curGroup.style.pagenumbering = "THAILETTER";
	}
	cmd$pgnthaib() {
		this.curGroup.style.pagenumbering = "THAIDIGIT";
	}
	cmd$pgnthaic() {
		this.curGroup.style.pagenumbering = "THAIDESCRIPTIVE";
	}
	cmd$pgnvieta() {
		this.curGroup.style.pagenumbering = "VIETNAMESEDESCRIPTIVE";
	}
	cmd$pgnid() {
		this.curGroup.style.pagenumbering = "KOREANDASH";
	}
	cmd$pgnhn(val) {
		this.curGroup.style.pagenumberheaderlevel = val;
	}
	cmd$pgnhnsh() {
		this.curGroup.style.pagenumberseparator = "-";
	}
	cmd$pgnhnsp() {
		this.curGroup.style.pagenumberseparator = ".";
	}
	cmd$pgnhnsc() {
		this.curGroup.style.pagenumberseparator = ":";
	}
	cmd$pgnhnsm() {
		this.curGroup.style.pagenumberseparator = "—";
	}
	cmd$pgnhnsn() {
		this.curGroup.style.pagenumberseparator = "–";
	}

	cmd$vertal() {
		this.curGroup.style.textalign = "bottom"; //Alias for vertalb. Why? Ask Microsoft.
	}
	cmd$vertalt() {
		this.curGroup.style.textalign = "top";
	}
	cmd$vertalb() {
		this.curGroup.style.textalign = "bottom";
	}
	cmd$vertalc() {
		this.curGroup.style.textalign = "center";
	}
	cmd$vertalj() {
		this.curGroup.style.textalign = "justified";
	}

	cmd$srauth(val) {
		this.curGroup.attributes.revisionauthor = val;
	}
	cmd$srdate(val) {
		this.curGroup.attributes.revisiondate = val;
	}

	cmd$rtlsect() {
		this.curGroup.style.snakecolumns = "rtl";
	}
	cmd$ltrsect() {
		this.curGroup.style.snakecolumns = "ltr";
	}

	cmd$horzsect() {
		this.curGroup.style.renderdirection = "horizontal";
	}
	cmd$vertsect() {
		this.curGroup.style.renderdirection = "vertical";
	}

	cmd$stextflow(val) {
		this.curGroup.style.textflow = val;
	}

	cmd$sectexpand(val) {
		this.curGroup.style.charspacebase = val;
	}
	cmd$sectlinegrid(val) {
		this.curGroup.style.linegrid = val;
	}
	cmd$sectdefaultcl() {
		this.curGroup.style.specify = false;
	}
	cmd$sectspecifycl() {
		this.curGroup.style.specify = "line";
	}
	cmd$sectspecifyl() {
		this.curGroup.style.specify = "both";
	}
	cmd$sectspecifygenN() {
		this.curGroup.style.chargridsnap = true;
	}
	
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
		this.curGroup.style.foreground = this.doc.tables.colourtable[val - 1];
	}
	cmd$cb(val) {
		this.curGroup.style.background = this.doc.tables.colourtable[val - 1];
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