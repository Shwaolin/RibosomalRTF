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
	ATab,
	Pn,
	Field,
	Fldrslt,
	Picture,
	DateGroup,
	NonGroup
} = require("./RTFGroups.js");

const util = require('util');

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
			fontSize:22,
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
		this.lastPar = {style:this.defParState, attributes:{}};
		this.lastSect = {style:{}, attributes:{}};
		this.doc = {};
		this.curGroup = {};
		this.paraTypes = ["paragraph", "list-item", "list-text"];
		this.textTypes = ["text", "list-text", "field", "fragment"];
		this.curRow = {style:{}, attributes:{}};
	}
	followInstruction(instruction) {
		//console.log(instruction)
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
				if (this.curGroup.type === "fragment") {
					this.endGroup();
				} else if (this.curGroup.type === "paragraph") {
					this.endGroup();
					this.newGroup("paragraph");
				}
				break;
			case "documentStart":
				this.doc = new RTFDoc;
				this.curGroup = this.doc;
				break;
			case "documentEnd":
				while (this.curGroup !== this.doc) {this.endGroup();}
				this.output = this.doc.dumpContents();
				break;
		}
	}
	parseControl(instruction) {
		if (typeof this.curGroup.type === "undefined") {
			this.newGroup("fragment");
		}
		if (this.curGroup.parent instanceof Stylesheet && !(this.curGroup instanceof Style)) {
			this.curGroup = new Style(this.curGroup.parent, instruction);
		} else {
			let val = null;
			if (instruction.substr(0, 3) === 'hex') {
				val = instruction.substr(3);
				instruction = 'hex';
			}
			else {
				const numPos = instruction.search(/\d|\-/);
				if (numPos !== -1) {
					val = parseFloat(instruction.substr(numPos).replace(/,/g,""));
					instruction = instruction.substr(0,numPos);
				}
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
		if (this.curGroup.dumpContents) {this.curGroup.dumpContents();}

		if (this.curGroup.type === "paragraph") {
			this.lastPar.style = this.curGroup.style;
			this.lastPar.attributes = this.curGroup.attributes;
		} else if (this.curGroup.type === "section") {
			this.lastSect.style = this.curGroup.style;
			this.lastSect.attributes = this.curGroup.attributes;
		}


		if (this.curGroup.parent) {
			this.curGroup = this.curGroup.parent;
		} else {
			this.curGroup = this.doc;
		}


	}

	/* TEMPORARY, FOR DEBUGGING */
	cmd$datastore(val) {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}

	/* Header */
	cmd$rtf(val) {
		this.doc.attributes.rtfVersion = val;
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
		this.doc.attributes.ansiPg = val;
	}
	cmd$fbidis() {
		this.doc.attributes.fBidis = true;
	}

	/* Default Fonts and Languages */
	cmd$fromtext() {
		this.doc.attributes.fromText = true;
	}
	cmd$fromhtml(val) {
		this.doc.attributes.fromHTML = val;
	}
	cmd$deff(val) {
		this.doc.attributes.defaultFont = val;
	}
	cmd$adeff(val) {
		this.doc.attributes.defaultBidiFont = val;
	}
	cmd$stshfdbch(val) {
		this.doc.attributes.defaultEastAsian = val;
	}
	cmd$stshfloch(val) {
		this.doc.attributes.defaultASCII = val;
	}
	cmd$stshfhich(val) {
		this.doc.attributes.defaultHighANSI = val;
	}
	cmd$stshfbi(val) {
		this.doc.attributes.defaultBidi = val;
	}
	cmd$deflang(val) {
		this.doc.attributes.defaultLanguage = val;
	}
	cmd$deflangfe(val) {
		this.doc.attributes.defaultLanguageEastAsia = val;
	}
	cmd$adeflang(val) {
		this.doc.attributes.defaultLanguageSouthAsia = val;
	}

	/*Themes */
	cmd$themedata() {
		this.curGroup = new ParameterGroup(this.doc, "themeData", false, this.curGroup.parent);
	}
	cmd$colorschememapping() {
		this.curGroup = new ParameterGroup(this.doc, "colorSchemeMapping", false, this.curGroup.parent);
	}
	cmd$flomajor() {
		this.doc.attributes.fMajor = "ascii";
	}
	cmd$fhimajor() {
		this.doc.attributes.fMajor = "default";
	}
	cmd$fdbmajor() {
		this.doc.attributes.fMajor = "east-asian";
	}
	cmd$fbimajor() {
		this.doc.attributes.fMajor = "complex-scripts";
	}
	cmd$flominor() {
		this.doc.attributes.fMinor = "ascii";
	}
	cmd$fhiminor() {
		this.doc.attributes.fMinor = "default";
	}
	cmd$fdbminor() {
		this.doc.attributes.fMinor = "east-asian";
	}
	cmd$fbiminor() {
		this.doc.attributes.fMinor = "complex-scripts";
	}

	/* Code Page */
	cmd$cpg(val) {
		this.curGroup.attributes.codePage = val;
	}

	/* File Table */
	cmd$filetbl() {
		this.curGroup = new FileTable(this.doc, this.curGroup.parent);
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
		this.curGroup.attributes.osNumber = val;
	}
	cmd$fvalidmac() {
		this.curGroup.attributes.fileSystem = "mac";
	}
	cmd$fvaliddos() {
		this.curGroup.attributes.fileSystem = "ms-dos";
	}
	cmd$fvalidntfs() {
		this.curGroup.attributes.fileSystem = "ntfs";
	}
	cmd$fvalidhpfs() {
		this.curGroup.attributes.fileSystem = "hpfs";
	}
	cmd$fnetwork() {
		this.curGroup.attributes.networkFileSystem = true;
	}
	cmd$fnonfilesys() {
		this.curGroup.attributes.nonFileSys = true;
	}

	/* Colour Table */
	cmd$colortbl() {
		this.curGroup = new ColourTable(this.doc, this.curGroup.parent);
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
		this.curGroup.attributes.themeColour = "main-dark-one";
	}
	cmd$cmaindarktwo() {
		this.curGroup.attributes.themeColour = "main-dark-two";
	}
	cmd$cmainlightone() {
		this.curGroup.attributes.themeColour = "main-light-one";
	}
	cmd$cmainlighttwo() {
		this.curGroup.attributes.themeColour = "main-light-two";
	}
	cmd$caccentone() {
		this.curGroup.attributes.themeColour = "accent-one";
	}
	cmd$caccenttwo() {
		this.curGroup.attributes.themeColour = "accent-two";
	}
	cmd$caccentthree() {
		this.curGroup.attributes.themeColour = "accent-three";
	}
	cmd$caccentfour() {
		this.curGroup.attributes.themeColour = "accent-four";
	}
	cmd$caccentfive() {
		this.curGroup.attributes.themeColour = "accent-five";
	}
	cmd$caccentsix() {
		this.curGroup.attributes.themeColour = "accent-six";
	}
	cmd$chyperlink() {
		this.curGroup.attributes.themeColour = "hyperlink";
	}
	cmd$cfollowedhyperlink() {
		this.curGroup.attributes.themeColour = "followed-hyperlink";
	}
	cmd$cbackgroundone() {
		this.curGroup.attributes.themeColour = "background-one";
	}
	cmd$cbackgroundtwo() {
		this.curGroup.attributes.themeColour = "background-two";
	}
	cmd$ctextone() {
		this.curGroup.attributes.themeColour = "text-one";
	}
	cmd$ctexttwo() {
		this.curGroup.attributes.themeColour = "text-two";
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
		this.curGroup = new Stylesheet(this.doc, this.curGroup.parent);
	}
	cmd$tsrowd() {
		this.curGroup.attributes.tsRowd = true;
	}
	cmd$additive() {
		this.curGroup.attributes.additive = true;
	}
	cmd$sbasedon(val) {
		this.curGroup.attributes.basedOn = val;
	}
	cmd$snext(val) {
		this.curGroup.attributes.next = val;
	}
	cmd$sautoupd() {
		this.curGroup.attributes.autoUpdate = true;
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
		this.curGroup.attributes.emailStyle = "personal";
	}
	cmd$scompose() {
		this.curGroup.attributes.emailStyle = "compose";
	}
	cmd$reply() {
		this.curGroup.attributes.emailStyle = "reply";
	}
	cmd$styrsid(val) {
		this.curGroup.attributes.rsid = val;
	}
	cmd$ssemihidden(val) {
		if (val === null) {val = 0}
		this.curGroup.attributes.semiHidden = val;
	}
	cmd$keycode() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "keyCode");
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
		this.curGroup.attributes.unhideUsed = val;
	}

	cmd$s(val) {
		this.curGroup.attributes.styleDesignation = "s" + val;
	}
	cmd$cs(val) {
		this.curGroup.attributes.styleDesignation = "cs" + val;
	}
	cmd$ds(val) {
		this.curGroup.attributes.styleDesignation = "ds" + val;
	}
	cmd$ts(val) {
		this.curGroup.attributes.styleDesignation = "ts" + val;
	}

	cmd$noqfpromote() {
		this.doc.attributes.noQfPromote = true;
	}

	/* Table Styles */
	cmd$tscellwidth(val) {
		this.curGroup.style.cellWidth = val;
	}
	cmd$tscellwidthfts(val) {
		this.curGroup.style.cellWidthFts = val;
	}
	cmd$tscellpaddt(val) {
		this.curGroup.style.topPadding = val;
	}
	cmd$tscellpaddl(val) {
		this.curGroup.style.leftPadding = val;
	}
	cmd$tscellpaddr(val) {
		this.curGroup.style.rightPadding = val;
	}
	cmd$tscellpaddb(val) {
		this.curGroup.style.bottomPadding = val;
	}
	cmd$tscellpaddft(val) {
		this.curGroup.style.topPaddingUnits = val;
	}
	cmd$tscellpaddfl(val) {
		this.curGroup.style.leftPaddingUnits = val;
	}
	cmd$tscellpaddfr(val) {
		this.curGroup.style.rightPaddingUnits = val;
	}
	cmd$tscellpaddfb(val) {
		this.curGroup.style.bottomPaddingUnits = val;
	}
	cmd$tsvertalt() {
		this.curGroup.style.cellAlignment = "top";
	}
	cmd$tsvertalc() {
		this.curGroup.style.cellAlignment = "center";
	}
	cmd$tsvertalb() {
		this.curGroup.style.cellAlignment = "bottom";
	}
	cmd$tsnowrap() {
		this.curGroup.style.noWrap = true;
	}
	cmd$tscellcfpat(val) {
		this.curGroup.style.foregroundShading = val;
	}
	cmd$tscellcbpat(val) {
		this.curGroup.style.backgroundShading = val;
	}
	cmd$tscellpct(val) {
		this.curGroup.style.shadingPercentage = val;
	}
	cmd$tsbgbdiag() {
		this.curGroup.style.shadingPattern = "backwards-diagonal";
	}
	cmd$tsbgfdiag() {
		this.curGroup.style.shadingPattern = "forwards-diagonal";
	}
	cmd$tsbgdkbdiag() {
		this.curGroup.style.shadingPattern = "dark-backwards-diagonal";
	}
	cmd$tsbgdkfdiag() {
		this.curGroup.style.shadingPattern = "dark-forwards-diagonal";
	}
	cmd$tsbgcross() {
		this.curGroup.style.shadingPattern = "cross";
	}
	cmd$tsbgdcross() {
		this.curGroup.style.shadingPattern = "diagonal-cross";
	}
	cmd$tsbgdkcross() {
		this.curGroup.style.shadingPattern = "dark-cross";
	}
	cmd$tsbgdkdcross() {
		this.curGroup.style.shadingPattern = "dark-diagonal-cross";
	}
	cmd$tsbghoriz() {
		this.curGroup.style.shadingPattern = "horizontal";
	}
	cmd$tsbgvert() {
		this.curGroup.style.shadingPattern = "vertical";
	}
	cmd$tsbgdkhor() {
		this.curGroup.style.shadingPattern = "dark-horizontal";
	}
	cmd$tsbgdkvert() {
		this.curGroup.style.shadingPattern = "dark-vertical";
	}
	cmd$tsbrdrt() {
		this.curGroup.style.cellBorder = "top";
	}
	cmd$tsbrdrb() {
		this.curGroup.style.cellBorder = "bottom";
	}
	cmd$tsbrdrl() {
		this.curGroup.style.cellBorder = "left";
	}
	cmd$tsbrdrr() {
		this.curGroup.style.cellBorder = "right";
	}
	cmd$tsbrdrh() {
		this.curGroup.style.cellBorder = "horizontal";
	}
	cmd$tsbrdrv() {
		this.curGroup.style.cellBorder = "vertical";
	}
	cmd$tsbrdrdgl() {
		this.curGroup.style.cellBorder = "diagonal-ul-lr";
	}
	cmd$tsbrdrdgr() {
		this.curGroup.style.cellBorder = "diagonal-ll-ur";
	}
	cmd$tscbandsh(val) {
		this.curGroup.style.rowBandCount = val;
	}
	cmd$tscbandsv(val) {
		this.curGroup.style.cellBandCount = val;
	}

	/* Style Restrictions */
	cmd$latentstyles() {
		this.curGroup = new StyleRestrictions(this.doc, this.curGroup.parent);
	}
	cmd$lsdstimax(val) {
		this.curGroup.attributes.dstiMax = val;
	}
	cmd$lsdlockeddef(val) {
		this.curGroup.attributes.lockedDef = val;
	}
	cmd$lsdlockedexcept() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "lockedExceptions");
	}
	cmd$lsdsemihiddendef(val) {
		this.curGroup.attributes.semiHiddenDefault = val;
	}
	cmd$lsdunhideuseddef(val) {
		this.curGroup.attributes.unhideUsedDefault = val;
	}
	cmd$lsdqformatdef(val) {
		this.curGroup.attributes.qFormatDefault= val;
	}
	cmd$lsdprioritydef(val) {
		this.curGroup.attributes.priorityDefault = val;
	}
	cmd$lsdpriority(val) {
		this.curGroup.attributes.priorityLatentDefault = val;
	}
	cmd$lsdunhideused(val) {
		this.curGroup.attributes.unhideUsedLatentDefault = val;
	}
	cmd$lsdsemihidden(val) {
		this.curGroup.attributes.semihiddenlatentdefault = val;
	}
	cmd$lsdqformat(val) {
		this.curGroup.attributes.qFormatLatentDefault = val;
	}
	cmd$lsdlocked(val) {
		this.curGroup.attributes.lockedLatentDefault = val;
	}

	/* Font Table */
	cmd$fonttbl() {
		this.curGroup = new FontTable(this.doc, this.curGroup.parent);
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
		this.curGroup = new ParameterGroup(this.curGroup.parent, "alternate", "attributes");
	}
	cmd$panose() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "panose", "attributes");
	}
	cmd$fname() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "taggedName", "attributes");
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
		this.curGroup = new ListTable(this.doc, this.curGroup.parent);
	}
	/*-- List --*/
	cmd$list() {
		this.curGroup = new List(this.curGroup.parent);
	}
	cmd$listid(val) {
		this.curGroup.id = val;
	}
	cmd$listtemplateid(val) {
		this.curGroup.templateID = val;
	}
	cmd$listsimple(val) {
		this.curGroup.attributes.simple = val;
	}
	cmd$listhybrid(val) {
		this.curGroup.attributes.hybrid = true;
	}
	cmd$listname() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "listName");
	}
	cmd$liststyleid(val) {
		this.curGroup.attributes.styleID = val;
	}
	cmd$liststylename(val) {
		this.curGroup.attributes.styleName = val;
	}
	cmd$liststartat(val) {
		this.curGroup.attributes.startAt = val;
	}
	cmd$lvltentative() {
		this.curGroup.attributes.lvlTentative = true;
	}
	/*-- List Level --*/
	cmd$listlevel() {
		this.curGroup = new ListLevel(this.curGroup.parent);
	}
	cmd$levelstartat(val) {
		this.curGroup.attributes.startAt = val;
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
		this.curGroup = new ParameterGroup(this.curGroup.parent, "levelText");
	}
	cmd$levelnumbers(val) {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "levelNumbers");
	}
	cmd$levelfollow(val) {
		this.curGroup.attributes.follow = val;
	}
	cmd$levellegal(val) {
		this.curGroup.attributes.legal = val;
	}
	cmd$levelnorestart(val) {
		this.curGroup.attributes.noRestart = val;
	}
	cmd$levelold(val) {
		this.curGroup.attributes.old = val;
	}
	cmd$levelprev(val) {
		this.curGroup.attributes.prev = val;
	}
	cmd$levelprevspace(val) {
		this.curGroup.attributes.prevSpace = val;
	}
	cmd$levelindent(val) {
		this.curGroup.attributes.indent = val;
	}
	cmd$levelspace(val) {
		this.curGroup.attributes.space = val;
	}

	/* List Override Table */
	cmd$listoverridetable() {
		this.curGroup = new ListOverrideTable(this.doc, this.curGroup.parent);
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
		this.curGroup.attributes.overrideCount = val;
	}
	cmd$listoverridestartat() {
		this.curGroup.attributes.overrideStartAt = true;
	}
	cmd$listoverrideformat(val) {
		this.curGroup.attributes.overrideFormat = val;
	}

	/* Paragraph Group Properties */
	cmd$pgptbl() {
		this.curGroup = new ParagraphGroupTable(this.doc, this.curGroup.parent);
	}
	cmd$pgp() {
		this.curGroup = new ParagraphGroup(this.curGroup.parent, this.curGroup.parent);
	}
	cmd$ipgp(val) {
		this.curGroup.attributes.id = val;
	}

	/* Revision Marks */
	cmd$revtbl() {
		this.curGroup = new RevisionTable(this.doc, this.curGroup.parent);
	}

	/* RSID */
	cmd$rsidtbl() {
		this.curGroup = new RSIDTable(this.doc, this.curGroup.parent);
	}
	cmd$rsid(val) {
		this.curGroup.parent.table.push(val);
	}
	cmd$insrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "insert";
	}
	cmd$rsidroot(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "root";
	}
	cmd$delrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "delete";
	}
	cmd$charrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "character-format";
	}
	cmd$sectrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "section-format";
	}
	cmd$pararsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "paragraph-format";
	}
	cmd$tblrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "table-format";
	}

	/* Old Properties */
	cmd$oldcProps() {
		this.curGroup.type = "old-c-prop";
	}
	cmd$oldpProps() {
		this.curGroup.type = "old-p-prop";
	}
	cmd$oldtProps() {
		this.curGroup.type = "old-t-prop";
	}
	cmd$oldsProps() {
		this.curGroup.type = "old-s-prop";
	}

	/* User Protection Information */
	cmd$protusertbl() {
		this.curGroup = new ProtectedUsersTable(this.doc, this.curGroup.parent);
	}

	/* Generator */
	cmd$generator() {
		this.curGroup = new ParameterGroup(this.doc, "generator", "attributes");
	}

	/* Information */
	cmd$info() {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}
	cmd$title() {
		this.curGroup = new ParameterGroup(this.doc, "title", "attributes");
	}
	cmd$subject() {
		this.curGroup = new ParameterGroup(this.doc, "subject", "attributes");
	}
	cmd$author() {
		this.curGroup = new ParameterGroup(this.doc, "author", "attributes");
	}
	cmd$manager() {
		this.curGroup = new ParameterGroup(this.doc, "manager", "attributes");
	}
	cmd$company() {
		this.curGroup = new ParameterGroup(this.doc, "company", "attributes");
	}
	cmd$operator() {
		this.curGroup = new ParameterGroup(this.doc, "operator", "attributes");
	}
	cmd$category() {
		this.curGroup = new ParameterGroup(this.doc, "category", "attributes");
	}
	cmd$keywords() {
		this.curGroup = new ParameterGroup(this.doc, "keywords", "attributes");
	}
	cmd$comment() {
		this.curGroup = new ParameterGroup(this.doc, "comment", "attributes");
	}
	cmd$version(val) {
		this.doc.attributes.version = val;
	}
	cmd$title() {
		this.curGroup = new ParameterGroup(this.doc, "docComment", "attributes");
	}
	cmd$hlinkbase() {
		this.curGroup = new ParameterGroup(this.doc, "hlinkBase", "attributes");
	}
	/*-- User Properties --*/
	cmd$userprops() {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}
	cmd$propname(val) {
		this.curGroup.parent = new UserProperty(this.doc.attributes);
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertyName");
	}
	cmd$proptype(val) {
		this.curGroup.propertyType = val;
	}
	cmd$staticval() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertyValue");
	}
	cmd$linkval() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "propertyLink");
	}

	cmd$vern(val) {
		this.doc.attributes.initialVersion = val;
	}
	cmd$creatim() {
		this.curGroup = new DateGroup(this.doc, "createTime", "attributes");
	}
	cmd$revtim() {
		this.curGroup = new DateGroup(this.doc, "revisionTime", "attributes");
	}
	cmd$printtim() {
		this.curGroup = new DateGroup(this.doc, "lastPrintTime", "attributes");
	}
	cmd$buptim() {
		this.curGroup = new DateGroup(this.doc, "backupTime", "attributes");
	}
	cmd$edmins(val) {
		this.doc.attributes.editingMinutes = val;
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
		this.doc.attributes.charsNoSpaces = val;
	}
	cmd$id(val) {
		this.doc.attributes.id = val;
	}
	/*-- Dates --*/
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
		this.curGroup = new ParameterGroup(this.doc, "passwordHash", "attributes");
	}

	/* XML Namespace Table */
	cmd$xmlnstbl() {
		this.curGroup = new XMLNamespaceTable(this.doc, this.curGroup.parent);
	}
	cmd$xmlns(val) {
		this.curGroup = new XMLNamespace(this.curGroup.parent, val);
	}

	/* Document Formatting Properties */
	cmd$deftab(val) {
		this.curGroup.style.defaultTab = val;
	}
	cmd$hyphhotz(val) {
		this.curGroup.style.hyphenHotzone = val;
	}
	cmd$hyphconsec(val) {
		this.curGroup.style.hyphenConsecutive = val;
	}
	cmd$hyphcaps(val) {
		this.curGroup.style.hyphenCaps = val !== 0;
	}
	cmd$hyphauto(val) {
		this.curGroup.style.hyphenAuto = val !== 0;
	}
	cmd$linestart(val) {
		this.curGroup.style.lineStart = val;
	}
	cmd$fracwidth(val) {
		this.curGroup.style.fractionalWidths = true;
	}
	cmd$nextfile() {
		this.curGroup = new ParameterGroup(this.doc, "nextFile", "attributes");
	}
	cmd$template() {
		this.curGroup = new ParameterGroup(this.curGroup, "template", "style");
	}
	cmd$makebackup() {
		this.doc.attributes.makeBackup = true;
	}
	cmd$muser() {
		this.doc.attributes.compatability = true;
	}
	cmd$defformat() {
		this.doc.attributes.defFormat = true;
	}
	cmd$psover() {
		this.doc.attributes.psOver = true;
	}
	cmd$doctemp() {
		this.doc.attributes.boilerplate = true;
	}
	cmd$windowcaption() {
		this.curGroup = new ParameterGroup(this.doc, "caption", "attributes");
	}
	cmd$doctype(val) {
		this.doc.attributes.doctype = val;
	}
	cmd$ilfomacatclnup(val) {
		this.doc.attributes.incompleteCleanup = val !== 0;
	}
	cmd$horzdoc() {
		this.curGroup.style.rendering = "horizontal";
	}
	cmd$vertdoc() {
		this.curGroup.style.rendering = "vertical";
	}
	cmd$jcompress() {
		this.curGroup.style.justification = "compressing";
	}
	cmd$jexpand() {
		this.curGroup.style.justification = "expanding";
	}
	cmd$lnongrid() {
		this.curGroup.style.lineOnGrid = true;
	}
	cmd$grfdocevents(val) {
		this.doc.attributes.grfDocEvents = val;
	}
	cmd$themelang(val) {
		this.curGroup.style.themeLanguage = val;
	}
	cmd$themelangfe(val) {
		this.curGroup.style.themeLanguageFE = val;
	}
	cmd$themelangcs(val) {
		this.curGroup.style.themeLanguageCS = val;
	}
	cmd$relyonvml(val) {
		this.doc.attributes.vml = val;
	}
	cmd$validatexml(val) {
		this.doc.attributes.validateXML = val;
	}
	cmd$xform() {
		this.curGroup = new ParameterGroup(this.doc, "xForm", "attributes");
	}
	cmd$donotembedsysfont(val) {
		this.doc.attributes.doNotEmbedSysFont = val;
	}
	cmd$donotembedlingdata(val) {
		this.doc.attributes.doNotEmbedLingData = val;
	}
	cmd$showplaceholdtext(val) {
		this.doc.attributes.showPlaceholderText = val;
	}
	cmd$trackmoves(val) {
		this.doc.attributes.trackMoves = val;
	}
	cmd$trackformatting(val) {
		this.doc.attributes.trackFormatting = val;
	}
	cmd$ignoreremixedcontent(val) {
		this.doc.attributes.ignoreRemixedContent = val;
	}
	cmd$saveinvalidxml(val) {
		this.doc.attributes.saveInvalidXML = val;
	}
	cmd$showxmlerrors(val) {
		this.doc.attributes.showXMLErrors = val;
	}
	cmd$stylelocktheme(val) {
		this.curGroup.style.lockTheme = true;
	}
	cmd$stylelockqfset(val) {
		this.curGroup.style.lockQfSet = true;
	}
	cmd$usenormstyforlist(val) {
		this.curGroup.style.useNormalListStyle = true;
	}
	cmd$wgrffmtfilter(val) {
		if (val) {
			this.curGroup.style.styleFilters = val;
		} else {
			this.curGroup = new ParameterGroup(this.doc, "styleFilters", "style");
		}
	}
	cmd$readonlyrecommended() {
		this.curGroup.style.readOnlyRecommended = true;
	}
	cmd$stylesortmethod(val) {
		this.curGroup.style.styleSortMethod = val;
	}
	cmd$writereservhash() {
		this.curGroup = new ParameterGroup(this.doc, "reserveHash", "attributes");
	}
	cmd$writereservation() {
		this.curGroup = new ParameterGroup(this.doc, "reservation", "attributes");
	}
	cmd$saveprevpict() {
		this.doc.attributes.savePrevPict = true;
	}
	cmd$viewkind(val) {
		this.curGroup.style.viewKind = val;
	}
	cmd$viewscale(val) {
		this.curGroup.style.viewScale = val;
	}
	cmd$viewzk(val) {
		this.curGroup.style.viewZK = val;
	}
	cmd$viewbksp(val) {
		this.curGroup.style.viewBksp = val;
	}

	/*-- Footnotes and Endnotes --*/
	cmd$fet(val) {
		this.curGroup.style.fet = val;
	}
	cmd$ftnsep() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footnoteSep", "style");
	}
	cmd$ftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footnotesEpc", "style");
	}
	cmd$ftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footnoteNotice", "style");
	}
	cmd$aftnsep() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "pageSep", "style");
	}
	cmd$aftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "pageSepc", "style");
	}
	cmd$aftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "pageNotice", "style");
	}
	cmd$pages() {
		this.doc.attributes.footnoteposition = "pages";
	}
	cmd$enddoc() {
		this.doc.attributes.footnoteposition = "end-doc";
	}
	cmd$ftntj() {
		this.doc.attributes.footnoteposition = "beneath-text-tj";
	}
	cmd$ftnbj() {
		this.doc.attributes.footnoteposition = "bottom-of-page-bj";
	}
	cmd$apages() {
		this.doc.attributes.pageposition = "pages";
	}
	cmd$aenddoc() {
		this.doc.attributes.pageposition = "endDoc";
	}
	cmd$aftntj() {
		this.doc.attributes.pageposition = "beneath-text-tj";
	}
	cmd$aftnbj() {
		this.doc.attributes.pageposition = "bottom-of-page-bj";
	}
	cmd$ftnstart(val) {
		this.curGroup.style.footnoteStart = val;
	}
	cmd$aftnstart(val) {
		this.curGroup.style.pageStart = val;
	}
	cmd$ftnrstpg() {
		this.curGroup.style.footnoteRestart = "page";
	}
	cmd$ftnrestart() {
		this.curGroup.style.footnoteRestart = "section";
	}
	cmd$ftnrstcont() {
		this.curGroup.style.footnoteRestart = "none";
	}
	cmd$aftnrestart() {
		this.curGroup.style.pageRestart = "section";
	}
	cmd$aftnrstcont() {
		this.curGroup.style.pageRestart = "none";
	}
	cmd$ftnnar() {
		this.curGroup.style.footnoteNumbering = "1";
	}
	cmd$ftnnalc() {
		this.curGroup.style.footnoteNumbering = "a";
	}
	cmd$ftnnauc() {
		this.curGroup.style.footnoteNumbering = "A";
	}
	cmd$ftnnrlc() {
		this.curGroup.style.footnoteNumbering = "i";
	}
	cmd$ftnnruc() {
		this.curGroup.style.footnoteNumbering = "I";
	}
	cmd$ftnnchi() {
		this.curGroup.style.footnoteNumbering = "*";
	}
	cmd$ftnnchosung() {
		this.curGroup.style.footnoteNumbering = "CHOSUNG";
	}
	cmd$ftnncnum() {
		this.curGroup.style.footnoteNumbering = "CIRCLENUM";
	}
	cmd$ftnndbnum() {
		this.curGroup.style.footnoteNumbering = "DBNUM1";
	}
	cmd$ftnndbnumd() {
		this.curGroup.style.footnoteNumbering = "DBNUM2";
	}
	cmd$ftnndbnumt() {
		this.curGroup.style.footnoteNumbering = "DBNUM3";
	}
	cmd$ftnndbnumk() {
		this.curGroup.style.footnoteNumbering = "DBNUM4";
	}
	cmd$ftnndbar() {
		this.curGroup.style.footnoteNumbering = "DBCHAR";
	}
	cmd$ftnnganada() {
		this.curGroup.style.footnoteNumbering = "GANADA";
	}
	cmd$ftnngbnum() {
		this.curGroup.style.footnoteNumbering = "GB1";
	}
	cmd$ftnngbnumd() {
		this.curGroup.style.footnoteNumbering = "GB2";
	}
	cmd$ftnngbnuml() {
		this.curGroup.style.footnoteNumbering = "GB3";
	}
	cmd$ftnngbnumk() {
		this.curGroup.style.footnoteNumbering = "GB4";
	}
	cmd$ftnnzodiac() {
		this.curGroup.style.footnoteNumbering = "ZODIAC1";
	}
	cmd$ftnnzodiacd() {
		this.curGroup.style.footnoteNumbering = "ZODIAC2";
	}
	cmd$ftnnzodiacl() {
		this.curGroup.style.footnoteNumbering = "ZODIAC3";
	}
	cmd$aftnnar() {
		this.curGroup.style.pageNumbering = "1";
	}
	cmd$aftnnalc() {
		this.curGroup.style.pageNumbering = "a";
	}
	cmd$aftnnauc() {
		this.curGroup.style.pageNumbering = "A";
	}
	cmd$aftnnrlc() {
		this.curGroup.style.pageNumbering = "i";
	}
	cmd$aftnnruc() {
		this.curGroup.style.pageNumbering = "I";
	}
	cmd$aftnnchi() {
		this.curGroup.style.pageNumbering = "*";
	}
	cmd$aftnnchosung() {
		this.curGroup.style.pageNumbering = "CHOSUNG";
	}
	cmd$aftnncnum() {
		this.curGroup.style.pageNumbering = "CIRCLENUM";
	}
	cmd$aftnndbnum() {
		this.curGroup.style.pageNumbering = "DBNUM1";
	}
	cmd$aftnndbnumd() {
		this.curGroup.style.pageNumbering = "DBNUM2";
	}
	cmd$aftnndbnumt() {
		this.curGroup.style.pageNumbering = "DBNUM3";
	}
	cmd$aftnndbnumk() {
		this.curGroup.style.pageNumbering = "DBNUM4";
	}
	cmd$aftnndbar() {
		this.curGroup.style.pageNumbering = "DBCHAR";
	}
	cmd$aftnnganada() {
		this.curGroup.style.pageNumbering = "GANADA";
	}
	cmd$aftnngbnum() {
		this.curGroup.style.pageNumbering = "GB1";
	}
	cmd$aftnngbnumd() {
		this.curGroup.style.pageNumbering = "GB2";
	}
	cmd$aftnngbnuml() {
		this.curGroup.style.pageNumbering = "GB3";
	}
	cmd$aftnngbnumk() {
		this.curGroup.style.pageNumbering = "GB4";
	}
	cmd$aftnnzodiac() {
		this.curGroup.style.pageNumbering = "ZODIAC1";
	}
	cmd$aftnnzodiacd() {
		this.curGroup.style.pageNumbering = "ZODIAC2";
	}
	cmd$aftnnzodiacl() {
		this.curGroup.style.pageNumbering = "ZODIAC3";
	}

	/*-- Page Information --*/
	cmd$paperw(val) {
		this.curGroup.style.paperWidth = val;
	}
	cmd$paperh(val) {
		this.curGroup.style.paperHeight = val;
	}
	cmd$margl(val) {
		this.curGroup.style.marginLeft = val;
	}
	cmd$margr(val) {
		this.curGroup.style.marginRight = val;
	}
	cmd$margt(val) {
		this.curGroup.style.marginTop = val;
	}
	cmd$margb(val) {
		this.curGroup.style.marginBottom = val;
	}
	cmd$facingp() {
		this.curGroup.style.facingPages = true;
	}
	cmd$gutter(val) {
		this.curGroup.style.gutterWidth = val;
	}
	cmd$ogutter(val) {
		this.curGroup.style.outsideGutterWidth = val;
	}
	cmd$rtlgutter() {
		this.curGroup.style.gutterRight = true;
	}
	cmd$gutterprl() {
		this.curGroup.style.gutterParallel = true;
	}
	cmd$margmirror() {
		this.curGroup.style.mirroredMargins = true;
	}
	cmd$landscape() {
		this.curGroup.style.landscape = true;
	}
	cmd$pgnstarts(val) {
		this.curGroup.style.pageNumberStart = val;
	}
	cmd$widowctrl() {
		this.curGroup.style.widowControl = true;
	}
	cmd$twoonone() {
		this.curGroup.style.twoOnOne = true;
	}
	cmd$bookfold() {
		this.curGroup.style.bookFold = true;
	}
	cmd$bookfoldrev() {
		this.curGroup.style.bookFoldReverse = true;
	}
	cmd$bookfoldsheets(val) {
		this.curGroup.style.bookFoldSheets = val;
	}

	/*-- Linked Styles --*/
	cmd$linkstyles() {
		this.doc.attributes.linkStyles = true;
	}

	/*-- Compatability Options --*/
	cmd$notabind() {
		this.curGroup.style.noTabIndent = true;
	}
	cmd$wraptrsp() {
		this.curGroup.style.wrapTrailingWhitespace = true;
	}
	cmd$prcolbl() {
		this.curGroup.style.printColourBlack = true;
	}
	cmd$noextrasprl() {
		this.curGroup.style.noExtraSpaceRl = true;
	}
	cmd$nocolbal() {
		this.curGroup.style.noColumnBalance = true;
	}
	cmd$cvmme() {
		this.curGroup.style.cvMailMergeEscape = true;
	}
	cmd$sprstsp() {
		this.curGroup.style.surpressextraline = true;
	}
	cmd$sprsspbf() {
		this.curGroup.style.surpressSpaceBefore = true;
	}
	cmd$otblrul() {
		this.curGroup.style.combineTableBorders = true;
	}
	cmd$transmf() {
		this.curGroup.style.transparentMetaFile = true;
	}
	cmd$swpbdr() {
		this.curGroup.style.swapBorder = true;
	}
	cmd$brkfrm() {
		this.curGroup.style.hardBreaks = true;
	}
	cmd$sprslnsp() {
		this.curGroup.style.surpressLineSpace = true;
	}
	cmd$subfontbysize() {
		this.curGroup.style.subFontBySize = true;
	}
	cmd$truncatefontheight() {
		this.curGroup.style.truncateFontHeight = true;
	}
	cmd$truncex() {
		this.curGroup.style.noLeadingSpace = true;
	}
	cmd$bdbfhdr() {
		this.curGroup.style.bodyBeforeHf = true;
	}
	cmd$dntblnsbdb() {
		this.curGroup.style.noBalanceSbdb = true;
	}
	cmd$expshrtn() {
		this.curGroup.style.expandCharSpace = true;
	}
	cmd$lytexcttp() {
		this.curGroup.style.noCenterLhLines = true;
	}
	cmd$lytprtmet() {
		this.curGroup.style.usePrinterMetrics = true;
	}
	cmd$msmcap() {
		this.curGroup.style.macSmallcaps = true;
	}
	cmd$nolead() {
		this.curGroup.style.noLeading = true;
	}
	cmd$nospaceforul() {
		this.curGroup.style.noUnderlineSpace = true;
	}
	cmd$noultrlspc() {
		this.curGroup.style.noUnderlineTrailing = true;
	}
	cmd$noxlattoyen() {
		this.curGroup.style.noYenTranslate = true;
	}
	cmd$oldlinewrap() {
		this.curGroup.style.oldLineWrap = true;
	}
	cmd$sprsbsp() {
		this.curGroup.style.surpressExtraSpaceB = true;
	}
	cmd$sprstsm() {
		//Does literally nothing. Why is it here? We don't know.
	}
	cmd$wpjst() {
		this.curGroup.style.wpJustify = true;
	}
	cmd$wpsp() {
		this.curGroup.style.wpSpaceWidth = true;
	}
	cmd$wptb() {
		this.curGroup.style.wpTabAdvance = true;
	}
	cmd$splytwnine() {
		this.curGroup.style.noLegacyAutoShape = true;
	}
	cmd$ftnlytwnine() {
		this.curGroup.style.noLegacyFootnote = true;
	}
	cmd$htmautsp() {
		this.curGroup.style.htmlAutoSpace = true;
	}
	cmd$useltbaln() {
		this.curGroup.style.noForgetLatTab = true;
	}
	cmd$alntblind() {
		this.curGroup.style.noIndependentRowAlign = true;
	}
	cmd$lytcalctblwd() {
		this.curGroup.style.noRawTableWidth = true;
	}
	cmd$lyttblrtgr() {
		this.curGroup.style.noTableRowApart = true;
	}
	cmd$oldas() {
		this.curGroup.style.ninetyFiveAutoSpace = true;
	}
	cmd$lnbrkrule() {
		this.curGroup.style.noLineBreakRule = true;
	}
	cmd$bdrrlswsix() {
		this.curGroup.style.useLegacyBorderRules = true;
	}
	cmd$nolnhtadjtbl() {
		this.curGroup.style.noAdjusttableLineHeight = true;
	}
	cmd$ApplyBrkRules() {
		this.curGroup.style.applyBreakRules = true;
	}
	cmd$rempersonalinfo() {
		this.curGroup.style.removePersonalInfo = true;
	}
	cmd$remdttm() {
		this.curGroup.style.removeDateTime = true;
	}
	cmd$snaptogridincell() {
		this.curGroup.style.snapTextToGrid = true;
	}
	cmd$wrppunct() {
		this.curGroup.style.hangingPunctuation = true;
	}
	cmd$asianbrkrule() {
		this.curGroup.style.asianBreakRules = true;
	}
	cmd$nobrkwrptbl() {
		this.curGroup.style.noBreakWrappedTable = true;
	}
	cmd$toplinepunct() {
		this.curGroup.style.topLinePunct = true;
	}
	cmd$viewnobound() {
		this.curGroup.style.hidePageBetweenSpace = true;
	}
	cmd$donotshowmarkup() {
		this.curGroup.style.noShowMarkup = true;
	}
	cmd$donotshowcomments() {
		this.curGroup.style.noShowComments = true;
	}
	cmd$donotshowinsdel() {
		this.curGroup.style.noShowInsDel = true;
	}
	cmd$donotshowprops() {
		this.curGroup.style.noShowFormatting = true;
	}
	cmd$allowfieldendsel() {
		this.curGroup.style.fieldEndSelect = true;
	}
	cmd$nocompatoptions() {
		this.curGroup.style.compatabilityDefaults = true;
	}
	cmd$nogrowautofit() {
		this.curGroup.style.noTableAutoFit = true;
	}
	cmd$newtblstyruls() {
		this.curGroup.style.newTableStyleRules = true;
	}
	cmd$background() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "docBackground", "style");
	}
	cmd$nouicompat() {
		this.curGroup.style.noUICompatability = true;
	}
	cmd$nofeaturethrottle(val) {
		this.curGroup.style.noFeatureThrottle = val;
	}
	cmd$forceupgrade() {
		this.curGroup.style.mayUpgrade = true;
	}
	cmd$noafcnsttbl() {
		this.curGroup.style.noTableAutoWrap = true;
	}
	cmd$noindnmbrts() {
		this.curGroup.style.bulletHangIndent = true;
	}
	cmd$felnbrelev() {
		this.curGroup.style.alternateBeginEnd = true;
	}
	cmd$indrlsweleven() {
		this.curGroup.style.ignoreFloatingObjectVector = true;
	}
	cmd$nocxsptable() {
		this.curGroup.style.noTableParSpace = true;
	}
	cmd$notcvasp() {
		this.curGroup.style.notablevectorvertical = true;
	}
	cmd$notvatxbx() {
		this.curGroup.style.noTextBoxVertical = true;
	}
	cmd$spltpgpar() {
		this.curGroup.style.splitPageParagraph = true;
	}
	cmd$hwelev() {
		this.curGroup.style.hangUlFixedWidth = true;
	}
	cmd$afelev() {
		this.curGroup.style.tableAutoFitMimic = true;
	}
	cmd$cachedcolbal() {
		this.curGroup.style.cachedColumnBalance = true;
	}
	cmd$utinl() {
		this.curGroup.style.underlineNumberedPar = true;
	}
	cmd$notbrkcnstfrctbl () {
		this.curGroup.style.noTableRowSplit = true;
	}
	cmd$krnprsnet() {
		this.curGroup.style.ansiKerning = true;
	}
	cmd$usexform() {
		this.curGroup.style.noXSLTransform = true;
	}

	/*-- Forms --*/
	cmd$formprot() {
		this.doc.attributes.protectedForms = true;
	}
	cmd$allprot() {
		this.doc.attributes.protectedAll = true;
	}
	cmd$formshade() {
		this.curGroup.style.formFieldShading = true;
	}
	cmd$formdisp() {
		this.doc.attributes.formBoxSelected = true;
	}
	cmd$formprot() {
		this.doc.attributes.printFormData = true;
	}

	/*-- Revision Marks --*/
	cmd$revprot() {
		this.doc.attributes.protectedRevisions = true;
	}
	cmd$revisions() {
		this.curGroup.style.revisions = true;
	}
	cmd$revprop(val) {
		this.curGroup.style.revisionTextDisplay = val;
	}
	cmd$revbar(val) {
		this.curGroup.style.revisionLineMarking = val;
	}

	/*-- Write Protection --*/
	cmd$readprot() {
		this.doc.attributes.protectedRead = true;
	}

	/*-- Comment Protection --*/
	cmd$annotprot() {
		this.doc.attributes.protectedComments = true;
	}

	/*-- Style Protection --*/
	cmd$stylelock() {
		this.doc.attributes.styleLock = true;
	}
	cmd$stylelockenforced() {
		this.doc.attributes.styleLockEnforced = true;
	}
	cmd$stylelockbackcomp() {
		this.doc.attributes.styleLockCompatability = true;
	}
	cmd$autofmtoverride() {
		this.doc.attributes.autoFormatLockOverride = true;
	}

	/*-- Style and Formatting Properties --*/
	cmd$enforceprot(val) {
		this.doc.attributes.enforceProtection = val;
	}
	cmd$protlevel(val) {
		this.doc.attributes.protectionLevel = val;
	}

	/*-- Tables --*/
	cmd$tsd(val) {
		this.curGroup.style.defaultTableStyle = val;
	}

	/*-- Bidirectional Controls --*/
	cmd$rtldoc() {
		this.doc.attributes.paginationDirection = "rtl";
	}
	cmd$ltrdoc() {
		this.doc.attributes.paginationDirection = "ltr";
	}

	/*-- Click and Type --*/
	cmd$cts(val) {
		this.doc.attributes.clickAndType = val;
	}

	/*-- Kinsoku Characters --*/
	cmd$jsksu() {
		this.doc.attributes.japaneseKinsoku = true;
	}
	cmd$ksulang(val) {
		this.doc.attributes.kinsokuLang = val;
	}
	cmd$fchars() {
		this.curGroup = new ParameterGroup(this.doc, "followingKinsoku", "attributes");
	}
	cmd$lchars() {
		this.curGroup = new ParameterGroup(this.doc, "leadingKinsoku", "attributes");
	}
	cmd$nojkernpunct() {
		this.doc.attributes.latinKerningOnly = true;
	}

	/*-- Drawing Grid --*/
	cmd$dghspace(val) {
		this.curGroup.style.drawGridHorizontal = val;
	}
	cmd$dgvspace(val) {
		this.curGroup.style.drawGridVertical= val;
	}
	cmd$dghorigin(val) {
		this.curGroup.style.drawGridHorizontalOrigin = val;
	}
	cmd$dgvorigin(val) {
		this.curGroup.style.drawGridVerticalOrigin = val;
	}
	cmd$dghshow(val) {
		this.curGroup.style.drawGridHorizontalShow = val;
	}
	cmd$dgvshow(val) {
		this.curGroup.style.drawGridVerticalShow = val;
	}
	cmd$dgsnap() {
		this.curGroup.style.drawGridSnap = true;
	}
	cmd$dgmargin() {
		this.curGroup.style.drawGridMargin = true;
	}

	/*-- Page Borders --*/
	cmd$pgbrdrhead() {
		this.curGroup.style.pageBorderSurroundsHeader = true;
	}
	cmd$pgbrdrfoot() {
		this.curGroup.style.pageBorderSurroundsFooter = true;
	}
	cmd$pgbrdrt() {
		this.curGroup.style.pageBorderTop = true;
	}
	cmd$pgbrdrb() {
		this.curGroup.style.pageBorderBottom = true;
	}
	cmd$pgbrdrl() {
		this.curGroup.style.pageBorderLeft = true;
	}
	cmd$pgbrdrr() {
		this.curGroup.style.pageBorderRight = true;
	}
	cmd$brdrart(val) {
		this.curGroup.style.pageBorderArt = val;
	}
	cmd$pgbrdropt(val) {
		this.curGroup.style.pageBorderOptions = val;
	}
	cmd$pgbrdrsnap() {
		this.curGroup.style.pageBorderSnap = true;
	}

	/* Mail Merge */
	cmd$mailmerge() {
		this.curGroup = new MailMergeTable(this.doc, this.curGroup.parent);
	}
	cmd$mmlinktoquery() {
		this.curGroup.attributes.linkToQuery = true;
	}
	cmd$mmdefaultsql() {
		this.curGroup.attributes.defaultSQL = true;
	}
	cmd$mmconnectstrdata() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "connectStringData", "attributes");
	}
	cmd$mmconnectstr() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "connectString", "attributes");
	}
	cmd$mmquery() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "connectStringData", "attributes");
	}
	cmd$mmdatasource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "dataSource", "attributes");
	}
	cmd$mmheadersource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "headerSource", "attributes");
	}
	cmd$mmblanklinks() {
		this.curGroup.attributes.blankLinks = true;
	}
	cmd$mmaddfieldname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "fieldName", "attributes");
	}
	cmd$mmmailsubject() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "subject", "attributes");
	}
	cmd$mmattatch() {
		this.curGroup.attributes.attach = true;
	}
	cmd$mmshowdata() {
		this.curGroup.attributes.showData = true;
	}
	cmd$mmreccur(val) {
		this.curGroup.attributes.reccur = val;
	}
	cmd$mmerrors(val) {
		this.curGroup.attributes.errorReporting = val;
	}
	cmd$mmodso() {
		this.curGroup = new Odso(this.curGroup.parent);
	}
	cmd$mmodsoudldata() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "udlData", "attributes");
	}
	cmd$mmodsoudl() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "udl", "attributes");
	}
	cmd$mmodsotable() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "table", "attributes");
	}
	cmd$mmodsosrc() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "source", "attributes");
	}
	cmd$mmodsofilter() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "filter", "attributes");
	}
	cmd$mmodsosort() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "sort", "attributes");
	}
	cmd$mmodsofldmpdata() {
		this.curGroup = new FieldMap(this.curGroup.parent);
	}
	cmd$mmodsoname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "name", "attributes");
	}
	cmd$mmodsomappedname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent, "mappedName", "attributes");
	}
	cmd$mmodsofmcolumn(val) {
		this.curGroup.attributes.columnIndex = val;
	}
	cmd$mmodsodynaddr(val) {
		this.curGroup.attributes.addressOrder = val;
	}
	cmd$mmodsosolid(val) {
		this.curGroup.attributes.language = val;
	}
	cmd$mmodsocoldelim(val) {
		this.curGroup.attributes.columnDelimiter = val;
	}
	cmd$mmjdsotype(val) {
		this.curGroup.attributes.dataSourceType = val;
	}
	cmd$mmodsofhdr(val) {
		this.curGroup.attributes.firstRowHeader = val;
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
		this.curGroup = new ParameterGroup(this.curgroup.parent, "uniqueTag", "attributes");
	}

	cmd$mmfttypenull() {
		this.curGroup.attributes.dataType = "null";
	}
	cmd$mmfttypedbcolumn() {
		this.curGroup.attributes.dataType = "database-column";
	}
	cmd$mmfttypeaddress() {
		this.curGroup.attributes.dataType = "address";
	}
	cmd$mmfttypesalutation() {
		this.curGroup.attributes.dataType = "salutation";
	}
	cmd$mmfttypemapped() {
		this.curGroup.attributes.datatype = "mapped";
	}
	cmd$mmfttypebarcode() {
		this.curGroup.attributes.datatype = "barcode";
	}

	cmd$mmdestnewdoc() {
		this.curGroup.attributes.destination = "new-document";
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
		this.curGroup.attributes.sourceType = "cataloq";
	}
	cmd$mmmaintypeenvelopes() {
		this.curGroup.attributes.sourceType = "envelope";
	}
	cmd$mmmaintypelabels() {
		this.curGroup.attributes.sourceType = "label";
	}
	cmd$mmmaintypeletters() {
		this.curGroup.attributes.sourceType = "letter";
	}
	cmd$mmmaintypeemail() {
		this.curGroup.attributes.sourceType = "email";
	}
	cmd$mmmaintypefax() {
		this.curGroup.attributes.sourceType = "fax";
	}

	cmd$mmdatatypeaccess() {
		this.curGroup.attributes.connectionType = "dde-access";
	}
	cmd$mmdatatypeexcel() {
		this.curGroup.attributes.connectionType = "dde-excel";
	}
	cmd$mmdatatypeqt() {
		this.curGroup.attributes.connectionType = "external-query";
	}
	cmd$mmdatatypeodbc() {
		this.curGroup.attributes.connectionType = "odbc";
	}
	cmd$mmdatatypeodso() {
		this.curGroup.attributes.connectionType = "odso";
	}
	cmd$mmdatatypefile() {
		this.curGroup.attributes.connectionType = "dde-text-file";
	}

	/* Sections */
	cmd$sect() {
		if (this.curGroup.type === "section") {
			this.endGroup();
		}
		this.newGroup("section");
		this.curGroup.style = this.lastSect.style;
		this.curGroup.attributes = this.lastSect.attributes;
	}
	cmd$sectd() {
		let defSectStyle = {};
		if (this.doc.style.paperWidth) {defSectStyle.paperWidth = this.doc.style.paperWidth}
		if (this.doc.style.paperHeight) {defSectStyle.paperHeight = this.doc.style.paperHeight}
		if (this.doc.style.marginLeft) {defSectStyle.marginLeft = this.doc.style.marginLeft}
		if (this.doc.style.marginRight) {defSectStyle.marginRight = this.doc.style.marginRight}
		if (this.doc.style.marginTop) {defSectStyle.marginTop = this.doc.style.marginTop}
		if (this.doc.style.marginBottom) {defSectStyle.marginBottom = this.doc.style.marginBottom}
		if (this.doc.style.gutterWidth) {defSectStyle.gutterWidth = this.doc.style.gutterWidth}
		if (this.curGroup.type === "section") {
			this.curGroup.style = defSectStyle;
		} else {
			this.newGroup("section");
			this.curGroup.style = defSectStyle;
		}
	}
	cmd$endnhere() {
		this.curGroup.attributes.pagesIncluded = true;
	}
	cmd$binfsxn(val) {
		this.curGroup.attributes.firstPrinterBin = val;
	}
	cmd$binsxn(val) {
		this.curGroup.attributes.printerBin = val;
	}
	cmd$pnseclvl(val) {
		this.curGroup = new Pn(this.curGroup.parent, val);
	}
	cmd$sectunlocked() {
		this.curGroup.attributes.unlocked = true;
	}
	/*-- Section Breaks --*/
	cmd$sbknone() {
		this.curGroup.style.sectionBreak = "none";
	}
	cmd$sbkcol() {
		this.curGroup.style.sectionBreak = "column";
	}
	cmd$sbkpage() {
		this.curGroup.style.sectionBreak = "page";
	}
	cmd$sbkeven() {
		this.curGroup.style.sectionBreak = "even";
	}
	cmd$sbkodd() {
		this.curGroup.style.sectionBreak = "odd";
	}
	/*-- Columns --*/
	cmd$cols(val) {
		this.curGroup.style.columns = val;
	}
	cmd$colsx(val) {
		this.curGroup.style.columnSpace = val;
	}
	cmd$colno(val) {
		this.curGroup.style.columnNumber = val;
	}
	cmd$colsr(val) {
		this.curGroup.style.columnRightSpace = val;
	}
	cmd$colw(val) {
		this.curGroup.style.columnWidth = val;
	}
	cmd$linebetcol() {
		this.curGroup.style.lineBetweenColumns = true;
	}
	/*-- Footnotes/Endnotes --*/
	cmd$sftntj() {
		this.curGroup.style.footnotePos = "beneath";
	}
	cmd$sftnbj() {
		this.curGroup.style.footnotePos = "bottom";
	}
	cmd$sftnstart(val) {
		this.curGroup.style.footnoteStart = val;
	}
	cmd$pgnnstart(val) {
		this.curGroup.style.pageStart = val;
	}
	cmd$sftnrstpg() {
		this.curGroup.style.footnoteRestart = "page";
	}
	cmd$sftnrestart() {
		this.curGroup.style.footnoteRestart = "section";
	}
	cmd$sftnrstcont() {
		this.curGroup.style.footnoteRestart = "none";
	}
	cmd$pgnnrstpg() {
		this.curGroup.style.pageRestart = "page";
	}
	cmd$pgnnrestart() {
		this.curGroup.style.pageRestart = "section";
	}
	cmd$pgnnrstcont() {
		this.curGroup.style.pageRestart = "none";
	}
	cmd$sftnnar() {
		this.curGroup.style.footnoteNumbering = "1";
	}
	cmd$sftnnalc() {
		this.curGroup.style.footnoteNumbering = "a";
	}
	cmd$sftnnauc() {
		this.curGroup.style.footnoteNumbering = "A";
	}
	cmd$sftnnrlc() {
		this.curGroup.style.footnoteNumbering = "i";
	}
	cmd$sftnnruc() {
		this.curGroup.style.footnoteNumbering = "I";
	}
	cmd$sftnnchi() {
		this.curGroup.style.footnoteNumbering = "*";
	}
	cmd$sftnnchosung() {
		this.curGroup.style.footnoteNumbering = "CHOSUNG";
	}
	cmd$sftnncnum() {
		this.curGroup.style.footnoteNumbering = "CIRCLENUM";
	}
	cmd$sftnndbnum() {
		this.curGroup.style.footnoteNumbering = "DBNUM1";
	}
	cmd$sftnndbnumd() {
		this.curGroup.style.footnoteNumbering = "DBNUM2";
	}
	cmd$sftnndbnumt() {
		this.curGroup.style.footnoteNumbering = "DBNUM3";
	}
	cmd$sftnndbnumk() {
		this.curGroup.style.footnoteNumbering = "DBNUM4";
	}
	cmd$sftnndbar() {
		this.curGroup.style.footnoteNumbering = "DBCHAR";
	}
	cmd$sftnnganada() {
		this.curGroup.style.footnoteNumbering = "GANADA";
	}
	cmd$sftnngbnum() {
		this.curGroup.style.footnoteNumbering = "GB1";
	}
	cmd$sftnngbnumd() {
		this.curGroup.style.footnoteNumbering = "GB2";
	}
	cmd$sftnngbnuml() {
		this.curGroup.style.footnoteNumbering = "GB3";
	}
	cmd$sftnngbnumk() {
		this.curGroup.style.footnoteNumbering = "GB4";
	}
	cmd$sftnnzodiac() {
		this.curGroup.style.footnoteNumbering = "ZODIAC1";
	}
	cmd$sftnnzodiacd() {
		this.curGroup.style.footnoteNumbering = "ZODIAC2";
	}
	cmd$sftnnzodiacl() {
		this.curGroup.style.footnoteNumbering = "ZODIAC3";
	}
	cmd$pgnnnar() {
		this.curGroup.style.pageNumbering = "1";
	}
	cmd$pgnnnalc() {
		this.curGroup.style.pageNumbering = "a";
	}
	cmd$pgnnnauc() {
		this.curGroup.style.pageNumbering = "A";
	}
	cmd$pgnnnrlc() {
		this.curGroup.style.pageNumbering = "i";
	}
	cmd$pgnnnruc() {
		this.curGroup.style.pageNumbering = "I";
	}
	cmd$pgnnnchi() {
		this.curGroup.style.pageNumbering = "*";
	}
	cmd$pgnnnchosung() {
		this.curGroup.style.pageNumbering = "CHOSUNG";
	}
	cmd$pgnnncnum() {
		this.curGroup.style.pageNumbering = "CIRCLENUM";
	}
	cmd$pgnnndbnum() {
		this.curGroup.style.pageNumbering = "DBNUM1";
	}
	cmd$pgnnndbnumd() {
		this.curGroup.style.pageNumbering = "DBNUM2";
	}
	cmd$pgnnndbnumt() {
		this.curGroup.style.pageNumbering = "DBNUM3";
	}
	cmd$pgnnndbnumk() {
		this.curGroup.style.pageNumbering = "DBNUM4";
	}
	cmd$pgnnndbar() {
		this.curGroup.style.pageNumbering = "DBCHAR";
	}
	cmd$pgnnnganada() {
		this.curGroup.style.pageNumbering = "GANADA";
	}
	cmd$pgnnngbnum() {
		this.curGroup.style.pageNumbering = "GB1";
	}
	cmd$pgnnngbnumd() {
		this.curGroup.style.pageNumbering = "GB2";
	}
	cmd$pgnnngbnuml() {
		this.curGroup.style.pageNumbering = "GB3";
	}
	cmd$pgnnngbnumk() {
		this.curGroup.style.pageNumbering = "GB4";
	}
	cmd$pgnnnzodiac() {
		this.curGroup.style.pageNumbering = "ZODIAC1";
	}
	cmd$pgnnnzodiacd() {
		this.curGroup.style.pageNumbering = "ZODIAC2";
	}
	cmd$pgnnnzodiacl() {
		this.curGroup.style.pageNumbering = "ZODIAC3";
	}

	cmd$linemod(val) {
		this.curGroup.style.lineModulus = val;
	}
	cmd$linex(val) {
		this.curGroup.style.lineDistance = val;
	}
	cmd$linestarts(val) {
		this.curGroup.style.lineStarts = val;
	}
	cmd$linerestart() {
		this.curGroup.style.lineRestart = "on-line-starts";
	}
	cmd$lineppage() {
		this.curGroup.style.lineRestart = "page";
	}
	cmd$lineppage() {
		this.curGroup.style.lineRestart = "none";
	}

	cmd$pgwsxn(val) {
		this.curGroup.style.pageWidth = val;
	}
	cmd$pghsxn(val) {
		this.curGroup.style.pageHeight = val;
	}
	cmd$marglsxn(val) {
		this.curGroup.style.marginLeft = val;
	}
	cmd$margrsxn(val) {
		this.curGroup.style.marginRight = val;
	}
	cmd$margtsxn(val) {
		this.curGroup.style.marginTop = val;
	}
	cmd$margbsxn(val) {
		this.curGroup.style.marginBottom = val;
	}
	cmd$guttersxn(val) {
		this.curGroup.style.gutterWidth = val;
	}
	cmd$margmirsxn() {
		this.curGroup.style.marginSwap = true;
	}
	cmd$lndscpsxn() {
		this.curGroup.style.landscape = true;
	}
	cmd$titlepg() {
		this.curGroup.style.titlePage = true;
	}
	cmd$headery(val) {
		this.curGroup.style.headerTop = val;
	}
	cmd$footery(val) {
		this.curGroup.style.footerBottom = val;
	}

	cmd$pgncont() {
		this.curGroup.style.pageNumberRestart = "none";
	}
	cmd$pgnrestart() {
		this.curGroup.style.pageNumberRestart = "on-page-number";
	}
	cmd$pgnx(val) {
		this.curGroup.style.pageNumberRight = val;
	}
	cmd$pgny(val) {
		this.curGroup.style.pageNumberTop = val;
	}
	cmd$pgndec() {
		this.curGroup.style.pageNumbering = "DECIMAL";
	}
	cmd$pgnucrm() {
		this.curGroup.style.pageNumbering = "I";
	}
	cmd$pgnlcrm() {
		this.curGroup.style.pageNumbering = "i";
	}
	cmd$pgnucltr() {
		this.curGroup.style.pageNumbering = "A";
	}
	cmd$pgnlcltr() {
		this.curGroup.style.pageNumbering = "a";
	}
	cmd$pgnbidia() {
		this.curGroup.style.pageNumbering = "BIDIA";
	}
	cmd$pgnbidib() {
		this.curGroup.style.pageNumbering = "BIDIB";
	}
	cmd$pgnchosung() {
		this.curGroup.style.pageNumbering = "CHOSUNG";
	}
	cmd$pgncnum() {
		this.curGroup.style.pageNumbering = "CIRCLENUM";
	}
	cmd$pgndbnum() {
		this.curGroup.style.pageNumbering = "KANJI-NO-DIGIT";
	}
	cmd$pgndbnumd() {
		this.curGroup.style.pageNumbering = "KANJI-DIGIT";
	}
	cmd$pgndbnumt() {
		this.curGroup.style.pageNumbering = "DBNUM3";
	}
	cmd$pgndbnumk() {
		this.curGroup.style.pageNumbering = "DBNUM4";
	}
	cmd$pgndecd() {
		this.curGroup.style.pageNumbering = "DOUBLE-BYTE";
	}
	cmd$pgnganada() {
		this.curGroup.style.pageNumbering = "GANADA";
	}
	cmd$pgngbnum() {
		this.curGroup.style.pageNumbering = "GB1";
	}
	cmd$pgngbnumd() {
		this.curGroup.style.pageNumbering = "GB2";
	}
	cmd$pgngbnuml() {
		this.curGroup.style.pageNumbering = "GB3";
	}
	cmd$pgngbnumk() {
		this.curGroup.style.pageNumbering = "GB4";
	}
	cmd$pgnzodiac() {
		this.curGroup.style.pageNumbering = "ZODIAC1";
	}
	cmd$pgnzodiacd() {
		this.curGroup.style.pageNumbering = "ZODIAC2";
	}
	cmd$pgnzodiacl() {
		this.curGroup.style.pageNumbering = "ZODIAC3";
	}
	cmd$pgnhindia() {
		this.curGroup.style.pageNumbering = "HINDI-VOWEL";
	}
	cmd$pgnhindib() {
		this.curGroup.style.pageNumbering = "HINDI-CONSONANT";
	}
	cmd$pgnhindic() {
		this.curGroup.style.pageNumbering = "HINDI-DIGIT";
	}
	cmd$pgnhindid() {
		this.curGroup.style.pageNumbering = "HINDI-DESCRIPTIVE";
	}
	cmd$pgnthaia() {
		this.curGroup.style.pageNumbering = "THAI-LETTER";
	}
	cmd$pgnthaib() {
		this.curGroup.style.pageNumbering = "THAI-DIGIT";
	}
	cmd$pgnthaic() {
		this.curGroup.style.pageNumbering = "THAIDESCRIPTIVE";
	}
	cmd$pgnvieta() {
		this.curGroup.style.pageNumbering = "VIETNAMESE-DESCRIPTIVE";
	}
	cmd$pgnid() {
		this.curGroup.style.pageNumbering = "KOREAN-DASH";
	}
	cmd$pgnhn(val) {
		this.curGroup.style.pageNumberHeaderLevel = val;
	}
	cmd$pgnhnsh() {
		this.curGroup.style.pageNumberSeparator = "-";
	}
	cmd$pgnhnsp() {
		this.curGroup.style.pageNumberSeparator = ".";
	}
	cmd$pgnhnsc() {
		this.curGroup.style.pageNumberSeparator = ":";
	}
	cmd$pgnhnsm() {
		this.curGroup.style.pageNumberSeparator = "—";
	}
	cmd$pgnhnsn() {
		this.curGroup.style.pageNumberSeparator = "–";
	}

	cmd$vertal() {
		this.curGroup.style.textAlign = "bottom"; //Alias for vertalb. Why? Ask Microsoft.
	}
	cmd$vertalt() {
		this.curGroup.style.textAlign = "top";
	}
	cmd$vertalb() {
		this.curGroup.style.textAlign = "bottom";
	}
	cmd$vertalc() {
		this.curGroup.style.textAlign = "center";
	}
	cmd$vertalj() {
		this.curGroup.style.textAlign = "justified";
	}

	cmd$srauth(val) {
		this.curGroup.attributes.revisionAuthor = val;
	}
	cmd$srdate(val) {
		this.curGroup.attributes.revisionDate = val;
	}

	cmd$rtlsect() {
		this.curGroup.style.snakeColumns = "rtl";
	}
	cmd$ltrsect() {
		this.curGroup.style.snakeColumns = "ltr";
	}

	cmd$horzsect() {
		this.curGroup.style.renderDirection = "horizontal";
	}
	cmd$vertsect() {
		this.curGroup.style.renderDirection = "vertical";
	}

	cmd$stextflow(val) {
		this.curGroup.style.textFlow = val;
	}

	cmd$sectexpand(val) {
		this.curGroup.style.charSpaceBase = val;
	}
	cmd$sectlinegrid(val) {
		this.curGroup.style.lineGrid = val;
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
		this.curGroup.style.charGridSnap = true;
	}

	/* Headers, Footers */
	cmd$header() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "header", "style");
	}
	cmd$footer() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footer", "style");
	}
	cmd$headerl() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "headerLeft", "style");
	}
	cmd$headerr() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "headerRight", "style");
	}
	cmd$headerf() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "headerFirst", "style");
	}
	cmd$footerl() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footerLeft", "style");
	}
	cmd$footerr() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footerRight", "style");
	}
	cmd$footerf() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "footerFirst", "style");
	}

	/* Paragraphs */
	cmd$par() {
		while(this.curGroup.type !== "document" && this.curGroup.type !== "section" && this.curGroup.type !== "table-row") {this.endGroup();}
		this.newGroup("paragraph");
		if (this.lastPar.style.contextualspace) {
			const spacings = [
			"spaceBefore",
			"spaceAfter",
			"autoSpaceBefore",
			"autoSpaceAfter",
			"spaceBeforeChar",
			"spaceAfterChar",
			];
			Object.keys(this.lastPar.style).forEach(key => {
				if (spacings.includes(key)) {
					this.lastPar.style[key] = false;
				}
			});
		}
		this.curGroup.style = this.lastPar.style;
		this.curGroup.attributes = this.lastPar.attributes;
	}
	cmd$pard() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.style = Object.assign(JSON.parse(JSON.stringify(this.defCharState)),JSON.parse(JSON.stringify(this.defParState)));
			this.curGroup.attributes = {};
		} else {
			while(this.curGroup.type !== "document" && this.curGroup.type !== "section" && this.curGroup.type !== "table-row") {this.endGroup();}
			this.newGroup("paragraph");
			this.curGroup.style = Object.assign(JSON.parse(JSON.stringify(this.defCharState)),JSON.parse(JSON.stringify(this.defParState)));
			this.curGroup.attributes = {};
		}
	}
	cmd$plain() {
		Object.keys(this.defCharState).forEach(key => {
			if (this.curGroup.style[key]) {
				this.curGroup.style[key] = this.defCharState[key];
			}
		});
		this.curGroup.attributes = {};
	}
	cmd$spv() {
		this.curGroup.style.styleSeparator = true;
	}
	cmd$hyphpar(val) {
		this.curGroup.style.autoHyphenate = val !== 0;
	}
	cmd$intbl() {
		this.curGroup.attributes.inTable = true;
	}
	cmd$itap(val) {
		this.curGroup.attributes.nestingDepth = val;
	}
	cmd$keep() {
		this.curGroup.style.keep = true;
	}
	cmd$keepn() {
		this.curGroup.style.keepNext = true;
	}
	cmd$level(val) {
		this.curGroup.style.outlineLevel = val;
	}
	cmd$noline() {
		this.curGroup.style.noLineNumbers = true;
	}
	cmd$nowidctlpar() {
		this.curGroup.style.widowControl = false;
	}
	cmd$outlinelevel(val) {
		this.curGroup.style.outlineLevel = val; //Don't ask.
	}
	cmd$pagebb() {
		this.curGroup.style.breakBefore = true;
	}
	cmd$sbys() {
		this.curGroup.style.sideBySide = true;
	}

	cmd$yts(val) {
		this.curGroup.style.tableStyleHandle = val;
	}
	cmd$tscfirstrow(val) {
		this.curGroup.attributes.firstRow = true;
	}
	cmd$tsclastrow(val) {
		this.curGroup.attributes.lastRow = true;
	}
	cmd$tscfirstcol(val) {
		this.curGroup.attributes.firstColumn = true;
	}
	cmd$tsclastcol(val) {
		this.curGroup.attributes.lastColumn = true;
	}
	cmd$tscbandhorzodd(val) {
		this.curGroup.attributes.rowEO = "odd";
	}
	cmd$tscbandhorzeven(val) {
		this.curGroup.attributes.rowEO = "even";
	}
	cmd$tscbandvertodd(val) {
		this.curGroup.attributes.columnEO = "odd";
	}
	cmd$tscbandverteven(val) {
		this.curGroup.attributes.columnEO = "even";
	}
	cmd$tscnwcell(val) {
		this.curGroup.attributes.northWestCell = true;
	}
	cmd$tscnecell(val) {
		this.curGroup.attributes.northEastCell = true;
	}
	cmd$tscswcell(val) {
		this.curGroup.attributes.southWestCell = true;
	}
	cmd$tscsecell(val) {
		this.curGroup.attributes.southEastCell = true;
	}

	cmd$qc() {
		this.curGroup.style.alignment = "center";
	}
	cmd$qj() {
		this.curGroup.style.alignment = "justified";
	}
	cmd$ql() {
		this.curGroup.style.alignment = "left";
	}
	cmd$qr() {
		this.curGroup.style.alignment = "right";
	}
	cmd$qd() {
		this.curGroup.style.alignment = "distributed";
	}
	cmd$qk(val) {
		this.curGroup.style.alignment = "kashida" + val;
	}
	cmd$qt() {
		this.curGroup.style.alignment = "thai";
	}

	cmd$faauto() {
		this.curGroup.style.fontAlignment = "auto";
	}
	cmd$fahang() {
		this.curGroup.style.fontAlignment = "hanging";
	}
	cmd$faroman() {
		this.curGroup.style.fontAlignment = "roman";
	}
	cmd$favar() {
		this.curGroup.style.fontAlignment = "variable";
	}
	cmd$fafixed() {
		this.curGroup.style.fontAlignment = "fixed";
	}

	cmd$fi(val) {
		this.curGroup.style.firstLineIndent = val;
	}
	cmd$cufi(val) {
		this.curGroup.style.firstLineIndentChar = val;
	}
	cmd$li(val) {
		this.curGroup.style.leftIndent = val;
	}
	cmd$lin(val) {
		this.curGroup.style.leftRightIndent = val;
	}
	cmd$culi(val) {
		this.curGroup.style.leftIndentChar = val;
	}
	cmd$ri(val) {
		this.curGroup.style.rightIndent = val;
	}
	cmd$rin(val) {
		this.curGroup.style.rightLeftIndent = val;
	}
	cmd$curi(val) {
		this.curGroup.style.rightIndentChar = val;
	}
	cmd$adjustright() {
		this.curGroup.style.adjustRightIndent = true;
	}
	cmd$indmirror() {
		this.curGroup.style.mirrorIndents = true;
	}

	cmd$sb(val) {
		this.curGroup.style.spaceBefore = val;
	}
	cmd$sa(val) {
		this.curGroup.style.spaceAfter = val;
	}
	cmd$sbauto(val) {
		this.curGroup.style.autoSpaceBefore = val;
	}
	cmd$saauto(val) {
		this.curGroup.style.autoSpaceAfter = val;
	}
	cmd$lisb(val) {
		this.curGroup.style.spaceBeforeChar = val;
	}
	cmd$lisa(val) {
		this.curGroup.style.spaceAfterChar = val;
	}
	cmd$sl(val) {
		this.curGroup.style.lineSpacing = val;
	}
	cmd$slmult(val) {
		this.curGroup.style.lineSpacingMultiple = val;
	}
	cmd$nosnaplinegrid() {
		this.curGroup.style.noSnapLineGrid = true;
	}
	cmd$contextualspace() {
		this.curGroup.style.contextualSpace = true;
	}

	cmd$subdocument(val) {
		this.curGroup.attributes.subdocument = val;
	}

	cmd$prauth(val) {
		this.curGroup.attributes.revisionAuthor = val;
	}
	cmd$prdate(val) {
		this.curGroup.attributes.revisionDate = val;
	}

	cmd$rtlpar() {
		this.curGroup.style.direction = "rtl";
	}
	cmd$ltrpar() {
		this.curGroup.style.direction = "ltr";
	}

	cmd$nocwrap() {
		this.curGroup.style.noCharWrap = true;
	}
	cmd$nowwrap() {
		this.curGroup.style.noWordWrap = true;
	}
	cmd$nooverflow() {
		this.curGroup.style.noOverflow = true;
	}
	cmd$aspalpha() {
		this.curGroup.style.dbcEngAutoSpace = true;
	}
	cmd$aspnum() {
		this.curGroup.style.dbcNumAutoSpace = true;
	}

	cmd$collapsed(val) {
		this.curGroup.style.collapsed = val !== 0;
	}

	cmd$txbxtwno() {
		this.curGroup.style.textboxwrap = false;
	}
	cmd$txbxtwalways() {
		this.curGroup.style.textboxWrap = "always";
	}
	cmd$txbxtwfirstlast() {
		this.curGroup.style.textboxWrap = "first-last";
	}
	cmd$txbxtwfirst() {
		this.curGroup.style.textboxWrap = "first";
	}
	cmd$txbxtwlast() {
		this.curGroup.style.textboxWrap = "last";
	}

	/* Tabs */
	cmd$tx(val) {
		this.curGroup.style.tabPosition = val;
	}
	cmd$tqr() {
		this.curGroup.style.tabType = "flush-right";
	}
	cmd$tqc() {
		this.curGroup.style.tabType = "center";
	}
	cmd$tqdec() {
		this.curGroup.style.tabType = "decimal";
	}
	cmd$tb(val) {
		this.curGroup.style.barTabPosition = val;
	}
	cmd$tldot() {
		this.curGroup.style.leader = "dots";
	}
	cmd$tlmdot() {
		this.curGroup.style.leader = "middle-dots";
	}
	cmd$tlhyph() {
		this.curGroup.style.leader = "hyphen";
	}
	cmd$tlul() {
		this.curGroup.style.leader = "underline";
	}
	cmd$tlth() {
		this.curGroup.style.leader = "thick-line";
	}
	cmd$tleq() {
		this.curGroup.style.leader = "equalsign";
	}

	/* Absolute Position Tabs */
	cmd$ptablnone() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.leading = false;
	}
	cmd$ptabldot() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.leading = ".";
	}
	cmd$ptablnone() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.leading = "-";
	}
	cmd$ptablnone() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.leading = "_";
	}
	cmd$ptablnone() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.leading = "·";
	}
	cmd$pmartabql() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "left-to-margin";
	}
	cmd$pmartabqc() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "center-to-margin";
	}
	cmd$pmartabqr() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "right-to-margin";
	}
	cmd$pindtabql() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "left-to-indent";
	}
	cmd$pindtabqc() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "center-to-indent";
	}
	cmd$pindtabqr() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "right-to-indent";
	}

	/* Bullets and Numbering */
	/*-- Older Word (6.0, 95) --*/
	cmd$pntext() {
		this.curGroup.type = "pn-list-text";
	}
	cmd$pn() {
		this.curGroup.type = "pn-list-item";
	}
	cmd$pnlvl(val) {
		this.curGroup.style.pnlvl = val;
	}
	cmd$pnlvlblt() {
		this.curGroup.style.pnBullet = true;
	}
	cmd$pnlvlbody() {
		this.curGroup.style.pnBody = true;
	}
	cmd$pnlvlcont() {
		this.curGroup.style.pnContinue = true;
	}
	cmd$pnnumonce() {
		this.curGroup.style.pnNumberOnce = true;
	}
	cmd$pnacross() {
		this.curGroup.style.pnNumberAcross = true;
	}
	cmd$pnhang() {
		this.curGroup.style.pnHangingIndent = true;
	}
	cmd$pnrestart() {
		this.curGroup.style.pnRestart = true;
	}
	cmd$pncard() {
		this.curGroup.style.pnNumbering = "One";
	}
	cmd$pndec() {
		this.curGroup.style.pnNumbering = "1";
	}
	cmd$pnucltr() {
		this.curGroup.style.pnNumbering = "A";
	}
	cmd$pnucrm() {
		this.curGroup.style.pnNumbering = "I";
	}
	cmd$pnlcltr() {
		this.curGroup.style.pnNumbering = "a";
	}
	cmd$pnlcrm() {
		this.curGroup.style.pnNumbering = "i";
	}
	cmd$pnord() {
		this.curGroup.style.pnNumbering = "1st";
	}
	cmd$pnordt() {
		this.curGroup.style.pnNumbering = "First";
	}
	cmd$pnbidia() {
		this.curGroup.style.pnNumbering = "BIDIA";
	}
	cmd$pnbidib() {
		this.curGroup.style.pnNumbering = "BIDIB";
	}
	cmd$pnaiu() {
		this.curGroup.style.pnNumbering = "AIUEO";
	}
	cmd$pnaiud() {
		this.curGroup.style.pnNumbering = "AIUEODBCHAR";
	}
	cmd$pnaiueo() {
		this.curGroup.style.pnNumbering = "AIUEO"; //Alias for \pnaiu
	}
	cmd$pnaiueod() {
		this.curGroup.style.pnNumbering = "AIUEO-DBCHAR"; //Alias for \pnaiud
	}
	cmd$pnchosung() {
		this.curGroup.style.pnNumbering = "CHOSUNG";
	}
	cmd$pncnum() {
		this.curGroup.style.pnNumbering = "CIRCLENUM";
	}
	cmd$pndbnum() {
		this.curGroup.style.pnNumbering = "DBNUM1";
	}
	cmd$pndbnumd() {
		this.curGroup.style.pnNumbering = "DBNUM2";
	}
	cmd$pndbnuml() {
		this.curGroup.style.pnNumbering = "DBNUM3";
	}
	cmd$pndbnumk() {
		this.curGroup.style.pnNumbering = "DBNUM4";
	}
	cmd$pndbnumt() {
		this.curGroup.style.pnNumbering = "DBNUM3"; //Alias for \pndbnuml
	}
	cmd$pndecd() {
		this.curGroup.style.pnNumbering = "DBCHAR";
	}
	cmd$pnganada() {
		this.curGroup.style.pnNumbering = "GANADA";
	}
	cmd$pngbnum() {
		this.curGroup.style.pnNumbering = "GB1";
	}
	cmd$pngbnumd() {
		this.curGroup.style.pnNumbering = "GB2";
	}
	cmd$pngbnuml() {
		this.curGroup.style.pnNumbering = "GB3";
	}
	cmd$pngbnumk() {
		this.curGroup.style.pnNumbering = "GB4";
	}
	cmd$pniroha() {
		this.curGroup.style.pnNumbering = "IROHA";
	}
	cmd$pnirohad() {
		this.curGroup.style.pnNumbering = "IROHA-DBCHAR";
	}
	cmd$pnzodiac() {
		this.curGroup.style.pnNumbering = "ZODIAC1";
	}
	cmd$pnzodiacd() {
		this.curGroup.style.pnNumbering = "ZODIAC2";
	}
	cmd$pnzodiacl() {
		this.curGroup.style.pnNumbering = "ZODIAC3";
	}
	cmd$pnb(val) {
		this.curGroup.style.pnBold = val !== 0;
	}
	cmd$pni(val) {
		this.curGroup.style.pnItalics = val !== 0;
	}
	cmd$pncaps(val) {
		this.curGroup.style.pnCaps = val !== 0;
	}
	cmd$pnscaps(val) {
		this.curGroup.style.pnSmallcaps = val !== 0;
	}
	cmd$pnul(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "continuous" : false;
	}
	cmd$pnuld(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "dot" : false;
	}
	cmd$pnuldash(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "dash" : false;
	}
	cmd$pnuldashdot(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "dash-dot" : false;
	}
	cmd$pnuldashdd(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "dash-dot-dot" : false;
	}
	cmd$pnulhair(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "hairline" : false;
	}
	cmd$pnulth(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "thick" : false;
	}
	cmd$pnulwave(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "wave" : false;
	}
	cmd$pnuldb(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "double" : false;
	}
	cmd$pnulw(val) {
		this.curGroup.style.pnUnderline = val !== 0 ? "word" : false;
	}
	cmd$pnulnone() {
		this.curGroup.style.pnUnderline = false;
	}
	cmd$pnstrike(val) {
		this.curGroup.style.pnStrikethrough = val !== 0;
	}
	cmd$pncf(val) {
		this.curGroup.style.pnForeground = val;
	}
	cmd$pnf(val) {
		this.curGroup.style.pnFont = val;
	}
	cmd$pnfs(val) {
		this.curGroup.style.pnFontSize = val;
	}
	cmd$pnindent(val) {
		this.curGroup.style.pnIndent = val;
	}
	cmd$pnsp(val) {
		this.curGroup.style.pnDistance = val;
	}
	cmd$pnprev() {
		this.curGroup.style.pnPrev = true;
	}
	cmd$pnqc() {
		this.curGroup.style.pnAlignment = "center";
	}
	cmd$pnql() {
		this.curGroup.style.pnAlignment = "left";
	}
	cmd$pnqr() {
		this.curGroup.style.pnAlignment = "right";
	}
	cmd$pnstart(val) {
		this.curGroup.style.pnStart = val;
	}
	cmd$pntxta() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "textAfter", "style");
	}
	cmd$pntxtb() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "textBefore", "style");
	}

	/*-- Newer Word (97-2007) --*/
	cmd$listtext(val) {
		this.curGroup.type = "list-text";
	}
	cmd$ilvl(val) {
		this.curGroup.attributes.ilvl = val;
		this.curGroup.type = "list-item";
	}

	/*-- Revision Marks --*/
	cmd$pnauth(val) {
		this.curGroup.attributes.pnRevisionAuthor = val;
	}
	cmd$pndate(val) {
		this.curGroup.attributes.pnRevisionDate = val;
	}
	cmd$pnrnot() {
		this.curGroup.style.pnInserted = true;
	}
	cmd$pnrxst(val) {
		if (!this.curGroup.attributes.pnRXST) {this.curGroup.attributes.pnRXST = [];}
		this.curGroup.attributes.pnRXST.push(val);
	}
	cmd$pnrrgb(val) {
		if (!this.curGroup.attributes.pnRRGB) {this.curGroup.attributes.pnRRGB = [];}
		this.curGroup.attributes.pnRRGB.push(val);
	}
	cmd$pnrnfc(val) {
		if (!this.curGroup.attributes.pnRNFC) {this.curGroup.attributes.pnRNFC = [];}
		this.curGroup.attributes.pnRNFC.push(val);
	}
	cmd$pnrpnbr(val) {
		if (!this.curGroup.attributes.pnRPnBR) {this.curGroup.attributes.pnRPnBR = [];}
		this.curGroup.attributes.pnRPnBR.push(val);
	}
	cmd$pnrstart(val) {
		//Uneeded.
	}
	cmd$pnrstop(val) {
		//Uneeded.
	}
	cmd$dfrauth(val) {
		this.curGroup.attributes.dfRevisionAuthor = val;
	}
	cmd$dfrdate(val) {
		this.curGroup.attributes.dfRevisionDate = val;
	}
	cmd$pnrxst(val) {
		if (!this.curGroup.attributes.pnRXST) {this.curGroup.attributes.dfRXST = [];}
		this.curGroup.attributes.pnRXST.push(val);
	}
	cmd$dfrstart(val) {
		//Uneeded.
	}
	cmd$dfrstop(val) {
		//Uneeded.
	}

	/* Paragraph Borders */
	cmd$brdrt() {
		this.curGroup.style.borderTop = true;
	}
	cmd$brdrb() {
		this.curGroup.style.borderBottom = true;
	}
	cmd$brdrl() {
		this.curGroup.style.borderLeft = true;
	}
	cmd$brdrr() {
		this.curGroup.style.borderRight = true;
	}
	cmd$brdrbtw() {
		this.curGroup.style.borderBetween = true;
	}
	cmd$brdrbar() {
		this.curGroup.style.borderOutside = true;
	}
	cmd$box() {
		this.curGroup.style.box = true;
	}
	cmd$brdrs() {
		this.curGroup.style.borderStyle = "single-thick";
	}
	cmd$brdrth() {
		this.curGroup.style.borderStyle = "double-thick";
	}
	cmd$brdrsh() {
		this.curGroup.style.borderStyle = "shadowed";
	}
	cmd$brdrdb() {
		this.curGroup.style.borderStyle = "double";
	}
	cmd$brdrdot() {
		this.curGroup.style.borderStyle = "dotted";
	}
	cmd$brdrdash() {
		this.curGroup.style.borderStyle = "dashed";
	}
	cmd$brdrhair() {
		this.curGroup.style.borderStyle = "hairline";
	}
	cmd$brdrdashsm() {
		this.curGroup.style.borderStyle = "dashed-small";
	}
	cmd$brdrdashd() {
		this.curGroup.style.borderStyle = "dot-dashed";
	}
	cmd$brdrdashdd() {
		this.curGroup.style.borderStyle = "dot-dot-dashed";
	}
	cmd$brdrdashdot() {
		this.curGroup.style.borderStyle = "dot-dashed"; //Alias for \brdrdashd
	}
	cmd$brdrdashdotdot() {
		this.curGroup.style.borderStyle = "dot-dot-dashed"; //Alias for \brdrdashdd
	}
	cmd$brdrinset() {
		this.curGroup.style.borderStyle = "inset";
	}
	cmd$brdrnone() {
		this.curGroup.style.borderStyle = "none";
	}
	cmd$brdroutset() {
		this.curGroup.style.borderStyle = "outset";
	}
	cmd$brdrtriple() {
		this.curGroup.style.borderStyle = "triple";
	}
	cmd$brdrtnthsg() {
		this.curGroup.style.borderStyle = "thick-thin-small";
	}
	cmd$brdrthtnsg() {
		this.curGroup.style.borderStyle = "thin-thick-small";
	}
	cmd$brdrtnthtnsg() {
		this.curGroup.style.borderStyle = "thin-thick-thin-small";
	}
	cmd$brdrtnthmg() {
		this.curGroup.style.borderStyle = "thick-thin-medium";
	}
	cmd$brdrthtnmg() {
		this.curGroup.style.borderStyle = "thin-thick-medium";
	}
	cmd$brdrtnthtnmg() {
		this.curGroup.style.borderStyle = "thin-thick-thin-medium";
	}
	cmd$brdrtnthlg() {
		this.curGroup.style.borderStyle = "thick-thin-large";
	}
	cmd$brdrthtnlg() {
		this.curGroup.style.borderStyle = "thin-thick-large";
	}
	cmd$brdrtnthtnlg() {
		this.curGroup.style.borderStyle = "thin-thick-thin-large";
	}
	cmd$brdrwavy() {
		this.curGroup.style.borderStyle = "wavy";
	}
	cmd$brdrwavydb() {
		this.curGroup.style.borderStyle = "wavy-double";
	}
	cmd$brdrdashdotstr() {
		this.curGroup.style.borderStyle = "striped";
	}
	cmd$brdremboss() {
		this.curGroup.style.borderStyle = "embossed";
	}
	cmd$brdrengrave() {
		this.curGroup.style.borderStyle = "engraved";
	}
	cmd$brdrframe() {
		this.curGroup.style.borderStyle = "framed";
	}
	cmd$brdrw(val) {
		this.curGroup.style.borderWidth = val;
	}
	cmd$brdrcf(val) {
		this.curGroup.style.borderForeground = val;
	}
	cmd$brsp(val) {
		this.curGroup.style.borderSpace = val;
	}
	cmd$brdrnil() {
		this.curGroup.style.borderNil = true;
	}
	cmd$brdrtbl() {
		this.curGroup.style.noBorderTable = true;
	}

	/* Paragraph Shading */
	cmd$shading(val) {
		this.curGroup.style.shading = val;
	}
	cmd$bghoriz() {
		this.curGroup.style.backgroundPattern = "horizontal";
	}
	cmd$bgvert() {
		this.curGroup.style.backgroundPattern = "vertical";
	}
	cmd$bgfdiag() {
		this.curGroup.style.backgroundPattern = "diagonal-forwards";
	}
	cmd$bgbdiag() {
		this.curGroup.style.backgroundPattern = "diagonal-backwards";
	}
	cmd$bgcross() {
		this.curGroup.style.backgroundPattern = "cross";
	}
	cmd$bgdcross() {
		this.curGroup.style.backgroundPattern = "cross-diagonal";
	}
	cmd$bgdkhoriz() {
		this.curGroup.style.backgroundPattern = "dark-horizontal";
	}
	cmd$bgdkvert() {
		this.curGroup.style.backgroundPattern = "dark-vertical";
	}
	cmd$bgdkfdiag() {
		this.curGroup.style.backgroundPattern = "dark-diagonal-forwards";
	}
	cmd$bgdkbdiag() {
		this.curGroup.style.backgroundPattern = "dark-diagonal-backwards";
	}
	cmd$bgdkcross() {
		this.curGroup.style.backgroundPattern = "dark-cross";
	}
	cmd$bgdkdcross() {
		this.curGroup.style.backgroundPattern = "dark-cross-diagonal";
	}
	cmd$cfpat(val) {
		this.curGroup.style.backgroundFill = val;
	}
	cmd$cfpat(val) {
		this.curGroup.style.backgroundPatternColour = val;
	}

	/* Positioned Objects and Frames */
	cmd$absw(val) {
		this.curGroup.style.frameWidth = val;
	}
	cmd$absh(val) {
		this.curGroup.style.frameHeight = val;
	}
	/*-- Horizontal Position --*/
	cmd$phmrg() {
		this.curGroup.style.horizontalRef = "margin";
	}
	cmd$phpg() {
		this.curGroup.style.horizontalRef = "page";
	}
	cmd$phcol() {
		this.curGroup.style.horizontalRef = "column";
	}
	cmd$posx(val) {
		this.curGroup.style.frameXPosition = val;
	}
	cmd$posnegx(val) {
		this.curGroup.style.frameXPosition = val; //Same as \posy but allows negative values
	}
	cmd$posxc() {
		this.curGroup.style.frameHorizontalPosition = "center";
	}
	cmd$posxi() {
		this.curGroup.style.frameHorizontalPosition = "inside";
	}
	cmd$posxo() {
		this.curGroup.style.frameHorizontalPosition = "outside";
	}
	cmd$posxr() {
		this.curGroup.style.frameHorizontalPosition = "right";
	}
	cmd$posxl() {
		this.curGroup.style.frameHorizontalPosition = "left";
	}
	/*-- Vertical Position --*/
	cmd$pvmrg() {
		this.curGroup.style.verticalRef = "margin";
	}
	cmd$pvpg() {
		this.curGroup.style.verticalRef = "page";
	}
	cmd$pvpara() {
		this.curGroup.style.verticalRef = "paragraph";
	}
	cmd$posy(val) {
		this.curGroup.style.frameYPosition = val;
	}
	cmd$posnegy(val) {
		this.curGroup.style.frameYPosition = val; //Same as \posy but allows negative values
	}
	cmd$posyl() {
		this.curGroup.style.frameVerticalPosition = "left";
	}
	cmd$posyt() {
		this.curGroup.style.frameVerticalPosition = "top";
	}
	cmd$posyc() {
		this.curGroup.style.frameVerticalPosition = "center";
	}
	cmd$posyb() {
		this.curGroup.style.frameVerticalPosition = "bottom";
	}
	cmd$posyin() {
		this.curGroup.style.frameVerticalPosition = "inside";
	}
	cmd$posyout() {
		this.curGroup.style.frameVerticalPosition = "outside";
	}
	cmd$abslock(val) {
		this.curGroup.style.frameLock = val === 1;
	}
	/*-- Text Wrapping --*/
	cmd$dxfrtext(val) {
		this.curGroup.style.distanceFromText = val;
	}
	cmd$dfrmtxtx(val) {
		this.curGroup.style.distanceFromTextX = val;
	}
	cmd$dfrmtxty(val) {
		this.curGroup.style.distanceFromTextY = val;
	}
	cmd$overlay() {
		this.curGroup.style.overlay = true;
	}
	cmd$wrapdefault() {
		this.curGroup.style.wrapType = "default";
	}
	cmd$wraparound() {
		this.curGroup.style.wrapType = "around";
	}
	cmd$wraptight() {
		this.curGroup.style.wrapType = "tight";
	}
	cmd$wrapthrough() {
		this.curGroup.style.wrapType = "through";
	}
	/*-- Drop Caps --*/
	cmd$dropcapli(val) {
		this.curGroup.style.dropCapLines = val;
	}
	cmd$dropcapt(val) {
		this.curGroup.style.dropCapType = val === 1 ? "in-text" : val === 2 ? "margin" : false;
	}
	/*-- Overlap --*/
	cmd$absnoovrlp(val) {
		this.curGroup.style.overlapAllowed = val === 0;
	}
	/*-- Text Flow --*/
	cmd$frmtxlrtb() {
		this.curGroup.style.frameFlow = "ltr-ttb";
	}
	cmd$frmtxtbrl() {
		this.curGroup.style.frameFlow = "rtl-ttb";
	}
	cmd$frmtxbtlr() {
		this.curGroup.style.frameFlow = "ltr-btt";
	}
	cmd$frmtxlrtbv() {
		this.curGroup.style.frameFlow = "ltr-ttb-v";
	}
	cmd$frmtxtbrlv() {
		this.curGroup.style.frameFlow = "rtl-ttb-v";
	}

	/* Tables */
	/* See issue https://github.com/EndaHallahan/RibosomalRTF/issues/1 */

	cmd$intbl() {
		this.curGroup.type = "table-entry";
		this.curGroup.style = {...this.curGroup.style, row: {...this.curRow.style}}
		this.curGroup.attributes = {...this.curGroup.attributes, row:{...this.curRow.attributes}}
	}

	cmd$trowd() {

	}
	cmd$irow(val) {
		this.curRow.attributes.irow = val;
	}
	cmd$irowband(val) {
		this.curRow.attributes.irowHeaderAdjusted = val;
	}
	cmd$row() {

	}
	cmd$lastrow() {
		this.curRow.lastRow = true;
	}
	cmd$tcelld() {

	}
	cmd$nestcell() {
		this.curGroup.type = "cell";
		this.endGroup()
		this.newGroup("cell");
	}
	cmd$nestrow() {

	}
	cmd$nesttableprops() {
		this.curGroup = new NonGroup(this.curGroup.parent);
	}
	cmd$nonesttables() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "noNestTables", "attributes")
	}
	cmd$trgaph(val) {
		this.curRow.style.rowGap = val;
	}
	cmd$cellx(val) {
		this.curGroup.style.rightCellBoundry = val;
	}
	cmd$cell() {
		this.curGroup.type = "cell";
		this.endGroup()
		this.newGroup("cell");
	}
	cmd$clmgf() {
		this.curGroup.style.firstMergeCell = true;
	}
	cmd$clmrg() {
		this.curGroup.style.mergeCell = true;
	}
	cmd$clvmgf() {
		this.curGroup.style.firstMergeCellVertical = true;
	}
	cmd$clvmrg() {
		this.curGroup.style.mergeCellVertical = true;
	}
	/*-- Table Row Revision Tracking --*/
	cmd$trauth(val) {
		this.curRow.attributes.revisionAuthor = val;
	}
	cmd$trdate(val) {
		this.curRow.attributes.revisionDate = val;
	}
	/*-- Autoformatting Tags --*/

	/*-- Row Formatting --*/

	/*-- Compared Table Cells --*/

	/*-- Bidirectional Controls --*/

	/*-- Row Borders --*/

	/*-- Cell Borders --*/

	/*-- Cell Shading and Background Pattern --*/

	/*-- Cell Vertical Text Alignment --*/

	/*-- Cell Text Flow --*/



	/* Mathematics */



	/* Character Text */




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



	/* Special Characters */
	cmd$emdash() {
		this.newGroup("text");
		this.curGroup.contents.push("—");
		this.endGroup();
	}
	cmd$endash() {
		this.newGroup("text");
		this.curGroup.contents.push("–");
		this.endGroup();
	}
	cmd$tab() {
		this.newGroup("text");
		this.curGroup.contents.push("\t");
		this.endGroup();
	}
	cmd$line() {
		this.newGroup("text");
		this.curGroup.contents.push("\n");
		this.endGroup();
	}
	cmd$hrule() {
		this.newGroup("hr");
		this.endGroup();
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
		this.curGroup.style.fontSize = val;
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
		this.curGroup.type = "shpict";
	}
	cmd$pict() {
		this.curGroup = new Picture(this.curGroup.parent);
	}
	cmd$nisusfilename() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "nisusFileName");
	}
	cmd$nonshppict() {
		this.curGroup.attributes.nonShpPict = false;
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
		this.curGroup.attributes.source = "OS/2-METAFILE";
		this.curGroup.attributes.sourceType = val;
	}
	cmd$wmetafile(val) {
		this.curGroup.attributes.source = "WINDOWS-METAFILE";
		this.curGroup.attributes.mappingMode = val;
	}
	cmd$dibitmap(val) {
		this.curGroup.attributes.source = "WINDOWS-DI-BITMAP";
		this.curGroup.attributes.sourceType = val;
	}
	cmd$wbitmap(val) {
		this.curGroup.attributes.source = "WINDOWS-DD-BITMAP";
		this.curGroup.attributes.sourceType = val;
	}
	cmd$wbmbitspixel(val) {
		this.curGroup.attributes.bitsPixel = val;
	}
	cmd$wbmplanes(val) {
		this.curGroup.attributes.planes = val;
	}
	cmd$wbmwidthbytes(val) {
		this.curGroup.attributes.widthBytes = val;
	}
	cmd$picw(val) {
		this.curGroup.style.width = val;
	}
	cmd$pich(val) {
		this.curGroup.style.height = val;
	}
	cmd$picwgoal(val) {
		this.curGroup.style.widthGoal = val;
	}
	cmd$pichgoal(val) {
		this.curGroup.style.heightGoal = val;
	}
	cmd$picscalex(val) {
		this.curGroup.style.scaleX = val;
	}
	cmd$picscaley(val) {
		this.curGroup.style.scaleY = val;
	}
	cmd$picscaled() {
		this.curGroup.style.scaleD = true;
	}
	cmd$piccropt(val) {
		this.curGroup.style.cropTop = val;
	}
	cmd$piccropb(val) {
		this.curGroup.style.cropBottom = val;
	}
	cmd$piccropl(val) {
		this.curGroup.style.cropLeft = val;
	}
	cmd$piccropr(val) {
		this.curGroup.style.cropRight = val;
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
