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
		this.curGroup = new ParameterGroup(this.curGroup.parent, "themeData");
	}
	cmd$colorschememapping() {
		this.curGroup = new ParameterGroup(this.curGroup.parent, "colorSchemeMapping");
	}
	cmd$flomajor() {
		this.doc.attributes.fMajor = "ascii";
	}
	cmd$fhimajor() {
		this.doc.attributes.fMajor = "default";
	}
	cmd$fdbmajor() {
		this.doc.attributes.fMajor = "eastasian";
	}
	cmd$fbimajor() {
		this.doc.attributes.fMajor = "complexscripts";
	}
	cmd$flominor() {
		this.doc.attributes.fMinor = "ascii";
	}
	cmd$fhiminor() {
		this.doc.attributes.fMinor = "default";
	}
	cmd$fdbminor() {
		this.doc.attributes.fMinor = "eastasian";
	}
	cmd$fbiminor() {
		this.doc.attributes.fMinor = "complexscripts";
	}

	/* Code Page */
	cmd$cpg(val) {
		this.curGroup.attributes.codePage = val;
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
		this.curGroup.attributes.themeColour = "maindarkone";
	}
	cmd$cmaindarktwo() {
		this.curGroup.attributes.themeColour = "maindarktwo";
	}
	cmd$cmainlightone() {
		this.curGroup.attributes.themeColour = "mainlightone";
	}
	cmd$cmainlighttwo() {
		this.curGroup.attributes.themeColour = "mainlighttwo";
	}
	cmd$caccentone() {
		this.curGroup.attributes.themeColour = "accentone";
	}
	cmd$caccenttwo() {
		this.curGroup.attributes.themeColour = "accenttwo";
	}
	cmd$caccentthree() {
		this.curGroup.attributes.themeColour = "accentthree";
	}
	cmd$caccentfour() {
		this.curGroup.attributes.themeColour = "accentfour";
	}
	cmd$caccentfive() {
		this.curGroup.attributes.themeColour = "accentfive";
	}
	cmd$caccentsix() {
		this.curGroup.attributes.themeColour = "accentsix";
	}
	cmd$chyperlink() {
		this.curGroup.attributes.themeColour = "hyperlink";
	}
	cmd$cfollowedhyperlink() {
		this.curGroup.attributes.themeColour = "followedhyperlink";
	}
	cmd$cbackgroundone() {
		this.curGroup.attributes.themeColour = "backgroundone";
	}
	cmd$cbackgroundtwo() {
		this.curGroup.attributes.themeColour = "backgroundtwo";
	}
	cmd$ctextone() {
		this.curGroup.attributes.themeColour = "textone";
	}
	cmd$ctexttwo() {
		this.curGroup.attributes.themeColour = "texttwo";
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
		this.curGroup.style.shadingPattern = "backwardsdiagonal";
	}
	cmd$tsbgfdiag() {
		this.curGroup.style.shadingPattern = "forwardsdiagonal";
	}
	cmd$tsbgdkbdiag() {
		this.curGroup.style.shadingPattern = "darkbackwardsdiagonal";
	}
	cmd$tsbgdkfdiag() {
		this.curGroup.style.shadingPattern = "darkforwardsdiagonal";
	}
	cmd$tsbgcross() {
		this.curGroup.style.shadingPattern = "cross";
	}
	cmd$tsbgdcross() {
		this.curGroup.style.shadingPattern = "diagonalcross";
	}
	cmd$tsbgdkcross() {
		this.curGroup.style.shadingPattern = "darkcross";
	}
	cmd$tsbgdkdcross() {
		this.curGroup.style.shadingPattern = "darkdiagonalcross";
	}
	cmd$tsbghoriz() {
		this.curGroup.style.shadingPattern = "horizontal";
	}
	cmd$tsbgvert() {
		this.curGroup.style.shadingPattern = "vertical";
	}
	cmd$tsbgdkhor() {
		this.curGroup.style.shadingPattern = "darkhorizontal";
	}
	cmd$tsbgdkvert() {
		this.curGroup.style.shadingPattern = "darkvertical";
	}
	cmd$tsbrdrt() {
		theis.curGroup.style.cellBorder = "top";
	}
	cmd$tsbrdrb() {
		theis.curGroup.style.cellBorder = "bottom";
	}
	cmd$tsbrdrl() {
		theis.curGroup.style.cellBorder = "left";
	}
	cmd$tsbrdrr() {
		theis.curGroup.style.cellBorder = "right";
	}
	cmd$tsbrdrh() {
		theis.curGroup.style.cellBorder = "horizontal";
	}
	cmd$tsbrdrv() {
		theis.curGroup.style.cellBorder = "vertical";
	}
	cmd$tsbrdrdgl() {
		theis.curGroup.style.cellBorder = "diagonalullr";
	}
	cmd$tsbrdrdgr() {
		theis.curGroup.style.cellBorder = "diagonalllur";
	}
	cmd$tscbandsh(val) {
		theis.curGroup.style.rowBandCount = val;
	}
	cmd$tscbandsv(val) {
		theis.curGroup.style.cellBandCount = val;
	}

	/* Style Restrictions */
	cmd$latentstyles() {
		this.curGroup = new StyleRestrictions(this.doc);
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
		this.curGroup = new ParameterGroup(this.curGroup.parent, "taggedName");
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
		this.curGroup.attributes.rsidType = "characterformat";
	}
	cmd$sectrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "sectionformat";
	}
	cmd$pararsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "paragraphformat";
	}
	cmd$tblrsid(val) {
		this.curGroup.attributes.rsid = val;
		this.curGroup.attributes.rsidType = "tableformat";
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "docComment");
	}
	cmd$hlinkbase() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "hlinkBase");
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
		this.curGroup = new DateGroup(this.doc.attributes, "createTime");
	}
	cmd$revtim() {
		this.curGroup = new DateGroup(this.doc.attributes, "revisionTime");
	}
	cmd$printtim() {
		this.curGroup = new DateGroup(this.doc.attributes, "lastPrintTime");
	}
	cmd$buptim() {
		this.curGroup = new DateGroup(this.doc.attributes, "backupTime");
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "passwordHash");
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "nextFile");
	}
	cmd$template() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "template");
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "caption");
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "xForm");
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
		this.curGroup.style.styleFilters = val;
	}
	cmd$readonlyrecommended() {
		this.curGroup.style.readOnlyRecommended = true;
	}
	cmd$stylesortmethod(val) {
		this.curGroup.style.styleSortMethod = val;
	}
	cmd$writereservhash() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "reserveHash");
	}
	cmd$writereservation() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "reservation");
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
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnoteSep");
	}
	cmd$ftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnotesEpc");
	}
	cmd$ftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "footnoteNotice");
	}
	cmd$aftnsep() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pageSep");
	}
	cmd$aftnsepc() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pageSepc");
	}
	cmd$aftncn() {
		this.curGroup = new ParameterGroup(this.curGroup.style, "pageNotice");
	}
	cmd$pages() {
		this.doc.attributes.footnoteposition = "pages";
	}
	cmd$enddoc() {
		this.doc.attributes.footnoteposition = "endDoc";
	}
	cmd$ftntj() {
		this.doc.attributes.footnoteposition = "beneathTextTj";
	}
	cmd$ftnbj() {
		this.doc.attributes.footnoteposition = "bottomOfPageBj";
	}
	cmd$apages() {
		this.doc.attributes.pageposition = "pages";
	}
	cmd$aenddoc() {
		this.doc.attributes.pageposition = "endDoc";
	}
	cmd$aftntj() {
		this.doc.attributes.pageposition = "beneathTextTj";
	}
	cmd$aftnbj() {
		this.doc.attributes.pageposition = "bottomOfPageBj";
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
	cmd$pgnstart(val) {
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
		this.curGroup = new ParameterGroup(this.curGroup.style, "docBackground");
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
		this.curGroup = new ParameterGroup(this.doc.attributes, "followingKinsoku");
	}
	cmd$lchars() {
		this.curGroup = new ParameterGroup(this.doc.attributes, "leadingKinsoku");
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
		this.curGroup = new MailMergeTable(this.doc);
	}
	cmd$mmlinktoquery() {
		this.curGroup.attributes.linkToQuery = true;
	}
	cmd$mmdefaultsql() {
		this.curGroup.attributes.defaultSQL = true;
	}
	cmd$mmconnectstrdata() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectStringData");
	}
	cmd$mmconnectstr() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectString");
	}
	cmd$mmquery() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "connectStringData");
	}
	cmd$mmdatasource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "dataSource");
	}
	cmd$mmheadersource() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "headerSource");
	}
	cmd$mmblanklinks() {
		this.curGroup.attributes.blankLinks = true;
	}
	cmd$mmaddfieldname() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "fieldName");
	}
	cmd$mmmailsubject() {
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "subject");
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
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "udlData");
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
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "mappedName");
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
		this.curGroup = new ParameterGroup(this.curgroup.parent.attributes, "uniqueTag");
	}

	cmd$mmfttypenull() {
		this.curGroup.attributes.dataType = "null";
	}
	cmd$mmfttypedbcolumn() {
		this.curGroup.attributes.dataType = "databasecolumn";
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
		this.curGroup.attributes.connectionType = "externalquery";
	}
	cmd$mmdatatypeodbc() {
		this.curGroup.attributes.connectionType = "odbc";
	}
	cmd$mmdatatypeodso() {
		this.curGroup.attributes.connectionType = "odso";
	}
	cmd$mmdatatypefile() {
		this.curGroup.attributes.connectionType = "dde-textfile";
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
		this.curGroup.style.listStyle = val;
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
		this.curGroup.style.lineRestart = "onlinestarts";
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
		this.curGroup.style.pageNumberRestart = "onpagenumber";
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
		this.curGroup.style.pageNumbering = "KANJINODIGIT";
	}
	cmd$pgndbnumd() {
		this.curGroup.style.pageNumbering = "KANJIDIGIT";
	}
	cmd$pgndbnumt() {
		this.curGroup.style.pageNumbering = "DBNUM3";
	}
	cmd$pgndbnumk() {
		this.curGroup.style.pageNumbering = "DBNUM4";
	}
	cmd$pgndecd() {
		this.curGroup.style.pageNumbering = "DOUBLEBYTE";
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
		this.curGroup.style.pageNumbering = "HINDIVOWEL";
	}	
	cmd$pgnhindib() {
		this.curGroup.style.pageNumbering = "HINDICONSONANT";
	}
	cmd$pgnhindic() {
		this.curGroup.style.pageNumbering = "HINDIDIGIT";
	}
	cmd$pgnhindid() {
		this.curGroup.style.pageNumbering = "HINDIDESCRIPTIVE";
	}
	cmd$pgnthaia() {
		this.curGroup.style.pageNumbering = "THAILETTER";
	}
	cmd$pgnthaib() {
		this.curGroup.style.pageNumbering = "THAIDIGIT";
	}
	cmd$pgnthaic() {
		this.curGroup.style.pageNumbering = "THAIDESCRIPTIVE";
	}
	cmd$pgnvieta() {
		this.curGroup.style.pageNumbering = "VIETNAMESEDESCRIPTIVE";
	}
	cmd$pgnid() {
		this.curGroup.style.pageNumbering = "KOREANDASH";
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
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "header");
	}
	cmd$footer() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "footer");
	}
	cmd$headerl() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "headerLeft");
	}
	cmd$headerr() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "headerRight");
	}
	cmd$headerf() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "headerFirst");
	}
	cmd$footerl() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "footerLeft");
	}
	cmd$footerr() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "footerRight");
	}
	cmd$footerf() {
		this.curGroup = new ParameterGroup(this.curGroup.parent.style, "footerFirst");
	}

	/* Paragraphs */
	cmd$par() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			let prevStyle = this.curGroup.curstyle;
			if (prevStyle.contextualspace) {
				const spacings = [
				"spaceBefore",
				"spaceAfter",
				"autoSpaceBefore",
				"autoSpaceAfter",
				"spaceBeforeChar",
				"spaceAfterChar",
				];
				Object.keys(prevStyle).forEach(key => {
					if (spacings.includes(key)) {
						prevStyle[key] = false;
					}	
				});
			}
			const prevAtt = this.curGroup.curattributes;
			this.endGroup();
			this.newGroup("paragraph");
			this.curGroup.style = prevStyle;
			this.curGroup.attributes = prevAtt;
		} else {
			this.newGroup("paragraph");
		}	
	}
	cmd$pard() {
		if (this.paraTypes.includes(this.curGroup.type)) {
			this.curGroup.style = Object.assign(JSON.parse(JSON.stringify(this.defCharState)),JSON.parse(JSON.stringify(this.defParState)));
			this.curGroup.attributes = {};
		} else {
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
		this.curGroup.style.textboxWrap = "firstlast";
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
		this.curGroup.style.tabType = "flushright";
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
		this.curGroup.style.leader = "middledots";
	}
	cmd$tlhyph() {
		this.curGroup.style.leader = "hyphen";
	}
	cmd$tlul() {
		this.curGroup.style.leader = "underline";
	}
	cmd$tlth() {
		this.curGroup.style.leader = "thickline";
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
		this.curGroup.relativeTo = "lefttomargin";
	}
	cmd$pmartabqc() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "centertomargin";
	}
	cmd$pmartabqr() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "righttomargin";
	}
	cmd$pindtabql() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "lefttoindent";
	}
	cmd$pindtabqc() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "centertoindent";
	}
	cmd$pindtabqr() {
		if (this.curGroup.type !== "atab") {this.curGroup = new ATab(this.curGroup.parent);}
		this.curGroup.relativeTo = "righttoindent";
	}

	/* Bullets and Numbering */





	



	/* Tables */
	cmd$trowd() {
		newGroup("tablerow");
	}
	cmd$row() {
		if (this.curGroup.type === "tablerow") {endGroup()}
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
		this.curGroup.attributes.source = "OS/2 METAFILE";
		this.curGroup.attributes.sourceType = val;
	}
	cmd$wmetafile(val) {
		this.curGroup.attributes.source = "WINDOWS METAFILE";
		this.curGroup.attributes.mappingMode = val;
	}
	cmd$dibitmap(val) {
		this.curGroup.attributes.source = "WINDOWS DI BITMAP";
		this.curGroup.attributes.sourceType = val;
	}
	cmd$wbitmap(val) {
		this.curGroup.attributes.source = "WINDOWS DD BITMAP";
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