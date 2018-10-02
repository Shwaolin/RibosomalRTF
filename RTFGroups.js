class RTFObj {
	constructor(parent) {
		this.parent = parent;
		this.style = {};
		this.attributes = {};
		this.contents = [];
		this.type = "";
	}
	get curstyle() {
		return JSON.parse(JSON.stringify(this.style));
	}
	get curattributes() {
		return JSON.parse(JSON.stringify(this.attributes));
	}
}

class RTFDoc extends RTFObj {
	constructor(parent) {
		super(null);
		this.colourTable = [];
		this.fontTable = [];
		this.listTable = [];
		this.listOverrideTable = [];
		this.type = "document";
	}
	dumpContents() {
		return {
			colourtable: this.colourTable,
			fonttable: this.fontTable,
			listtable: this.listTable,
			listoverridetable: this.listOverrideTable,
			style: this.curstyle,
			attributes: this.curattributes,
			contents: this.contents
		};
	}
}

class RTFGroup extends RTFObj {
	constructor(parent, type) {
		super(parent);
		this.type = type;
	}
	dumpContents() {
		if (this.contents[0] && this.contents.every(entry => typeof entry === "string") && this.contents.every (entry => entry.style === this.contents[0].style)) {
			this.contents = this.contents.join("");
			if (this.type === "span") {this.type = "text";}
		}
		this.parent.contents.push({
			contents: this.contents,
			style: this.curstyle,
			attributes: this.curattributes,
			type: this.type
		});
	}
}

class ParameterGroup extends RTFObj {
	constructor (parent, parameter) {
		super(parent);
		this.param = parameter;
	}
	dumpContents() {
		if (this.contents[1] && this.contents.every(entry => typeof entry === "string")) {
			this.contents = this.contents.join("");
		}
		if (this.contents[0]) {
			this.parent[this.param] = this.contents[0].replace(/[;"]/g,"");
		}		
	}
}

class DocTable {
	constructor(doc) {
		this.doc = doc;
		this.table = [];
	}
}

class ColourTable extends DocTable {
	constructor(doc) {
		super(doc);
		this.rgb = {};
	}
	addColour(colour, value) {
		this.rgb[colour] = value;
		if (Object.keys(this.rgb).length === 3) {
			this.table.push(this.rgb);
			this.rgb = {};
		}
	}
	dumpContents() {
		this.doc.colourTable = this.table;
	}
}

class FontTable extends DocTable {
	constructor(doc) {
		super(doc);
		this.attributes = {};
		this.contents = [];
	}
	dumpContents() {
		if (!this.table[0] && this.contents[0]) {
			this.table.push ({
				fontname: this.contents[0].replace(";",""),
				attributes: this.attributes
			});
		}
		this.doc.fontTable = this.table;	
	}
}

class Font extends RTFObj{
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		this.parent.table.push({
			fontname: this.contents[0].replace(";",""),
			attributes: this.curattributes
		});
	}
}

class ListTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.listTable = this.table;
	}
}

class List extends RTFObj {
	constructor (parent) {
		super(parent);
		this.templateid = null;
		this.id = null;
		this.listname = "";
	}
	dumpContents() {
		this.attributes.listname = this.listname;
		this.parent.table.push({
			templateid: this.templateid,
			id: this.id,
			levels: this.contents,
			attributes: this.curattributes,
		});
	}
}

class ListLevel extends RTFObj{
	constructor (parent) {
		super(parent);
	}
	dumpContents() {
		this.attributes.leveltext = this.leveltext;
		this.attributes.levelnumbers = this.levelnumbers;
		this.parent.contents.push({
			style:this.curstyle,
			attributes: this.curattributes,
		});
	}
}

class ListOverrideTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.listOverrideTable = this.table;
	}
}

class ListOverride extends RTFObj {
	constructor(parent) {
		super(parent);
		this.id = null;
		this.ls = null;
	}
	dumpContents() {
		this.parent.table.push({
			attributes: this.curattributes,
			id: this.id,
			ls: this.ls
		});
	}
}

class Field extends RTFObj {
	constructor(parent) {
		super(parent);
		this.fieldInst = "";
		this.contents = "";
		this.type = "field";
	}
	dumpContents() {
		const fieldInstProps = this.fieldInst.split(" ");
		this.attributes.fieldtype = fieldInstProps[0];
		this.attributes.fieldvalue = fieldInstProps[1];
		this.parent.contents.push({
			attributes: this.curattributes,
			contents: this.contents,
			style: this.curstyle,
			type: this.type
		});
	}
}

class Fldrslt extends RTFObj {
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		this.parent.style = this.curstyle;
		this.parent.contents = this.contents[0];
	}
}

class Picture extends RTFObj {
	constructor(parent) {
		super(parent);
		this.contents = [];
		this.type = "picture"
	}
	dumpContents() {
		this.parent.contents.push({
			attributes: this.curattributes,
			image: this.contents,
			style: this.curstyle,
			type: this.type
		});
	}
}

module.exports = {
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
}