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
		this.tables = {};
		this.type = "document";
	}
	dumpContents() {
		return {
			tables:this.tables,
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
			this.parent[this.param] = this.contents[0].replace(/["]/g,"");
		}		
	}
}

class Default extends RTFObj {
	constructor(parent, writer, styletype) {
		super(parent);
	}
	dumpContents() {
		this.parent.tables.defaults[styletype] = {
			style: this.curstyle,
			attributes: this.curattributes
		}
		writer = Object.assign(this.curstyle, this.curattributes)
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
		this.listType = true;
		this.red = null;
		this.blue = null;
		this.green = null;
		this.attributes = {};
	}
	flush() {
		if (this.red !== null) {
			this.table.push({
				red: this.red,
				green: this.green,
				blue: this.blue,
				attributes: this.attributes
			});
		}
		this.red = null;
		this.blue = null;
		this.green = null;
		this.attributes = {};
	}
	dumpContents() {
		this.doc.tables.colourTable = this.table;
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
		this.doc.tables.fontTable = this.table;	
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

class FileTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.fileTable = this.table;	
	}
}

class File extends RTFObj {
	constructor (parent) {
		super(parent);
		this.attributes = {};
	}
	dumpContents() {
		this.parent.table.push({
			attributes: this.attributes,
			filename: this.contents[0].replace(";","")
		});
	}
}

class Stylesheet extends DocTable {
	constructor(doc) {
		super(doc);
		this.sheet = {};
		this.contents = [];
	}
	dumpContents() {
		this.doc.tables.styleSheet = this.sheet;
	}
}

class Style extends RTFObj {
	constructor(parent, designation) {
		super(parent);
		this.designation = designation;
	}
	dumpContents() {
		this.parent.sheet[this.designation] = ({
			style: this.curstyle,
			attributes: this.curattributes,
			name: this.contents[0].replace(/;/g,"")
		});
	}
}

class StyleRestrictions extends DocTable {
	constructor(doc) {
		super(doc);
		this.attributes = {};
		this.lockedexceptions = "";
	}
	dumpContents() {
		if (typeof this.lockedexceptions === "string") {
			this.lockedexceptions = this.lockedexceptions.split(";");
		}
		this.doc.tables.styleRestrictions = {
			attributes: this.attributes,
			lockedexceptions: this.lockedexceptions
		}
	}
}

class ListTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.listTable = this.table;
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
		this.doc.tables.listOverrideTable = this.table;
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
	Field, 
	Fldrslt, 
	Picture
}