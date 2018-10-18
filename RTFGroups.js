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
	constructor (destination, parameter) {
		super(destination);
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
				fontName: this.contents[0].replace(";",""),
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
			fontName: this.contents[0].replace(";",""),
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
			fileName: this.contents[0].replace(";","")
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
		this.doc.tables.stylesheet = this.sheet;
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
			lockedExceptions: this.lockedexceptions
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
			templateID: this.templateid,
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

class ParagraphGroupTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.paragraphGroupTable = this.table;
	}
}

class ParagraphGroup extends RTFObj {
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		this.parent.table.push({
			attributes: this.curattributes,
			style: this.curstyle
		});
	}
}

class RevisionTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.revisionTable = this.table;
	}
}

class RSIDTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.rsidTable = this.table;
	}
}

class ProtectedUsersTable extends DocTable {
	constructor(doc) {
		super(doc);
		this.contents = [];
	}
	dumpContents() {
		this.doc.tables.protectedUsersTable = this.contents;
	}
}

class UserProperty extends RTFObj {
	constructor(doc) {
		super(doc);
		this.propertyname = "";
		this.propertytype = null;
		this.propertyvalue = "";
		this.propertylink = false;
	}
	dumpContents() {
		this.doc.attributes.userAttributes[this.propertyname] = {
			type: this.propertytype,
			value: this.propertyvalue,
			link: this.propertylink,
		}
	}
}

class XMLNamespaceTable extends DocTable {
	constructor(doc) {
		super(doc);
	}
	dumpContents() {
		this.doc.tables.xmlNamespaceTable = this.table;
	}
}

class XMLNamespace extends RTFObj {
	constructor(parent, id) {
		super(parent);
		this.id = id;
	}
	dumpContents() {
		this.parent.table.push({
			id: this.id,
			namespace: this.contents[0]
		});
	}
}

class MailMergeTable extends DocTable {
	constructor(doc) {
		super(doc);
		this.attributes = {};
		this.mmodso = [];
		this.mmodrecip = [];
	}
	dumpContents() {
		this.doc.mailMergeTable = {
			odso: this.mmodso,
			recipData: this.mmodsorecip,
			attributes: this.attributes
		}
	}
}

class Odso extends DocTable {
	constructor(parent) {
		super(parent);
		this.attributes = {};
	}
	dumpContents() {
		this.parent.mmodso.push({
			attributes: this.attributes,
			fieldMapData: this.table
		});
	}
}

class FieldMap extends RTFObj {
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		this.parent.table.push(this.attributes);
	}
}

class OdsoRecip extends RTFObj {
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		this.parent.mmodsorecip.push(this.attributes);
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

class DateGroup extends RTFObj {
	constructor(parent, propname) {
		super(parent);
		this.propname = propname;
		this.year = null;
		this.month = null;
		this.day = null;
		this.hour = null;
		this.minute = null;
		this.second = null;
	}
	dumpContents() {
		let epoch = Date.parse(this.year+"-"+this.month+"-"+this.day+"T"+this.hour+":"+this.minute+":"+this.second+"Z");
		this.parent[this.propname] = {
			year: this.year,
			month: this.month,
			day: this.day,
			hour: this.hour,
			minute: this.minute,
			second: this.second,
			sinceEpoch: epoch
		}
	}
}

class NonGroup extends RTFObj {
	constructor(parent) {
		super(parent);
	}
	dumpContents() {
		return;
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
}