const Writable = require("stream").Writable;
const Readable = require("stream").Readable;

class SmallRTFSubunit extends Writable {
	constructor() {
		super({
			write(chunk, encoding, callback) {
				chunk = chunk.toString()
				if (chunk !== "Finished!") {
					this.curChar = chunk;
					this.operation(this.curChar);
				} else {
					this.setInstruction({type:"documentEnd"});
					this.output.push(null);
				}
				callback();
			}
		});
		this.curInstruction = {type: "", value: ""};
		this.curChar = "";
		this.output = new Readable({read() {}});
		this.operation = this.parseText;
		this.docGroup = false;
	}
	parseText(char) {
		switch(char) {
			case "\\": 
				this.operation = this.parseEscape;
				break;
			case "{": 
				this.setInstruction();
				if (this.docGroup) {
					this.setInstruction({type:"groupStart"});
				} else {
					this.setInstruction({type:"documentStart"});
					this.docGroup = true;
				}
				break;
			case "}": 
				this.setInstruction();
				this.setInstruction({type:"groupEnd"});
				break;
			case "\n": 
				this.setInstruction();
				this.setInstruction({type:"break"});
				break;
			case "\r":
				this.setInstruction();
				this.setInstruction({type:"break"});
				break;
			default: 
				this.curInstruction.type = "text";
				this.curInstruction.value += char;
		}
	}
	parseEscape(char) {
		if (char.search(/[ \\{}\n\r]/) === -1) {
			this.setInstruction();
			this.operation = this.parseControl;
			this.parseControl(char);
		} else if (char.search(/[\n\r]/) !== -1){
			this.curInstruction.value += char + char;
			this.parseText(char);
			this.operation = this.parseText;
		} else {
			this.operation = this.parseText;
			this.curInstruction.type = "text";
			this.curInstruction.value += char;
		}
	}
	parseControl(char) {
		if (char.search(/[ \\{}\t'\n\r;*]/) === -1) {
			this.curInstruction.type = "control";
			this.curInstruction.value += char;
		} else if (char === "*") {
			this.setInstruction({type:"ignorable"});
		} else if (char === "'") {
			this.operation = this.parseHex;
			this.curInstruction.type = "control";
			this.curInstruction.value += "hex";
		} else if (char === " ") {
			this.setInstruction();
			this.operation = this.parseText;
		} else if (char === ";") {
			this.setInstruction();
			this.setInstruction({type:"listBreak"});
			this.operation = this.parseText;
		} else {
			this.setInstruction();
			this.operation = this.parseText;
			this.parseText(char);
		}
	}
	parseHex(char) {
		if (this.curInstruction.value.length >= 5) {
			this.setInstruction();
			this.operation = this.parseText;
			this.parseText(char);
		} else {
			this.curInstruction.value += char;
		}
	}
	setInstruction(instruction = this.curInstruction) {
		if (instruction.type !== "") {
			this.output.push(JSON.stringify(instruction));
		}
		this.curInstruction = {type: "", value: ""};
	}
}

module.exports = SmallRTFSubunit;