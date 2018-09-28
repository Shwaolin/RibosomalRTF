const fs = require("fs");
const Readable = require("stream").Readable;
const Transform = require("stream").Transform;
const SmallRTFSubunit = require("./smallRTFSubunit.js");
const LargeRTFSubunit = require("./largeRTFSubunit.js");

exports.parseString = function(rtf) {
	let str = new Readable();
	str.push(rtf);
	str.push(null);
	return parse(str);
}

exports.parseFile = function(rtf) {
	return parse(fs.createReadStream(rtf));
}

exports.parseStream = function(rtf) {
	return parse(rtf);
}

function parse(rtf) {
	const reader = new SmallRTFSubunit;
	const writer = new LargeRTFSubunit;
	const charStreamSplitter = new Transform({
		transform(chunk, encoding, callback) {
			chunk.toString().split("").forEach(char => this.push(char));
			callback();
		}
	})

	function errorHandler(error) {
		console.log("Error encountered! Shutting down!");
		writer.destroy(error);
		reader.destroy(error);
		charStreamSplitter.destroy(error);
	}

	charStreamSplitter.on("error", errorHandler);
	reader.on("error", errorHandler);
	writer.on("error", errorHandler);

	rtf.on("end", () => {
		console.log("RTF!");
		charStreamSplitter.end("waa");
	});

	charStreamSplitter.on("end", () => {
		console.log("Splitter!");
		reader.end("Finished!");
	});

	reader.on("finish", () => {
		console.log("Reader!");
	});

	reader.output.on("end", () => {
		console.log("Reader Output!");
	});

	writer.on("finish", () => {
		console.log("Writer!");
	});

	rtf.pipe(charStreamSplitter, {end:false})
		.pipe(reader).output
		.pipe(writer)
		.on('error', errorHandler)
		.on("finish",()=>console.dir(writer.output));
}

