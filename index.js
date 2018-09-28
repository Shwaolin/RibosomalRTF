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
	});

	function errorHandler(error) {
		console.error(error);
		writer.destroy(Error("Error encountered! Shutting down Writer!"));
		reader.destroy(Error("Error encountered! Shutting down Reader!"));
		charStreamSplitter.destroy(Error("Error encountered! Shutting down Splitter!"));
	}

	charStreamSplitter.on("error", errorHandler);
	reader.on("error", errorHandler);
	writer.on("error", errorHandler);

	charStreamSplitter.on("end", () => {
		reader.end("Finished!");
	});

	return new Promise((resolve, reject) => {
		try {
			rtf.pipe(charStreamSplitter)
			.pipe(reader).output
			.pipe(writer)
			.on("finish",()=>resolve(writer.output));
		}
		catch(err) {
			console.log(err);
		}
		
	});
}

