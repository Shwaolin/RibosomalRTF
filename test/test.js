const RibosomalRTF = require("../index");
const util = require('util');
const { performance } = require('perf_hooks');

function test() {
	const start = performance.now();
	RibosomalRTF.parseFile("./test/234.rtf").then(obj => {
		const timeElapsed = performance.now() - start;
		console.log(util.inspect(obj,{depth:null,colors:true,maxArrayLength:null,compact:false}));
		console.log("(" + timeElapsed + " milliseconds)");
		
	});
}

test();
