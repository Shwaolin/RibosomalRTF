const RibosomalRTF = require("../index");
const util = require('util');

function test() {
	RibosomalRTF.parseFile("./test/234.rtf").then(obj => {
		console.log(util.inspect(obj,{depth:null,colors:true,maxArrayLength:null,compact:false}));
	});
}

test();
