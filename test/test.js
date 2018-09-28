const RibosomalRTF = require("../index");

function test() {
	RibosomalRTF.parseFile("./test/234.rtf").then(obj=>console.dir(obj));
	
}

test();
