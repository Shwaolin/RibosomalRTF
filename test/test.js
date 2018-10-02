const RibosomalRTF = require("../index");
const util = require('util');
const { performance } = require('perf_hooks');

const stylesheet = `{\\stylesheet{\\ql \\li0\\ri0\\widctlpar\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 \\fs24\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033 \\snext0 Normal;}
{\\*\\cs10 \\additive Default Paragraph Font;}{\\*\\cs15 \\additive \\b\\ul\\cf6 \\sbasedon10 UNDERLINE;} {\\*\\ts11\\tsrowd\\trftsWidthB3\\trpaddl108\\trpaddr108\\trpaddfl3 \\trpaddft3\\trpaddfb3\\trpaddfr3\\tscellwidthfts0\\tsvertalt\\tsbrdrt\\tsbrdrl\\tsbrdrb\\tsbrdrr\\tsbrdrdgl\\tsbrdrdgr\\tsbrdrh\\tsbrdrv \\ql \\li0\\ri0\\widctlpar\\aspalpha\\aspnum\\faauto\\adjustright\\rin0 \\lin0\\itap0 \\fs20\\lang1024\\langfe1024\\cgrid\\langnp1024 \\langfenp1024 \\snext11 \\ssemihidden Normal Table; }{\\s16\\qc \\li0\\ri0\\widctlpar\\aspalpha\\aspnum\\faauto\\adjustright\\rin0\\lin0\\itap0 \\b\\fs24\\cf2\\lang1033\\langfe1033\\cgrid\\langnp1033\\langfenp1033 \\sbasedon0 \\snext16 \\sautoupd CENTER;}}
`

function test() {
	const start = performance.now();
	/*RibosomalRTF.parseFile("./test/234.rtf").then(obj => {
		const timeElapsed = performance.now() - start;
		console.log(util.inspect(obj,{depth:null,colors:true,maxArrayLength:null,compact:false}));
		console.log("(" + timeElapsed + " milliseconds)");	
	});*/
	RibosomalRTF.parseString(stylesheet).then(obj => {
		const timeElapsed = performance.now() - start;
		console.log(util.inspect(obj.stylesheet,{depth:null,colors:true,maxArrayLength:null,compact:false}));
		console.log("(" + timeElapsed + " milliseconds)");	
	});
}

test();
