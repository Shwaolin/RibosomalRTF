<h3 align="center"> --Current specification progress: 89/278 pages-- </h3>

## What is RibosomalRTF?
RibosomalRTF is a parser designed to convert RTF documents into navigable document objects. This by itself is not that useful; however, it is a crucial step in extracting data from the format or converting it into other languages. Unlike HTML or XML, RTF is not a markup language, which can make it annoying to work with programmatically. RibosomalRTF aims to ease this process, and create a base on which other parsers can be built.

## How does it work?
RibosomalRTF works in two parts—the small subunit and the large subunit. Similar to the manner in which a ribosome reads RNA and produces proteins, the small subunit reads an RTF document one character at a time and produces a stream of instructions which are fed to the large subunit. The large subunit then builds a document object based on these instructions.

This design is largely based upon the structure used here: [rtf-parser](https://github.com/iarna/rtf-parser).

## How do I use it?
RibosomalRTF exports three functions: parseString, parseFile, and parseStream. Each returns a promise with the document object.
```javascript
const RibosomalRTF = require("RibosomalRTF");

//parseString - takes a string; returns a promise.
RibosomalRTF.parseString("{\\rtf1\\ansi Example}").then(obj => ...);

//parseFile - takes a filepath as a string; returns a promise.
RibosomalRTF.parseFile("./example.rtf").then(obj => ...);

//parseStream - takes a bytestream; returns a promise.
RibosomalRTF.parseStream(myStream).then(obj => ...);
```

## Is it done yet?
**No!** But you can help us get it done sooner by making a pull request, picking something from the RTF Specification that looks interesting, and adding support for it. Or, if writing documentation is more your thing, by contributing to our wiki here on Github.

