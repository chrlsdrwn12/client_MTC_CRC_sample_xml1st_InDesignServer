/*
Copyright 2017 Manila Typesetting Company
www.mtcstm.com
All rights reserved

Developer: Darwin Lopena (info@darwinlopena.com).
Date: Nov 23, 2017
Custom JSX script file for automatically importing xml and producing a print pdf
*/

var dummyFolder = new Folder("c:/CRC/");
var indesignTemplateFile = new File(dummyFolder + '/template.indd');
importXML(indesignTemplateFile);

function importXML(template){
        var myDoc = app.open(template);
        setXMLPrefs(myDoc);
        var xmlfile = File(getXMLFile(myDoc));
        // app.consoleout(xmlfile.filePath);
        if (xmlfile != null) {
            myDoc.importXML(xmlfile);
            var xmlRoot = myDoc.xmlElements[0];
            var textFrame = myDoc.pages[0].textFrames[0];
            xmlRoot.placeXML(textFrame);
            mapTagsToStyles(myDoc);
            saveAs(xmlfile, myDoc);
        }
}

function setXMLPrefs(myDoc){
    with(myDoc.xmlImportPreferences){
        importStyle = XMLImportStyles.MERGE_IMPORT;
        createLinkToXML = false;
        repeatTextElements = true;
        ignoreUnmatchedIncoming = false;
        importTextIntoTables = true;
        ignoreWhitespace = false;
        removeUnmatchedExisting = false;
        importCALSTables = true;
        //allowTransform = true;
        //transformFilename = File("path/to/file.xsl");
    }
}

function mapTagsToStyles(myDoc) {
    //var mappings = {"Heads:BH": "bh", "Text:T": "p"};
    var mappings = readMappings(myDoc);
    for (var key in mappings){
        pStyle = key.split(':');
        xTag = mappings[key];
        try{
            myDoc.xmlImportMaps.add(xTag, myDoc.paragraphStyleGroups.item(pStyle[0]).paragraphStyles.item(pStyle[1]));
        }
        catch(err) {}
    }
    var charMappings = {"cItalic": "italic", "cBold": "bold", "cItalic-ref":"source"};
    for (var key in charMappings){
        xTag = charMappings[key];
        //alert('tag: '+key+' | style: '+xTag);
        try{
            myDoc.xmlImportMaps.add(xTag, myDoc.characterStyles.item(key));
        }
        catch(err) {}
    }
    myDoc.mapXMLTagsToStyles();
}

function readMappings(myDoc){
    var mapFile = File(myDoc.filePath+"/CRCStyleMaps.txt");
    mapFile.open("r");
    var str = mapFile.read();
    mapFile.close();
    var array = str.split(/[\r\n]+/);
    var dict = {};
    for(var i=0; i<array.length; i++) {//
        k = array[i].split('|')[0];
        v = array[i].split('|')[1];
        dict[k] = (v);
    }
    return dict;
}

function getXMLFile(myDoc){
    if(myDoc.filePath.getFiles(filterFile).length > 0){
        return myDoc.filePath.getFiles(filterFile)[0];
    }
}

function filterFile(currentFile){
    var strExt = currentFile.fullName.toLowerCase().substr(currentFile.fullName.lastIndexOf("."));
    return strExt == ".xml";
}

function saveAs(xmlfile, myDoc){
    var filename = xmlfile.name.replace('.xml','');
    var myNewFile = new File(myDoc.filePath+'/'+ filename + '.indd');
    myDoc.save(myNewFile);
    var outputFile = new File(myDoc.filePath+'/'+ filename + '.pdf');
    myDoc.exportFile(ExportFormat.pdfType, outputFile);
    myDoc.close();
}