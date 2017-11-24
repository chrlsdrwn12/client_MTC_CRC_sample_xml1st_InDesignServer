/*
Copyright 2017 Manila Typesetting Company
www.mtcstm.com
All rights reserved

Developer: Darwin Lopena (info@darwinlopena.com).
Date: Nov 23, 2017
Save as cc2017 document to cs5.5
*/

var dummyFolder = new Folder("c:/CRC/");
var indesignTemplateFile = new File(dummyFolder + '/template.idml');
var template = new File(dummyFolder + '/cs5.indd');
var myDoc = app.open(indesignTemplateFile);
myDoc.save(template);
myDoc.close();