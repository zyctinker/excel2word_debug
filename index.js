var excel2Json = require('node-excel-to-json');
var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');

var content = fs.readFileSync(path.resolve(__dirname, 'input.docx'), 'binary');
var zip = new JSZip(content);
var doc = new Docxtemplater();
doc.loadZip(zip);
excel2Json('../../../sample.xlsx', {
    'convert_all_sheet': false,
    'return_type': 'Object',
    'sheetName': 'phonenumber'
}, function(err, output) {
    //输出单人
    var i = 0;
    var buf = new Array();
    var nowPerson = {
     name : output[1].author,
     data : {"record": []}
     }
    for(var singleRecord in output) {
        console.log(output[singleRecord].author);
        if((output[singleRecord].author===nowPerson.name)){//同一个人放到同一个数组里
            //console.log(output[singleRecord]);
            nowPerson.data.record.push(output[singleRecord]);
        }
        else{
            console.log(nowPerson.data);
            doc.setData(nowPerson.data);
            try {
                doc.render();
            }
            catch (error) {
                var e = {
                    message: error.message,
                    name: error.name,
                    stack: error.stack,
                    properties: error.properties,
                }
                console.log(JSON.stringify({error: e}));
                // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
                throw error;
            }
            buf = doc.getZip()
                .generate({type: 'nodebuffer'});
            fs.writeFileSync(path.resolve(__dirname, nowPerson.name+'.docx'), buf);
            nowPerson.name = output[singleRecord].author;
            nowPerson.data.record = [];
            nowPerson.data.record.push(output[singleRecord]);
        }
    }
    doc.setData(nowPerson.data);
    try {
        doc.render()
    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({error: e}));
        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
        throw error;
    }
    var buf = doc.getZip()
        .generate({type: 'nodebuffer'});
    fs.writeFileSync(path.resolve(__dirname, nowPerson.name+'.docx'), buf);

});