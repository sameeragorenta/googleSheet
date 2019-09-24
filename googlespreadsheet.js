const express = require('express');
const multer = require('multer');
const upload = multer({
  dest: 'uploads/' // this saves your file into a directory called "uploads"
});
const readline = require('readline');
const {google} = require('googleapis');
const bodyParser = require('body-parser');

const SCOPES = ['https://www.googleapis.com/auth/documents','https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'token.json';

const app = express();
var fs = require('fs');

function getCredentials(){
    return new Promise((resolve, reject) => {
        fs.readFile('credentials.json', (err, content) => {
        if (err) return console.log('Error loading client secret file:', err);
        resolve(JSON.parse(content));
        });
    });
}

function authorize(credentials) {
    return new Promise((resolve, reject) => {
        const {client_secret, client_id, redirect_uris} = credentials.installed;
        const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
        fs.readFile(TOKEN_PATH, (err, token) => {
        oAuth2Client.setCredentials(JSON.parse(token));
        resolve(oAuth2Client);
        });
    });
}

app.use(bodyParser.json({limit: '50mb',type: 'application/json'}));
app.use(bodyParser.urlencoded({ extended: false }));

var type = upload.single('fileToUpload');
app.get('/index', function (req, res) {
    res.sendFile( __dirname + "/" + "index.htm" );
})

app.post('/createFolder', type, function (req,res) {
    getCredentials()
        .then(function(creds) {
            return authorize(creds);
        })
        .then(function(oAuth2Client) {
            return createFolder(req,oAuth2Client);
        })
        .then(function(resid) {
            res.end(resid);
        })
        .catch(function(error) {
            console.log(error);
        });
});

app.post('/createDocInGoogleDrive', type, function (req,res) {
  //console.log(req.body.redirectUrl);
    // if(req.file.mimetype!='application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
    //     res.end( JSON.stringify( 'Unsupported format: docx files are accepted' ) );
    // }else{
        getCredentials()
            .then(function(creds) {
                return authorize(creds);
            })
            .then(function(oAuth2Client) {
                return createDocInGoogleDrive(req,oAuth2Client);
            })
            .then(function(resid) {
                //res.end(resid);
                var folderID=req.body.folderId;
                var templateName=req.body.templateName;
                res.redirect(req.body.redirectUrl+'/?fileid=' + resid+"&folderId="+folderID+"&templateName="+templateName);
            })
            .catch(function(error) {
                console.log(error);
                res.end('failed');
            });
    // }
});

app.post('/createEmptyDocInGoogleDrive', type, function (req,res) {
  //console.log(req.body.redirectUrl);
    // if(req.file.mimetype!='application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
    //     res.end( JSON.stringify( 'Unsupported format: docx files are accepted' ) );
    // }else{
        getCredentials()
            .then(function(creds) {
                return authorize(creds);
            })
            .then(function(oAuth2Client) {
                return createEmptyDocInGoogleDrive(req,oAuth2Client);
            })
            .then(function(resid) {
                res.end(resid);
            })
            .catch(function(error) {
                console.log(error);
                res.end('failed');
            });
    // }
});

app.post('/mergeTagwithFileId', function (req,res) {
    // console.log(req.header('Content-Type'))
    // console.log(req.body); // <====
    // res.end();
    //res.status(200); // will give { name: 'Lorem',age:18'} in response
    getCredentials()
        .then(function(creds) {
            return authorize(creds);
        })
        .then(function(oAuth2Client) {
            return mergeTagwithTagId(req,oAuth2Client);
        })
        .then(function(resid) {
            res.end(resid);
        })
        .catch(function(error) {
            console.log(error);
            res.end('failed');
        });
});

app.post('/duplicateStandardTemplate', function (req,res) {
    // console.log(req.header('Content-Type'))
    // console.log(req.body); // <====
    // res.end();
    //res.status(200); // will give { name: 'Lorem',age:18'} in response
    getCredentials()
        .then(function(creds) {
            return authorize(creds);
        })
        .then(function(oAuth2Client) {
            return duplicateStandardTemplate(req,oAuth2Client);
        })
        .then(function(resid) {
            res.end(resid);
        })
        .catch(function(error) {
            console.log(error);
            res.end('failed');
        });
});

function createFolder(req,auth){
  return new Promise((resolve, reject) => {
    const drive = google.drive({version: 'v3', auth});
    console.log(req);
    var fileMetadata = {
        'name': req.body.folderName,
        'mimeType': 'application/vnd.google-apps.folder'
    };
    drive.files.create({
        resource: fileMetadata,
        fields: 'id'
    }, function (err, file) {
        if (err) {
            // Handle error
            console.log(err);
            reject('failed');
        } else {
            console.log('Folder Id: ', file.data.id);
            resolve(file.data.id);
        }
    });
  });
}

function createDocInGoogleDrive(req,auth) {
    return new Promise((resolve, reject) => {
        const drive = google.drive({version: 'v3', auth});
        const docs = google.docs({version: 'v1', auth});
        //console.log(req.body.folderId);
        if(req.body.folderId==""){
            resolve('folder is missing');
        }
        var fileMetadata = {
          'title': req.file.originalname,
          'name': req.file.originalname,
          'parents': [req.body.folderId],
          'mimeType': 'application/vnd.google-apps.document'
        };
        var media = {
            //mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            body: fs.createReadStream(req.file.path)
        };
        drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id'
        }, function (err, file) {
            if (err) {
                console.error("error from create file"+err);
                reject('failed');
            } else {
                console.log('File Id:', JSON.stringify(file));
                drive.permissions.create({
                    fileId: file.data.id,
                    resource:{
                        "role": "writer",
                        "type": "anyone"
                    }
                }, function(err,result){
                    if(err){ console.log(err);reject('failed');}
                      else console.log(result)
                });
                console.log(file.data);
                resolve(file.data.id);
            }
        });
    });
}

function createEmptyDocInGoogleDrive(req,auth) {
    return new Promise((resolve, reject) => {
        const drive = google.drive({version: 'v3', auth});
        const docs = google.docs({version: 'v1', auth});
        //console.log(req.body.folderId);
        if(req.body.folderId==""){
            resolve('folder is missing');
        }
        var fileMetadata = {
          'parents': [req.body.folderId],
          'mimeType': 'application/vnd.google-apps.document'
        };
        var media = {
            //mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            // body: fs.createReadStream(req.file.path)
        };
        drive.files.create({
            resource: fileMetadata,
            media: media,
            fields: 'id'
        }, function (err, file) {
            if (err) {
                console.error("error from create file"+err);
                reject('failed');
            } else {
                console.log('File Id:', JSON.stringify(file));
                drive.permissions.create({
                    fileId: file.data.id,
                    resource:{
                        "role": "writer",
                        "type": "anyone"
                    }
                }, function(err,result){
                    if(err){ console.log(err);reject('failed');}
                      else console.log(result)
                });
                console.log(file.data);
                resolve(file.data.id);
            }
        });
    });
}

app.get('/files', function(req, res){
  console.log(req.query.file);
  var file = 'files/'+req.query.file;
  res.download(file); // Set disposition and send it.
});

async function mergeTagwithTagId(req,auth) {
    const drive = google.drive({version: 'v3', auth});
    const sheets = google.sheets({version: 'v4', auth});
    console.log(req);
    let reqData=req.body.data;
    var fileId=await copytheId(drive,req.body.fileID);

    /*first Batch links Start*/
    var batchupdateTable=await batchUpdate(sheets,reqData,fileId);
    /*first Batch links End*/

    /*final exportPDF Start*/
        var exportFileURL=await exportFile(drive,req,fileId);
        console.log(exportFileURL);
        var deleteFileCopy=await deleteFile(drive,fileId);
        //console.log(deleteFileCopy);
        return exportFileURL;
    /*final exportPDF End*/
}

function copytheId(drive,fileID){
    return new Promise((resolve, reject) => {
        drive.files.copy({
            fileId:fileID
        }, (err, res) => {
            if (err){console.log('The COPY API returned an error: ' + err);reject('failed');}
            resolve(res.data.id);
        });
    });
}

function batchUpdate(sheets,requests,fileID){
    return new Promise((resolve, reject) => {
        //console.log(JSON.stringify(requests));
        console.log('batchUpdate');

        sheets.spreadsheets.batchUpdate({
            spreadsheetId: fileID,
            resource: requests,
        },function(err, response) {
            if (err){ console.log('The batchUpdate API returned an error: ' + err); reject('failed');}
            resolve(response);
        });
    });
}

function exportFile(drive,req,fileID){
    // console.log(req.body);
    return new Promise((resolve, reject) => {
        if(req.body.applicationType=='application/vnd.google-apps.spreadsheet'){
            var typecon='.csv';
        }else{
            var typecon='.pdf';
        }
        var fileName=(Math.random().toString(36).substring(2, 15))+typecon;
        const dest = fs.createWriteStream("files/"+fileName);
        drive.files.export({
            fileId: fileID,
            //mimeType: 'application/pdf'
            mimeType: req.body.applicationType
        },{
            responseType: 'stream'
        },
        function(err, response){
            if(err){ console.log(err);reject('failed');}
            response.data.on('error', err => {
                console.log(err);
            }).on('end', ()=>{

            })
            .pipe(dest);
            // resolve('http://18.217.49.28:5000/files?file='+fileName);
            resolve('http://3.14.154.251:5000/files?file='+fileName);
        });
        //console.log(dest);
    });
    //file=requiredoc.pdf
}

function deleteFile(drive,fileID){
    return new Promise((resolve, reject) => {
        drive.files.delete({
            fileId:fileID,
        }, (err, res) => {
            if (err){ console.log('The COPY API returned an error: ' + err); reject('failed');};
            resolve(res);
        });
    });
}

function duplicateStandardTemplate(req,auth) {
  return new Promise((resolve, reject) => {
      const drive = google.drive({version: 'v3', auth});
      const docs = google.docs({version: 'v1', auth});

      var fileMetadata = {
        'parents': [req.body.folderId],
        'mimeType': 'application/vnd.google-apps.document'
      };

      drive.files.copy({
          resource:fileMetadata,
          fileId:req.body.fileID
      }, (err, res) => {
          if (err){ console.log('The COPY API returned an error: ' + err); reject('failed');};
          let datap=req.body.data;
          let requests=datap;
          var copyfileID=res.data.id;
          resolve(copyfileID);
      });
  });
}

function getFileDatawithId(fileID,auth) {
    return new Promise((resolve, reject) => {
        const drive = google.drive({version: 'v3', auth});
        const docs = google.docs({version: 'v1', auth});
        //console.log(fileID);
        docs.documents.get({
            documentId: fileID,
        },(err, resp) => {
            if(err){console.log('api getfiledatawithid return'+err);reject('failed')}
            //console.log(JSON.stringify(resp));
            resolve(JSON.stringify(resp));
        });
    });
}

app.listen(5100);