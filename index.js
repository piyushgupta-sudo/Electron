const { notEqual } = require("assert");
const fs = require("fs");
const path = require("path");
const ipc = require("electron").ipcRenderer;
const XLSX = require('xlsx');
const nodemailer = require("nodemailer");

fileName = document.getElementById("fileName");

fileContents = document.getElementById("fileContents");
fileContents2 = document.getElementById("fileContents2");
btnRead = document.getElementById("btnRead");

let pathName = path.join(__dirname, "Files");

// Upload Function
function upload(event) {
  let path1 = "";
  let Data = [];
  // console.log("The Upload Function is called")
  ipc.send("open-file-dialog-for-file");
  ipc.on("selected-file", function (event, path) {
    // var path1 = path.basename(pathName);
    path1 = path;
    path1 = path1.toString();
    let index = path1.lastIndexOf("\\");
    let fileName = path1.slice(index + 1);
    const workbook = XLSX.readFile(path1);
    const sheet_name_list = workbook.SheetNames;
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    // console.log(data.filter(obj => obj['E mail']));
    Data = data.filter(obj => obj['E mail']);
    console.log(Data);
    var filePath = 'C:/Users/Admin/Downloads/dir/';

    Data.forEach(iterate)
    // Data = undefined;
    // console.log(Data);
    function iterate(item) {
      filePath = filePath + item['CODE1'] + '.pdf';
      var transporter = nodemailer.createTransport({
        // host: "smtp.mailtrap.io",
        // port: 2525,
        service: "hotmail",
        auth: {
          user: "piyush@infocusin.com",
          pass: "Sad21445@2"
        }
      });
      var mailOptions = {
        from: 'piyush@infocusin.com',
        to: item['E mail'],
        subject: 'Trail2',
        html: '<h1>Hello, This is techsolutionstuff !!</h1><p>This is test mail..!</p>',
        attachments: [
          {
            filename: item['CODE1'] + '.pdf',
            path: filePath
          }
        ]
      };
      transporter.sendMail(mailOptions, function (error, info) {
        if (error) {
          console.log(error);
        } else {
          console.log('Email sent: ' + info.response);
        }
      });


      filePath = 'C:/Users/Admin/Downloads/dir/';

      // Data.shift();
      // console.log(Data);



    }

    // Data.shift();
    // console.log(Data);
    location.reload();
  });

  // location.reload();
}

// Download Function
function download(event) {
  ipc.send("open-folder-dialog-for-file");
  fileName = fileName.value;
  ipc.on("selected-folder", function (event, path) {
    console.log(path[0]);
    let file = path[0] + "\\" + fileName;
    console.log(file);
    console.log(fileName);
    let contents = fileContents2.value;
    console.log(contents);
    fs.writeFile(file, contents, function (err) {
      if (err) {
        return console.log(err);
      }
      location.reload();
      console.log("The File was Created");
    });

  });
}
