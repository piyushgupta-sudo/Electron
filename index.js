const { notEqual } = require("assert");
const fs = require("fs");
const path = require("path");
const ipc = require("electron").ipcRenderer;
const XLSX = require('xlsx');
const nodemailer = require("nodemailer");
// const archiver = require('archiver');
const AdmZip = require('adm-zip');
const { electron } = require("process");
const { dialog, ipcMain, ipcRenderer } = require('electron');

fileName = document.getElementById("fileName");

fileContents = document.getElementById("fileContents");
fileContents2 = document.getElementById("fileContents2");
btnRead = document.getElementById("btnRead");

let pathName = path.join(__dirname, "Files");
var userId = "";
var password = "";
var mailSubject = "";
var mailBody = "";
var dirPath = "";
var Data = [];
var path1 = "";
// Upload Function
// userId = document.getElementById('email').value;
// // ('email').value;
// console.log("userId", userId);

// password = document.getElementById('password').value;
// console.log("password", password);

// send files
const element = document.getElementById("sendFiles");
element.addEventListener("click", sendFiles);

function upload(event) {



  // console.log("The Upload Function is called")
  ipc.send("open-file-dialog-for-file");
  ipc.on("selected-file", function (event, path) {
    //event.preventDefault();
    // var path1 = path.basename(pathName);
    path1 = path;
    path1 = path1.toString();
    let index = path1.lastIndexOf("\\");
    let fileName = path1.slice(index + 1);
    document.getElementById('excel').value = fileName;




    // Data.shift();
    // console.log(Data);
    // location.reload();
  });

  // location.reload();
}
function sendFiles(event) {

  const loadingElement = document.getElementById("loading");
  const successMessageElement = document.getElementById("successMessage");

  loadingElement.style.display = "block"; // Show the loading animation

  const workbook = XLSX.readFile(path1);
  const sheet_name_list = workbook.SheetNames;
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
  // console.log(data.filter(obj => obj['E mail']));
  Data = data.filter(obj => obj['E mail']);
  console.log(Data);
  var filePath = dirPath;

  const zip = new AdmZip();

  Data.forEach(iterate);

  // Data = undefined;
  // console.log(Data);
  // var date = new Date();
  // date = date.toString().replace(/ /g, "_");
  // console.log(date);
  function iterate(item) {
    for (var i = 1; i <= 2; i++) {
      zip.addLocalFile(filePath + item[`CODE${i}`] + '.pdf');

    }
    const downloadName = `${Date.now()}.zip`;
    const data1 = zip.toBuffer();
    zip.writeZip(filePath + downloadName);

    // const output = fs.createWriteStream(filePath + item['E mail'] + '.zip');
    // const archive = archiver('zip', {
    //   zlib: { level: 9 }
    // });
    // output.on('close', () => {
    //   console.log('Archive finished.');
    // });
    // archive.on('error', (err) => {
    //   throw err;
    // });
    // archive.pipe(output);
    // for (let i = 1; i <= 2; i++) {
    //   try {
    //     const text = path.join(filePath, item[`CODE${i}`] + '.pdf');
    //     archive.append(fs.createReadStream(text), { name: item[`CODE${i}`] + '.pdf' });
    //   } catch (error) {
    //     console.log(error);
    //   }

    // }
    // archive.finalize();
    // // filePath = filePath + item['CODE1'] + '.pdf';
    var transporter = nodemailer.createTransport({
      // host: "smtp.mailtrap.io",
      // port: 2525,
      service: "hotmail",
      auth: {
        user: userId,
        pass: password
      }
    });
    // console.log(item['E mail'] + '_' + date + '.zip')
    var mailOptions = {
      from: userId,
      to: item['E mail'],
      subject: 'Trail4',
      // html: '<h1>Hello, This is techsolutionstuff !!</h1><p>This is test mail..!</p>',
      attachments: [
        {
          filename: downloadName,
          path: filePath + downloadName
        }
      ]
    };
    transporter.sendMail(mailOptions, function (error, info) {
      loadingElement.style.display = "none"; // Hide the loading animation

      if (error) {
        console.log(error);
      } else {
        console.log('Email sent: ' + info.response);

        document.getElementById('formBox').reset();
        // Show the success message for 3 seconds
        successMessageElement.style.display = "block";
        setTimeout(function () {
          successMessageElement.style.display = "none";
        }, 3000);
        // location.reload();
      }
    });


    filePath = dirPath;

    // Data.shift();
    // console.log(Data);



  }
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

//Credential Save
function submitForm(event) {
  userId = document.getElementById('email').value;
  // ('email').value;
  console.log("userId", userId);

  password = document.getElementById('password').value;
  console.log("password", password);

}

//Directory Path
function onDirectorySelected(event) {

  // dialog.showOpenDialog({ properties: ['openDirectory'], defaultPath: 'C:/Users/Admin/Downloads/dir/' }).then(result => {
  //   if (!result.canceled) {
  //     const directoryPath = result.filePaths[0];
  //     console.log(`Selected directory: ${directoryPath}`);
  //   }
  // }).catch(err => {
  //   console.log(err);
  // });



  ipcRenderer.send('open-directory-dialog');

  ipcRenderer.on('selected-directory', (event, path) => {
    // event.preventDefault();
    console.log(`Selected directory: ${path}`);
    dirPath = `${path}`.toString() + '/';

    console.log(dirPath);
    document.getElementById('dirPath').value = dirPath;
  });
  // document.getElementById('dirPath').insertAdjacentText = dirPath;


}

