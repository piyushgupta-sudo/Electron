const { app, BrowserWindow, ipcMain } = require('electron');
const fs = require('fs-extra');
const path = require('path');
const archiver = require('archiver');
const nodemailer = require('nodemailer');

let mainWindow;

app.on('ready', () => {
    mainWindow = new BrowserWindow({ width: 800, height: 600 });
    mainWindow.loadFile('index.html');
});

ipcMain.on('sendFiles', async (event, data) => {
    try {
        const { directoryPath, emailAddresses } = data;

        // Step 1: Get the list of files in the static directory
        const filePaths = await getFilesInDirectory(directoryPath);

        // Step 2: Compress the files into a zip
        const zipPath = await zipFiles(filePaths);

        // Step 3: Send the zip file via email
        await sendEmail(zipPath, emailAddresses);

        // Notify the renderer process that the operation is complete
        event.sender.send('sendFilesResponse', 'Files sent successfully.');
    } catch (error) {
        event.sender.send('sendFilesResponse', `Error sending files: ${error.message}`);
    }
});

async function getFilesInDirectory(directoryPath) {
    return new Promise((resolve, reject) => {
        fs.readdir(directoryPath, (err, files) => {
            if (err) {
                reject(err);
            } else {
                // Filter out directories (if any) and return the file paths
                const filePaths = files
                    .filter((file) => fs.statSync(path.join(directoryPath, file)).isFile())
                    .map((file) => path.join(directoryPath, file));

                resolve(filePaths);
            }
        });
    });
}

async function zipFiles(filePaths) {
    return new Promise((resolve, reject) => {
        const output = fs.createWriteStream(path.join(app.getPath('temp'), 'files.zip'));
        const archive = archiver('zip');

        output.on('close', () => resolve(output.path));
        archive.on('error', (err) => reject(err));

        archive.pipe(output);

        for (const filePath of filePaths) {
            archive.file(filePath, { name: path.basename(filePath) });
        }

        archive.finalize();
    });
}

async function sendEmail(zipPath, emailAddresses) {
    const transporter = nodemailer.createTransport({
        service: 'Gmail',
        auth: {
            user: 'your_email@gmail.com', // Replace with your Gmail email address
            pass: 'your_gmail_password', // Replace with your Gmail password
        },
    });

    const mailOptions = {
        from: 'your_email@gmail.com', // Replace with your Gmail email address
        to: emailAddresses.join(','),
        subject: 'Files from Electron App',
        text: 'Please find the attached files.',
        attachments: [
            {
                filename: 'files.zip',
                path: zipPath,
            },
        ],
    };

    return new Promise((resolve, reject) => {
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                reject(error);
            } else {
                resolve(info);
            }
        });
    });
}
