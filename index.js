const fs = require('fs');
const nodemailer = require('nodemailer');
const path = require('path');
const xlsx = require('xlsx');

// Certificate directory
const certificateDir = './certificates';

// Load Excel data with trimmed column names
function loadExcelData(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert worksheet to JSON and map to trim headers
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Trim the headers in the first row
    const headers = jsonData[0].map(header => header.trim());
    
    // Map rows to an array of objects with trimmed headers as keys
    return jsonData.slice(1).map(row => {
        const record = {};
        headers.forEach((header, index) => {
            record[header] = row[index];
        });
        return record;
    });
}


// Get list of certificate files
function getCertificateFiles(directory) {
    return fs.readdirSync(directory);
}

// Helper function to normalize names (lowercase, replace spaces and underscores)
function normalizeName(name) {
    return name
        .toLowerCase()
        .replace(/[_\s]+/g, '') // Remove spaces and underscores
        .replace(/\./g, ''); // Remove dots if any
}

// Function to extract number before .pdf in a file name
function extractNumberFromFileName(fileName) {
    const match = fileName.match(/(\d+)\.pdf$/);
    return match ? match[1] : null;
}

// Main function for renaming files with names from Excel
async function renameFilesWithNames() {
    const excelData = loadExcelData('nd.xlsx');
    const certificateFiles = fs.readdirSync(certificateDir);

    for (let i = 0; i < excelData.length; i++) {
        const record = excelData[i];
        const name = record['Name'] || record['Name ']; // Adjust for potential whitespace issues

        if (!name) {
            console.error("No Name provided for a record:", record);
            continue;
        }

        const normalizedNewName = normalizeName(name);

        // Ensure we're within bounds of available certificate files
        if (i < certificateFiles.length) {
            const matchingFile = certificateFiles[i];
            const fileNumber = extractNumberFromFileName(matchingFile);

            if (fileNumber) {
                const newFileName = `${normalizedNewName}.pdf`;
                const oldFilePath = path.join(certificateDir, matchingFile);
                const newFilePath = path.join(certificateDir, newFileName);

                fs.renameSync(oldFilePath, newFilePath);
                console.log(`Renamed ${matchingFile} to ${newFileName}`);
            } else {
                console.error(`No identifiable number found in ${matchingFile}. Skipping.`);
            }
        } else {
            console.error(`No matching file available for ${name}.`);
        }
    }
}
// Configure Nodemailer transport
function createTransporter() {
    return nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'your-gmail@gmail.com', // Replace with your email
            pass: '***',    // Replace with your app password
        }
    });
}

// Function to send an email with the certificate attachment
async function sendEmail(transporter, email, attachmentPath, name) {
    const mailOptions = {
        from: 'your-gmail@gmail.com',
        to: email,
        subject: 'Congratulations on Participating ',
        html: `
        <div style="font-family: Arial, sans-serif; color: black; line-height: 1.5;">
            <p style="font-size: 18px;">Dear ${name},</p>

            <p>Congratulations on your participation in Event, organized by the your-organization. Your enthusiasm and commitment to innovation truly reflect the spirit of this national-level hackathon.</p>

            <p>Attached is your participation certificate, acknowledging your effort throughout the event. We appreciate your contribution to our your-organization vibrant community of creators and problem-solvers.</p>

            <p>Thank you for being a part of Event. We hope this experience motivates you to keep exploring new ideas.</p>

            <p>Best regards,<br>your-organization</p>
        </div>
        <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
        <div style="text-align: center; background-color: #d3d3d3; padding: 20px 0;">
            <p style="font-size: 16px; font-weight: bold;">Follow your-organization on:</p>
            <div style="display: flex; align-items: center; justify-content: space-between; max-width: 600px; margin:0 auto; padding: 0 10px;">
                <a href="link" style="margin: 0 15px;">
                    <img src="https://upload.wikimedia.org/wikipedia/commons/d/d1/FacebookIcon.png" alt="Facebook" style="width: 25px;">
                </a>
                <a href="link" style="margin: 0 15px;">
                    <img src="https://upload.wikimedia.org/wikipedia/commons/c/ca/LinkedIn_logo_initials.png" alt="LinkedIn" style="width: 25px;">
                </a>
                <a href="link" style="margin: 0 15px;">
                    <img src="https://img.freepik.com/premium-psd/black-brand-new-twitter-x-logo-icon_1129635-1.jpg?size=626&ext=jpg" alt="X" style="width: 25px;">
                </a>
                <a href="link" target="_blank" style="margin: 0 15px;">
                    <img src="https://upload.wikimedia.org/wikipedia/commons/4/42/YouTube_icon_%282013-2017%29.png" alt="YouTube" style="width: 32px;">
                </a>
                <a href="link" style="margin: 0 15px;">
                    <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Instagram_icon.png" alt="Instagram" style="width: 25px;">
                </a>
                <a href="link" style="margin: 0 15px;">
                    <img src="https://static-00.iconduck.com/assets.00/github-icon-256x244-6n017jrr.png" alt="GitHub" style="width: 25px;">
                </a>
            </div>
        </div>
        <div style="background-color: #f4f4f4; padding: 20px 0; text-align: center; font-size: 12px; color: #555;">
            <p>©2024 your-organization. All rights reserved.</p>
        </div>
        `,
        attachments: [
            {
                filename: path.basename(attachmentPath),
                path: attachmentPath
            }
        ]
    };

    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${email} successfully.`);    
}

// Function to find matching certificate file by ID or Name
function findMatchingFileByIdOrName(certificateFiles, id, name) {
    const normalizedId = id ? id.trim() : null;
    const normalizedName = normalizeName(name);

    let matchingFile = certificateFiles.find(file => {
        const fileNameWithoutExt = path.basename(file, '.pdf');
        return normalizedId && fileNameWithoutExt.includes(normalizedId);
    });

    if (!matchingFile) {
        matchingFile = certificateFiles.find(file => {
            const fileNameWithoutExt = path.basename(file, '.pdf');
            const normalizedFileName = normalizeName(fileNameWithoutExt);
            return normalizedFileName.includes(normalizedName);
        });
    }

    return matchingFile;
}

// Delay function to pause between sending emails
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Process each record in the Excel data for sending emails
async function processCertificates(excelData, certificateFiles, transporter) {
    for (let nd of excelData) {
        const id = nd.ID ? String(nd.ID).trim() : '';
        const name = nd['Name'] ? String(nd['Name']).trim() : '';
        const email = nd['Mail'] ? String(nd['Mail']).trim() : ''; // If "Mail" is the correct column name        

        if (!name && !id) {
            console.error("No ID or Name provided for a record:", nd);
            continue;
        }

        if (!email) {
            console.warn(`Email is empty for ${name}. Skipping this entry.`);
            continue;
        }

        try {
            console.log(`Processing certificate for ${name}...`);
            const matchingFile = findMatchingFileByIdOrName(certificateFiles, id, name);

            if (!matchingFile) {
                console.error(`Certificate file not found for ${name}.`);
                continue;
            }

            const filePath = path.join(certificateDir, matchingFile);
            await sendEmail(transporter, email, filePath, name);
            await delay(1000);
            console.log(`Email sent to ${email}`);
        } catch (error) {
            console.error(`Error processing certificate for ${name}:`, error);
        }
    }
}

// Main function for creating sequential IDs
function createSequentialIds() {
    const filePath = 'nd.xlsx';
    const excelData = loadExcelData(filePath);

    let idCounter = 1;
    excelData.forEach(nd => {
        nd.ID = idCounter++;
    });

    const workbook = xlsx.readFile(filePath); // Load the existing workbook
    const worksheet = xlsx.utils.json_to_sheet(excelData); // Create a new sheet with updated data
    workbook.Sheets[workbook.SheetNames[0]] = worksheet; // Replace the first sheet

    xlsx.writeFile(workbook, filePath); // Write back to the same file
    console.log('Sequential IDs created and saved to nd.xlsx');
}


// Main function for sending emails with certificates
async function sendEmails() {
    const excelData = loadExcelData('nd.xlsx');
    const certificateFiles = getCertificateFiles(certificateDir);
    const transporter = createTransporter();

    await processCertificates(excelData, certificateFiles, transporter);
}

// Uncomment the function you want to execute

//renameFilesWithNames();
//createSequentialIds();
//sendEmails();
