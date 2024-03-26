const Excel = require('exceljs');
const nodemailer = require('nodemailer');

// Read Excel Sheet
const workbook = new Excel.Workbook();
workbook.xlsx.readFile('participants.xlsx')
    .then(function() {
        const worksheet = workbook.getWorksheet(1);
        
        // Extract Email Addresses
        let emails = [];
        worksheet.eachRow(function(row, rowNumber) {
            // Assuming email addresses are in the first column
            emails.push(row.getCell(1).value);
        });

        // Compose Email
        const subject = 'Your Subject Here';
        const text = 'Your email content here';

        // Configure SMTP transporter
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                user: 'your-email@gmail.com',
                pass: 'your-password'
            }
        });

        // Send Email to each participant
        emails.forEach(email => {
            const mailOptions = {
                from: 'your-email@gmail.com',
                to: email,
                subject: subject,
                text: text
            };

            transporter.sendMail(mailOptions, function(error, info) {
                if (error) {
                    console.log('Error sending email:', error);
                } else {
                    console.log('Email sent:', info.response);
                }
            });
        });
    })
    .catch(function(error) {
        console.log('Error reading Excel file:', error);
    });