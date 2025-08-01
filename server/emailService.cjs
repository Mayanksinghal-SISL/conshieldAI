require('dotenv').config(); // Load environment variables

const nodemailer = require('nodemailer');

// Configure your email transporter
// You can use different services (Gmail, Outlook, SendGrid, etc.)
// For simplicity, let's start with a generic SMTP setup using environment variables.
// NOTE: For production, consider a dedicated email service (SendGrid, Mailgun, AWS SES)
//       or using OAuth2 for services like Gmail/Outlook for better security.

const transporter = nodemailer.createTransport({
    host: process.env.EMAIL_HOST,        // e.g., 'smtp.office365.com' for Outlook, 'smtp.gmail.com' for Gmail
    port: process.env.EMAIL_PORT,        // e.g., 587 for TLS, 465 for SSL
    secure: process.env.EMAIL_SECURE === 'true', // Use 'true' if port is 465, 'false' if port is 587
    auth: {
        user: process.env.EMAIL_USER,    // Your email address (e.g., noreply@yourdomain.com)
        pass: process.env.EMAIL_PASS     // Your email password or app-specific password
    },
    tls: {
        // Do not fail on invalid certs - USE ONLY IN DEVELOPMENT
        // For production, ensure valid certificates and remove this line or set to false
        rejectUnauthorized: false
    }
});

/**
 * Sends an email notification.
 * @param {string} to - Recipient email address(es), comma-separated for multiple.
 * @param {string} subject - Subject line of the email.
 * @param {string} text - Plain text body of the email.
 * @param {string} [html] - HTML body of the email (optional, overrides text if provided).
 */
async function sendEmail(to, subject, text, html) {
    const mailOptions = {
        from: process.env.EMAIL_FROM || process.env.EMAIL_USER, // Sender address, default to user
        to: to,
        subject: subject,
        text: text,
        html: html || text // Use HTML if provided, otherwise plain text
    };

    try {
        let info = await transporter.sendMail(mailOptions);
        console.log(`Email sent: ${info.messageId}`);
        console.log(`Preview URL (if available): ${nodemailer.getTestMessageUrl(info)}`);
        return true;
    } catch (error) {
        console.error('Error sending email:', error);
        console.error('Email send details:', { to, subject });
        // Log more details about the error if needed
        if (error.response) {
            console.error('SMTP Response:', error.response);
        }
        return false;
    }
}

module.exports = {
    sendEmail
};

//npm run start:dev