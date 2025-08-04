// server/sharepoint-api.cjs

require('dotenv').config({ path: '../.env' });
// TEMPORARY DEBUGGING LOGS - You can remove these later
console.log("-----------------------------------------");
console.log("Debugging Environment Variables:");
console.log("TENANT_ID:", process.env.TENANT_ID ? "Found" : "Not Found");
console.log("CLIENT_ID:", process.env.CLIENT_ID ? "Found" : "Not Found");
console.log("CLIENT_SECRET:", process.env.CLIENT_SECRET ? "Found" : "Not Found");
console.log("-----------------------------------------");
const express = require('express');
const cors = require('cors');
const path = require('path');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
const fetch = require('node-fetch'); // Ensure node-fetch is installed if you use it directly

// Import the new modules
const { processExcelDataServer } = require('./excelProcessor.cjs');
const cron = require('node-cron');
const { sendEmail } = require('./emailService.cjs'); 

const app = express();
const port = process.env.PORT || 3001;

// Ensure the port is a number
const normalizedPort = parseInt(port, 10);

// Middleware to parse JSON bodies
const corsOptions = {
  origin: 'http://localhost:3032', // Explicitly allow your frontend's origin
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'], // Allow common HTTP methods and OPTIONS for preflight
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'], // CRITICAL: Now includes Authorization
  credentials: true // Allow cookies/auth headers to be sent
};
app.use(cors(corsOptions)); 


// --- Microsoft Graph Authentication Helper Functions ---
class MyAuthProvider {
  constructor(credential) {
    this.credential = credential;
  }

  async getAccessToken() {
    try {
      const token = await this.credential.getToken('https://graph.microsoft.com/.default');
      return token.token;
    } catch (error) {
      console.error('Error getting token for Graph API:', error);
      throw error;
    }
  }
}

async function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.TENANT_ID,
    process.env.CLIENT_ID,
    process.env.CLIENT_SECRET
  );

  const authProvider = new MyAuthProvider(credential);
  const client = Client.initWithMiddleware({
    authProvider
  });

  return client;
}


// --- CORE EXCEL PROCESSING AND EMAIL SENDING LOGIC (wrapped in a function) ---
async function runExcelProcessingAndEmailTask() {
    console.log('--- Scheduled Task: Initiating Excel processing and email check ---');
    try {
        const client = await getGraphClient();
        const driveId = process.env.SHAREPOINT_DRIVE_ID;
        const fileId = process.env.SHAREPOINT_FILE_ID;   
        const fileName = process.env.SHAREPOINT_FILE_NAME; 

        if (!driveId || !fileId || !fileName) {
            console.error("Scheduled Task Error: Missing SharePoint environment variables (SHAREPOINT_DRIVE_ID, SHAREPOINT_FILE_ID, SHAREPOINT_FILE_NAME). Please check your .env file.");
            return; // Exit if critical env vars are missing
        }
        
        console.log(`Scheduled Task: Attempting to fetch file: ${fileName} from drive ID: ${driveId}`);
        
        // Get the file content stream
        const fileStream = await client
          .api(`/drives/${driveId}/items/${fileId}/content`) 
          .getStream();

        // Convert stream to buffer
        const chunks = [];
        for await (const chunk of fileStream) {
          chunks.push(chunk);
        }
        const buffer = Buffer.concat(chunks);

        console.log('Scheduled Task: Successfully fetched Excel file as buffer. Processing data...');

        const processedData = processExcelDataServer(buffer);
        console.log('Scheduled Task: Excel data processed. Checking for email triggers...');

        const recipients = process.env.EMAIL_RECIPIENTS ? process.env.EMAIL_RECIPIENTS.split(',') : []; // No default if env not set
        if (recipients.length === 0) {
            console.warn('Scheduled Task Warning: No EMAIL_RECIPIENTS defined in .env. Skipping email sending.');
            return; // Exit if no recipients
        }
        
        const triggeredComments = [];

        processedData.clientDataForDisplay.forEach(clientData => {
            const comments = clientData.comments;
            const clientName = clientData.client;

            if (comments.includes("Monthly difference more than 5%")) {
                triggeredComments.push(`${clientName}: Monthly difference > 5%`);
            }
            if (comments.includes("Fortnightly difference more than 5%")) {
                triggeredComments.push(`${clientName}: Fortnightly difference > 5%`);
            }
            if (comments.includes("Weekly difference more than 5%")) {
                triggeredComments.push(`${clientName}: Weekly difference > 5%`);
            }
            // Add other conditions for comments like "Yesterday Data Blank" if needed
            if (comments.includes("Yesterday Data Blank") && !comments.includes("difference more than 5%")) {
                triggeredComments.push(`${clientName}: Yesterday Data Blank (No other diff)`);
            }
        });

        if (triggeredComments.length > 0) {
            const subject = `ACTION REQUIRED: SharePoint Consumption Alerts - ${new Date().toLocaleDateString('en-GB')}`;
            const textBody = `The following alerts were detected:\n\n${triggeredComments.join('\n')}\n\nPlease check the SharePoint dashboard for details.`;
            const htmlBody = `
                <p>The following alerts were detected:</p>
                <ul>
                    ${triggeredComments.map(comment => `<li>${comment}</li>`).join('')}
                </ul>
                <p>Please check the SharePoint dashboard for details.</p>
                <p>Access the dashboard here: <a href="http://localhost:3032">http://localhost:3032</a></p>
                <br>
                <p>This is an automated notification. Please do not reply.</p>
            `;

            console.log(`Scheduled Task: Sending email to ${recipients.join(', ')} for ${triggeredComments.length} alerts.`);
            await sendEmail(recipients.join(','), subject, textBody, htmlBody);
            console.log('Scheduled Task: Email sending process initiated.');
        } else {
            console.log('Scheduled Task: No actionable comments found. No email sent.');
        }

    } catch (error) {
        console.error('Scheduled Task Error (Detailed):', JSON.stringify(error, Object.getOwnPropertyNames(error), 2));
        console.error('Scheduled Task Error:', error.message);
    }
    console.log('--- Scheduled Task: Completed ---');
}

// --- SCHEDULE THE EMAIL TASK ---
// Cron string for 6:00 AM (0 minutes past the 6th hour), every day
// '0 6 * * *'
// Ensure the timezone is set to Asia/Kolkata for IST
cron.schedule('0 16 * * *', () => {
    runExcelProcessingAndEmailTask();
}, {
    timezone: "Asia/Kolkata" 
});

// Optional: Run the task immediately when the server starts for testing
// runExcelProcessingAndEmailTask();


// --- API ENDPOINT FOR FRONTEND (DOES NOT TRIGGER EMAILS UNLESS ALERTS ARE FOUND) ---
// This endpoint only sends the processed data back to the frontend
app.get('/api/file/excel', async (req, res) => {
  console.log('API endpoint hit: /api/file/excel (from frontend request)');
  try {
    const client = await getGraphClient();
    
    const driveId = process.env.SHAREPOINT_DRIVE_ID;
    const fileId = process.env.SHAREPOINT_FILE_ID;   
    const fileName = process.env.SHAREPOINT_FILE_NAME; 

    if (!driveId || !fileId || !fileName) {
        throw new Error("Missing SharePoint environment variables for API request.");
    }
    
    const fileStream = await client
      .api(`/drives/${driveId}/items/${fileId}/content`) 
      .getStream();

    const chunks = [];
    for await (const chunk of fileStream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);

    const processedData = processExcelDataServer(buffer);
    console.log('API endpoint: Excel data processed for frontend. Sending response.');

    // The email sending logic is removed from here for frontend requests
    // and is now only in the scheduled task (runExcelProcessingAndEmailTask).
    // The frontend only needs the processed data.

    res.json(processedData);

  } catch (error) {
    console.error('Raw error object from Graph API call or processing:', JSON.stringify(error, Object.getOwnPropertyNames(error), 2));
    console.error('Error in /api/file/excel (frontend request):', error.message);
    
    let detailedError = "Unknown error during Graph API call or processing.";
    if (error.response && error.response.statusMessage) {
        detailedError = `Graph API Response: ${error.response.status} - ${error.response.statusMessage} - ${error.response.body ? JSON.stringify(error.response.body) : ''}`;
    } else if (error.statusCode) {
        detailedError = `Graph API Status Code: ${error.statusCode} - ${error.message || 'No specific message'}`;
    } else if (error instanceof Error) {
        detailedError = error.message;
    }

    res.status(500).json({
      error: 'Failed to fetch or process Excel file',
      details: detailedError,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// Fallback route for SPA with API route guard - Keep this as is
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api/')) {
    console.error(`API route not found: ${req.path}`);
    return res.status(404).json({ 
      error: 'API route not found',
      path: req.path 
    });
  }
  console.log(`Serving SPA for path: ${req.path}`);
  res.sendFile(path.join(__dirname, '../dist/index.html'));
});

// Start the server
const server = app.listen(normalizedPort, () => {
  console.log(`Server running at http://localhost:${normalizedPort}`);
});

// Handle server errors
server.on('error', (error) => {
  if (error.syscall !== 'listen') {
    throw error;
  }

  switch (error.code) {
    case 'EACCES':
      console.error(`Port ${normalizedPort} requires elevated privileges`);
      process.exit(1);
      break;
    case 'EADDRINUSE':
      console.error(`Port ${normalizedPort} is already in use`);
      process.exit(1);
      break;
    default:
      throw error;
  }
});