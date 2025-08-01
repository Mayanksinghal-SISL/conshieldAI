# SharePoint File Viewer

This application automatically fetches and displays an Excel file from SharePoint using Microsoft Graph API with server-side authentication.

## Prerequisites

1. Node.js (v14 or later)
2. npm (v7 or later)
3. Azure AD App Registration with the following API permissions:
   - `Sites.Read.All` (Application permission)
   - `Files.Read.All` (Application permission)

## Setup Instructions

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd <repository-folder>
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Configure environment variables**
   - Copy `.env.example` to `.env`
   - Update the values in `.env` with your Azure AD and SharePoint details

4. **Build the application**
   ```bash
   npm run build:all
   ```

## Configuration

Update the following environment variables in the `.env` file:

```env
# Server Configuration
PORT=3000
NODE_ENV=development

# Azure AD App Registration
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_ID=your_client_id_here
AZURE_CLIENT_SECRET=your_client_secret_here

# SharePoint Details
SHAREPOINT_SITE_ID=your_sharepoint_site_id
SHAREPOINT_DRIVE_ID=your_sharepoint_drive_id
SHAREPOINT_FILE_ID=your_sharepoint_file_id
SHAREPOINT_FILE_NAME=your_file_name.xlsx

# Frontend Configuration
VITE_API_BASE_URL=http://localhost:3000
```

## Running Locally

1. Start the development server:
   ```bash
   npm run start:dev
   ```

2. Open your browser and navigate to:
   ```
   http://localhost:3000
   ```

## Production Deployment

### Option 1: Using PM2 (Recommended for Linux servers)

1. Install PM2 globally:
   ```bash
   npm install -g pm2
   ```

2. Start the application in production mode:
   ```bash
   NODE_ENV=production pm2 start server/index.js --name "sharepoint-viewer"
   ```

3. Set up PM2 to start on boot:
   ```bash
   pm2 startup
   pm2 save
   ```

### Option 2: Using Systemd (For Linux servers with systemd)

1. Create a systemd service file at `/etc/systemd/system/sharepoint-viewer.service`:
   ```ini
   [Unit]
   Description=SharePoint File Viewer
   After=network.target

   [Service]
   User=your_username
   WorkingDirectory=/path/to/your/app
   Environment="NODE_ENV=production"
   ExecStart=/usr/bin/node server/index.js
   Restart=always
   RestartSec=10
   StandardOutput=syslog
   StandardError=syslog
   SyslogIdentifier=sharepoint-viewer

   [Install]
   WantedBy=multi-user.target
   ```

2. Reload systemd and start the service:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable sharepoint-viewer
   sudo systemctl start sharepoint-viewer
   ```

## FileZilla Deployment

1. Connect to your server using FileZilla
2. Upload the following files and directories to `/var/www/qa`:
   - `dist/` (generated after build)
   - `server/`
   - `package.json`
   - `package-lock.json`
   - `.env`

3. On the server, install dependencies and start the application:
   ```bash
   cd /var/www/qa
   npm install --production
   npm run build:all
   ```bash
NODE_ENV=production node server/sharepoint-api.cjs
   ```

## Setting up a Reverse Proxy (Nginx)

1. Install Nginx:
   ```bash
   sudo apt update
   sudo apt install nginx
   ```

2. Create a new Nginx configuration file at `/etc/nginx/sites-available/sharepoint-viewer`:
   ```nginx
   server {
       listen 80;
       server_name qa.anplabs.in;

       location / {
           proxy_pass http://localhost:3000;
           proxy_http_version 1.1;
           proxy_set_header Upgrade $http_upgrade;
           proxy_set_header Connection 'upgrade';
           proxy_set_header Host $host;
           proxy_cache_bypass $http_upgrade;
           proxy_set_header X-Real-IP $remote_addr;
           proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
       }
   }
   ```

3. Enable the site and restart Nginx:
   ```bash
   sudo ln -s /etc/nginx/sites-available/sharepoint-viewer /etc/nginx/sites-enabled/
   sudo nginx -t
   sudo systemctl restart nginx
   ```

## Troubleshooting

- **Application not starting**: Check the logs using `pm2 logs` or `journalctl -u sharepoint-viewer`
- **Authentication errors**: Verify your Azure AD App Registration has the correct permissions and the client secret is valid
- **File not found**: Check the SharePoint file ID and ensure the service account has access

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
