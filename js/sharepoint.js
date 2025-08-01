import { getAccessToken } from './auth.js';

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';

// Function to get file content from SharePoint
export async function getFileContent(fileId) {
    try {
        const accessToken = await getAccessToken();
        const response = await fetch(
            `${GRAPH_ENDPOINT}/sites/${import.meta.env.VITE_SHAREPOINT_SITE_ID}` +
            `/drives/${import.meta.env.VITE_SHAREPOINT_DRIVE_ID}` +
            `/items/${fileId}/content`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/octet-stream'
                }
            }
        );

        if (!response.ok) {
            throw new Error(`Failed to get file content: ${response.statusText}`);
        }

        return await response.arrayBuffer();
    } catch (error) {
        console.error('Error getting file content:', error);
        throw error;
    }
}

// Function to search for a file in SharePoint
export async function findFile(fileName) {
    try {
        const accessToken = await getAccessToken();
        const response = await fetch(
            `${GRAPH_ENDPOINT}/sites/${import.meta.env.VITE_SHAREPOINT_SITE_ID}` +
            `/drives/${import.meta.env.VITE_SHAREPOINT_DRIVE_ID}` +
            `/root/search(q='${encodeURIComponent(fileName)}')`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        if (!response.ok) {
            throw new Error(`Failed to search for file: ${response.statusText}`);
        }

        const data = await response.json();
        return data.value[0]; // Return first match
    } catch (error) {
        console.error('Error finding file:', error);
        throw error;
    }
}

// Function to get file metadata
export async function getFileMetadata(fileId) {
    try {
        const accessToken = await getAccessToken();
        const response = await fetch(
            `${GRAPH_ENDPOINT}/sites/${import.meta.env.VITE_SHAREPOINT_SITE_ID}` +
            `/drives/${import.meta.env.VITE_SHAREPOINT_DRIVE_ID}` +
            `/items/${fileId}`,
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        if (!response.ok) {
            throw new Error(`Failed to get file metadata: ${response.statusText}`);
        }

        return await response.json();
    } catch (error) {
        console.error('Error getting file metadata:', error);
        throw error;
    }
}
