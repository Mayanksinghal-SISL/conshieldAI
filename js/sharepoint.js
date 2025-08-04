// This file is in your frontend directory, e.g., src/sharepoint.js

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || window.location.origin;
const EXCEL_API_ENDPOINT = `${API_BASE_URL}/api/file/excel`;

/**
 * Fetches the processed Excel data (as a JSON object) from the backend.
 * The backend handles all SharePoint API calls securely.
 * @returns {Promise<Object>} An object containing the processed Excel data.
 */
export async function getProcessedDataFromBackend() {
    try {
        const response = await fetch(EXCEL_API_ENDPOINT);

        if (!response.ok) {
            let errorMessage = `Backend API error! Status: ${response.status}`;
            const errorBody = await response.json();
            errorMessage += ` - Details: ${errorBody.error?.details || JSON.stringify(errorBody)}`;
            throw new Error(errorMessage);
        }

        const processedData = await response.json();
        return processedData;

    } catch (error) {
        console.error('An error occurred during backend fetch:', error);
        throw error;
    }
}