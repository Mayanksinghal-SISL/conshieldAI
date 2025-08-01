// Remove any 'import XLSX' or 'require("xlsx")' from this file.
// The Excel processing is now done on the backend.

/**
 * Renders the processed Excel data (received from the backend) to build the UI.
 * @param {Object} processedData An object containing clientDataForDisplay, combinedDateHeaders, and visibility flags.
 */
function renderExcelData(processedData) { // Renamed from processExcelData
    const { clientDataForDisplay, combinedDateHeaders, hasNonZeroMonthlyCount, hasNonZeroFortnightlyCount, hasNonZeroWeeklyCount, error } = processedData;

    document.getElementById('loadingPlaceholder')?.remove(); // Remove loading placeholder immediately here

    if (error) {
        console.error("Backend reported an error during Excel processing:", error);
        alert(`Error from server: ${error}`);
        return;
    }

    if (!clientDataForDisplay || clientDataForDisplay.length === 0) {
        console.warn("No client data received from backend to display.");
        alert("No relevant client data received from the server. Please check backend logs or Excel file sheet names.");
        return;
    }

    // Helper function for safe numeric conversion - Keep this here if you use it for rendering
    const getNumericValueForDisplay = (text) => {
        const raw = String(text).trim().replace(/[$,%]/g, '');
        return (!isNaN(raw) && raw !== '' && raw.toUpperCase() !== 'NA') ? Number(raw) : null;
    };

    // --- BUILD THE HTML TABLE FOR DESKTOP DISPLAY ---
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    // Create 'Account Name' header with toggle functionality
    const thClient = document.createElement('th');
    thClient.classList.add('account-name');

    const headerContent = document.createElement('div');
    headerContent.style.display = 'flex';
    headerContent.style.justifyContent = 'space-between';
    headerContent.style.alignItems = 'center';
    headerContent.style.width = '100%';

    const spanText = document.createElement('span');
    spanText.textContent = 'Account Name';

    const toggleIcon = document.createElement('span');
    toggleIcon.textContent = '+';
    toggleIcon.classList.add('toggle-icon');

    toggleIcon.style.position = 'static';
    toggleIcon.style.transform = 'none';
    toggleIcon.style.margin = '0';
    toggleIcon.style.padding = '0';
    toggleIcon.style.minWidth = '1em';
    toggleIcon.style.cursor = 'pointer';

    toggleIcon.addEventListener('click', (e) => {
        e.stopPropagation();
        const expanding = toggleIcon.textContent === '+';
        toggleIcon.textContent = expanding ? '−' : '+';

        document.querySelectorAll('.date-col, .accordion-col-summary').forEach(cell => {
            cell.classList.toggle('expanded', expanding);
        });
    });

    headerContent.appendChild(spanText);
    headerContent.appendChild(toggleIcon);
    thClient.appendChild(headerContent);
    headerRow.appendChild(thClient);

    // Add combined date headers
    combinedDateHeaders.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        th.classList.add('date-col');
        headerRow.appendChild(th);
    });

    // --- Conditionally Add Summary Headers based on flags from backend ---
    if (hasNonZeroMonthlyCount) {
        const th = document.createElement('th');
        th.textContent = 'Total Monthly Count';
        th.classList.add('accordion-col-summary');
        headerRow.appendChild(th);
    }
    if (hasNonZeroFortnightlyCount) {
        const th = document.createElement('th');
        th.textContent = 'Total Fortnightly Count';
        th.classList.add('accordion-col-summary');
        headerRow.appendChild(th);
    }
    if (hasNonZeroWeeklyCount) {
        const th = document.createElement('th');
        th.textContent = 'Total Weekly Count';
        th.classList.add('accordion-col-summary');
        headerRow.appendChild(th);
    }

    // Add always-visible summary headers
    const alwaysVisibleSummaryHeaders = [
        'Difference %',
        'Rolling Monthly Avg',
        'Rolling Fortnightly Avg',
        'Rolling Weekly Avg',
        "Yesterday's Data",
        'Comments'
    ];
    alwaysVisibleSummaryHeaders.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        if (header === 'Comments') {
            th.classList.add('fixed-comment');
        }
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    const mobileAccordionContainer = document.getElementById('mobileAccordion');
    mobileAccordionContainer.innerHTML = ''; // Clear previous content
    mobileAccordionContainer.classList.remove('visible'); // Hide initially

    // Populate table body with client data and prepare mobile accordion data
    clientDataForDisplay.forEach(clientData => {
        // Only append row if it has valid daily data for a client
        if (!clientData.dailyValues.some(val => val !== '' && val !== null && String(val).toUpperCase() !== 'NA')) {
            return; // Skip this client if no valid daily data
        }

        const tr = document.createElement('tr');
        tr.classList.add('highlight-row');

        const tdClient = document.createElement('td');
        tdClient.textContent = clientData.client;
        tr.appendChild(tdClient);

        // Add daily values
        clientData.dailyValues.forEach(val => {
            const td = document.createElement('td');
            td.classList.add('date-col');
            const numericVal = getNumericValueForDisplay(val);
            if (numericVal !== null) {
                td.textContent = Math.round(numericVal);
            } else if (typeof val === 'string' && val.toUpperCase() !== 'NA' && val !== '') {
                td.textContent = val;
            } else {
                td.textContent = '';
            }
            tr.appendChild(td);
        });

        // --- Conditionally Add Summary Cells for Desktop Table ---
        if (hasNonZeroMonthlyCount) {
            const tdTotal = document.createElement('td');
            tdTotal.textContent = clientData.monthlyCount;
            tdTotal.classList.add('accordion-col-summary');
            tr.appendChild(tdTotal);
        }
        if (hasNonZeroFortnightlyCount) {
            const tdFortnightCount = document.createElement('td');
            tdFortnightCount.textContent = clientData.fortnightlyCount;
            tdFortnightCount.classList.add('accordion-col-summary');
            tr.appendChild(tdFortnightCount);
        }
        if (hasNonZeroWeeklyCount) {
            const tdWeeklyCount = document.createElement('td');
            tdWeeklyCount.textContent = clientData.weeklyCount;
            tdWeeklyCount.classList.add('accordion-col-summary');
            tr.appendChild(tdWeeklyCount);
        }

        // Add Difference %, Rolling Avgs, Yesterday's Data, Comments
        const tdDiff = document.createElement('td');
        tdDiff.textContent = clientData.diffPct;
        tr.appendChild(tdDiff);

        const tdMonthly = document.createElement('td');
        const tooltipWrapperMonthly = document.createElement('div');
        tooltipWrapperMonthly.classList.add('tooltip-wrapper');
        const spanValueMonthly = document.createElement('span');
        spanValueMonthly.textContent = clientData.monthlyAvg;
        const tooltipTextMonthly = document.createElement('div');
        tooltipTextMonthly.classList.add('tooltip-text');
        tooltipTextMonthly.textContent = clientData.monthlyAvg === 'N/A' ? 'No valid numeric data in last 30 days' : `Sum of ${clientData.monthlyCount} days ÷ ${clientData.monthlyCount}`;
        tooltipWrapperMonthly.appendChild(spanValueMonthly);
        tooltipWrapperMonthly.appendChild(tooltipTextMonthly);
        tdMonthly.appendChild(tooltipWrapperMonthly);
        tr.appendChild(tdMonthly);

        const tdFortnightly = document.createElement('td');
        const fortnightWrapper = document.createElement('div');
        fortnightWrapper.classList.add('tooltip-wrapper');
        const spanFortnight = document.createElement('span');
        spanFortnight.textContent = clientData.fortnightlyAvg;
        const tooltipFortnight = document.createElement('div');
        tooltipFortnight.classList.add('tooltip-text');
        tooltipFortnight.textContent = clientData.fortnightlyAvg === 'N/A' ? 'No valid data in last 15 days' : `Sum of ${clientData.fortnightlyCount} days ÷ ${clientData.fortnightlyCount}`;
        fortnightWrapper.appendChild(spanFortnight);
        fortnightWrapper.appendChild(tooltipFortnight);
        tdFortnightly.appendChild(fortnightWrapper);
        tr.appendChild(tdFortnightly);

        const tdWeeklyAvg = document.createElement('td');
        const weeklyWrapper = document.createElement('div');
        weeklyWrapper.classList.add('tooltip-wrapper');
        const spanWeekly = document.createElement('span');
        spanWeekly.textContent = clientData.weeklyAvg;
        const tooltipWeekly = document.createElement('div');
        tooltipWeekly.classList.add('tooltip-text');
        tooltipWeekly.textContent = clientData.weeklyAvg === 'N/A' ? 'No valid data in last 7 days' : `Sum of ${clientData.weeklyCount} days ÷ ${clientData.weeklyCount}`;
        weeklyWrapper.appendChild(spanWeekly);
        weeklyWrapper.appendChild(tooltipWeekly);
        tdWeeklyAvg.appendChild(weeklyWrapper);
        tr.appendChild(tdWeeklyAvg);

        const tdYesterday = document.createElement('td');
        tdYesterday.classList.add('yesterday-data');
        tdYesterday.textContent = getNumericValueForDisplay(clientData.yesterdayData) !== null ? Math.round(getNumericValueForDisplay(clientData.yesterdayData)) : (clientData.yesterdayData === '' ? '' : clientData.yesterdayData);
        tr.appendChild(tdYesterday);

        const tdComments = document.createElement('td');
        tdComments.textContent = clientData.comments;
        tdComments.classList.add(clientData.commentClass);
        tr.appendChild(tdComments);

        tbody.appendChild(tr);

        // Create mobile accordion item
        const accordionItem = document.createElement('div');
        accordionItem.classList.add('accordion-item');
        accordionItem.classList.add(`${clientData.commentClass}-bg`);

        const accordionHeader = document.createElement('div');
        accordionHeader.classList.add('accordion-header');
        const headerTextSpan = document.createElement('span');
        headerTextSpan.textContent = clientData.client;
        accordionHeader.appendChild(headerTextSpan);

        const accordionToggleIcon = document.createElement('span');
        accordionToggleIcon.textContent = '+';
        accordionToggleIcon.classList.add('accordion-toggle-icon');
        accordionHeader.appendChild(accordionToggleIcon);

        accordionItem.appendChild(accordionHeader);

        const accordionContent = document.createElement('div');
        accordionContent.classList.add('accordion-content');
        accordionContent.innerHTML = `
            ${hasNonZeroMonthlyCount ? `<p><strong>Total Monthly Count:</strong> ${clientData.monthlyCount}</p>` : ''}
            ${hasNonZeroFortnightlyCount ? `<p><strong>Total Fortnightly Count:</strong> ${clientData.fortnightlyCount}</p>` : ''}
            ${hasNonZeroWeeklyCount ? `<p><strong>Total Weekly Count:</strong> ${clientData.weeklyCount}</p>` : ''}
            <p><strong>Difference %:</strong> ${clientData.diffPct}</p>
            <p><strong>Rolling Monthly Avg:</strong> ${clientData.monthlyAvg}</p>
            <p><strong>Rolling Fortnightly Avg:</strong> ${clientData.fortnightlyAvg}</p>
            <p><strong>Rolling Weekly Avg:</strong> ${clientData.weeklyAvg}</p>
            <p><strong>Yesterday's Data:</strong> ${tdYesterday.textContent}</p>
            <p><strong>Comments:</strong> <span class="${clientData.commentClass}">${clientData.comments}</span></p>
        `;
        accordionItem.appendChild(accordionContent);
        mobileAccordionContainer.appendChild(accordionItem);

        accordionHeader.addEventListener('click', () => {
            document.querySelectorAll('.accordion-item.active').forEach(item => {
                if (item !== accordionItem) {
                    item.classList.remove('active');
                    item.querySelector('.accordion-content').style.display = 'none';
                    item.querySelector('.accordion-toggle-icon').textContent = '+';
                }
            });

            accordionItem.classList.toggle('active');
            const isExpanded = accordionContent.style.display === 'block';
            accordionContent.style.display = isExpanded ? 'none' : 'block';
            accordionToggleIcon.textContent = isExpanded ? '+' : '−';
        });
    });

    table.appendChild(tbody);

    const container = document.getElementById('tableContainer');
    container.innerHTML = '';
    container.appendChild(table);

    if (mobileAccordionContainer.childElementCount > 0) {
        mobileAccordionContainer.classList.add('visible');
    }
    document.getElementById('tableWrapper').classList.add('visible');

    // This line is already removed by the top of this function now
    // document.getElementById('loadingPlaceholder')?.remove();
}


// --- Main execution starts when the page loads ---
document.addEventListener('DOMContentLoaded', async function () {
    console.log('Page loaded, initiating SharePoint authentication...');

    const loadingPlaceholder = document.createElement('div');
    loadingPlaceholder.id = 'loadingPlaceholder';
    loadingPlaceholder.textContent = 'Loading...';
    loadingPlaceholder.style.color = '#d3d3d3';
    loadingPlaceholder.style.fontStyle = 'italic';
    loadingPlaceholder.style.margin = '10px 0';

    const heading = document.querySelector('h2');
    if (heading) {
        heading.parentElement.insertBefore(loadingPlaceholder, heading.nextSibling);
    } else {
        document.body.insertBefore(loadingPlaceholder, document.body.firstChild);
    }

// 1. MSAL Configuration
// 1. MSAL Configuration
    const msalConfig = {
        auth: {
            clientId: 'aa32d781-5b43-46ef-9f1e-e5761d749a20',
            // --- CHANGE THIS LINE ---
            authority: 'https://login.microsoftonline.com/2cf8b742-8a28-4ecc-9256-f1ccf9a6381b',
            // -----------------------
            redirectUri: window.location.origin
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: false
        }
    };

    const myMSALObj = new msal.PublicClientApplication(msalConfig);
    let account = null;

    try {
        // 2. Handle Login & Get Account
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0) {
            account = accounts[0];
            console.log('User already signed in:', account.username);
        } else {
            console.log('No active account found, prompting for login...');
            loadingPlaceholder.textContent = 'Authenticating with Microsoft... (Popup may appear)';
            const loginResponse = await myMSALObj.loginPopup({
                scopes: ['User.Read', 'Files.Read.All', 'Sites.Read.All']
            });
            account = loginResponse.account;
            console.log('Login successful for:', account.username);
        }

        if (account) {
            console.log('Authentication successful. Acquiring token for Graph API...');
            loadingPlaceholder.textContent = 'Acquiring access token...';

            // 3. Get Access Token for Graph API
            const tokenRequest = {
                scopes: ['User.Read', 'Files.Read.All', 'Sites.Read.All'],
                account: account
            };
            const tokenResponse = await myMSALObj.acquireTokenSilent(tokenRequest);
            console.log('Access token acquired.');

            // 4. Fetch the file from SharePoint using YOUR Backend API
            const backendApiUrl = `${import.meta.env.VITE_API_BASE_URL}/api/file/excel`;
            console.log('Attempting to fetch processed data from backend API:', backendApiUrl);
            loadingPlaceholder.textContent = 'Fetching and processing Excel data...';

            const response = await fetch(backendApiUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${tokenResponse.accessToken}` // Pass the token to your backend
                }
            });

            if (!response.ok) {
                let errorMessage = `Backend API error! Status: ${response.status}`;
                try {
                    const errorBody = await response.json();
                    errorMessage += ` - Details: ${errorBody.details || errorBody.error || JSON.stringify(errorBody)}. Please check console for details.`;
                } catch (jsonError) {
                    errorMessage += ` - Details: ${await response.text()}. (Response was not JSON).`;
                }
                throw new Error(errorMessage);
            }

            const processedData = await response.json(); // !!! THIS IS THE CORRECT WAY NOW !!!
            console.log('Processed data received from backend. Rendering UI...');

            renderExcelData(processedData); // Call the new rendering function

            console.log('UI updated with processed data.');
            // The loading placeholder is removed inside renderExcelData on success

        } else {
            throw new Error('Authentication failed: No account available after login attempt.');
        }

    } catch (error) {
        console.error('An error occurred during API call or UI rendering:', error);
        if (loadingPlaceholder) {
            loadingPlaceholder.textContent = `Error: ${error.message}. Please check console for details.`;
            loadingPlaceholder.style.color = 'red';
        }
        alert(`An error occurred: ${error.message}. Please check the console for more details.`);
    }
});