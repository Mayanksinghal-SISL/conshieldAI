// Helper function for safe numeric conversion - Keep this here
// This is still needed for tooltip calculations in the frontend
const getNumericValue = (text) => {
    const raw = String(text).trim().replace(/[$,%]/g, '');
    return (!isNaN(raw) && raw !== '' && raw.toUpperCase() !== 'NA') ? Number(raw) : null;
};

/**
 * Renders the processed Excel data (JSON object from backend) into the HTML table and accordions.
 * @param {Object} processedData An object containing clientDataForDisplay, combinedDateHeaders, and flags.
 */
function renderProcessedData(processedData) {
    const {
        clientDataForDisplay,
        combinedDateHeaders,
        hasNonZeroMonthlyCount,
        hasNonZeroFortnightlyCount,
        hasNonZeroWeeklyCount
    } = processedData;

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

    // If wrapper doesn't exist (first run), create it
    let accordionWrapper = mobileAccordionContainer.querySelector('.accordionWrapper');
    if (!accordionWrapper) {
        accordionWrapper = document.createElement('div');
        accordionWrapper.classList.add('accordionWrapper');
        mobileAccordionContainer.appendChild(accordionWrapper);
    }
    // Clear only the accordion items, not the logo
    accordionWrapper.innerHTML = '';
    accordionWrapper.classList.remove('visible');

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
            const numericVal = getNumericValue(val);
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
        tdYesterday.textContent = getNumericValue(clientData.yesterdayData) !== null ? Math.round(getNumericValue(clientData.yesterdayData)) : (clientData.yesterdayData === '' ? '' : clientData.yesterdayData);
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
        accordionWrapper.appendChild(accordionItem);

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

    // Remove loading placeholder on successful rendering
    document.getElementById('loadingPlaceholder')?.remove();

    // Update the last updated time
    const lastUpdatedSpan = document.getElementById('lastUpdated');
    if (lastUpdatedSpan) {
        const now = new Date();
        const options = {
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: true
        };
        lastUpdatedSpan.textContent = now.toLocaleString('en-US', options);
    }
}

// --- Main execution starts when the page loads ---
document.addEventListener('DOMContentLoaded', async function () {
    console.log('Page loaded, initiating data fetch from backend...');

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

    // API Configuration
    const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || window.location.origin;
    const EXCEL_API_ENDPOINT = `${API_BASE_URL}/api/file/excel`;

    // Function to load data and process it
    async function loadAndProcessExcelData() {
        try {
            loadingPlaceholder.textContent = 'Fetching data from backend...';
            console.log('Fetching from:', EXCEL_API_ENDPOINT);

            const response = await fetch(EXCEL_API_ENDPOINT);

            if (!response.ok) {
                let errorMessage = `Backend API error! Status: ${response.status}`;
                try {
                    const errorBody = await response.json();
                    errorMessage += ` - Details: ${errorBody.error?.message || JSON.stringify(errorBody)}`;
                } catch (jsonError) {
                    const errorText = await response.text();
                    errorMessage += ` - Raw Response: ${errorText.substring(0, 200)}...`;
                }
                throw new Error(errorMessage);
            }

            // IMPORTANT: Now we expect JSON data from the backend, not the raw ArrayBuffer!
            const processedData = await response.json();

            // The backend's console.log "Raw Excel file fetched from backend..." actually occurs on the server.
            // On the frontend, we are now just receiving the *processed* data.
            console.log('Processed data received from backend. Keys:', Object.keys(processedData));

            loadingPlaceholder.textContent = 'Building display...';
            renderProcessedData(processedData); // Call the new render function

            console.log('Data processing complete and UI updated.');

        } catch (error) {
            console.error('An error occurred during backend fetch or processing:', error);
            if (loadingPlaceholder) {
                loadingPlaceholder.textContent = `Error: ${error.message}. Please check console for details.`;
                loadingPlaceholder.style.color = 'red';
            }
            alert(`An error occurred: ${error.message}. Please check the browser console (F12) for more details.`);
        }
    }

    // Initial load
    loadAndProcessExcelData();

    // Set up auto-refresh every 5 minutes (300,000 milliseconds)
    // setInterval(loadAndProcessExcelData, 5 * 60 * 1000);
});