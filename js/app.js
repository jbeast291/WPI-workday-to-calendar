const fileInput = document.getElementById('fileInput');
const parseBtn = document.getElementById('parseBtn');
const statusDiv = document.getElementById('status');
const downloadBtn = document.getElementById('downloadBtn');

let selectedFile = null;

// reset file input on page reload
ResetFileInput();

function ResetFileInput(){

    fileInput.value = '';
    selectedFile = null;
    parseBtn.disabled = true;
    downloadBtn.disabled = true;
    statusDiv.textContent = 'Waiting for file...';
}

fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];

    if (file) {
        selectedFile = file;

        //Once a file is uploaded, allow parsing
        parseBtn.disabled = false;
        downloadBtn.disabled = true;
        statusDiv.textContent = `File selected: ${file.name}`;
    } else {
        selectedFile = null;
        parseBtn.disabled = true;
        downloadBtn.disabled = true;
        statusDiv.textContent = 'Waiting for file...';
    }
});

// ======= BEGIN EXCEL PARSE =======

let lastParseJsonData = null;

parseBtn.addEventListener('click', () => {
    if (!selectedFile) return;

    statusDiv.textContent = 'Parsing...';

    const reader = new FileReader();

    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            worksheet['!ref'] = `A1:N${GetWorksheetMaxRow(worksheet)}`;
            
            // Convert worksheet to raw 2D array
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Find row index containing 'Meeting Patterns' or 'Section'
            const headerRowIndex = rawData.findIndex(row =>
                row && row.some(cell => typeof cell === 'string' && cell.includes('Meeting Patterns'))
            );

            if (headerRowIndex === -1) {
                statusDiv.textContent = "Could not locate Workday header row.";
                return;
            }

            const headers = rawData[headerRowIndex];
            const courseRows = rawData.slice(headerRowIndex + 1);

            // Map course rows to JS objects
            const jsonData = courseRows
                .filter(row => row && row.length > 0)
                .map(row => {
                    let obj = {};
                    headers.forEach((header, colIdx) => {
                        if (header) {
                            obj[header] = row[colIdx] !== undefined ? row[colIdx] : "";
                        }
                    });
                    return obj;
                });

            console.log("Final Parsed Courses:", jsonData);
            statusDiv.textContent = `Successfully parsed ${jsonData.length} courses!`;
            lastParseJsonData = jsonData;
            downloadBtn.disabled = false;

        } catch (error) {
            console.error("Parsing Error:", error);
            statusDiv.textContent = "Error parsing Excel file.";
            downloadBtn.disabled = true;
        }
    };

    reader.readAsArrayBuffer(selectedFile);
});

function GetWorksheetMaxRow(worksheet) {
    // Automatically recalculate the bounding box based on existing keys
    const keys = Object.keys(worksheet).filter(k => !k.startsWith('!'));
    const rows = keys.map(k => parseInt(k.replace(/[^\d]/g, ''), 10)).filter(Boolean);
    const maxRow = Math.max(...rows);
    return maxRow;
}

// ======= BEGIN ICS GENERATION =======

// ======= UPDATED ICS GENERATION =======

function generateICS(courses) {
    let icsContent = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Workday Course Schedule//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH"
    ];

    const dayMap = { 'M': 'MO', 'T': 'TU', 'W': 'WE', 'R': 'TH', 'F': 'FR', 'S': 'SA', 'U': 'SU' };

    courses.forEach((course, idx) => {
        const pattern = course["Meeting Patterns"];
        const startDateRaw = course["Start Date"];
        const endDateRaw = course["End Date"];

        if (!pattern || !startDateRaw || !endDateRaw) return;

        // Pattern format: "M-T-R-F | 12:00 PM - 12:50 PM | Salisbury Labs 115 Kinnicutt Hall"
        const parts = pattern.split('|').map(s => s.trim());
        if (parts.length < 3) return;

        const daysStr = parts[0];       // e.g., "M-T-R-F"
        const timesStr = parts[1];      // e.g., "12:00 PM - 12:50 PM"
        const location = parts[2];      // e.g., "Salisbury Labs 115"

        // Map day letters to BYDAY codes
        const byDays = daysStr.split('-').map(d => dayMap[d.trim()]).filter(Boolean).join(',');

        // Convert times to HHMMSS format
        const [startTimeStr, endTimeStr] = timesStr.split('-').map(t => t.trim());
        const startTimeFormatted = formatTime24(startTimeStr);
        const endTimeFormatted = formatTime24(endTimeStr);

        // Find the FIRST actual class meeting date based on meeting pattern days
        const firstOccurrenceDate = getFirstOccurrenceDate(startDateRaw, daysStr);
        const startDateFormatted = formatDateYYYYMMDD(firstOccurrenceDate);
        const endDateFormatted = formatDateYYYYMMDD(endDateRaw);

        icsContent.push(
            "BEGIN:VEVENT",
            `UID:course-${idx}-${Date.now()}@workday`,
            `SUMMARY:${course["Section"] || course["Course Listing"]}`,
            `LOCATION:${location}`,
            `DESCRIPTION:Instructor: ${course["Instructor"] || "N/A"}\\nFormat: ${course["Instructional Format"] || "N/A"}`,
            `DTSTART;TZID=America/New_York:${startDateFormatted}T${startTimeFormatted}`,
            `DTEND;TZID=America/New_York:${startDateFormatted}T${endTimeFormatted}`,
            `RRULE:FREQ=WEEKLY;UNTIL=${endDateFormatted}T235959Z;BYDAY=${byDays}`,
            "END:VEVENT"
        );
    });

    icsContent.push("END:VCALENDAR");
    return icsContent.join("\r\n");
}

function getFirstOccurrenceDate(startDateRaw, daysStr) {
    const dayToJsNum = { 'M': 1, 'T': 2, 'W': 3, 'R': 4, 'F': 5, 'S': 6, 'U': 0 };

    let dt;
    if (typeof startDateRaw === 'number') {
        const parsed = XLSX.SSF.parse_date_code(startDateRaw);
        dt = new Date(parsed.y, parsed.m - 1, parsed.d);
    } else {
        dt = new Date(startDateRaw);
    }

    // Convert days string like "T-F" to array of JS weekday numbers [2, 5]
    const targetDays = daysStr.split('-').map(d => dayToJsNum[d.trim()]).filter(d => d !== undefined);
    if (targetDays.length === 0) return dt;

    // Advance date day-by-day until matching a meeting day
    let checkDate = new Date(dt);
    for (let i = 0; i < 7; i++) {
        if (targetDays.includes(checkDate.getDay())) {
            return checkDate;
        }
        checkDate.setDate(checkDate.getDate() + 1);
    }

    return dt;
}

function formatTime24(timeStr) {
    const [time, modifier] = timeStr.split(' ');
    let [hours, minutes] = time.split(':');
    if (modifier === 'PM' && hours !== '12') hours = parseInt(hours, 10) + 12;
    if (modifier === 'AM' && hours === '12') hours = '00';
    return `${String(hours).padStart(2, '0')}${minutes}00`;
}

function formatDateYYYYMMDD(val) {
    let dt;
    if (val instanceof Date) {
        dt = val;
    } else if (typeof val === 'number') {
        const parsed = XLSX.SSF.parse_date_code(val);
        dt = new Date(parsed.y, parsed.m - 1, parsed.d);
    } else {
        dt = new Date(val);
    }
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}${m}${d}`;
}

downloadBtn.addEventListener('click', () => {
    if(lastParseJsonData == null) return;
    const icsString = generateICS(lastParseJsonData);
    const blob = new Blob([icsString], { type: 'text/calendar;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'My_Course_Schedule.ics';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});