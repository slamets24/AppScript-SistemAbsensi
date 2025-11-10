# AppScript-SistemAbsensi

A Google Apps Script based system for attendance tracking and work hour management.

## Key Features & Benefits

*   **Simple Attendance Tracking:** Easily record attendance using a web interface.
*   **Work Hour Management:**  Track start and end times for accurate work hour calculation.
*   **Google Apps Script Integration:** Leverages the power of Google Sheets and other Google services.
*   **Customizable Interface:** Adaptable HTML-based interface for user-friendly interaction.

## Prerequisites & Dependencies

*   **Google Account:** Required to access Google Apps Script and related services.
*   **Google Sheets:**  Attendance data is stored in a Google Sheet.
*   **Basic understanding of HTML and JavaScript:** Helpful for customization.

## Installation & Setup Instructions

1.  **Create a new Google Sheet:** This sheet will store the attendance data.
2.  **Open Google Apps Script:** In the Google Sheet, go to "Tools" > "Script editor".
3.  **Copy `code.gs` contents:** Paste the code from `code.gs` into the script editor.
4.  **Copy `index.html` contents:**
    *   Create a new HTML file in the script editor by clicking the "+" icon next to "Files" and selecting "HTML".
    *   Name the file `index`.
    *   Paste the contents of `index.html` into this new file.
5.  **Deploy as a web app:**
    *   In the script editor, go to "Deploy" > "New deployment".
    *   Configure the deployment settings:
        *   **Type:** Web app
        *   **Description:** (Optional)
        *   **Web app:**
            *   **Execute as:** `your Google account`
            *   **Who has access:** `Anyone with a Google account` or `Anyone` (depending on your needs)
    *   Click "Deploy".
6.  **Authorize the script:** You will be prompted to authorize the script to access your Google account and data.  Grant the necessary permissions.
7.  **Copy the web app URL:** After deployment, you will be given a web app URL. This is the link you will use to access the attendance system.
8. **Set up spreadsheet:** Ensure the Google Sheet has the necessary headers.  For example, you might need columns for: Timestamp, Employee Name, Check-In Time, Check-Out Time, Total Hours. Modify the `code.gs` file to match your spreadsheet structure if needed.

## Usage Examples & API Documentation

### Example: Opening the Attendance System

Open the web app URL obtained during the deployment process. This will display the HTML interface defined in `index.html`.

### Interacting with the System

The `index.html` provides the front-end interface. When a user interacts with this interface (e.g., clicks a "Check-In" button), it will trigger functions defined in `code.gs` which then interact with the Google Sheet.

### API (Indirect)

While there isn't a direct "API" in the traditional sense, the `code.gs` file exposes functions that can be triggered from the HTML interface or even other Apps Scripts.  For example:

```javascript
// Example function in code.gs
function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Sistem Absensi & Jam Kerja');
}

function saveData(data) {
  // Logic to save data to Google Sheet
  // 'data' would contain the attendance information
}
```

The `doGet(e)` function handles initial requests to the web app. Other functions like `saveData(data)` would be called from the HTML interface using `google.script.run`.

## Configuration Options

*   **Spreadsheet ID:**  You **must** configure the `SPREADSHEET_ID` variable in `code.gs` to point to the ID of your Google Sheet. You can find the Spreadsheet ID in the URL of your Google Sheet (e.g., `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`).

    ```javascript
    const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID"; // Replace with your actual Spreadsheet ID
    ```

*   **Sheet Name:** You may also need to adjust the `SHEET_NAME` variable in `code.gs` if your sheet is not named "Sheet1".
    ```javascript
    const SHEET_NAME = "Sheet1"; // Replace with your sheet's name if it is not called "Sheet1"
    ```

*   **Timezone:** Ensure the timezone is correctly configured in your Google Apps Script project settings to reflect the correct time for attendance records.
    *   Go to "File" -> "Project properties" -> "Time zone".

*   **HTML Customization:**  You can modify the `index.html` file to customize the look and feel of the attendance system.

## Contributing Guidelines

1.  Fork the repository.
2.  Create a new branch for your feature or bug fix.
3.  Make your changes and commit them with descriptive messages.
4.  Submit a pull request to the main branch.

## License Information

License not specified.  All rights reserved by slamets24.

## Acknowledgments

This project utilizes Google Apps Script and its associated services.
