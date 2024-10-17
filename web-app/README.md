Web App for Data Management

This web application provides a streamlined interface for managing and processing records with advanced user authentication and role-based permissions. It integrates Google Apps Script to handle backend operations and Vue.js for a dynamic frontend experience. The app is designed to allow users to view, add, update, and, for admin users, delete records from a connected Google Spreadsheet.

Key Features

User Authentication: Secure login system ensuring that only authorized users can access the application.
Role-Based Access Control: Only admin users have the ability to delete records, while regular users can view, create, and update entries.
Dynamic Record Management: Users can create, edit, and delete records, which are dynamically reflected in the app's data table.
Form Validation: Ensures all required fields are filled and validates specific fields like phone numbers before submission.
Snackbar Notifications: Provides real-time feedback to users after key actions like saving, deleting, and validation errors.
Google Apps Script Integration: The app interacts with Google Sheets as a database through Google Apps Script, providing seamless data management.
Vue.js Framework: The frontend is built using Vue.js, offering a reactive and efficient user interface.

Usage

Once the app is deployed:

Admin users can create, edit, and delete records.
Regular users can only view and edit records but do not have access to delete functionality.
All changes to records will be synced with the connected Google Spreadsheet in real-time.

Technology Stack

Frontend: Vue.js, Vuetify for UI components
Backend: Google Apps Script, Google Sheets as the database
Authentication: Google OAuth

Contributions

Feel free to open an issue or submit a pull request for improvements, new features, or bug fixes.
