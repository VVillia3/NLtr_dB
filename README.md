# Newsletter Recipients dB

Newsletter Recipients dB is a desktop application for managing newsletter mailing and email lists. It stores recipient information locally on the user's computer and provides tools for maintaining, reviewing, importing, backing up, and exporting recipient records.

The application was designed for organizations that need a straightforward way to manage both printed-newsletter recipients and email-newsletter recipients without relying on an online database service.

## Features

* Maintain mailing-list and email-list recipient records
* Store names, contact information, mailing addresses, list assignments, and notes
* Search records by last name
* Edit or delete existing records
* Identify possible duplicate records during data entry
* Combine duplicate records while choosing which information to preserve
* Import recipient records from CSV files
* Review possible CSV duplicates before replacing or selectively updating existing data
* Add legitimate additional household members as separate records when recipients share an address
* Automatically skip exact CSV matches while documenting them in an exceptions report
* Export a complete CSV backup
* Export Excel spreadsheets for all recipients, mailing-list recipients, or email-list recipients
* Export mailing labels for Avery 5162 labels (14 per page)
* Export mailing labels for Avery 5160-compatible labels (30 per page)
* Store the database locally on the user's computer

## Download

Download the latest version from the Releases page.

- Windows: download the `.exe` installer
- macOS: download the `.dmg` disk image

## Download the Windows Version

Download the newest Windows ZIP file from the [Releases](../../releases/latest) page.

After downloading the ZIP file:

1. Extract the ZIP file to a folder on your computer.
2. Open the extracted folder.
3. Double-click **Newsletter Recipients dB.exe**.

The entire extracted folder should be kept together. Do not move the `.exe` file by itself, because the application requires the supporting files included in the folder.

Windows Defender or another security tool may display a warning the first time the application is opened because the downloadable build is not digitally signed.

## Data Storage

The application stores its SQLite database locally. Recipient records are not uploaded to an online service.

On Windows, the database is stored under the current user's application-data folder:

```text
%APPDATA%\NewsletterRecipients\
```

The application creates a separate database filename based on the computer name or user environment.

## Backups

Use the application's **Export CSV** feature regularly to create a portable backup of the recipient database.

A CSV backup can also be imported into another installation of Newsletter Recipients dB.

## CSV Import and Duplicate Review

When a CSV file is imported, the application checks for possible duplicates using available identifying information such as email address, phone number, last name, and mailing address.

Exact matches are skipped automatically and noted in the CSV exceptions report.

When a possible partial match is found, the application opens a review window. The user can:

* Replace the entire existing record
* Update selected fields only
* Add the CSV entry as a separate new record
* Skip the CSV entry
* Cancel the import

The **Add as Separate New Record** option is useful when several members of the same household share an Address 1 value but have different Address 2 values or other individual information.

## Mailing Labels

The application can create Microsoft Word documents formatted for:

* **Avery 5162** labels: 14 labels per page
* **Avery 5160-compatible** labels: 30 labels per page

Additional pages are generated automatically when the mailing list exceeds one sheet.

## Running from Source

Newsletter Recipients dB is written in Python and uses Tkinter with a local SQLite database.

Required Python packages include:

```text
openpyxl
python-docx
```

To run the source code:

```text
python Newsletter_Recipients.py
```

To build a distributable Windows folder with PyInstaller, run the build process on a Windows computer.

## License

Newsletter Recipients dB is distributed under the MIT License.

See [`LICENSE.txt`](LICENSE.txt) for the full license text.

## Version

Current release: **1.0.5**

## Author

Copyright © 2026 William Tinney
