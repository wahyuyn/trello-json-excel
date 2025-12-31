## JSON to Excel Report Converter
This Laravel-based utility is designed to transform complex Trello-style JSON data into a professionally formatted Excel spreadsheet. It handles filtering, data mapping, and automated styling through a custom Artisan command.

## 🛠 Features
Automated Mapping: Converts Trello card attributes into a standard "Ticket" format.

Smart Filtering: Automatically excludes closed cards and specific list IDs.

Optimized Performance: Uses PhpSpreadsheet with memory-efficient streaming to handle large datasets.

Professional Formatting: Generates auto-sized columns and styled headers for immediate use in reporting.

## 🚀 Installation
Follow these steps to set up the application locally:

Clone the repository:

Bash

git clone <your-repo-url>
cd <your-project-directory>
Install PHP Dependencies:

Bash

composer install
Environment Setup:

Bash

cp .env.example .env
php artisan key:generate
Database & Storage:

Bash

php artisan migrate
php artisan storage:link
## 📂 Preparation
Before running the command, you must provide the source data.

Navigate to storage/app/private/ (create the private folder if it doesn't exist).

Paste your source file and rename it to report.json.

## 📊 Usage
To trigger the conversion process, run the following command in your terminal:

Bash

php artisan report:convert-json
What happens during execution:
Validation: The system checks for the existence and validity of report.json.

Transformation: A progress bar will appear while the system filters and maps the JSON data.

Optimization: The spreadsheet columns are auto-resized for readability.

Output: The final file will be saved as report_export.xlsx inside storage/app/private/.

## ⚙️ Logic & Filtering
The command currently applies the following rules:

Excluded Lists: Cards belonging to List IDs 67e23... and 6826e... are ignored.

Status Filter: Only "Open" cards (where dateClosed is null) are processed.

Default Values: Hardcoded values for Department, Priority, and Agent Assigned are applied to maintain report consistency.