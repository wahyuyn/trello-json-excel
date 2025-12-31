# JSON to Excel Report Converter
This Laravel-based utility is designed to transform complex Trello-style JSON data into a professionally formatted Excel spreadsheet. It handles filtering, data mapping, and automated styling through a custom Artisan command.

## 🛠 Features
* **Automated Mapping:** Maps Trello card schema to a standardized "Ticket" format.
* **Smart Filtering:** Automatically excludes closed cards and specific Trello List IDs.
* **Memory Efficiency:** Implements optimized row-writing to handle large JSON datasets without crashing.
* **UI/UX Optimized:** Generates auto-sized columns and high-contrast styled headers.

## 🚀 Installation
Follow these steps to set up the application locally:

### Clone the repository:
```bash
git clone <your-repo-url>
cd <your-project-directory>
```

### Install PHP Dependencies:
```bash
composer install
```

### Environment Setup:
```bash
cp .env.example .env
php artisan key:generate
```

### Database & Storage:
```bash
php artisan migrate
php artisan storage:link
```

## 📂 Preparation
Before running the converter, you must place your data source in the correct directory:
1. Locate your JSON export file.
2. Move it to: storage/app/private/report.json

## 📊 Usage
To trigger the conversion process, run the following command in your terminal:

```bash
php artisan report:convert-json
```

What happens during execution:
1. Validation: The system checks for the existence and validity of report.json.
2. Transformation: A progress bar will appear while the system filters and maps the JSON data.
3. Optimization: The spreadsheet columns are auto-resized for readability.
4. Output: The final file will be saved as report_export.xlsx inside storage/app/private/.

## ⚙️ Logic & Filtering
The command currently applies the following rules:

* Excluded Lists: Cards belonging to List IDs 67e23... and 6826e... are ignored.
* Status Filter: Only "Open" cards (where dateClosed is null) are processed.
* Default Values: Hardcoded values for Department, Priority, and Agent Assigned are applied to maintain report consistency.
