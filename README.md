**DataLink** is a  WPF desktop application for managing, filtering, and exporting user data. It supports CSV import, database storage, Excel/XML export, multilingual UI, and efficient data navigation.

## Features

* **CSV Import:** Import large CSV files with validation for dates and text fields. Supports progress reporting and cancellation.
* **Filtering:** Filter records by date, first name, last name, surname, city, and country.
* **Pagination:** Display records in a paginated DataGrid for easy navigation.
* **Export:** Export filtered data to Excel (.xlsx) or XML format with streaming to handle large datasets efficiently.
* **Multilingual UI:** English and Russian language support, dynamically switchable at runtime.
* **Modern UI:** Uses MahApps.Metro for a clean, stylish interface with custom color scheme and hover effects.

## Technology Stack

* **WPF (.NET 8):** Desktop UI framework.
* **MahApps.Metro:** Modern window and control styling.
* **Entity Framework Core:** ORM for SQL Server database operations.
* **CSVHelper:** Efficient CSV parsing with validation.
* **OpenXML SDK:** Excel export without requiring Microsoft Excel installation.
* **XmlWriter:** Streaming XML export.

## Architecture

* **Data Layer:** `AppDbContext` manages database operations using EF Core.
* **Models:** `Record` represents a single data entry.
* **Services:**

  * `CsvService` — handles CSV import, validation, and batch inserts.
  * `ExportService` — handles export to Excel and XML with progress reporting.
* **UI Layer:** `MainWindow` implements user interface, filtering, pagination, and triggers service operations.

## Usage

1. Clone the repository.
2. Open the solution in Visual Studio.
3. Connect your server and DataBase in AppDbContext.cs
4. Build and run the project.
5. Use the **Load CSV** button to import data.
6. Apply filters, navigate pages, and export to Excel or XML.
7. Switch UI language using the dropdown menu.

## Notes

* The application uses asynchronous operations with cancellation support for long-running tasks.
* Suitable for managing large (> 1_000_000) datasets efficiently with batch processing.
