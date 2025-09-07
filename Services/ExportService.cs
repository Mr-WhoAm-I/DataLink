using DataLink.Data;
using DataLink.Models;
using DataLink.Resources;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.Windows;
using System.Xml;

namespace DataLink.Services
{
    // Service responsible for exporting records to Excel or XML
    public class ExportService(AppDbContext dbContext)
    {
        private readonly AppDbContext _dbContext = dbContext;

        // Maximum number of rows per Excel sheet
        private const int ExcelMaxRowsPerSheet = 1_000_000;

        // Number of rows after which to report progress
        private const int ProgressBatchSize = 1000;

        // Public method to export query results to Excel
        public static async Task ExportToExcel(IQueryable<Record> query, string filePath, IProgress<int>? progress = null, CancellationToken token = default)
        {
            await ExecuteExportAsync(query, filePath, progress, ExportFormat.Excel, token);
        }

        // Public method to export query results to XML
        public static async Task ExportToXml(IQueryable<Record> query, string filePath, IProgress<int>? progress = null, CancellationToken token = default)
        {
            await ExecuteExportAsync(query, filePath, progress, ExportFormat.Xml, token);
        }

        // Core method handling export logic for both Excel and XML formats
        private static async Task ExecuteExportAsync(IQueryable<Record> query, string filePath, IProgress<int>? progress, ExportFormat format, CancellationToken token)
        {
            string tempPath = filePath + ".tmp"; // Temporary file during export

            try
            {
                // Count total records to export
                int totalRecords = await query.CountAsync(token);
                if (totalRecords == 0)
                {
                    ShowError(Resource.Error_ExportNoData);
                    return;
                }

                // Perform export in a background task
                await Task.Run(async () =>
                {
                    int processed = 0; // Tracks number of processed records
                    int sheetIndex = 1; // Excel sheet index
                    int rowIndex = 2;   // Current row in sheet

                    if (format == ExportFormat.Excel)
                    {
                        // Create Excel document
                        using var document = SpreadsheetDocument.Create(tempPath, SpreadsheetDocumentType.Workbook);
                        var workbookPart = document.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                        WorksheetPart? worksheetPart = null;
                        OpenXmlWriter? writer = null;

                        // Iterate through records asynchronously
                        await foreach (var record in query.AsNoTracking().AsAsyncEnumerable().WithCancellation(token))
                        {
                            token.ThrowIfCancellationRequested();

                            // Start new worksheet if needed
                            if (worksheetPart == null || rowIndex > ExcelMaxRowsPerSheet)
                            {
                                writer?.Close();

                                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                                string sheetName = sheetIndex == 1 ? "Records" : $"Records_{sheetIndex}";
                                sheets.Append(new Sheet
                                {
                                    Id = workbookPart.GetIdOfPart(worksheetPart),
                                    SheetId = (uint)sheetIndex,
                                    Name = sheetName
                                });

                                writer = OpenXmlWriter.Create(worksheetPart);
                                writer.WriteStartElement(new Worksheet());
                                writer.WriteStartElement(new SheetData());

                                // Write header row
                                WriteRow(writer,
                                [
                                    Resource.Header_Date,
                                    Resource.Header_FirstName,
                                    Resource.Header_LastName,
                                    Resource.Header_SurName,
                                    Resource.Header_City,
                                    Resource.Header_Country
                                ]);

                                rowIndex = 2;
                                sheetIndex++;
                            }

                            // Write data row
                            WriteRow(writer!,
                            [
                                record.Date.ToString("yyyy-MM-dd"),
                                record.FirstName,
                                record.LastName,
                                record.SurName,
                                record.City,
                                record.Country
                            ]);

                            rowIndex++;
                            processed++;

                            // Report progress every batch
                            if (processed % ProgressBatchSize == 0)
                                progress?.Report(processed);
                        }

                        // Close writer and save workbook
                        writer?.Close();
                        workbookPart.Workbook.Save();
                    }
                    else if (format == ExportFormat.Xml)
                    {
                        // Configure XML writer
                        var settings = new XmlWriterSettings { Indent = true };
                        using var writer = XmlWriter.Create(tempPath, settings);
                        writer.WriteStartDocument();
                        writer.WriteStartElement("TestProgram"); // Root element

                        // Write each record as XML
                        await foreach (var record in query.AsNoTracking().AsAsyncEnumerable().WithCancellation(token))
                        {
                            token.ThrowIfCancellationRequested();

                            writer.WriteStartElement("Record");
                            writer.WriteAttributeString("id", (processed + 1).ToString());
                            writer.WriteElementString("Date", record.Date.ToString("yyyy-MM-dd"));
                            writer.WriteElementString("FirstName", record.FirstName);
                            writer.WriteElementString("LastName", record.LastName);
                            writer.WriteElementString("SurName", record.SurName);
                            writer.WriteElementString("City", record.City);
                            writer.WriteElementString("Country", record.Country);
                            writer.WriteEndElement();

                            processed++;
                            if (processed % ProgressBatchSize == 0)
                                progress?.Report(processed);
                        }

                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }

                    progress?.Report(totalRecords);
                }, token);

                // Replace original file with exported file after successful export
                File.Copy(tempPath, filePath, true);
                File.Delete(tempPath);

                // Notify user of successful export
                Application.Current.Dispatcher.Invoke(() =>
                    MessageBox.Show(Resource.Export_Success, Resource.Export_Success, MessageBoxButton.OK, MessageBoxImage.Information));
            }
            catch (OperationCanceledException)
            {
                // Clean up temporary file if operation was canceled
                if (File.Exists(tempPath))
                    File.Delete(tempPath);

                Application.Current.Dispatcher.Invoke(() =>
                    MessageBox.Show(Resource.Export_Canceled, Resource.Export_CanceledTitle, MessageBoxButton.OK, MessageBoxImage.Information));
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException || ex is OutOfMemoryException)
            {
                // Handle common file-related exceptions
                if (File.Exists(tempPath))
                    File.Delete(tempPath);

                if (ex is UnauthorizedAccessException)
                    ShowError(Resource.Error_ExportFileAccess);
                else if (ex is IOException ioEx && ioEx.Message.Contains("used by another process"))
                    ShowError(Resource.Error_ExportFileOpen);
                else if (ex is IOException)
                    ShowError(Resource.Error_ExportFileAccess);
                else if (ex is OutOfMemoryException)
                    ShowError(Resource.Error_ExportMemory);
            }
            catch (Exception ex)
            {
                // Catch-all for other exceptions
                if (File.Exists(tempPath))
                    File.Delete(tempPath);

                ShowError(Resource.Error_ExportFile, ex.Message);
            }
        }

        // Helper method to write a row into Excel using OpenXmlWriter
        private static void WriteRow(OpenXmlWriter writer, string[] values)
        {
            writer.WriteStartElement(new Row());
            foreach (var value in values)
            {
                writer.WriteElement(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(value ?? string.Empty)
                });
            }
            writer.WriteEndElement();
        }

        // Helper method to show error messages on UI thread
        private static void ShowError(string message, string? detail = null)
        {
            string fullMessage = detail != null ? string.Format(message, detail) : message;
            Application.Current.Dispatcher.Invoke(() =>
                MessageBox.Show(fullMessage, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Error));
        }

        // Enum to specify export format
        private enum ExportFormat
        {
            Excel,
            Xml
        }
    }
}
