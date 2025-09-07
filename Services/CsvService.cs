using CsvHelper;
using CsvHelper.Configuration;
using DataLink.Data;
using DataLink.Models;
using DataLink.Resources;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace DataLink.Services
{
    // Service responsible for CSV import operations
    public class CsvService(AppDbContext dbContext)
    {
        private readonly AppDbContext _dbContext = dbContext;
        private const int ProgressBatchSize = 1000;

        // Imports CSV file into the database with progress reporting and cancellation support
        public async Task ImportCsv(string filePath, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
        {
            // Check if the file exists
            if (!File.Exists(filePath))
            {
                MessageBox.Show(Resource.Error_CsvFileAccess, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // CSV reader configuration
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                HasHeaderRecord = false,
                BadDataFound = null
            };

            const int batchSize = 10000; // Number of records to insert in a single batch
            var batch = new List<Record>();

            // Flags for invalid data detection
            int invalidColumnsCount = 0;
            bool hasInvalidTextField = false;
            bool hasInvalidDate = false;

            // Begin a database transaction
            await using var transaction = await _dbContext.Database.BeginTransactionAsync(cancellationToken);

            try
            {
                long totalLines = File.ReadLines(filePath).LongCount(); // Total lines in CSV
                long processed = 0; // Lines processed

                using var reader = new StreamReader(filePath);
                using var csv = new CsvReader(reader, config);

                // Main CSV reading loop
                while (await csv.ReadAsync())
                {
                    cancellationToken.ThrowIfCancellationRequested(); // Check for cancellation
                    processed++;

                    // Validate number of columns
                    if (csv.Parser.Count != 6)
                    {
                        invalidColumnsCount++;
                        continue;
                    }

                    // Validate date field
                    if (!DateTime.TryParse(csv.GetField<string>(0), out DateTime date))
                    {
                        hasInvalidDate = true;
                        continue;
                    }

                    // Validate text fields (no digits allowed)
                    for (int i = 1; i <= 5; i++)
                    {
                        string field = csv.GetField<string>(i) ?? string.Empty;
                        if (Regex.IsMatch(field, @"\d"))
                        {
                            hasInvalidTextField = true;
                            break;
                        }
                    }

                    // Create record object
                    var record = new Record
                    {
                        Date = date,
                        FirstName = csv.GetField<string>(1) ?? string.Empty,
                        LastName = csv.GetField<string>(2) ?? string.Empty,
                        SurName = csv.GetField<string>(3) ?? string.Empty,
                        City = csv.GetField<string>(4) ?? string.Empty,
                        Country = csv.GetField<string>(5) ?? string.Empty
                    };

                    batch.Add(record);

                    // Batch insert into database
                    if (batch.Count >= batchSize)
                    {
                        await _dbContext.Records.AddRangeAsync(batch, cancellationToken);
                        await _dbContext.SaveChangesAsync(cancellationToken);
                        batch.Clear();
                    }

                    // Report progress every 1000 lines
                    if (processed % ProgressBatchSize == 0 && progress != null)
                    {
                        int percent = (int)((double)processed / totalLines * 100);
                        progress.Report(percent);
                    }
                }

                // Insert remaining records
                if (batch.Count != 0)
                {
                    await _dbContext.Records.AddRangeAsync(batch, cancellationToken);
                    await _dbContext.SaveChangesAsync(cancellationToken);
                }

                // Commit transaction
                await transaction.CommitAsync(cancellationToken);

                progress?.Report(100); // Report 100% progress

                // Show warnings for invalid data
                if (invalidColumnsCount > 0)
                {
                    MessageBox.Show(
                        string.Format(Resource.Error_CsvPartialWarning, invalidColumnsCount),
                        Resource.Error_Generic,
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }

                if (hasInvalidDate)
                    MessageBox.Show(Resource.Error_InvalidDate, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Error);

                if (hasInvalidTextField)
                    MessageBox.Show(Resource.Error_InvalidTextField, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (OperationCanceledException)
            {
                // Rollback transaction if operation was canceled
                await transaction.RollbackAsync(cancellationToken);
                MessageBox.Show(Resource.Import_Canceled, Resource.Export_CanceledTitle,
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // Rollback transaction for any other exception
                await transaction.RollbackAsync(cancellationToken);

                if (ex is UnauthorizedAccessException || ex is IOException)
                {
                    MessageBox.Show(Resource.Error_CsvFileAccess, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show(string.Format(Resource.Error_Generic, ex.Message),
                        Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
