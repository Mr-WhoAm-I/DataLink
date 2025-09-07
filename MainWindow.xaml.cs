using DataLink.Data;
using DataLink.Models;
using DataLink.Resources;
using DataLink.Services;
using MahApps.Metro.Controls;
using Microsoft.EntityFrameworkCore;
using Microsoft.Win32;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace DataLink
{
    // Main window for the application. Handles CSV import, filtering, pagination, and export.
    public partial class MainWindow : MetroWindow
    {
        private readonly AppDbContext _dbContext;
        private readonly CsvService _csvService;
        private readonly ExportService _exportService;
        private CancellationTokenSource? _cts; // Token for cancelling long-running operations

        private int _currentPage = 1;      // Tracks current page in DataGrid
        private const int _pageSize = 50;  // Number of records per page
        private int _totalPages = 1;       // Total number of pages

        public MainWindow()
        {
            InitializeComponent();

            _dbContext = new AppDbContext();
            _csvService = new CsvService(_dbContext);
            _exportService = new ExportService(_dbContext);

            // Hide navigation and progress controls until data is loaded
            PageNavPanel.Visibility = Visibility.Collapsed;
            ProgressBarMain.Visibility = Visibility.Collapsed;
            ProgressText.Visibility = Visibility.Collapsed;
        }

        // Event handler for language selection change
        private void LangSelector_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (LangSelector.SelectedItem is ComboBoxItem selected)
            {
                string lang = selected.Tag?.ToString() ?? "en";
                Thread.CurrentThread.CurrentUICulture = new CultureInfo(lang);
                Thread.CurrentThread.CurrentCulture = new CultureInfo(lang);

                string source = lang == "ru" ? "Resources/Dictionary.ru.xaml" : "Resources/Dictionary.xaml";

                // Update application resources for selected language
                Application.Current.Resources.MergedDictionaries.Clear();
                var dict = new ResourceDictionary() { Source = new Uri(source, UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(dict);

                UpdatePageLabel();
            }
        }

        // Event handler to load CSV file
        private async void BtnLoadCsv_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "CSV Files (*.csv)|*.csv" };
            if (openFileDialog.ShowDialog() != true) return;

            DisableUI();
            ProgressBarMain.Visibility = Visibility.Visible;
            ProgressText.Visibility = Visibility.Visible;

            _cts = new CancellationTokenSource();
            BtnCancel.Visibility = Visibility.Visible;

            try
            {
                // Import CSV with progress reporting
                await _csvService.ImportCsv(openFileDialog.FileName, new Progress<int>(count =>
                {
                    ProgressBarMain.Maximum = 100;
                    ProgressBarMain.Value = count;
                    ProgressText.Text = $"{count}%";
                }), _cts.Token);

                _currentPage = 1;
                await LoadPage();
            }
            catch (OperationCanceledException)
            {
                // Handle cancellation
            }
            finally
            {
                EnableUI();
                ProgressBarMain.Visibility = Visibility.Collapsed;
                ProgressText.Visibility = Visibility.Collapsed;
                ProgressText.Text = string.Empty;
                BtnCancel.Visibility = Visibility.Collapsed;
                _cts = null;
            }
        }

        // Apply filters and reload first page
        private async void BtnApplyFilters_Click(object sender, RoutedEventArgs e)
        {
            _currentPage = 1;
            await LoadPage();
        }

        // Export filtered data to Excel
        private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx" };
            if (saveFileDialog.ShowDialog() != true) return;

            var query = GetFilteredQuery();
            DisableUI();
            ProgressBarMain.Visibility = Visibility.Visible;
            ProgressText.Visibility = Visibility.Visible;

            int totalRecords = await query.CountAsync();

            var progress = new Progress<int>(count =>
            {
                ProgressBarMain.Maximum = totalRecords;
                ProgressBarMain.Value = count;
                ProgressText.Text = $"{count} / {totalRecords}";
            });

            _cts = new CancellationTokenSource();
            BtnCancel.Visibility = Visibility.Visible;

            try
            {
                await ExportService.ExportToExcel(query, saveFileDialog.FileName, progress, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                // Handle export cancellation
            }
            finally
            {
                _cts.Dispose();
                _cts = null;
                BtnCancel.Visibility = Visibility.Collapsed;
                ProgressBarMain.Value = 0;
                ProgressText.Text = string.Empty;
                ProgressBarMain.Visibility = Visibility.Collapsed;
                ProgressText.Visibility = Visibility.Collapsed;
                EnableUI();
            }
        }

        // Export filtered data to XML
        private async void BtnExportXml_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog { Filter = "XML Files (*.xml)|*.xml" };
            if (saveFileDialog.ShowDialog() != true) return;

            var query = GetFilteredQuery();
            DisableUI();
            ProgressBarMain.Visibility = Visibility.Visible;
            ProgressText.Visibility = Visibility.Visible;

            int totalRecords = await query.CountAsync();

            var progress = new Progress<int>(count =>
            {
                ProgressBarMain.Maximum = totalRecords;
                ProgressBarMain.Value = count;
                ProgressText.Text = $"{count} / {totalRecords}";
            });

            _cts = new CancellationTokenSource();
            BtnCancel.Visibility = Visibility.Visible;

            try
            {
                await ExportService.ExportToXml(query, saveFileDialog.FileName, progress, _cts.Token);
            }
            catch (OperationCanceledException)
            {
                // Handle export cancellation
            }
            finally
            {
                _cts.Dispose();
                _cts = null;
                BtnCancel.Visibility = Visibility.Collapsed;
                ProgressBarMain.Value = 0;
                ProgressText.Text = string.Empty;
                ProgressBarMain.Visibility = Visibility.Collapsed;
                ProgressText.Visibility = Visibility.Collapsed;
                EnableUI();
            }
        }

        // Cancel ongoing operations
        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cts?.Cancel();
        }

        // Navigate to previous page
        private async void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage <= 1) return;
            _currentPage--;
            await LoadPage();
        }

        // Navigate to next page
        private async void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage >= _totalPages) return;
            _currentPage++;
            await LoadPage();
        }

        // Build filtered query based on user input
        private IQueryable<Record> GetFilteredQuery()
        {
            var query = _dbContext.Records.AsQueryable();

            // Filter by start date
            if (FilterStartDatePicker.SelectedDate.HasValue)
            {
                var startDate = FilterStartDatePicker.SelectedDate.Value.Date;
                query = query.Where(r => r.Date.Date >= startDate);
            }

            // Filter by end date
            if (FilterEndDatePicker.SelectedDate.HasValue)
            {
                var endDate = FilterEndDatePicker.SelectedDate.Value.Date;
                query = query.Where(r => r.Date.Date <= endDate);
            }

            // Validate date range
            if (FilterStartDatePicker.SelectedDate.HasValue && FilterEndDatePicker.SelectedDate.HasValue)
            {
                if (FilterStartDatePicker.SelectedDate > FilterEndDatePicker.SelectedDate)
                {
                    MessageBox.Show(Resource.Error_InvalidDateRange, Resource.Error_Generic, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            // Filter by text fields if provided
            if (!string.IsNullOrWhiteSpace(FilterFirstName.Text))
            {
                string fn = FilterFirstName.Text;
                query = query.Where(r => r.FirstName != null && r.FirstName.Contains(fn));
            }

            if (!string.IsNullOrWhiteSpace(FilterLastName.Text))
            {
                string ln = FilterLastName.Text;
                query = query.Where(r => r.LastName != null && r.LastName.Contains(ln));
            }

            if (!string.IsNullOrWhiteSpace(FilterSurName.Text))
            {
                string sn = FilterSurName.Text;
                query = query.Where(r => r.SurName != null && r.SurName.Contains(sn));
            }

            if (!string.IsNullOrWhiteSpace(FilterCity.Text))
            {
                string city = FilterCity.Text;
                query = query.Where(r => r.City != null && r.City.Contains(city));
            }

            if (!string.IsNullOrWhiteSpace(FilterCountry.Text))
            {
                string country = FilterCountry.Text;
                query = query.Where(r => r.Country != null && r.Country.Contains(country));
            }

            return query;
        }

        // Load data for the current page and update UI
        private async Task LoadPage()
        {
            DisableUI();

            var query = GetFilteredQuery();
            int totalRecords = await query.CountAsync();
            _totalPages = Math.Max(1, (int)Math.Ceiling((double)totalRecords / _pageSize));

            if (_currentPage > _totalPages)
                _currentPage = _totalPages;

            // Load current page of records
            var page = await query
                .OrderBy(r => r.Id)
                .Skip((_currentPage - 1) * _pageSize)
                .Take(_pageSize)
                .ToListAsync();

            RecordsDataGrid.ItemsSource = page;
            PageNavPanel.Visibility = totalRecords > 0 ? Visibility.Visible : Visibility.Collapsed;

            UpdatePageLabel();
            BtnPrevPage.IsEnabled = _currentPage > 1;
            BtnNextPage.IsEnabled = _currentPage < _totalPages;

            EnableUI();
        }

        // Update page label with current/total pages
        private void UpdatePageLabel()
        {
            if (PageNavPanel == null || PageLabel == null) return;

            if (PageNavPanel.Visibility != Visibility.Visible)
            {
                PageLabel.Content = string.Empty;
                return;
            }

            var format = (string)Application.Current.Resources["Label_PageInfo"];
            PageLabel.Content = string.Format(format, _currentPage, _totalPages);
        }

        // Disable all main UI controls
        private void DisableUI()
        {
            BtnLoadCsv.IsEnabled = false;
            BtnApplyFilters.IsEnabled = false;
            BtnExportExcel.IsEnabled = false;
            BtnExportXml.IsEnabled = false;
            BtnPrevPage.IsEnabled = false;
            BtnNextPage.IsEnabled = false;
        }

        // Enable all main UI controls
        private void EnableUI()
        {
            BtnLoadCsv.IsEnabled = true;
            BtnApplyFilters.IsEnabled = true;
            BtnExportExcel.IsEnabled = true;
            BtnExportXml.IsEnabled = true;
            BtnPrevPage.IsEnabled = _currentPage > 1;
            BtnNextPage.IsEnabled = _currentPage < _totalPages;
        }
    }
}
