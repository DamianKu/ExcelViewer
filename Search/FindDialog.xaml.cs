using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Search
{
    public partial class FindDialog : Window
    {
        private readonly MainWindow _mainWindow;
        private readonly SpreadsheetSearchService _searchService;

        private List<SearchResult> _results = new();
        private int _currentIndex = -1;

        public FindDialog(
            MainWindow mainWindow,
            IReadOnlyList<SpreadsheetModel> sheets)
        {
            InitializeComponent();

            _mainWindow = mainWindow;
            _searchService = new SpreadsheetSearchService(sheets);
        }

        // ====== ZNAJDŹ WSZYSTKIE ======
        private void OnFindAll(object sender, RoutedEventArgs e)
        {
            string text = FindTextBox.Text;

            if (string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("Wpisz tekst do wyszukania.");
                return;
            }

            _results = _searchService
                .FindAll(text, false, false)
                .ToList();

            _currentIndex = -1;

            ResultsListBox.ItemsSource = _results;
            ResultsInfoText.Text = $"Znaleziono: {_results.Count}";

            if (_results.Count > 0)
            {
                GoToIndex(0);
            }
        }

        // ====== ZNAJDŹ NASTĘPNY ======
        private void OnFindNext(object sender, RoutedEventArgs e)
        {
            if (_results.Count == 0)
            {
                OnFindAll(sender, e);
                return;
            }

            _currentIndex++;
            if (_currentIndex >= _results.Count)
                _currentIndex = 0;

            GoToIndex(_currentIndex);
        }

        // ====== KLIK NA LIŚCIE ======
        private void ResultsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ResultsListBox.SelectedItem is SearchResult result)
            {
                _mainWindow.GoToSearchResult(result);
            }
        }

        // ====== WSPÓLNA LOGIKA ======
        private void GoToIndex(int index)
        {
            if (index < 0 || index >= _results.Count)
                return;

            _currentIndex = index;
            var result = _results[index];

            ResultsListBox.SelectedItem = result;
            ResultsListBox.ScrollIntoView(result);

            _mainWindow.GoToSearchResult(result);
        }
    }
}
