using System;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using SsmsResultsGrid.Core.Models;
using SsmsResultsGrid.Core.ViewModels;

namespace SsmsResultsGrid.Views
{
    /// <summary>
    /// View for <see cref="ResultsViewModel"/>. Code-behind is limited to view
    /// concerns that WPF cannot express in bindings: explicit column generation,
    /// sort-gesture interception (sorting runs on a background thread in the VM),
    /// row-number headers, the save dialog, and clipboard access.
    /// </summary>
    public partial class ResultsViewControl : UserControl
    {
        private ResultsViewModel _viewModel;
        private ResultSetViewModel _attachedSet;

        public ResultsViewControl()
        {
            InitializeComponent();
            DataContextChanged += OnDataContextChanged;
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
            }

            _viewModel = e.NewValue as ResultsViewModel;

            if (_viewModel != null)
            {
                _viewModel.PropertyChanged += OnViewModelPropertyChanged;
            }
            AttachResultSet(_viewModel?.SelectedResultSet);
        }

        private void OnViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ResultsViewModel.SelectedResultSet))
            {
                AttachResultSet(_viewModel?.SelectedResultSet);
            }
        }

        private void AttachResultSet(ResultSetViewModel set)
        {
            if (_attachedSet != null)
            {
                _attachedSet.ColumnsChanged -= OnColumnsChanged;
            }

            _attachedSet = set;

            if (_attachedSet != null)
            {
                _attachedSet.ColumnsChanged += OnColumnsChanged;
            }
            RebuildColumns();
        }

        private void OnColumnsChanged(object sender, EventArgs e) => RebuildColumns();

        // ---- column generation ----

        private void RebuildColumns()
        {
            SaveColumnWidths();
            ResultsGrid.Columns.Clear();
            var set = _attachedSet;
            if (set == null) return;

            var columns = set.ColumnNames;
            for (int i = 0; i < columns.Count; i++)
            {
                var column = new DataGridTemplateColumn
                {
                    Header = columns[i],
                    CellTemplate = BuildCellTemplate(i),
                    ClipboardContentBinding = new Binding("[" + i.ToString(CultureInfo.InvariantCulture) + "]"),
                    SortMemberPath = "[" + i.ToString(CultureInfo.InvariantCulture) + "]",
                    CanUserSort = true,
                    MinWidth = 60
                };

                if (set.ColumnWidthMemory.TryGetValue(WidthKey(i, columns[i]), out var width))
                {
                    column.Width = new DataGridLength(width);
                }
                ResultsGrid.Columns.Add(column);
            }
        }

        /// <summary>
        /// Cell template: a TextBlock rendered through the TextHighlight attached
        /// properties. Highlight state binds to the APPLIED filter on the selected
        /// result set so highlights always agree with the visible rows.
        /// </summary>
        private static DataTemplate BuildCellTemplate(int columnIndex)
        {
            var factory = new FrameworkElementFactory(typeof(TextBlock));
            factory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            factory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);

            factory.SetBinding(TextHighlight.SourceTextProperty, new Binding("[" + columnIndex.ToString(CultureInfo.InvariantCulture) + "]")
            {
                Mode = BindingMode.OneWay,
                FallbackValue = string.Empty,
                TargetNullValue = string.Empty
            });
            factory.SetBinding(TextHighlight.HighlightTextProperty, GridScopedBinding("DataContext.SelectedResultSet.AppliedFilterText"));
            factory.SetBinding(TextHighlight.MatchModeProperty, GridScopedBinding("DataContext.SelectedResultSet.AppliedFilterModeIndex"));
            factory.SetBinding(TextHighlight.CaseSensitiveProperty, GridScopedBinding("DataContext.SelectedResultSet.AppliedFilterCaseSensitive"));

            var template = new DataTemplate { VisualTree = factory };
            template.Seal();
            return template;
        }

        private static Binding GridScopedBinding(string path) => new Binding(path)
        {
            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor)
            {
                AncestorType = typeof(DataGrid)
            }
        };

        private void SaveColumnWidths()
        {
            var set = _attachedSet;
            if (set == null) return;
            for (int i = 0; i < ResultsGrid.Columns.Count; i++)
            {
                var column = ResultsGrid.Columns[i];
                if (column.Width.IsAbsolute)
                {
                    set.ColumnWidthMemory[WidthKey(i, column.Header?.ToString())] = column.Width.Value;
                }
            }
        }

        private static string WidthKey(int index, string name) =>
            index.ToString(CultureInfo.InvariantCulture) + "|" + (name ?? string.Empty);

        // ---- view event handlers ----

        private void ResultsGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = e.Row.Item is ResultRow row
                ? row.SourceRowNumber.ToString("N0", CultureInfo.CurrentCulture)
                : (e.Row.GetIndex() + 1).ToString("N0", CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Intercepts the header-click sort gesture: the VM sorts on a threadpool
        /// thread and republishes VisibleRows, avoiding a UI-thread collection-view
        /// sort over very large row sets. Cycles ascending → descending → unsorted.
        /// </summary>
        private void ResultsGrid_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;
            var set = _attachedSet;
            if (set == null) return;

            int columnIndex = ResultsGrid.Columns.IndexOf(e.Column);
            if (columnIndex < 0) return;

            var next = e.Column.SortDirection == null
                ? ListSortDirection.Ascending
                : e.Column.SortDirection == ListSortDirection.Ascending
                    ? ListSortDirection.Descending
                    : (ListSortDirection?)null;

            foreach (var column in ResultsGrid.Columns)
            {
                column.SortDirection = null;
            }
            e.Column.SortDirection = next;

            if (next == null)
            {
                set.ClearSort();
            }
            else
            {
                set.SetSort(columnIndex, next == ListSortDirection.Descending);
            }
        }

        private void ClearFilterButton_Click(object sender, RoutedEventArgs e)
        {
            if (_attachedSet != null)
            {
                _attachedSet.FilterText = string.Empty;
            }
            FilterBox.Focus();
        }

        private void CopySelectedButton_Click(object sender, RoutedEventArgs e)
        {
            if (ResultsGrid.Items.Count == 0) return;
            ApplicationCommands.Copy.Execute(null, ResultsGrid);
            if (_viewModel != null)
            {
                _viewModel.StatusText = Core.Resources.Strings.CopyComplete;
            }
        }

        private async void ExportCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var set = _attachedSet;
            if (set == null || set.VisibleRows.Count == 0) return;

            var dialog = new SaveFileDialog
            {
                Filter = Resources.Strings.CsvDialogFilter,
                FileName = Resources.Strings.CsvDefaultFileName,
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                int rows = await set.ExportVisibleRowsAsync(dialog.FileName, CancellationToken.None);
                if (_viewModel != null)
                {
                    _viewModel.StatusText = Core.Resources.Strings.Format(Core.Resources.Strings.ExportComplete, rows);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Resources.Strings.ViewTitle, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
