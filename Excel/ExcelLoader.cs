using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClosedXML.Excel;
using ExcelViewerV2.Model;
using ExcelViewerV2.Utils;

namespace ExcelViewerV2.Excel
{
    public sealed class ExcelLoader
    {
        private XLWorkbook _workbook;
        private readonly Logger _logger = new Logger();
        private readonly Dictionary<System.Drawing.Color, SolidColorBrush> _brushCache = new();
        private readonly Dictionary<string, SpreadsheetModel> _sheetModels = new();
        private string _currentPath;

        // =============================
        // Workbook – NEW / LOAD / SAVE
        // =============================

        public bool HasWorkbook => _workbook != null;
        public string CurrentPath => _currentPath;

        /// <summary>
        /// NOWY DOKUMENT
        /// </summary>
        public void CreateNewWorkbook()
        {
            _workbook = new XLWorkbook();
            _sheetModels.Clear();
            _brushCache.Clear();

            // Excel zawsze tworzy 1 arkusz
            var ws = _workbook.AddWorksheet("Arkusz1");

            // minimalny rozmiar 50x26 (A–Z)
            var model = new SpreadsheetModel(50, 26)
            {
                Name = ws.Name,
                IsVisible = true
            };

            // domyślne szerokości / wysokości
            for (int c = 0; c < model.Columns; c++)
                model.ColumnWidths[c] = 80;

            for (int r = 0; r < model.Rows; r++)
                model.RowHeights[r] = 20;

            _sheetModels[ws.Name] = model;

            _currentPath = null;
        }

        /// <summary>
        /// OTWÓRZ
        /// </summary>
        public void LoadWorkbook(string path)
        {
            if (!File.Exists(path))
            {
                _logger.Error($"[LoadWorkbook] Plik nie istnieje: {path}");
                return;
            }

            try
            {
                _workbook = new XLWorkbook(path);
                _currentPath = path;

                _sheetModels.Clear();
                _brushCache.Clear();

                _logger.Info($"[LoadWorkbook] Wczytano plik: {path}");
            }
            catch (Exception ex)
            {
                _logger.Error($"[LoadWorkbook] Błąd: {ex.Message}");
            }
        }

        /// <summary>
        /// ZAPISZ
        /// </summary>
        public bool Save()
        {
            if (_workbook == null)
            {
                _logger.Error("[Save] Brak workbooka");
                return false;
            }

            if (string.IsNullOrWhiteSpace(_currentPath))
            {
                _logger.Error("[Save] Brak ścieżki – użyj SaveAs()");
                return false;
            }

            try
            {
                this.UpdateWorkbookFromModels(); // synchronizacja modeli z workbookiem
                _workbook.SaveAs(_currentPath);
                _logger.Info($"[Save] Zapisano: {_currentPath}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"[Save] {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ZAPISZ JAKO
        /// </summary>
        public bool SaveAs(string path)
        {
            if (_workbook == null)
            {
                _logger.Error("[SaveAs] Brak workbooka");
                return false;
            }

            try
            {
                this.UpdateWorkbookFromModels(); // synchronizacja modeli z workbookiem
                _workbook.SaveAs(path);
                _currentPath = path;
                _logger.Info($"[SaveAs] Zapisano jako: {path}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"[SaveAs] {ex.Message}");
                return false;
            }
        }

        // =============================
        // Sheets
        // =============================

        public void LoadAllSheets()
        {
            if (_workbook == null)
            {
                _logger.Error("[LoadAllSheets] Workbook == null");
                return;
            }

            _sheetModels.Clear();

            foreach (var ws in _workbook.Worksheets)
                LoadSheet(ws.Name);
        }

        public SpreadsheetModel GetSheetModel(string sheetName) =>
            _sheetModels.TryGetValue(sheetName, out var model) ? model : null;

        public Dictionary<string, SpreadsheetModel> GetAllSheetModels() => _sheetModels;

        // =============================
        // Sheet loading
        // =============================

        private void LoadSheet(string sheetName)
        {
            try
            {
                var ws = _workbook.Worksheet(sheetName);
                if (ws == null) return;

                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

                if (lastRow == 0 || lastCol == 0)
                {
                    _logger.Info($"[LoadSheet] Arkusz '{sheetName}' pusty");
                    return;
                }

                var model = new SpreadsheetModel(lastRow, lastCol)
                {
                    Name = sheetName
                };

                // Kolumny
                for (int c = 1; c <= lastCol; c++)
                    model.ColumnWidths[c - 1] = ExcelColumnWidthToPixels(ws.Column(c).Width);

                // Wiersze
                for (int r = 1; r <= lastRow; r++)
                    model.RowHeights[r - 1] = ws.Row(r).Height;

                // Komórki
                for (int r = 1; r <= lastRow; r++)
                {
                    for (int c = 1; c <= lastCol; c++)
                    {
                        var xlCell = ws.Cell(r, c);
                        var cell = model[r - 1, c - 1];

                        cell.Text = xlCell.GetFormattedString();
                        cell.FontFamilyName = string.IsNullOrWhiteSpace(xlCell.Style.Font.FontName)
                            ? "Calibri"
                            : xlCell.Style.Font.FontName;

                        cell.FontSize = xlCell.Style.Font.FontSize > 0
                            ? xlCell.Style.Font.FontSize
                            : 14;

                        cell.Bold = xlCell.Style.Font.Bold;
                        cell.Italic = xlCell.Style.Font.Italic;
                        cell.Underline = xlCell.Style.Font.Underline != XLFontUnderlineValues.None;
                        cell.Strikeout = xlCell.Style.Font.Strikethrough;

                        cell.Foreground = GetBrush(xlCell.Style.Font.FontColor, false);
                        cell.Background = GetBrush(xlCell.Style.Fill.BackgroundColor, true);

                        cell.Alignment = ConvertAlignment(xlCell.Style.Alignment.Horizontal);
                        cell.VerticalAlignment = ConvertVerticalAlignment(xlCell.Style.Alignment.Vertical);
                        cell.Wrapping = xlCell.Style.Alignment.WrapText
                            ? System.Windows.TextWrapping.Wrap
                            : System.Windows.TextWrapping.NoWrap;

                        cell.TextRotation = ConvertTextRotation(xlCell.Style.Alignment.TextRotation);

                        // Hyperlink
                        if (xlCell.HasHyperlink)
                        {
                            try
                            {
                                var link = xlCell.GetHyperlink();
                                string raw = link?.ExternalAddress?.ToString() ?? link?.InternalAddress;

                                if (!string.IsNullOrWhiteSpace(raw))
                                {
                                    string resolved = HyperlinkResolver.Resolve(_currentPath, raw);
                                    if (!string.IsNullOrEmpty(resolved) && File.Exists(resolved))
                                    {
                                        cell.Hyperlink = new Uri(resolved);
                                        cell.Foreground = Brushes.Blue;
                                        cell.Underline = true;
                                    }
                                }
                            }
                            catch (Exception exLink)
                            {
                                _logger.Error($"[Hyperlink] {exLink.Message}");
                            }
                        }
                    }
                }

                // Watermarki
                model.WatermarkImages.Clear();
                ExcelWatermarkLoader.LoadWatermarks(_currentPath, model, ws.Name);

                foreach (var pic in ws.Pictures)
                {
                    try
                    {
                        pic.ImageStream.Position = 0;
                        var img = LoadImageFromStream(pic.ImageStream);
                        if (img != null)
                        {
                            model.WatermarkImages.Add(new WatermarkImage
                            {
                                Image = img,
                                Position = HeaderFooterPosition.Center,
                                Type = WatermarkType.Drawing
                            });
                        }
                    }
                    catch (Exception exPic)
                    {
                        _logger.Error($"[Picture] {exPic.Message}");
                    }
                }

                _sheetModels[sheetName] = model;
                _logger.Info($"[LoadSheet] '{sheetName}' OK ({lastRow}x{lastCol})");
            }
            catch (Exception ex)
            {
                _logger.Error($"[LoadSheet] {sheetName}: {ex}");
            }
        }

        // =============================
        // Helpers
        // =============================

        private double ExcelColumnWidthToPixels(double width) =>
            Math.Max(20, width * 7);

        private SolidColorBrush GetBrush(XLColor color, bool background)
        {
            if (color.ColorType == XLColorType.Color)
            {
                var c = color.Color;
                if (!_brushCache.TryGetValue(c, out var brush))
                {
                    brush = new SolidColorBrush(Color.FromArgb(c.A, c.R, c.G, c.B));
                    brush.Freeze();
                    _brushCache[c] = brush;
                }
                return brush;
            }

            return background ? Brushes.White : Brushes.Black;
        }

        private System.Windows.TextAlignment ConvertAlignment(XLAlignmentHorizontalValues v) =>
            v switch
            {
                XLAlignmentHorizontalValues.Center => System.Windows.TextAlignment.Center,
                XLAlignmentHorizontalValues.Right => System.Windows.TextAlignment.Right,
                _ => System.Windows.TextAlignment.Left
            };

        private System.Windows.VerticalAlignment ConvertVerticalAlignment(XLAlignmentVerticalValues v) =>
            v switch
            {
                XLAlignmentVerticalValues.Top => System.Windows.VerticalAlignment.Top,
                XLAlignmentVerticalValues.Bottom => System.Windows.VerticalAlignment.Bottom,
                _ => System.Windows.VerticalAlignment.Center
            };

        private double ConvertTextRotation(int excelRotation) =>
            excelRotation == 255 ? -90 : excelRotation;

        private ImageSource LoadImageFromStream(Stream stream)
        {
            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = stream;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        // =============================
        // Synchronizacja modeli z workbookiem
        // =============================
        public void UpdateWorkbookFromModels()
        {
            if (_workbook == null) return;

            foreach (var kvp in _sheetModels)
            {
                string sheetName = kvp.Key;
                var model = kvp.Value;

                // Pobierz istniejący arkusz lub utwórz nowy
                var ws = _workbook.Worksheet(sheetName) ?? _workbook.AddWorksheet(sheetName);

                for (int r = 0; r < model.Rows; r++)
                {
                    for (int c = 0; c < model.Columns; c++)
                    {
                        var cellModel = model[r, c];
                        var xlCell = ws.Cell(r + 1, c + 1);
                        xlCell.Value = cellModel.Text;

                        // zachowaj formatowanie
                        xlCell.Style.Font.FontName = cellModel.FontFamilyName ?? "Calibri";
                        xlCell.Style.Font.FontSize = cellModel.FontSize > 0 ? cellModel.FontSize : 14;
                        xlCell.Style.Font.Bold = cellModel.Bold;
                        xlCell.Style.Font.Italic = cellModel.Italic;
                        xlCell.Style.Font.Underline = cellModel.Underline ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
                        xlCell.Style.Font.Strikethrough = cellModel.Strikeout;
                        xlCell.Style.Alignment.Horizontal = cellModel.Alignment switch
                        {
                            System.Windows.TextAlignment.Center => XLAlignmentHorizontalValues.Center,
                            System.Windows.TextAlignment.Right => XLAlignmentHorizontalValues.Right,
                            _ => XLAlignmentHorizontalValues.Left
                        };
                        xlCell.Style.Alignment.Vertical = cellModel.VerticalAlignment switch
                        {
                            System.Windows.VerticalAlignment.Top => XLAlignmentVerticalValues.Top,
                            System.Windows.VerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
                            _ => XLAlignmentVerticalValues.Center
                        };
                        xlCell.Style.Alignment.WrapText = cellModel.Wrapping == System.Windows.TextWrapping.Wrap;
                        xlCell.Style.Alignment.TextRotation = (int)cellModel.TextRotation;
                    }
                }
            }
        }
    }

    // =============================
    // Logger
    // =============================
    public sealed class Logger
    {
        public void Info(string msg) => Console.WriteLine(msg);
        public void Error(string msg) => Console.WriteLine(msg);
    }
}
