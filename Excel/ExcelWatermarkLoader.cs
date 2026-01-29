using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelViewerV2.Model;

namespace ExcelViewerV2.Excel
{
    /// <summary>
    /// Robust loader for header/footer (&G), worksheet drawings, and sheet background images.
    /// Usage: ExcelWatermarkLoader.LoadWatermarks(filePath, model, sheetName);
    /// </summary>
    public static class ExcelWatermarkLoader
    {
        public static void LoadWatermarks(string filePath, SpreadsheetModel model, string sheetName)
        {
            if (!File.Exists(filePath) || model == null)
                return;

            try
            {
                using var doc = SpreadsheetDocument.Open(filePath, false);
                var wbPart = doc.WorkbookPart;

                var sheet = wbPart.Workbook.Sheets
                    .Elements<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                    return;

                var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);

                LoadHeaderFooterImages(wsPart, model);
                LoadDrawingImages(wsPart, model);
                LoadSheetBackground(wsPart, model);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExcelWatermarkLoader] ERROR: {ex}");
            }
        }

        // =========================================================
        // HEADER / FOOTER (&G) – oddzielnie nagłówek i stopka
        // =========================================================
        private static void LoadHeaderFooterImages(WorksheetPart wsPart, SpreadsheetModel model)
        {
            var headerFooter = wsPart.Worksheet.Elements<HeaderFooter>().FirstOrDefault();
            if (headerFooter == null) return;

            var added = new HashSet<string>();

            foreach (var vmlPart in wsPart.VmlDrawingParts)
            {
                try
                {
                    using var sr = new StreamReader(vmlPart.GetStream());
                    string xml = sr.ReadToEnd();

                    // 1️⃣ r:id="rIdX"
                    foreach (Match m in Regex.Matches(xml, "r:id=\"(rId[^\"]+)\""))
                    {
                        var rid = m.Groups[1].Value;
                        if (string.IsNullOrEmpty(rid) || added.Contains(rid)) continue;

                        var imgPart = vmlPart.GetPartById(rid) as ImagePart;
                        if (imgPart != null)
                        {
                            using var s = imgPart.GetStream();
                            var img = LoadBitmap(s);
                            if (img != null)
                            {
                                // Określamy nagłówek lub stopkę
                                bool isHeader = (headerFooter.OddHeader?.Text?.Contains("&G") ?? false) ||
                                                (headerFooter.EvenHeader?.Text?.Contains("&G") ?? false) ||
                                                (headerFooter.FirstHeader?.Text?.Contains("&G") ?? false);

                                var watermark = new WatermarkImage
                                {
                                    Image = img,
                                    Position = HeaderFooterPosition.Unknown,
                                    Type = WatermarkType.HeaderFooter
                                };

                                if (isHeader)
                                    model.HeaderImages.Add(watermark);
                                else
                                    model.FooterImages.Add(watermark);

                                added.Add(rid);
                            }
                        }
                    }

                    // 2️⃣ fallback: obrazy inline w VML
                    foreach (var part in vmlPart.Parts)
                    {
                        if (part.OpenXmlPart is ImagePart ip)
                        {
                            var key = ip.Uri.ToString();
                            if (added.Contains(key)) continue;

                            using var s2 = ip.GetStream();
                            var img2 = LoadBitmap(s2);
                            if (img2 != null)
                            {
                                var watermark = new WatermarkImage
                                {
                                    Image = img2,
                                    Position = HeaderFooterPosition.Unknown,
                                    Type = WatermarkType.HeaderFooter
                                };

                                model.HeaderImages.Add(watermark); // domyślnie nagłówek
                                added.Add(key);
                            }
                        }
                    }
                }
                catch { }
            }

            // 3️⃣ fallback: obrazy jako części arkusza
            foreach (var rel in wsPart.Parts)
            {
                if (rel.OpenXmlPart is ImagePart imgPart)
                {
                    var key = imgPart.Uri.ToString();
                    if (added.Contains(key)) continue;

                    try
                    {
                        using var s = imgPart.GetStream();
                        var img = LoadBitmap(s);
                        if (img != null)
                        {
                            var watermark = new WatermarkImage
                            {
                                Image = img,
                                Position = HeaderFooterPosition.Unknown,
                                Type = WatermarkType.HeaderFooter
                            };

                            model.HeaderImages.Add(watermark);
                            added.Add(key);
                        }
                    }
                    catch { }
                }
            }

            // 4️⃣ fallback: DrawingPart
            try
            {
                var dp = wsPart.DrawingsPart;
                if (dp != null)
                {
                    foreach (var ip in dp.ImageParts)
                    {
                        var key = ip.Uri.ToString();
                        if (added.Contains(key)) continue;

                        using var s = ip.GetStream();
                        var img = LoadBitmap(s);
                        if (img != null)
                        {
                            var watermark = new WatermarkImage
                            {
                                Image = img,
                                Position = HeaderFooterPosition.Unknown,
                                Type = WatermarkType.HeaderFooter
                            };

                            model.HeaderImages.Add(watermark);
                            added.Add(key);
                        }
                    }
                }
            }
            catch { }
        }

        // =========================================================
        // DRAWINGS (obrazy w arkuszu)
        // =========================================================
        private static void LoadDrawingImages(WorksheetPart wsPart, SpreadsheetModel model)
        {
            var drawingsPart = wsPart.DrawingsPart;
            if (drawingsPart == null) return;

            foreach (var imgPart in drawingsPart.ImageParts)
            {
                try
                {
                    using var s = imgPart.GetStream();
                    var img = LoadBitmap(s);
                    if (img != null)
                    {
                        model.WatermarkImages.Add(new WatermarkImage
                        {
                            Image = img,
                            Position = HeaderFooterPosition.Unknown,
                            Type = WatermarkType.Drawing
                        });
                    }
                }
                catch { }
            }
        }

        // =========================================================
        // BACKGROUND (tło arkusza)
        // =========================================================
        private static void LoadSheetBackground(WorksheetPart wsPart, SpreadsheetModel model)
        {
            foreach (var vmlPart in wsPart.VmlDrawingParts)
            {
                try
                {
                    using var sr = new StreamReader(vmlPart.GetStream());
                    var xml = XDocument.Load(sr);
                    XNamespace v = "urn:schemas-microsoft-com:vml";
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                    foreach (var fill in xml.Descendants(v + "fill"))
                    {
                        var ridAttr = fill.Attribute(r + "id");
                        if (ridAttr == null) continue;

                        try
                        {
                            var imgPart = vmlPart.GetPartById(ridAttr.Value) as ImagePart;
                            if (imgPart != null)
                            {
                                using var s = imgPart.GetStream();
                                var img = LoadBitmap(s);
                                if (img != null)
                                {
                                    model.WatermarkImages.Add(new WatermarkImage
                                    {
                                        Image = img,
                                        Position = HeaderFooterPosition.Unknown,
                                        Type = WatermarkType.Background
                                    });
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }

        // =========================================================
        // Bitmap → WPF ImageSource
        // =========================================================
        private static ImageSource LoadBitmap(Stream stream)
        {
            try
            {
                using var mem = new MemoryStream();
                stream.CopyTo(mem);
                mem.Position = 0;

                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = mem;
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch
            {
                return null;
            }
        }
    }
}
