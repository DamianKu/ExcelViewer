using System;
using System.IO;

namespace ExcelViewerV2.Utils
{
    public static class HyperlinkResolver
    {
        public static string Resolve(string excelFilePath, string hyperlink)
        {
            if (string.IsNullOrWhiteSpace(hyperlink))
                return hyperlink;

            hyperlink = hyperlink.Replace('/', '\\');

            // ============================
            // 1. URL (http / https / mailto)
            // ============================
            if (Uri.TryCreate(hyperlink, UriKind.Absolute, out var uri)
                && uri.Scheme != Uri.UriSchemeFile)
            {
                return hyperlink;
            }

            // ============================
            // 2. Absolutna ścieżka
            // ============================
            if (Path.IsPathRooted(hyperlink) && File.Exists(hyperlink))
                return hyperlink;

            string excelDir = Path.GetDirectoryName(excelFilePath)!;

            // ============================
            // 3. Względna względem Excela
            // ============================
            string fromExcel = Path.GetFullPath(
                Path.Combine(excelDir, hyperlink)
            );

            if (File.Exists(fromExcel))
                return fromExcel;

            // ============================
            // 4. FALLBACK → Public\Rysunki\SC
            // ============================
            string fileName = Path.GetFileName(hyperlink);
            string publicRoot = FindPublicRoot(excelDir);

            if (publicRoot != null)
            {
                return Path.Combine(
                    publicRoot,
                    "Rysunki",
                    "SC",
                    fileName
                );
            }

            // ============================
            // 5. Ostatecznie to co mamy
            // ============================
            return fromExcel;
        }

        private static string FindPublicRoot(string startDir)
        {
            var dir = new DirectoryInfo(startDir);

            while (dir != null)
            {
                if (dir.Name.Equals("Public", StringComparison.OrdinalIgnoreCase))
                    return dir.FullName;

                dir = dir.Parent;
            }

            return null;
        }
    }
}
