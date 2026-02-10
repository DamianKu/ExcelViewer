using System;

namespace ExcelViewerV2.Model
{
    public static class CellAddress
    {
        public static (int row, int col) Parse(string addr)
        {
            int i = 0;
            while (i < addr.Length && char.IsLetter(addr[i])) i++;

            string colPart = addr[..i].ToUpper();
            string rowPart = addr[i..];

            int col = 0;
            foreach (char c in colPart)
                col = col * 26 + (c - 'A' + 1);

            col--;

            int row = int.Parse(rowPart) - 1;

            return (row, col);
        }

        public static string ToName(int row, int col)
        {
            col++;
            string name = "";

            while (col > 0)
            {
                int r = (col - 1) % 26;
                name = (char)('A' + r) + name;
                col = (col - 1) / 26;
            }

            return name + (row + 1);
        }
    }
}
