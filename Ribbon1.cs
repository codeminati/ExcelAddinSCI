using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddinSCI
{
    public partial class Ribbon1
    {
        private Dictionary<Excel.Range, string> _originalValues = new Dictionary<Excel.Range, string>();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Ribbon loaded!");
        }

        private void ConvertToAlphanumeric_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null) return;

            foreach (Excel.Range cell in selectedRange.Cells)
            {
                string original = Convert.ToString(cell.Value2);
                if (string.IsNullOrEmpty(original)) continue;

                if (!_originalValues.ContainsKey(cell))
                    _originalValues[cell] = original;

                string sanitized = Regex.Replace(original, "[^a-zA-Z0-9]", "");
                cell.Value2 = sanitized;
            }
        }

        private void RevertToOriginal_Click(object sender, RibbonControlEventArgs e)
        {
            foreach (var kvp in _originalValues)
            {
                kvp.Key.Value2 = kvp.Value;
            }
            _originalValues.Clear();
        }
    }
}
