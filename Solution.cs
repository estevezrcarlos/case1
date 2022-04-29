using Microsoft.Office.Interop.Excel;

namespace case1
{
    internal class Solution
    {
        private Workbook excelFile;
        private Application excelApp;
        public Solution(string filePath)
        {
            excelApp = new Application();
            excelFile = excelApp.Workbooks.Open(filePath);
        }
        
        public void sortSheetThreeColumnAscending(string sheetName)
        {
            Worksheet excelSheet = (Worksheet)excelFile.Sheets[sheetName];
            var usedRange = excelSheet.UsedRange;
            var rgKey1 = excelSheet.Columns["A:A"];
            var rgKey2 = excelSheet.Columns["B:B"];
            var rgKey3 = excelSheet.Columns["C:C"];

            usedRange.Sort(Key1: rgKey1, Order1: XlSortOrder.xlAscending, Key2: rgKey2, Order2: XlSortOrder.xlAscending, Key3: rgKey3, Order3: XlSortOrder.xlAscending, Header: XlYesNoGuess.xlYes);
        }

        public void saveAndClose()
        {
            excelFile.Save();
            excelFile.Close(0);
            excelApp.Quit();
        }
    }
}
