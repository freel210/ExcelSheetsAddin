using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private Excel.Workbook _activeBook;

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if(tb1.Checked)
            {
                Globals.ThisAddIn.PaneDictionary[_activeBook].IsCheked = true;
                Globals.ThisAddIn.PaneDictionary[_activeBook].Pane.Visible = true;
            }
            else
            {
                Globals.ThisAddIn.PaneDictionary[_activeBook].IsCheked = false;
                Globals.ThisAddIn.PaneDictionary[_activeBook].Pane.Visible = false;
            }
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _activeBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActivate;
        }

        private void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            _activeBook = Wb;
            tb1.Checked = Globals.ThisAddIn.PaneDictionary[_activeBook].IsCheked;
        }
    }
}
