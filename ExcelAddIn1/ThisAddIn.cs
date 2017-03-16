using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        internal Dictionary<Excel.Workbook, PaneInfoRecord> PaneDictionary { get; private set; }

        public event System.Action ChangeSheetVisibility;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            PaneDictionary = new Dictionary<Excel.Workbook, PaneInfoRecord>();

            //подписываемся на событие создания новой книги
            ((Excel.AppEvents_Event)Application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(TaskPaneCreate);

            //подписываемся на событие открытия книги
            Application.WorkbookOpen += TaskPaneCreate;

            //для созданной по умолчанию первой книги запускаем создание панели вручную
            //if(Application.ActiveWorkbook != null)
            TaskPaneCreate(Application.ActiveWorkbook);

            //подписываемся на событие при закрытии книги
            Application.WorkbookBeforeClose += ThisAddIn_WorkbookBeforeClose;
        }

        private void ThisAddIn_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            //удаляем из словаря ссылку на книгу
            if(PaneDictionary.ContainsKey(Wb))
                PaneDictionary.Remove(Wb);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private void TaskPaneCreate(Excel.Workbook wb)
        {
            try
            {                
                //создаем панель
                CustomTaskPane pane = CustomTaskPanes.Add(new PaneControl(), "Листы книги");

                //сохраняем в каталоге ссылку на панель и статус отображения кноки
                if (wb != null)
                    PaneDictionary.Add(wb, new PaneInfoRecord
                    {
                        Pane = pane,
                        IsCheked = false
                    });
            }
            catch(Exception ex)
            {
                MessageBox.Show(string.Format("TargetSite = {0}\n\n HelpLink = {1}\n\n StackTrace = {2}\n\n Source = {3}"
                    ,ex.TargetSite.ToString(), ex.HelpLink, ex.StackTrace, ex.Source));
            }
        }

        internal List<SheetInfo> GetSheetList()
        {
            List<SheetInfo> lst = new List<SheetInfo>();

            //считаем количество листов в книге
            int listsCount;
            try
            {
                //листов в книге может еще не быть
                listsCount = Application.Sheets.Count;
            }
            catch
            {
                listsCount = 0;
            }
            
            //заполняем список
            for(int i = 1; i <= listsCount; i++)
            {
                bool isHidden, isVeryHidden;

                switch ((Application.Sheets[i] as Excel.Worksheet).Visible)
                {

                    case Excel.XlSheetVisibility.xlSheetHidden:
                        isHidden = true;
                        isVeryHidden = false;
                        break;

                    case Excel.XlSheetVisibility.xlSheetVeryHidden:
                        isHidden = false;
                        isVeryHidden = true;
                        break;

                    case Excel.XlSheetVisibility.xlSheetVisible:
                    default:
                        isHidden = false;
                        isVeryHidden = false;
                        break;
                }

                string SheetName = (Application.Sheets[i] as Excel.Worksheet).Name;
                int sheetIndex = i;

                lst.Add(new SheetInfo(SheetName, isHidden, isVeryHidden, sheetIndex));
            }

            //ищем и отмечаем (если нашли) последний нескрытый лист
            var c = from item in lst
                    where item.IsHidden == false && item.IsVeryHidden == false
                    select item;

            if(c.ToList().Count == 1)
                c.First().IsLastVisible = true;

            return lst;
        }

        internal void AddSheet()
        {
            Application.Sheets.Add(After: Application.ActiveSheet);
        }

        internal void HideSheet(string sheetName)
        {
            (Application.Sheets[sheetName] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetHidden;

            if(ChangeSheetVisibility != null)
                ChangeSheetVisibility();
        }

        internal void HideAllSheets()
        {
            for(int i = 1; i <= Application.Sheets.Count; i++)
            {
                bool condition1 = (Application.Sheets[i] as Excel.Worksheet) == Application.ActiveSheet;
                bool condition2 = (Application.Sheets[i] as Excel.Worksheet).Visible == Excel.XlSheetVisibility.xlSheetHidden;

                if(condition1 || condition2) continue;

                (Application.Sheets[i] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        internal void VeryHideSheet(string sheetName)
        {
            (Application.Sheets[sheetName] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;

            if(ChangeSheetVisibility != null)
                ChangeSheetVisibility();
        }

        internal void VeryHideAllSheets()
        {
            for(int i = 1; i <= Application.Sheets.Count; i++)
            {
                bool condition1 = (Application.Sheets[i] as Excel.Worksheet) == Application.ActiveSheet;
                bool condition2 = (Application.Sheets[i] as Excel.Worksheet).Visible == Excel.XlSheetVisibility.xlSheetVeryHidden;

                if(condition1 || condition2) continue;

                (Application.Sheets[i] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            }
        }

        internal void ShowSheet(string sheetName)
        {
            (Application.Sheets[sheetName] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetVisible;

            if(ChangeSheetVisibility != null)
                ChangeSheetVisibility();
        }

        internal void ShowAllSheets()
        {
            for(int i = 1; i <= Application.Sheets.Count; i++)
            {
                bool condition1 = (Application.Sheets[i] as Excel.Worksheet) == Application.ActiveSheet;
                bool condition2 = (Application.Sheets[i] as Excel.Worksheet).Visible == Excel.XlSheetVisibility.xlSheetVisible;

                if(condition1 || condition2) continue;

                (Application.Sheets[i] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
        }

        internal void DeleteSheet(string sheetName)
        {
            bool condition = (Application.Sheets[sheetName] as Excel.Worksheet).Visible == Excel.XlSheetVisibility.xlSheetVeryHidden;

            //если удалять сильно скрытый лист - Excel кидает исключение
            if(condition)
                (Application.Sheets[sheetName] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetHidden;

            (Application.Sheets[sheetName] as Excel.Worksheet).Delete();
        }

        internal void DeleteAllSheets()
        {
            List<string> dlNames = new List<string>();

            //переписываем имена листов поделажщих удалению
            for(int i = 1; i <= Application.Sheets.Count; i++)
            {
                bool condition = (Application.Sheets[i] as Excel.Worksheet) == Application.ActiveSheet;

                if(condition) continue;

                string name = (Application.Sheets[i] as Excel.Worksheet).Name;
                dlNames.Add(name);
            }

            //удаляем листы поименно
            foreach(string item in dlNames)
            {
                bool condition = (Application.Sheets[item] as Excel.Worksheet).Visible == Excel.XlSheetVisibility.xlSheetVeryHidden;

                //если удалять сильно скрытый лист - Excel кидает исключение
                if(condition)
                    (Application.Sheets[item] as Excel.Worksheet).Visible = Excel.XlSheetVisibility.xlSheetHidden;

                (Application.Sheets[item] as Excel.Worksheet).Delete();
            }
        }
    }
}
