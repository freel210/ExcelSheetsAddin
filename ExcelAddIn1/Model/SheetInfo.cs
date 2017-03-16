using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    class SheetInfo
    {
        private readonly ThisAddIn model;

        public string SheetName { get; set; }

        private bool _isHidden;
        public bool IsHidden
        {
            get { return _isHidden; }
            set
            {
                if(_isHidden == value) return;

                _isHidden = value;
                //прячем или показываем листы в книге
                if(_isHidden)
                    model.HideSheet(SheetName);
                else
                    model.ShowSheet(SheetName);
            }
        }

        private bool _isVeryHidden;
        public bool IsVeryHidden
        {
            get { return _isVeryHidden; }
            set
            {
                if(_isVeryHidden == value) return;

                _isVeryHidden = value;
                //прячем или показываем листы в книге
                if(_isVeryHidden)
                    model.VeryHideSheet(SheetName);
                else
                    model.HideSheet(SheetName);
            }
        }

        public int SheetIndex { get; set; }

        public bool IsLastVisible { get; set; }

        internal SheetInfo()
        {
            model = Globals.ThisAddIn;
        }

        internal SheetInfo(string sn, bool h, bool vh, int i) : this ()
        {
            SheetName = sn;
            _isHidden = h;
            _isVeryHidden = vh;
            SheetIndex = i;
        }
    }

    internal class PaneInfoRecord
    {
        internal CustomTaskPane Pane { get; set; }
        internal bool IsCheked { get; set; }
    }
}
