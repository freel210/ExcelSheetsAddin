using System.Collections.Generic;
using System.Windows.Input;
using System.Linq;

namespace ExcelAddIn1.ViewModel
{
    class PaneWPFControlViewModel : NotifyPropertyChanger
    {
        private const string HideButtonLabelA = "Скрыть";
        private const string HideButtonLabelB = "Скрыть все";

        private const string VeryHideButtonLabelA = "Сильно скрыть";
        private const string VeryHideButtonLabelB = "Сильно скрыть все";

        private const string ShowButtonLabelA = "Отобразить";
        private const string ShowButtonLabelB = "Отобразить все";

        private const string DeleteButtonLabelA = "Удалить";
        private const string DeleteButtonLabelB = "Удалить все";

        private readonly ThisAddIn _model;

        private readonly RelayCommand _addSheetCommand;
        private readonly RelayCommand _hideSheetCommand;
        private readonly RelayCommand _veryHideSheetCommand;
        private readonly RelayCommand _showSheetCommand;
        private readonly RelayCommand _deleteSheetCommand;

        private string _hideButtonLabel = HideButtonLabelB;
        public string HideButtonLabel
        {
            get { return _hideButtonLabel; }
            set
            {
                if(_hideButtonLabel == value) return;
                _hideButtonLabel = value;
                OnPropertyChanged();
            }
        }

        private string _veryHideButtonLabel = VeryHideButtonLabelB;
        public string VeryHideButtonLabel
        {
            get { return _veryHideButtonLabel; }
            set
            {
                if(_veryHideButtonLabel == value) return;
                _veryHideButtonLabel = value;
                OnPropertyChanged();
            }
        }

        private string _showButtonLabel = ShowButtonLabelB;
        public string ShowButtonLabel
        {
            get { return _showButtonLabel; }
            set
            {
                if(_showButtonLabel == value) return;
                _showButtonLabel = value;
                OnPropertyChanged();
            }
        }

        private string _deleteButtonLabel = DeleteButtonLabelB;
        public string DeleteButtonLabel
        {
            get { return _deleteButtonLabel; }
            set
            {
                if(_deleteButtonLabel == value) return;
                _deleteButtonLabel = value;
                OnPropertyChanged();
            }
        }

        private List<SheetInfo> _lst;
        public List<SheetInfo> lst
        {
            get { return _lst; }
            set
            {
                if(_lst == value) return;
                _lst = value;
                OnPropertyChanged();
            }
        }

        public List<SheetInfo> SelectedSheets { get; set; }

        public ICommand AddSheetCommand
        {
            get { return _addSheetCommand; }
        }

        public ICommand HideSheetCommand
        {
            get { return _hideSheetCommand; }
        }

        public ICommand VeryHideSheetCommand
        {
            get { return _veryHideSheetCommand; }
        }

        public ICommand ShowSheetCommand
        {
            get { return _showSheetCommand; }
        }

        public ICommand DeleteSheetCommand
        {
            get { return _deleteSheetCommand; }
        }

        internal PaneWPFControlViewModel()
        {
            _model = Globals.ThisAddIn;
            lst = _model.GetSheetList();
            SelectedSheets = new List<SheetInfo>();

            _addSheetCommand = new RelayCommand(o => AddSheet(), o => true);
            _hideSheetCommand = new RelayCommand(o => HideSheet(), o => CanHideSheet());
            _veryHideSheetCommand = new RelayCommand(o => VeryHideSheet(), o => CanVeryHideSheet());
            _showSheetCommand = new RelayCommand(o => ShowSheet(), o => CanShowSheet());
            _deleteSheetCommand = new RelayCommand(o => DeleteSheet(), o => CanDeleteSheet());

            //событие возникает как следствие добавления, удаления, скрытия и отображение листов
            _model.Application.SheetActivate += (s) => lst = _model.GetSheetList();

            ////костыль для отслеживания перетаскивания листов в Excel
            //_model.Application.SheetChange += (s, o) => lst = _model.GetSheetList();

            //событие возникает при переключении на другую книгу
            _model.Application.WorkbookActivate += (s) => lst = _model.GetSheetList();

            //костыль для отслеживания переименовывания листа
            _model.Application.SheetSelectionChange += Application_SheetSelectionChange;

            //событие возникает когда "галкой" меняют видимость листа
            _model.ChangeSheetVisibility += () => lst = _model.GetSheetList();
        }

        private void Application_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            _model.GetSheetList();
        }

        internal void AddSheet()
        {
            _model.AddSheet();
            lst = _model.GetSheetList();
        }

        internal void HideSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                foreach(var item in SelectedSheets)
                {
                    if(!item.IsHidden)
                        _model.HideSheet(item.SheetName);
                }
            }
            else
            {
                _model.HideAllSheets();
            }

            lst = _model.GetSheetList();
        }

        internal void VeryHideSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                foreach(var item in SelectedSheets)
                {
                    if(!item.IsVeryHidden)
                        _model.VeryHideSheet(item.SheetName);
                }
            }
            else
            {
                _model.VeryHideAllSheets();
            }

            lst = _model.GetSheetList();
        }

        internal void ShowSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                foreach(var item in SelectedSheets)
                {
                    if(item.IsHidden || item.IsVeryHidden)
                        _model.ShowSheet(item.SheetName);
                }
            }
            else
            {
                _model.ShowAllSheets();
            }

            lst = _model.GetSheetList();
        }

        internal void DeleteSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                foreach(var item in SelectedSheets)
                    _model.DeleteSheet(item.SheetName);
            }
            else
            {
                _model.DeleteAllSheets();
            }

            lst = _model.GetSheetList();
        }

        internal bool CanHideSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                int notHiddenSelectedCount = (from item in SelectedSheets
                                              where item.IsHidden == false
                                              select item).Count();

                if(notHiddenSelectedCount > 0 && NotHiddenNotSelectedCount() > 0) return true;

                return false;
            }
            else
            {
                //если есть еще не скрытые листы, помимо активного
                if(NotHiddenNotSelectedCount() > 1) return true;

                return false;
            }
        }
        internal bool CanVeryHideSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                int notVeryHiddenSelectedCount = (from item in SelectedSheets
                                                  where item.IsVeryHidden == false
                                                  select item).Count();


                if(notVeryHiddenSelectedCount > 0 && NotHiddenNotSelectedCount() > 0) return true;

                return false;
            }
            else
            {
                int notVeryHiddenNotSelected = (from item in lst.Except(SelectedSheets)
                                                where item.IsVeryHidden == false
                                                select item).Count();

                //если есть еще не сильно скрытые листы, помимо активного
                if(notVeryHiddenNotSelected > 1) return true;

                return false;
            }
        }

        internal bool CanShowSheet()
        {
            if(SelectedSheets.Count != 0)
            {
                int hiddenSelectedCount = (from item in SelectedSheets
                                           where item.IsHidden == true
                                           select item).Count();

                int veryHiddenSelectedCount = (from item in SelectedSheets
                                               where item.IsVeryHidden == true
                                               select item).Count();

                if(hiddenSelectedCount > 0 || veryHiddenSelectedCount > 0) return true;

                return false;
            }
            else
            {
                int notVisibleNotSelected = (from item in lst
                                             where item.IsHidden == true || item.IsVeryHidden == true
                                             select item).Count();

                //если есть скрытие любым образом листы
                if(notVisibleNotSelected > 0) return true;

                return false;
            }

        }

        internal bool CanDeleteSheet()
        {
            if(SelectedSheets.Count() > 0 && NotHiddenNotSelectedCount() > 0) return true;

            //если есть еще листы помимо активного
            if(lst.Count() > 1) return true;

            return false;
        }

        private int NotHiddenNotSelectedCount()
        {
            return (from item in lst.Except(SelectedSheets)
                    where item.IsHidden == false
                    select item).Count();
        }

        internal void ChangeButtonsNames()
        {
            if (SelectedSheets.Count != 0)
            {
                HideButtonLabel     = HideButtonLabelA;
                VeryHideButtonLabel = VeryHideButtonLabelA;
                ShowButtonLabel     = ShowButtonLabelA;
                DeleteButtonLabel   = DeleteButtonLabelA;
            }
            else
            {
                HideButtonLabel     = HideButtonLabelB;
                VeryHideButtonLabel = VeryHideButtonLabelB;
                ShowButtonLabel     = ShowButtonLabelB;
                DeleteButtonLabel   = DeleteButtonLabelB;
            }
        }
    }
}
