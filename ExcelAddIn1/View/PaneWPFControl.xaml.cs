using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelAddIn1.ViewModel;

namespace ExcelAddIn1.View
{
    /// <summary>
    /// Логика взаимодействия для PaneWPFControl.xaml
    /// </summary>
    public partial class PaneWPFControl : UserControl
    {
        PaneWPFControlViewModel vm;
        public PaneWPFControl()
        {
            InitializeComponent();

            vm = new PaneWPFControlViewModel();
            DataContext = vm;
        }

        private void lb_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            vm.SelectedSheets = lb.SelectedItems.Cast<SheetInfo>().ToList();
            vm.ChangeButtonsNames();
        }
    }
}
