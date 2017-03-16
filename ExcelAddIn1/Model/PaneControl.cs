using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelAddIn1.View;

namespace ExcelAddIn1
{
    public partial class PaneControl : UserControl
    {
        public PaneControl()
        {
            InitializeComponent();

            //создаем хост-элемент и заполняем им всю форму
            ElementHost host = new ElementHost();
            host.Dock = DockStyle.Fill;

            //создаем пользовательский WPF-контрол, помещаем его в хост и добавляем хост в коллекцию контролов формы
            PaneWPFControl c = new PaneWPFControl();
            host.Child = c;
            this.Controls.Add(host);
        }
    }
}
