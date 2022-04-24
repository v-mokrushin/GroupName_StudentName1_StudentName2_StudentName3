using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Threading;
using System.Data;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Project
{

    public class ApplicationCore
    {

        public MainWindow window; // ссылка на экземпляр окна

        //////////////////////////////////////////////

        public void SetWindowReference(MainWindow window)
        {
            this.window = window;
        }

        public void InitializeObjects()
        {
            window.xappearance = new InterfaceАppearance();
            window.xappearance.SetWindowReference(window);
            window.xappearance.ApplyDarkTheme();

            window.xinterface = new InterfaceController();
            window.xinterface.SetWindowReference(window);
            window.xinterface.SetAppearanceBuilderReference(window.xappearance);
            window.xinterface.Initialize();
        }

        public void RunProgram()
        {
            window.xinterface.OpenStartWindow();
        }

        //////////////////////////////////////////////
              
    }
}
