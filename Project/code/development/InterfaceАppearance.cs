using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Controls;

namespace Project
{

    public class InterfaceАppearance
    {

        private MainWindow window; // ссылка на экземпляр окна
        private Theme actualTheme;

        private SolidColorBrush mouseEnterColor;
        private SolidColorBrush mouseLeaveColor;

        //////////////////////////////////////////////

        public void SetWindowReference(MainWindow window)
        {
            this.window = window;
        }

        public void SwitchTheme()
        {
            if (actualTheme == Theme.Light)
            {
                ApplyDarkTheme();
                return;
            }
            if(actualTheme == Theme.Dark)
            {
                ApplyLightTheme();
                return;
            }
        }


        public SolidColorBrush GetMouseEnterColor()
        {
            return mouseEnterColor;
        }

        public SolidColorBrush GetMouseLeaveColor()
        {
            return mouseLeaveColor;
        }

        //////////////////////////////////////////////

        public void ApplyDarkTheme()
        {
            actualTheme = Theme.Dark;

            window.Resources["ColorGridBackground"] = new SolidColorBrush(Color.FromRgb(30, 30, 40));               // #1e1e28
            window.Resources["ColorMainFont"] = new SolidColorBrush(Color.FromRgb(241, 241, 241));                  // #f1f1f1
            window.Resources["ColorBlockBackground"] = new SolidColorBrush(Color.FromRgb(50, 50, 60));           // #eeebeb 
            window.Resources["ColorBlockBackgroundMouseOn"] = new SolidColorBrush(Color.FromRgb(60, 60, 70));    // #c9def5

            mouseEnterColor = new SolidColorBrush(Color.FromRgb(60, 60, 70));
            mouseLeaveColor = new SolidColorBrush(Color.FromRgb(50, 50, 60));

            // BAG SOLUTION
            window.gridChangeTheme.Background = new SolidColorBrush(Color.FromRgb(60, 60, 70));
            window.gridStartOpenSolition.Background = new SolidColorBrush(Color.FromRgb(50, 50, 60));
            window.gridStartCreateSolition.Background = new SolidColorBrush(Color.FromRgb(50, 50, 60));
            window.gridStartHelp.Background = new SolidColorBrush(Color.FromRgb(50, 50, 60));
            window.gridStartExitProgram.Background = new SolidColorBrush(Color.FromRgb(50, 50, 60));
        }

        public void ApplyLightTheme()
        {
            actualTheme = Theme.Light;

            window.Resources["ColorGridBackground"] = new SolidColorBrush(Color.FromRgb(251, 251, 251));            // #fbfbfb
            window.Resources["ColorMainFont"] = new SolidColorBrush(Color.FromRgb(10, 10, 10));                     // #0a0a0a 
            window.Resources["ColorBlockBackground"] = new SolidColorBrush(Color.FromRgb(238, 235, 235));           // #eeebeb 
            window.Resources["ColorBlockBackgroundMouseOn"] = new SolidColorBrush(Color.FromRgb(201, 222, 245));    // #c9def5

            mouseEnterColor = new SolidColorBrush(Color.FromRgb(201, 222, 245));
            mouseLeaveColor = new SolidColorBrush(Color.FromRgb(238, 235, 235));

            // BAG SOLUTION
            window.gridChangeTheme.Background = new SolidColorBrush(Color.FromRgb(201, 222, 245));
            window.gridStartOpenSolition.Background = new SolidColorBrush(Color.FromRgb(238, 235, 235));
            window.gridStartCreateSolition.Background = new SolidColorBrush(Color.FromRgb(238, 235, 235));
            window.gridStartHelp.Background = new SolidColorBrush(Color.FromRgb(238, 235, 235));
            window.gridStartExitProgram.Background = new SolidColorBrush(Color.FromRgb(238, 235, 235));

        }

        //////////////////////////////////////////////

        public void ApplyGridMouseEnterColor(Object gridSender)
        {
            Grid grid = gridSender as Grid;
            grid.Background = mouseEnterColor;
        }

        public void ApplyGridMouseLeaveColor(Object gridSender)
        {
            Grid grid = gridSender as Grid;
            grid.Background = mouseLeaveColor;
        }


        public void ApplyLabelMouseEnterColor(Object labelSender)
        {
            Label label = (Label)labelSender;
            label.Background = mouseEnterColor;
        }

        public void ApplyLabelMouseLeaveColor(Object labelSender)
        {
            Label label = (Label)labelSender; ;
            label.Background = mouseLeaveColor;
        }

    }

    enum Theme
    {
        Light,
        Dark
    }

}