﻿using System;
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

    /*
       Класс - точка входа проекта
     */
    public partial class MainWindow : Window
    {

        public ApplicationCore xcore; // экземпляр класса Core
        public InterfaceАppearance xappearance; // экземпляр класса АppearanceBuilder
        public ProjectSolutionBuilder xsolution; // экземпляр класса SolutionBuilder
        public InterfaceController xinterface; // экземпляр класса UserInterface

        /////////////////////////////////////////////////

        /*
           Точка входа проекта
         */
        public MainWindow()
        {
            InitializeComponent();
            InitializeCore();
            RunCore();
        }

        /*
           Инициализируем экземпляра класса Core
         */
        public void InitializeCore()
        {
            xcore = new ApplicationCore();
            xcore.SetWindowReference(this);
        }

        /*
           Запускаем Core
         */
        public void RunCore()
        {
            xcore.InitializeObjects();
            xcore.RunProgram();
        }

    }

}