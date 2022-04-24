using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Controls;
using System.Windows.Input;

namespace Project
{

    public class InterfaceController
    {

        public MainWindow window;                                           // ссылка на экземпляр окна
        private InterfaceАppearance аppearance;                             // ссылка на экземпляр класса InterfaceАppearance

        private SolutionActiveSection solutionActiveSection;                // переменная отображает, какой раздел окрыт в окне редактирования проекта
        private SolutionActiveReportSection solutionActiveReportSection;    // переменная отображает, какой раздел окрыт в окне редактирования проекта в разделе "Создание отчета"

        //////////////////////////////////////////////

        public void SetWindowReference(MainWindow window)
        {
            this.window = window;
        }

        public void SetAppearanceBuilderReference(InterfaceАppearance аppearanceBuilder)
        {
            this.аppearance = аppearanceBuilder;
        }

        //////////////////////////////////////////////

        public void Initialize()
        {
            InitializeGridsEvents();
            InitializeTextBoxEvents();
            InitializeLabelEvents();
            InitializeComboboxEvents();
        }

        public void InitializeGridsEvents()
        {
            window.KeyDown += this.Window_KeyDown;

            window.gridStartOpenSolition.MouseEnter += this.Grid_MouseEnter;
            window.gridStartCreateSolition.MouseEnter += this.Grid_MouseEnter;
            window.gridStartCopySolution.MouseEnter += this.Grid_MouseEnter;
            window.gridStartSettings.MouseEnter += this.Grid_MouseEnter;
            window.gridStartHelp.MouseEnter += this.Grid_MouseEnter;
            window.gridStartExitProgram.MouseEnter += this.Grid_MouseEnter;
            window.gridChangeTheme.MouseEnter += this.Grid_MouseEnter;
            window.gridCreateName.MouseEnter += this.Grid_MouseEnter;
            window.gridCreatePath.MouseEnter += this.Grid_MouseEnter;
            window.gridCreateComment.MouseEnter += this.Grid_MouseEnter;
            window.gridCreateResult.MouseEnter += this.Grid_MouseEnter;
            window.gridCreateFinish.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionSave.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionHelp.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionExit.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionInput.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionReport.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionInfo.MouseEnter += this.Grid_MouseEnter;

            window.gridStartOpenSolition.MouseLeave += this.Grid_MouseLeave;
            window.gridStartCreateSolition.MouseLeave += this.Grid_MouseLeave;
            window.gridStartCopySolution.MouseLeave += this.Grid_MouseLeave;
            window.gridStartSettings.MouseLeave += this.Grid_MouseLeave;
            window.gridStartHelp.MouseLeave += this.Grid_MouseLeave;
            window.gridStartExitProgram.MouseLeave += this.Grid_MouseLeave;
            window.gridChangeTheme.MouseLeave += this.Grid_MouseLeave;
            window.gridCreateName.MouseLeave += this.Grid_MouseLeave;
            window.gridCreatePath.MouseLeave += this.Grid_MouseLeave;
            window.gridCreateComment.MouseLeave += this.Grid_MouseLeave;
            window.gridCreateResult.MouseLeave += this.Grid_MouseLeave;
            window.gridCreateFinish.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionSave.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionHelp.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionExit.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionInput.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionReport.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionInfo.MouseLeave += this.Grid_MouseLeave;

            window.gridStartOpenSolition.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridStartCreateSolition.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridStartCopySolution.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridStartHelp.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridStartExitProgram.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridChangeTheme.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridCreateName.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridCreatePath.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridCreateComment.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridCreateResult.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridCreateFinish.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionSave.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionHelp.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionExit.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionInput.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionReport.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionInfo.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;

            //////////////////////////////////////////////

            window.gridSolutionSecReportUploadWord.MouseEnter += this.Grid_MouseEnter;
            window.gridSolutionSecReportUploadTxt.MouseEnter += this.Grid_MouseEnter;

            window.gridSolutionSecReportUploadWord.MouseLeave += this.Grid_MouseLeave;
            window.gridSolutionSecReportUploadTxt.MouseLeave += this.Grid_MouseLeave;

            window.gridSolutionSecReportUploadWord.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionSecReportUploadTxt.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionSecReport2.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;

            //////////////////////////////////////////////

            window.gridCrit1_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit1_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit1_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit1_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit1_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit1_6.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit2_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit2_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit2_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit3_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit3_8.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit4_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit4_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit4_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit4_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit5_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit5_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit5_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit5_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit5_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit5_6.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit6_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit6_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit6_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit6_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit6_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit6_6.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit7_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit7_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit7_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit7_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit8_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit8_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit8_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit9_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit9_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit9_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit9_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit9_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit9_6.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit10_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit10_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit10_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit10_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit11_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit11_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit11_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit11_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit12_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit12_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit12_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit12_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit12_5.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit13_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit13_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit13_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit13_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit14_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit14_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit14_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit15_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit15_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit15_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit16_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit16_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit16_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit17_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit17_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit17_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit18_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit18_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit18_3.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit19_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit19_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit19_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit19_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit19_5.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit20_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_8.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_9.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit20_10.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit21_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_8.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_9.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_10.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_11.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_12.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_13.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_14.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_15.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_16.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_17.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_18.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_19.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_20.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_21.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_22.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_23.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_24.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_25.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_26.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_27.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit21_28.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit22_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit22_8.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit23_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit23_7.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit24_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit24_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit24_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit24_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit24_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit24_6.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit25_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_8.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit25_9.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit26_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_8.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_9.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit26_10.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit27_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit27_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit27_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit27_4.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit28_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit28_8.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit29_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit29_7.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit30_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit30_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit30_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit30_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit30_5.MouseEnter += this.Grid_MouseEnter;

            window.gridCrit31_1.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_2.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_3.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_4.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_5.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_6.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_7.MouseEnter += this.Grid_MouseEnter;
            window.gridCrit31_8.MouseEnter += this.Grid_MouseEnter;

            //////////////////////////////////////////////

            window.gridCrit1_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit1_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit1_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit1_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit1_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit1_6.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit2_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit2_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit2_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit3_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit3_8.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit4_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit4_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit4_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit4_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit5_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit5_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit5_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit5_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit5_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit5_6.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit6_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit6_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit6_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit6_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit6_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit6_6.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit7_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit7_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit7_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit7_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit8_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit8_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit8_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit9_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit9_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit9_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit9_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit9_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit9_6.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit10_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit10_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit10_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit10_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit11_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit11_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit11_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit11_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit12_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit12_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit12_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit12_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit12_5.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit13_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit13_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit13_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit13_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit14_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit14_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit14_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit15_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit15_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit15_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit16_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit16_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit16_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit17_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit17_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit17_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit18_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit18_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit18_3.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit19_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit19_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit19_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit19_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit19_5.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit20_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_8.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_9.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit20_10.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit21_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_8.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_9.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_10.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_11.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_12.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_13.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_14.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_15.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_16.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_17.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_18.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_19.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_20.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_21.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_22.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_23.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_24.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_25.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_26.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_27.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit21_28.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit22_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit22_8.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit23_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit23_7.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit24_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit24_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit24_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit24_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit24_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit24_6.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit25_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_8.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit25_9.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit26_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_8.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_9.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit26_10.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit27_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit27_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit27_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit27_4.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit28_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit28_8.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit29_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit29_7.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit30_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit30_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit30_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit30_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit30_5.MouseLeave += this.Grid_MouseLeave;

            window.gridCrit31_1.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_2.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_3.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_4.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_5.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_6.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_7.MouseLeave += this.Grid_MouseLeave;
            window.gridCrit31_8.MouseLeave += this.Grid_MouseLeave;

        }

        public void InitializeTextBoxEvents()
        {
            window.textBoxProjectName.TextChanged += TextBox_TextChanged;
            window.textBoxProjectComment.TextChanged += TextBox_TextChanged;
        }

        public void InitializeLabelEvents()
        {
            window.labelReport1.MouseEnter += this.Label_MouseEnter;
            window.labelReport2.MouseEnter += this.Label_MouseEnter;
            window.labelReportValue1.MouseEnter += this.Label_MouseEnter;
            window.labelReportValue2.MouseEnter += this.Label_MouseEnter;

            window.labelReport1.MouseLeave += this.Label_MouseLeave;
            window.labelReport2.MouseLeave += this.Label_MouseLeave;
            window.labelReportValue1.MouseLeave += this.Label_MouseLeave;
            window.labelReportValue2.MouseLeave += this.Label_MouseLeave;

            window.labelReport1.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.labelReportValue1.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.labelReport2.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.labelReportValue2.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;
            window.gridSolutionSecReport2.MouseLeftButtonDown += this.Grid_MouseLeftButtonDown;

            //Test
            window.labelSolutionSectionView.MouseLeftButtonDown += this.Label_MouseLeftButtonDown;
        }

        public void InitializeComboboxEvents()
        {

            window.comboboxCrit1_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit1_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit1_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit1_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit1_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit1_6.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit2_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit2_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit2_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit3_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit3_8.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit4_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit4_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit4_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit4_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit5_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit5_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit5_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit5_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit5_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit5_6.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit6_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit6_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit6_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit6_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit6_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit6_6.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit7_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit7_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit7_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit7_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit8_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit8_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit8_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit9_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit9_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit9_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit9_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit9_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit9_6.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit10_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit10_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit10_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit10_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit11_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit11_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit11_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit11_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit12_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit12_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit12_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit12_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit12_5.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit13_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit13_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit13_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit13_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit14_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit14_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit14_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit15_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit15_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit15_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit16_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit16_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit16_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit17_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit17_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit17_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit18_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit18_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit18_3.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit19_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit19_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit19_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit19_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit19_5.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit20_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_8.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_9.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit20_10.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit21_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_8.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_9.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_10.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_11.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_12.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_13.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_14.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_15.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_16.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_17.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_18.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_19.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_20.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_21.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_22.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_23.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_24.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_25.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_26.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_27.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit21_28.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit22_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit22_8.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit23_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit23_7.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit24_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit24_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit24_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit24_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit24_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit24_6.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit25_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_8.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit25_9.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit26_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_8.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_9.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit26_10.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit27_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit27_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit27_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit27_4.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit28_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit28_8.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit29_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit29_7.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit30_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit30_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit30_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit30_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit30_5.SelectionChanged += Combobox_SelectionChanged;

            window.comboboxCrit31_1.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_2.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_3.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_4.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_5.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_6.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_7.SelectionChanged += Combobox_SelectionChanged;
            window.comboboxCrit31_8.SelectionChanged += Combobox_SelectionChanged;


        }

        //////////////////////////////////////////////

        public void OpenStartWindow()
        {
            window.gridSolutionCreate.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolution.Visibility = System.Windows.Visibility.Hidden;
            window.Height = 470;
            window.Width = 920;
            window.gridRoot.Visibility = System.Windows.Visibility.Visible;
            window.gridStart.Visibility = System.Windows.Visibility.Visible; ;
        }

        public void OpenCreateSolutionWindow()
        {
            window.gridCreatePath.Focus();
            window.gridStart.Visibility = System.Windows.Visibility.Hidden;
            window.Height = 500;
            window.Width = 920;
            window.gridSolutionCreate.Visibility = System.Windows.Visibility.Visible;
        }

        public void OpenSolutionWindow()
        {
            solutionActiveSection = SolutionActiveSection.Input;
            solutionActiveReportSection = SolutionActiveReportSection.Part1;

            аppearance.ApplyGridMouseEnterColor(window.gridSolutionInput);

            window.gridStart.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolutionSecInfo.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolutionCreate.Visibility = System.Windows.Visibility.Hidden;
            window.gridSolutionSecReportUpload.Visibility = System.Windows.Visibility.Hidden;
            window.Height = 750;
            window.Width = 1150;
            window.gridSolutionSecInput.Visibility = System.Windows.Visibility.Visible;
            window.gridSolution.Visibility = System.Windows.Visibility.Visible;

            window.xsolution.FillInfoSection();
        }

        //////////////////////////////////////////////

        private void Grid_MouseEnter(object sender, MouseEventArgs e)
        {
            аppearance.ApplyGridMouseEnterColor(sender);
        }

        private void Grid_MouseLeave(object sender, MouseEventArgs e)
        {
            аppearance.ApplyGridMouseLeaveColor(sender);
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender == window.gridStartOpenSolition)
            {
                Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
                dialog.Filter = "Project Files (*.xsf)|*.xsf";
                dialog.FilterIndex = 1;

                Nullable<bool> result = dialog.ShowDialog();

                if (result == true)
                {
                    window.xsolution = new ProjectSolutionBuilder();
                    window.xsolution.SetWindowReference(window);
                    window.xsolution.OpenProject(dialog.FileName);
                    OpenSolutionWindow();
                }

            }

            if (sender == window.gridStartCreateSolition)
            {
                window.xsolution = new ProjectSolutionBuilder();
                window.xsolution.SetWindowReference(window);
                window.xsolution.CreateSolution();

                window.labelDateSolutionCreate.Content = window.xsolution.GetDate();
                window.labelTimeSolutionCreate.Content = window.xsolution.GetTime();
                window.labelUserSolutionCreate.Content = window.xsolution.GetUsername();

                window.labelProjectPath.Content = "";
                window.labelPathSolutionCreate.Content = "";

                window.textBoxProjectName.Text = "project" + "_" + DateTime.Now.ToShortDateString();
                //window.labelNameSolution.Content = "project" + "__" + DateTime.Now.ToShortDateString();
                window.labelCreateTitle.Content = "Создание проекта";

                window.textBoxProjectComment.Text = "";
                window.labelCommentSolutionCreate.Content = "";

                this.OpenCreateSolutionWindow();
            }

            if (sender == window.gridStartCopySolution)
            {

            }

            if (sender == window.gridStartSettings)
            {

            }

            if (sender == window.gridStartHelp)
            {

            }

            if (sender == window.gridStartExitProgram)
            {
                window.Close();
            }

            if (sender == window.gridChangeTheme)
            {
                аppearance.SwitchTheme();
            }


            //////////////////////////////////////////////

            if (sender == window.gridCreateName)
            {

            }

            if (sender == window.gridCreatePath)
            {
                System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (Directory.Exists(folderBrowserDialog.SelectedPath.ToString()))
                    {
                        window.labelProjectPath.Content = folderBrowserDialog.SelectedPath.ToString();
                        window.labelPathSolutionCreate.Content = window.labelProjectPath.Content;
                    }
                }

                Keyboard.ClearFocus();
            }

            if (sender == window.gridCreateComment)
            {
            }

            if (sender == window.gridCreateResult)
            {
            }

            if (sender == window.gridCreateFinish)
            {
                string projectNameString = window.textBoxProjectName.Text.ToString();
                projectNameString = projectNameString.Trim();

                if (Directory.Exists(window.labelProjectPath.Content.ToString()) && window.labelProjectPath.Content.ToString() != ""
                    && !File.Exists(window.labelProjectPath.Content.ToString() + "/" + window.textBoxProjectName.Text.ToString() + ".xsf") && projectNameString != "")
                {
                    window.xsolution.FinishToCreateProject();
                    window.labelSolutionTitle.Content = "Проект \"" + StringLibrary.AppendChar(window.xsolution.GetName()) + "\"";
                    OpenSolutionWindow();
                }
                else if (projectNameString == "")
                {
                    MessageBox.Show("Необходимо ввести название проекта.", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                    window.textBoxProjectName.Text = "";
                }
                else if (window.labelProjectPath.Content.ToString() == "")
                {
                    MessageBox.Show("Необходимо ввести путь проекта.", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else if (!Directory.Exists(window.labelProjectPath.Content.ToString()))
                {
                    MessageBox.Show("Введен несуществующий путь проекта.", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else if (File.Exists(window.labelProjectPath.Content.ToString() + "/" + window.labelNameSolutionCreate.Content.ToString() + ".xsf"))
                {
                    MessageBox.Show("Файл с таким названием уже существует в указанной папке.", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            //////////////////////////////////////////////

            if (sender == window.gridSolutionSave)
            {
                window.xsolution.UpdateSolutionFile();
                MessageBox.Show("Изменения сохранены.", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            if (sender == window.gridSolutionHelp)
            {

            }

            if (sender == window.gridSolutionExit)
            {

                if (MessageBoxResult.Yes == MessageBox.Show("При выходе к главное меню несохранённые изменения будут утеряны." + "\n" + "\n" + "Продолжить выход?", "Выход", MessageBoxButton.YesNo, MessageBoxImage.Information))
                {
                    window.xsolution.ClearCriteriasComboboxes();
                    window.xsolution = null;
                    OpenStartWindow();
                }
                else
                {

                }

            }

            if (sender == window.gridSolutionInput)
            {
                if (solutionActiveSection != SolutionActiveSection.Input)
                {
                    solutionActiveSection = SolutionActiveSection.Input;
                    window.labelSolutionSectionView.Content = "Введите значения критериев";

                    аppearance.ApplyGridMouseEnterColor(window.gridSolutionInput);
                    аppearance.ApplyGridMouseLeaveColor(window.gridSolutionReport);
                    аppearance.ApplyGridMouseLeaveColor(window.gridSolutionInfo);

                    window.gridSolutionSecReportUpload.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecInfo.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecInput.Visibility = System.Windows.Visibility.Visible;
                }
            }

            if (sender == window.gridSolutionInfo)
            {
                if (solutionActiveSection != SolutionActiveSection.Info)
                {
                    solutionActiveSection = SolutionActiveSection.Info;
                    window.labelSolutionSectionView.Content = "Информация о проекте";

                    аppearance.ApplyGridMouseLeaveColor(window.gridSolutionInput);
                    аppearance.ApplyGridMouseLeaveColor(window.gridSolutionReport);
                    аppearance.ApplyGridMouseEnterColor(window.gridSolutionInfo);

                    window.gridSolutionSecReportUpload.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecInfo.Visibility = System.Windows.Visibility.Visible;
                }
            }

            if (sender == window.gridSolutionReport)
            {
                if (solutionActiveSection != SolutionActiveSection.Report)
                {
                    if (window.xsolution.IsCriteriasComboboxesSelected())
                    {
                        window.xsolution.CountSecurityAssessmentVar1();
                        window.xsolution.CreateLabelReport();
                        solutionActiveSection = SolutionActiveSection.Report;

                        аppearance.ApplyGridMouseLeaveColor(window.gridSolutionInput);
                        аppearance.ApplyGridMouseEnterColor(window.gridSolutionReport);
                        аppearance.ApplyGridMouseLeaveColor(window.gridSolutionInfo);

                        window.gridSolutionSecInput.Visibility = System.Windows.Visibility.Hidden;
                        window.gridSolutionSecInfo.Visibility = System.Windows.Visibility.Hidden;

                        if (solutionActiveReportSection == SolutionActiveReportSection.Part1)
                        {
                            window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Hidden;
                            window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Visible;
                        }
                        if (solutionActiveReportSection == SolutionActiveReportSection.Part2)
                        {
                            window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Hidden;
                            window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Visible;
                        }

                        window.gridSolutionSecReportUpload.Visibility = System.Windows.Visibility.Visible;
                        window.labelSolutionSectionView.Content = "Отчет";

                    }
                    else
                    {
                        MessageBox.Show("Не все значения критериев заполнены в разделе \"Введение данных\"." + "\n" + "\n" + "Заполните оставшиеся значения, затем можете перейти к этому разделу.",
                            "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }

            if (sender == window.gridSolutionSecReportUploadWord)
            {
                System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                    stopwatch.Start();
                    window.xsolution.CreateDocxReport(folderBrowserDialog.SelectedPath.ToString());
                    stopwatch.Stop();
                    MessageBox.Show("Отчёт в формате docx создан по указанному пути за " + stopwatch.ElapsedMilliseconds + " миллисекунд.", "Создание отчёта", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

            if (sender == window.gridSolutionSecReportUploadTxt)
            {
                System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
                    stopwatch.Start();
                    window.xsolution.CreateTxtReport(folderBrowserDialog.SelectedPath.ToString());
                    stopwatch.Stop();
                    MessageBox.Show("Отчёт в формате txt создан по указанному пути за " + stopwatch.ElapsedMilliseconds + " миллисекунд.", "Создание отчёта", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

            if (sender == window.labelReport1 || sender == window.labelReportValue1 || sender == window.labelReport2 || sender == window.labelReportValue2)
            {
                if (window.gridSolutionSecReport1.Visibility == System.Windows.Visibility.Visible)
                {
                    solutionActiveReportSection = SolutionActiveReportSection.Part2;
                    window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Visible;
                    return;
                }
                if (window.gridSolutionSecReport2.Visibility == System.Windows.Visibility.Visible)
                {
                    solutionActiveReportSection = SolutionActiveReportSection.Part1;
                    window.gridSolutionSecReport2.Visibility = System.Windows.Visibility.Hidden;
                    window.gridSolutionSecReport1.Visibility = System.Windows.Visibility.Visible;
                    return;
                }
            }

        }

        //////////////////////////////////////////////

        private void Label_MouseEnter(object sender, MouseEventArgs e)
        {
            аppearance.ApplyLabelMouseEnterColor(sender);
        }

        private void Label_MouseLeave(object sender, MouseEventArgs e)
        {
            аppearance.ApplyLabelMouseLeaveColor(sender);
        }

        private void Label_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender == window.labelSolutionSectionView)
            {
                window.xsolution.UpdateAsTestFill();
                MessageBox.Show("Критерии защищенности были заполнены в случайном порядке.", "Тестовый режим", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        //////////////////////////////////////////////

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape && window.gridSolutionCreate.Visibility == System.Windows.Visibility.Visible)
            {
                OpenStartWindow();
                window.xsolution = null;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender == window.textBoxProjectName)
            {            
                string projectNameString = window.textBoxProjectName.Text;

                if (projectNameString.IndexOf('/') != -1)
                {
                    projectNameString = projectNameString.Trim('/');
                    MessageBox.Show("В название проекта введен недопустимый символ \"/\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('\\') != -1)
                {
                    projectNameString = projectNameString.Trim('\\');
                    MessageBox.Show("В название проекта введен недопустимый символ \"\\\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('*') != -1)
                {
                    projectNameString = projectNameString.Trim('*');
                    MessageBox.Show("В название проекта введен недопустимый символ \"*\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf(':') != -1)
                {
                    projectNameString = projectNameString.Trim(':');
                    MessageBox.Show("В название проекта введен недопустимый символ \":\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('?') != -1)
                {
                    projectNameString = projectNameString.Trim('?');
                    MessageBox.Show("В название проекта введен недопустимый символ \"?\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('"') != -1)
                {
                    projectNameString = projectNameString.Trim('"');
                    //ShowMessageBoxError("Ошибка ввода", "В название проекта введен недопустимый символ \"");
                    MessageBox.Show("В название проекта введен недопустимый символ \".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('|') != -1)
                {
                    projectNameString = projectNameString.Trim('|');
                    MessageBox.Show("В название проекта введен недопустимый символ \"|\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('<') != -1)
                {
                    projectNameString = projectNameString.Trim('<');
                    MessageBox.Show("В название проекта введен недопустимый символ \"<\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                if (projectNameString.IndexOf('>') != -1)
                {
                    projectNameString = projectNameString.Trim('>');
                    MessageBox.Show("В название проекта введен недопустимый символ \">\".", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                window.textBoxProjectName.Text = projectNameString;
                window.labelNameSolutionCreate.Content = StringLibrary.AppendChar(window.textBoxProjectName.Text);
            }
            if (sender == window.textBoxProjectComment)
            {
                window.labelCommentSolutionCreate.Content = window.textBoxProjectComment.Text;
            }
        }

        private void Combobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            solutionActiveReportSection = SolutionActiveReportSection.Part1;
        }

    }

    enum SolutionActiveSection
    {
        Input,
        Info,
        Report
    }

    enum SolutionActiveReportSection
    {
        Part1,
        Part2
    }

}