using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using System.Text;
using System.Threading.Tasks;

namespace Project
{

    public class ProjectSolutionBuilder
    {

        public MainWindow window;                   // ссылка на экземпляр окна

        private DateTime dateClassSolutionCreation; // экземпляр DateTime создания проекта (дата, время)
        private string dateSolutionCreation;        // дата создания проекта
        private string timeSolutionCreation;        // время создания проекта
        private string path;                        // путь папки (директории) расположения файла
        private string name;                        // имя файла проекта (без рарешения)
        private string comment;                     // комментарий проекта
        private string fullFileName;                // полный путь проекта (директория + имя + разрешение)
        private string usernameSolutionCreation;    // имя учетной записи создателя проекта

        private List<string> partNames;
        private List<string> sectionNames;
        private List<List<ComboBox>> comboboxList;
        private List<List<string>> criteriasPseudonims;

        private double generalAssessment;
        private List<List<double>> criteriaAssessment;
        private List<double> sectionAssessment;
        private List<double> partAssessment;

        //////////////////////////////////////////////

        public void SetWindowReference(MainWindow window)
        {
            this.window = window;
        }

        private void SetFullProjectPath(string fullPath)
        {
            fullFileName = fullPath;
        }


        public string GetUsername()
        {
            return usernameSolutionCreation;
        }

        public string GetName()
        {
            return name;
        }

        public string GetDate()
        {
            return dateSolutionCreation;
        }

        public string GetTime()
        {
            return timeSolutionCreation;
        }

        public string GetTimeNow()
        {
            DateTime date = new DateTime();
            date = DateTime.Now;
            string ram = date.ToString();
            string timestr;
            if (ram[12] == ':') { timestr = "" + "0" + ram[11] + ram[12] + ram[13] + ram[14] + ram[15] + ram[16] + ram[17]; }
            else { timestr = "" + ram[11] + ram[12] + ram[13] + ram[14] + ram[15] + ram[16] + ram[17] + ram[18]; }

            return timestr;
        }

        //////////////////////////////////////////////

        private void FillLists()
        {

            partNames = new List<string>();
            partNames.Add("Ч1  Соответствие документации");
            partNames.Add("Ч2  Обеспечение защиты информации в ходе эксплуатации");
            partNames.Add("Ч3  Соответствие систем и средств защиты");

            sectionNames = new List<string>();
            sectionNames.Add("NULL");
            sectionNames.Add("Р1  Соответствие нормативно-правовым актам ");
            sectionNames.Add("Р2  Соответствие формуляру ");
            sectionNames.Add("Р3  Соответствие проектной документации ");
            sectionNames.Add("Р4  Соответствие эксплуатационной документации");
            sectionNames.Add("Р5  Соответствие организационно-распорядительной документации");
            sectionNames.Add("Р6  Обеспечение защищенности коммерческой тайны");
            sectionNames.Add("Р7  Планирование мероприятий по защите информации в информационной системе");
            sectionNames.Add("Р8  Анализ угроз безопасности информации в информационной системе");
            sectionNames.Add("Р9  Управление системой защиты информации информационной системы");
            sectionNames.Add("Р10 Управление конфигурацией информационной системы и ее системой ЗИ");
            sectionNames.Add("Р11 Реагирование на инциденты");
            sectionNames.Add("Р12 Информирование и обучение персонала информационной системы");
            sectionNames.Add("Р13 Контроль за обеспечением уровня защищенности информации, содержащейся в ИС");
            sectionNames.Add("Р14 Антивирусная защита");
            sectionNames.Add("Р15 Система обнаружения вторжений");
            sectionNames.Add("Р16 Сканер уязвимости");
            sectionNames.Add("Р17 Межсетевой экран");
            sectionNames.Add("Р18 Криптографическая защита");
            sectionNames.Add("Р19 Защита технических средств");
            sectionNames.Add("Р20 Защита среды виртуализации");
            sectionNames.Add("Р21 Защита информационной системы, ее средств и систем связи и передачи данных");
            sectionNames.Add("Р22 Защита машинных носителей информации");
            sectionNames.Add("Р23 Защита виртуальной инфраструктуры");
            sectionNames.Add("Р24 Защита физических носителей информации");
            sectionNames.Add("Р25 Идентификация и аутентификация в ИС");
            sectionNames.Add("Р26 Управление доступом");
            sectionNames.Add("Р27 Ограничение программной среды");
            sectionNames.Add("Р28 Обеспечение целостности информационной системы и информации");
            sectionNames.Add("Р29 Обеспечение доступности информации");
            sectionNames.Add("Р30 Анализ защищенности информации");
            sectionNames.Add("Р31 Регистрация событий безопасности");

        }

        private void FillCriteriasPseudonims()
        {
            criteriasPseudonims = new List<List<string>>();
            criteriasPseudonims.Add(new List<string>()); // NULL
            for (int i = 0; i < 31; i++) criteriasPseudonims.Add(new List<string>()); // ADD ALL SECTIONS 1 - 31

            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.1");
            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.2");
            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.3");
            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.4");
            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.5");
            criteriasPseudonims[1].Add("InformationSecurity.Criteria_1.6");

            criteriasPseudonims[2].Add("InformationSecurity.Criteria_2.1");
            criteriasPseudonims[2].Add("InformationSecurity.Criteria_2.2");
            criteriasPseudonims[2].Add("InformationSecurity.Criteria_2.3");

            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.1");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.2");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.3");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.4");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.5");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.6");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.7");
            criteriasPseudonims[3].Add("InformationSecurity.Criteria_3.8");

            criteriasPseudonims[4].Add("InformationSecurity.Criteria_4.1");
            criteriasPseudonims[4].Add("InformationSecurity.Criteria_4.2");
            criteriasPseudonims[4].Add("InformationSecurity.Criteria_4.3");
            criteriasPseudonims[4].Add("InformationSecurity.Criteria_4.4");

            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.1");
            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.2");
            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.3");
            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.4");
            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.5");
            criteriasPseudonims[5].Add("InformationSecurity.Criteria_5.6");

            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.1");
            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.2");
            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.3");
            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.4");
            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.5");
            criteriasPseudonims[6].Add("InformationSecurity.Criteria_6.6");

            criteriasPseudonims[7].Add("InformationSecurity.Criteria_7.1");
            criteriasPseudonims[7].Add("InformationSecurity.Criteria_7.2");
            criteriasPseudonims[7].Add("InformationSecurity.Criteria_7.3");
            criteriasPseudonims[7].Add("InformationSecurity.Criteria_7.4");

            criteriasPseudonims[8].Add("InformationSecurity.Criteria_8.1");
            criteriasPseudonims[8].Add("InformationSecurity.Criteria_8.2");
            criteriasPseudonims[8].Add("InformationSecurity.Criteria_8.3");

            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.1");
            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.2");
            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.3");
            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.4");
            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.5");
            criteriasPseudonims[9].Add("InformationSecurity.Criteria_9.6");

            criteriasPseudonims[10].Add("InformationSecurity.Criteria_10.1");
            criteriasPseudonims[10].Add("InformationSecurity.Criteria_10.2");
            criteriasPseudonims[10].Add("InformationSecurity.Criteria_10.3");
            criteriasPseudonims[10].Add("InformationSecurity.Criteria_10.4");

            criteriasPseudonims[11].Add("InformationSecurity.Criteria_11.1");
            criteriasPseudonims[11].Add("InformationSecurity.Criteria_11.2");
            criteriasPseudonims[11].Add("InformationSecurity.Criteria_11.3");
            criteriasPseudonims[11].Add("InformationSecurity.Criteria_11.4");

            criteriasPseudonims[12].Add("InformationSecurity.Criteria_12.1");
            criteriasPseudonims[12].Add("InformationSecurity.Criteria_12.2");
            criteriasPseudonims[12].Add("InformationSecurity.Criteria_12.3");
            criteriasPseudonims[12].Add("InformationSecurity.Criteria_12.4");
            criteriasPseudonims[12].Add("InformationSecurity.Criteria_12.5");

            criteriasPseudonims[13].Add("InformationSecurity.Criteria_13.1");
            criteriasPseudonims[13].Add("InformationSecurity.Criteria_13.2");
            criteriasPseudonims[13].Add("InformationSecurity.Criteria_13.3");
            criteriasPseudonims[13].Add("InformationSecurity.Criteria_13.4");

            criteriasPseudonims[14].Add("InformationSecurity.Criteria_14.1");
            criteriasPseudonims[14].Add("InformationSecurity.Criteria_14.2");
            criteriasPseudonims[14].Add("InformationSecurity.Criteria_14.3");

            criteriasPseudonims[15].Add("InformationSecurity.Criteria_15.1");
            criteriasPseudonims[15].Add("InformationSecurity.Criteria_15.2");
            criteriasPseudonims[15].Add("InformationSecurity.Criteria_15.3");

            criteriasPseudonims[16].Add("InformationSecurity.Criteria_16.1");
            criteriasPseudonims[16].Add("InformationSecurity.Criteria_16.2");
            criteriasPseudonims[16].Add("InformationSecurity.Criteria_16.3");

            criteriasPseudonims[17].Add("InformationSecurity.Criteria_17.1");
            criteriasPseudonims[17].Add("InformationSecurity.Criteria_17.2");
            criteriasPseudonims[17].Add("InformationSecurity.Criteria_17.3");

            criteriasPseudonims[18].Add("InformationSecurity.Criteria_18.1");
            criteriasPseudonims[18].Add("InformationSecurity.Criteria_18.2");
            criteriasPseudonims[18].Add("InformationSecurity.Criteria_18.3");

            criteriasPseudonims[19].Add("InformationSecurity.Criteria_19.1");
            criteriasPseudonims[19].Add("InformationSecurity.Criteria_19.2");
            criteriasPseudonims[19].Add("InformationSecurity.Criteria_19.3");
            criteriasPseudonims[19].Add("InformationSecurity.Criteria_19.4");
            criteriasPseudonims[19].Add("InformationSecurity.Criteria_19.5");

            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.1");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.2");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.3");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.4");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.5");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.6");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.7");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.8");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.9");
            criteriasPseudonims[20].Add("InformationSecurity.Criteria_20.10");

            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.1");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.2");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.3");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.4");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.5");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.6");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.7");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.8");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.9");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.10");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.11");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.12");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.13");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.14");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.15");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.16");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.17");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.18");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.19");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.20");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.21");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.22");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.23");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.24");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.25");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.26");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.27");
            criteriasPseudonims[21].Add("InformationSecurity.Criteria_21.28");

            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.1");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.2");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.3");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.4");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.5");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.6");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.7");
            criteriasPseudonims[22].Add("InformationSecurity.Criteria_22.8");

            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.1");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.2");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.3");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.4");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.5");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.6");
            criteriasPseudonims[23].Add("InformationSecurity.Criteria_23.7");

            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.1");
            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.2");
            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.3");
            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.4");
            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.5");
            criteriasPseudonims[24].Add("InformationSecurity.Criteria_24.6");

            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.1");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.2");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.3");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.4");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.5");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.6");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.7");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.8");
            criteriasPseudonims[25].Add("InformationSecurity.Criteria_25.9");

            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.1");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.2");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.3");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.4");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.5");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.6");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.7");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.8");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.9");
            criteriasPseudonims[26].Add("InformationSecurity.Criteria_26.10");

            criteriasPseudonims[27].Add("InformationSecurity.Criteria_27.1");
            criteriasPseudonims[27].Add("InformationSecurity.Criteria_27.2");
            criteriasPseudonims[27].Add("InformationSecurity.Criteria_27.3");
            criteriasPseudonims[27].Add("InformationSecurity.Criteria_27.4");

            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.1");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.2");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.3");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.4");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.5");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.6");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.7");
            criteriasPseudonims[28].Add("InformationSecurity.Criteria_28.8");

            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.1");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.2");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.3");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.4");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.5");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.6");
            criteriasPseudonims[29].Add("InformationSecurity.Criteria_29.7");

            criteriasPseudonims[30].Add("InformationSecurity.Criteria_30.1");
            criteriasPseudonims[30].Add("InformationSecurity.Criteria_30.2");
            criteriasPseudonims[30].Add("InformationSecurity.Criteria_30.3");
            criteriasPseudonims[30].Add("InformationSecurity.Criteria_30.4");
            criteriasPseudonims[30].Add("InformationSecurity.Criteria_30.5");

            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.1");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.2");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.3");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.4");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.5");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.6");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.7");
            criteriasPseudonims[31].Add("InformationSecurity.Criteria_31.8");

        }

        private void FillCriteriasComboboxes()
        {
            comboboxList = new List<List<ComboBox>>();
            comboboxList.Add(new List<ComboBox>()); // NULL
            for (int i = 0; i < 31; i++) comboboxList.Add(new List<ComboBox>()); // ADD ALL SECTIONS 1 - 31

            comboboxList[1].Add(window.comboboxCrit1_1);
            comboboxList[1].Add(window.comboboxCrit1_2);
            comboboxList[1].Add(window.comboboxCrit1_3);
            comboboxList[1].Add(window.comboboxCrit1_4);
            comboboxList[1].Add(window.comboboxCrit1_5);
            comboboxList[1].Add(window.comboboxCrit1_6);

            comboboxList[2].Add(window.comboboxCrit2_1);
            comboboxList[2].Add(window.comboboxCrit2_2);
            comboboxList[2].Add(window.comboboxCrit2_3);

            comboboxList[3].Add(window.comboboxCrit3_1);
            comboboxList[3].Add(window.comboboxCrit3_2);
            comboboxList[3].Add(window.comboboxCrit3_3);
            comboboxList[3].Add(window.comboboxCrit3_4);
            comboboxList[3].Add(window.comboboxCrit3_5);
            comboboxList[3].Add(window.comboboxCrit3_6);
            comboboxList[3].Add(window.comboboxCrit3_7);
            comboboxList[3].Add(window.comboboxCrit3_8);

            comboboxList[4].Add(window.comboboxCrit4_1);
            comboboxList[4].Add(window.comboboxCrit4_2);
            comboboxList[4].Add(window.comboboxCrit4_3);
            comboboxList[4].Add(window.comboboxCrit4_4);

            comboboxList[5].Add(window.comboboxCrit5_1);
            comboboxList[5].Add(window.comboboxCrit5_2);
            comboboxList[5].Add(window.comboboxCrit5_3);
            comboboxList[5].Add(window.comboboxCrit5_4);
            comboboxList[5].Add(window.comboboxCrit5_5);
            comboboxList[5].Add(window.comboboxCrit5_6);

            comboboxList[6].Add(window.comboboxCrit6_1);
            comboboxList[6].Add(window.comboboxCrit6_2);
            comboboxList[6].Add(window.comboboxCrit6_3);
            comboboxList[6].Add(window.comboboxCrit6_4);
            comboboxList[6].Add(window.comboboxCrit6_5);
            comboboxList[6].Add(window.comboboxCrit6_6);

            comboboxList[7].Add(window.comboboxCrit7_1);
            comboboxList[7].Add(window.comboboxCrit7_2);
            comboboxList[7].Add(window.comboboxCrit7_3);
            comboboxList[7].Add(window.comboboxCrit7_4);

            comboboxList[8].Add(window.comboboxCrit8_1);
            comboboxList[8].Add(window.comboboxCrit8_2);
            comboboxList[8].Add(window.comboboxCrit8_3);

            comboboxList[9].Add(window.comboboxCrit9_1);
            comboboxList[9].Add(window.comboboxCrit9_2);
            comboboxList[9].Add(window.comboboxCrit9_3);
            comboboxList[9].Add(window.comboboxCrit9_4);
            comboboxList[9].Add(window.comboboxCrit9_5);
            comboboxList[9].Add(window.comboboxCrit9_6);

            comboboxList[10].Add(window.comboboxCrit10_1);
            comboboxList[10].Add(window.comboboxCrit10_2);
            comboboxList[10].Add(window.comboboxCrit10_3);
            comboboxList[10].Add(window.comboboxCrit10_4);

            comboboxList[11].Add(window.comboboxCrit11_1);
            comboboxList[11].Add(window.comboboxCrit11_2);
            comboboxList[11].Add(window.comboboxCrit11_3);
            comboboxList[11].Add(window.comboboxCrit11_4);

            comboboxList[12].Add(window.comboboxCrit12_1);
            comboboxList[12].Add(window.comboboxCrit12_2);
            comboboxList[12].Add(window.comboboxCrit12_3);
            comboboxList[12].Add(window.comboboxCrit12_4);
            comboboxList[12].Add(window.comboboxCrit12_5);

            comboboxList[13].Add(window.comboboxCrit13_1);
            comboboxList[13].Add(window.comboboxCrit13_2);
            comboboxList[13].Add(window.comboboxCrit13_3);
            comboboxList[13].Add(window.comboboxCrit13_4);

            comboboxList[14].Add(window.comboboxCrit14_1);
            comboboxList[14].Add(window.comboboxCrit14_2);
            comboboxList[14].Add(window.comboboxCrit14_3);

            comboboxList[15].Add(window.comboboxCrit15_1);
            comboboxList[15].Add(window.comboboxCrit15_2);
            comboboxList[15].Add(window.comboboxCrit15_3);

            comboboxList[16].Add(window.comboboxCrit16_1);
            comboboxList[16].Add(window.comboboxCrit16_2);
            comboboxList[16].Add(window.comboboxCrit16_3);

            comboboxList[17].Add(window.comboboxCrit17_1);
            comboboxList[17].Add(window.comboboxCrit17_2);
            comboboxList[17].Add(window.comboboxCrit17_3);

            comboboxList[18].Add(window.comboboxCrit18_1);
            comboboxList[18].Add(window.comboboxCrit18_2);
            comboboxList[18].Add(window.comboboxCrit18_3);

            comboboxList[19].Add(window.comboboxCrit19_1);
            comboboxList[19].Add(window.comboboxCrit19_2);
            comboboxList[19].Add(window.comboboxCrit19_3);
            comboboxList[19].Add(window.comboboxCrit19_4);
            comboboxList[19].Add(window.comboboxCrit19_5);

            comboboxList[20].Add(window.comboboxCrit20_1);
            comboboxList[20].Add(window.comboboxCrit20_2);
            comboboxList[20].Add(window.comboboxCrit20_3);
            comboboxList[20].Add(window.comboboxCrit20_4);
            comboboxList[20].Add(window.comboboxCrit20_5);
            comboboxList[20].Add(window.comboboxCrit20_6);
            comboboxList[20].Add(window.comboboxCrit20_7);
            comboboxList[20].Add(window.comboboxCrit20_8);
            comboboxList[20].Add(window.comboboxCrit20_9);
            comboboxList[20].Add(window.comboboxCrit20_10);

            comboboxList[21].Add(window.comboboxCrit21_1);
            comboboxList[21].Add(window.comboboxCrit21_2);
            comboboxList[21].Add(window.comboboxCrit21_3);
            comboboxList[21].Add(window.comboboxCrit21_4);
            comboboxList[21].Add(window.comboboxCrit21_5);
            comboboxList[21].Add(window.comboboxCrit21_6);
            comboboxList[21].Add(window.comboboxCrit21_7);
            comboboxList[21].Add(window.comboboxCrit21_8);
            comboboxList[21].Add(window.comboboxCrit21_9);
            comboboxList[21].Add(window.comboboxCrit21_10);
            comboboxList[21].Add(window.comboboxCrit21_11);
            comboboxList[21].Add(window.comboboxCrit21_12);
            comboboxList[21].Add(window.comboboxCrit21_13);
            comboboxList[21].Add(window.comboboxCrit21_14);
            comboboxList[21].Add(window.comboboxCrit21_15);
            comboboxList[21].Add(window.comboboxCrit21_16);
            comboboxList[21].Add(window.comboboxCrit21_17);
            comboboxList[21].Add(window.comboboxCrit21_18);
            comboboxList[21].Add(window.comboboxCrit21_19);
            comboboxList[21].Add(window.comboboxCrit21_20);
            comboboxList[21].Add(window.comboboxCrit21_21);
            comboboxList[21].Add(window.comboboxCrit21_22);
            comboboxList[21].Add(window.comboboxCrit21_23);
            comboboxList[21].Add(window.comboboxCrit21_24);
            comboboxList[21].Add(window.comboboxCrit21_25);
            comboboxList[21].Add(window.comboboxCrit21_26);
            comboboxList[21].Add(window.comboboxCrit21_27);
            comboboxList[21].Add(window.comboboxCrit21_28);

            comboboxList[22].Add(window.comboboxCrit22_1);
            comboboxList[22].Add(window.comboboxCrit22_2);
            comboboxList[22].Add(window.comboboxCrit22_3);
            comboboxList[22].Add(window.comboboxCrit22_4);
            comboboxList[22].Add(window.comboboxCrit22_5);
            comboboxList[22].Add(window.comboboxCrit22_6);
            comboboxList[22].Add(window.comboboxCrit22_7);
            comboboxList[22].Add(window.comboboxCrit22_8);

            comboboxList[23].Add(window.comboboxCrit23_1);
            comboboxList[23].Add(window.comboboxCrit23_2);
            comboboxList[23].Add(window.comboboxCrit23_3);
            comboboxList[23].Add(window.comboboxCrit23_4);
            comboboxList[23].Add(window.comboboxCrit23_5);
            comboboxList[23].Add(window.comboboxCrit23_6);
            comboboxList[23].Add(window.comboboxCrit23_7);

            comboboxList[24].Add(window.comboboxCrit24_1);
            comboboxList[24].Add(window.comboboxCrit24_2);
            comboboxList[24].Add(window.comboboxCrit24_3);
            comboboxList[24].Add(window.comboboxCrit24_4);
            comboboxList[24].Add(window.comboboxCrit24_5);
            comboboxList[24].Add(window.comboboxCrit24_6);

            comboboxList[25].Add(window.comboboxCrit25_1);
            comboboxList[25].Add(window.comboboxCrit25_2);
            comboboxList[25].Add(window.comboboxCrit25_3);
            comboboxList[25].Add(window.comboboxCrit25_4);
            comboboxList[25].Add(window.comboboxCrit25_5);
            comboboxList[25].Add(window.comboboxCrit25_6);
            comboboxList[25].Add(window.comboboxCrit25_7);
            comboboxList[25].Add(window.comboboxCrit25_8);
            comboboxList[25].Add(window.comboboxCrit25_9);

            comboboxList[26].Add(window.comboboxCrit26_1);
            comboboxList[26].Add(window.comboboxCrit26_2);
            comboboxList[26].Add(window.comboboxCrit26_3);
            comboboxList[26].Add(window.comboboxCrit26_4);
            comboboxList[26].Add(window.comboboxCrit26_5);
            comboboxList[26].Add(window.comboboxCrit26_6);
            comboboxList[26].Add(window.comboboxCrit26_7);
            comboboxList[26].Add(window.comboboxCrit26_8);
            comboboxList[26].Add(window.comboboxCrit26_9);
            comboboxList[26].Add(window.comboboxCrit26_10);

            comboboxList[27].Add(window.comboboxCrit27_1);
            comboboxList[27].Add(window.comboboxCrit27_2);
            comboboxList[27].Add(window.comboboxCrit27_3);
            comboboxList[27].Add(window.comboboxCrit27_4);

            comboboxList[28].Add(window.comboboxCrit28_1);
            comboboxList[28].Add(window.comboboxCrit28_2);
            comboboxList[28].Add(window.comboboxCrit28_3);
            comboboxList[28].Add(window.comboboxCrit28_4);
            comboboxList[28].Add(window.comboboxCrit28_5);
            comboboxList[28].Add(window.comboboxCrit28_6);
            comboboxList[28].Add(window.comboboxCrit28_7);
            comboboxList[28].Add(window.comboboxCrit28_8);

            comboboxList[29].Add(window.comboboxCrit29_1);
            comboboxList[29].Add(window.comboboxCrit29_2);
            comboboxList[29].Add(window.comboboxCrit29_3);
            comboboxList[29].Add(window.comboboxCrit29_4);
            comboboxList[29].Add(window.comboboxCrit29_5);
            comboboxList[29].Add(window.comboboxCrit29_6);
            comboboxList[29].Add(window.comboboxCrit29_7);

            comboboxList[30].Add(window.comboboxCrit30_1);
            comboboxList[30].Add(window.comboboxCrit30_2);
            comboboxList[30].Add(window.comboboxCrit30_3);
            comboboxList[30].Add(window.comboboxCrit30_4);
            comboboxList[30].Add(window.comboboxCrit30_5);

            comboboxList[31].Add(window.comboboxCrit31_1);
            comboboxList[31].Add(window.comboboxCrit31_2);
            comboboxList[31].Add(window.comboboxCrit31_3);
            comboboxList[31].Add(window.comboboxCrit31_4);
            comboboxList[31].Add(window.comboboxCrit31_5);
            comboboxList[31].Add(window.comboboxCrit31_6);
            comboboxList[31].Add(window.comboboxCrit31_7);
            comboboxList[31].Add(window.comboboxCrit31_8);
        }

        private void FillCriteriasGridsTitles()
        {

            window.labelSectionName1.Content = "Раздел 1. Соответствие нормативно-правовым актам ";
            window.critLabel1_1.Text = "Оценка класса средств защиты информации";
            window.critLabel1_2.Text = "Оценка уровня недокументированных возможностей";
            window.critLabel1_3.Text = "Оценка уровня доверия";
            window.critLabel1_4.Text = "Модель угроз";
            window.critLabel1_5.Text = "Техническое задание";
            window.critLabel1_6.Text = "Анализ уязвимостей";

            window.labelSectionName2.Content = "Раздел 2. Соответствие формуляру ";
            window.critLabel2_1.Text = "Применение средств защиты информации в соответствии с его назначением";
            window.critLabel2_2.Text = "Соответствие установленных средств защиты информации документации";
            window.critLabel2_3.Text = "Соответствие программных и технических характеристик";

            window.labelSectionName3.Content = "Раздел 3. Соответствие проектной документации ";
            window.critLabel3_1.Text = "Определение типа учетных записей и контейнеров с защищаемой информацией";
            window.critLabel3_2.Text = "Разрешительная система доступа";
            window.critLabel3_3.Text = "Выбор мер защиты, подлежащих реализации в системе защиты информации информационной системы";
            window.critLabel3_4.Text = "Определение видов и типов средств защиты информации, реализующих техническую защиту информации";
            window.critLabel3_5.Text = "Определение структуры системы ЗИ ИС";
            window.critLabel3_6.Text = "Выбор сертифицированных СЗИ, с учетом программно-аппаратной совместимости, особенностей реализации, класса защищенности ИС и стоимости";
            window.critLabel3_7.Text = "Определение требований к параметрам настройки ПО и СЗИ, а также устранение возможных уязвимостей";
            window.critLabel3_8.Text = "Определение мер ЗИ при информационном воздействии с иными ИС";

            window.labelSectionName4.Content = "Раздел 4. Соответствие эксплуатационной документации";
            window.critLabel4_1.Text = "Описание структуры системы безопасности состава, мест установки, параметров и порядка настройки средств защиты информации";
            window.critLabel4_2.Text = "Описание структуры системы безопасности состава, мест установки, параметров и порядка настройки программного обеспечения";
            window.critLabel4_3.Text = "Описание структуры системы безопасности состава, мест установки, параметров и порядка настройки технических средств";
            window.critLabel4_4.Text = "Описание правил эксплуатации системы защиты информации информационной системы";

            window.labelSectionName5.Content = "Раздел 5. Соответствие организационно-распорядительной документации";
            window.critLabel5_1.Text = "Контроль инцидентов безопасности";
            window.critLabel5_2.Text = "Разработка мероприятий при возникновении инцидентов безопасности";
            window.critLabel5_3.Text = "Поддержание БК ИС и ее системы ЗИ";
            window.critLabel5_4.Text = "Наделение полномочиями работников, осуществляющих изменения в администрировании подсистемами защиты";
            window.critLabel5_5.Text = "Контроль появления УБИ в ходе эксплуатации";
            window.critLabel5_6.Text = "Документирование мероприятий периодического аудита кибербезопасности";

            window.labelSectionName6.Content = "Раздел 6. Обеспечение защищенности коммерческой тайны";
            window.critLabel6_1.Text = "Определение перечня информации, которая составляет коммерческую тайну";
            window.critLabel6_2.Text = "Установление режима коммерческой тайны";
            window.critLabel6_3.Text = "Установление порядка обращения с коммерческой тайной и контроля за соблюдением этого порядка";
            window.critLabel6_4.Text = "Разработка системы учета лиц, которым был предоставлен доступ к такого рода сведениям";
            window.critLabel6_5.Text = "Договорное регулирование отношений по использованию информации";
            window.critLabel6_6.Text = "Использование грифа «коммерческая тайна»";

            window.labelSectionName7.Content = "Раздел 7. Планирование мероприятий по защите информации в информационной системе";
            window.critLabel7_1.Text = "Определение лиц, ответственных за планирование и контроль мероприятий по защите информации в информационной системе";
            window.critLabel7_2.Text = "Определение лиц, ответственных за выявление инцидентов и реагирование на них";
            window.critLabel7_3.Text = "Разработка, утверждение и актуализация плана мероприятий по защите информации в информационной системе";
            window.critLabel7_4.Text = "Определение порядка контроля выполнения мероприятий по обеспечению защиты информации в информационной системе, предусмотренных утвержденным планом";

            window.labelSectionName8.Content = "Раздел 8. Анализ угроз безопасности информации в информационной системе";
            window.critLabel8_1.Text = "Выявление, анализ и устранение уязвимостей информационной системы";
            window.critLabel8_2.Text = "Анализ изменения угроз безопасности информации в информационной системе";
            window.critLabel8_3.Text = "Оценка возможных последствий реализации угроз безопасности информации в информационной системе";

            window.labelSectionName9.Content = "Раздел 9. Управление системой защиты информации информационной системы";
            window.critLabel9_1.Text = "Определение лиц, ответственных за управление (администрирование) системой защиты информации информационной системы";
            window.critLabel9_2.Text = "Управление учетными записями пользователей и поддержание в актуальном состоянии правил разграничения доступа в информационной системе";
            window.critLabel9_3.Text = "Управление средствами защиты информации информационной системы";
            window.critLabel9_4.Text = "Централизованное управление системой защиты информации информационной системы";
            window.critLabel9_5.Text = "Мониторинг и анализ зарегистрированных событий в информационной системе, связанных с обеспечением безопасности";
            window.critLabel9_6.Text = "Обеспечение функционирования системы защиты информации информационной системы в ходе ее эксплуатации, включая ведение эксплуатационной документации и организационно-распорядительных документов по защите информации";

            window.labelSectionName10.Content = "Раздел 10. Управление конфигурацией информационной системы и ее системой защиты информации";
            window.critLabel10_1.Text = "Определение лиц, которым разрешены действия по внесению изменений в конфигурацию информационной системы и ее системы защиты информации, и их полномочий";
            window.critLabel10_2.Text = "Определение компонентов информационной системы и ее системы защиты информации, подлежащих изменению в рамках управления конфигурацией (идентификация объектов управления конфигурацией): программно-аппаратные, программные средства, включая средства защиты информации, их настройки и программный код, эксплуатационная документация, интерфейсы, файлы и иные компоненты, подлежащие изменению и контролю";
            window.critLabel10_3.Text = "Управление изменениями информационной системы и ее системы защиты информации: разработка параметров настройки, обеспечивающих защиту информации, анализ потенциального воздействия планируемых изменений на обеспечение защиты информации, санкционирование внесения изменений в информационную систему и ее систему защиты информации, документирование действий по внесению изменений в информационную систему и сохранение данных об изменениях конфигурации";
            window.critLabel10_4.Text = "Контроль действий по внесению изменений в информационную систему и ее систему защиты информации";

            window.labelSectionName11.Content = "Раздел 11. Реагирование на инциденты";
            window.critLabel11_1.Text = "Своевременное информирование пользователями и администраторами лиц, ответственных за выявление инцидентов и реагирование на них, о возникновении инцидентов в информационной системе";
            window.critLabel11_2.Text = "Анализ инцидентов, в том числе определение источников и причин возникновения инцидентов, а также оценка их последствий";
            window.critLabel11_3.Text = "Планирование и принятие мер по устранению инцидентов, в том числе по восстановлению информационной системы и ее сегментов в случае отказа в обслуживании или после сбоев, устранению последствий нарушения правил разграничения доступа, неправомерных действий по сбору информации, внедрения вредоносных компьютерных программ (вирусов) и иных событий, приводящих к возникновению инцидентов";
            window.critLabel11_4.Text = "Планирование и принятие мер по предотвращению повторного возникновения инцидентов";

            window.labelSectionName12.Content = "Раздел 12. Информирование и обучение персонала информационной системы";
            window.critLabel12_1.Text = "Информирование персонала информационной системы о появлении актуальных угрозах безопасности информации, о правилах безопасной эксплуатации информационной системы";
            window.critLabel12_2.Text = "Доведение до персонала информационной системы требований по защите информации, а также положений организационно-распорядительных документов по защите информации с учетом внесенных в них изменений";
            window.critLabel12_3.Text = "Обучение персонала информационной системы правилам эксплуатации отдельных средств защиты информации";
            window.critLabel12_4.Text = "Проведение практических занятий и тренировок с персоналом информационной системы по блокированию угроз безопасности информации и реагированию на инциденты";
            window.critLabel12_5.Text = "Контроль осведомленности персонала информационной системы об угрозах безопасности информации и уровня знаний персонала по вопросам обеспечения защиты информации";

            window.labelSectionName13.Content = "Раздел 13. Контроль за обеспечением уровня защищенности информации, содержащейся в информационной системе";
            window.critLabel13_1.Text = "Контроль (анализ) защищенности информации с учетом особенностей функционирования информационной системы";
            window.critLabel13_2.Text = "Анализ и оценка функционирования информационной системы и ее системы защиты информации, включая анализ и устранение уязвимостей и иных недостатков в функционировании системы защиты информации информационной системы";
            window.critLabel13_3.Text = "Документирование процедур и результатов контроля за обеспечением уровня защищенности информации, содержащейся в информационной системе";
            window.critLabel13_4.Text = "Принятие решения по результатам контроля за обеспечением уровня защищенности информации, содержащейся в информационной системе, о необходимости доработки (модернизации) ее системы защиты информации";

            window.labelSectionName14.Content = "Раздел 14. Антивирусная защита";
            window.critLabel14_1.Text = "Реализация системы антивирусной защиты";
            window.critLabel14_2.Text = "Администрирование системы антивирусной защиты";
            window.critLabel14_3.Text = "Актуальность баз данных системы антивирусной защиты";

            window.labelSectionName15.Content = "Раздел 15. Система обнаружения вторжений";
            window.critLabel15_1.Text = "Реализация системы обнаружения вторжений";
            window.critLabel15_2.Text = "Администрирование системы обнаружения вторжений";
            window.critLabel15_3.Text = "Поддержание актуальности сигнатур системы обнаружения вторжений";

            window.labelSectionName16.Content = "Раздел 16. Сканер уязвимости";
            window.critLabel16_1.Text = "Реализация сканера уязвимости";
            window.critLabel16_2.Text = "Администрирование сканера уязвимости";
            window.critLabel16_3.Text = "Актуальность баз данных выявляемых уязвимостей";

            window.labelSectionName17.Content = "Раздел 17. Межсетевой экран";
            window.critLabel17_1.Text = "Реализация межсетевого экрана";
            window.critLabel17_2.Text = "Администрирования межсетевого экрана";
            window.critLabel17_3.Text = "Актуальность правил фильтрации трафика межсетевого экрана";

            window.labelSectionName18.Content = "Раздел 18. Криптографическая защита";
            window.critLabel18_1.Text = "Реализация криптографической защиты информации";
            window.critLabel18_2.Text = "Администрирование средства криптографической защиты информации";
            window.critLabel18_3.Text = "Ведение документации на средства криптографической защиты";

            window.labelSectionName19.Content = "Раздел 19. Защита технических средств";
            window.critLabel19_1.Text = "Защита информации, обрабатываемой техническими средствами, от ее утечки по техническим каналам";
            window.critLabel19_2.Text = "Организация контролируемой зоны, в пределах которой постоянно размещаются стационарные технические средства, обрабатывающие информацию, и средства защиты информации, а также средства обеспечения функционирования";
            window.critLabel19_3.Text = "Контроль и управление физическим доступом к техническим средствам, средствам защиты информации";
            window.critLabel19_4.Text = "Размещение устройств вывода (отображения) информации, исключающее ее несанкционированный просмотр";
            window.critLabel19_5.Text = "Защита от внешних воздействий (воздействий окружающей среды, нестабильности электроснабжения, кондиционирования и иных внешних факторов)";

            window.labelSectionName20.Content = "Раздел 20. Защита среды виртуализации";
            window.critLabel20_1.Text = "Идентификация и аутентификация субъектов доступа и объектов доступа в виртуальной инфраструктуре, в том числе администраторов управления средствами виртуализации";
            window.critLabel20_2.Text = "Управление доступом субъектов доступа к объектам доступа в виртуальной инфраструктуре, в том числе внутри виртуальных машин";
            window.critLabel20_3.Text = "Регистрация событий безопасности в виртуальной инфраструктуре";
            window.critLabel20_4.Text = "Управление (фильтрация, маршрутизация, контроль соединения, однонаправленная передача) потоками информации между компонентами виртуальной инфраструктуры, а также по периметру виртуальной инфраструктуры";
            window.critLabel20_5.Text = "Доверенная загрузка серверов виртуализации, виртуальной машины (контейнера), серверов управления виртуализацией";
            window.critLabel20_6.Text = "Управление перемещением виртуальных машин (контейнеров) и обрабатываемых на них данных";
            window.critLabel20_7.Text = "Контроль целостности виртуальной инфраструктуры и ее конфигураций";
            window.critLabel20_8.Text = "Резервное копирование данных, резервирование технических средств, программного обеспечения виртуальной инфраструктуры, а также каналов связи внутри виртуальной инфраструктуры";
            window.critLabel20_9.Text = "Реализация и управление антивирусной защитой в виртуальной инфраструктуре";
            window.critLabel20_10.Text = "Разбиение виртуальной инфраструктуры на сегменты (сегментирование виртуальной инфраструктуры) для обработки информации отдельным пользователем и (или) группой пользователей";

            window.labelSectionName21.Content = "Раздел 21. Защита информационной системы, ее средств и систем связи и передачи данных";
            window.critLabel21_1.Text = "Разделение в информационной системе функций по управлению информационной системой, управлению системой защиты информации, функций по обработке информации и иных функций информационной системы";
            window.critLabel21_2.Text = "Предотвращение задержки или прерывания выполнения процессов с высоким приоритетом со стороны процессов с низким приоритетом";
            window.critLabel21_3.Text = "Обеспечение защиты информации от раскрытия, модификации и навязывания (ввода ложной информации) при ее передаче (подготовке к передаче) по каналам связи, имеющим выход за пределы контролируемой зоны";
            window.critLabel21_4.Text = "Обеспечение доверенных канала, маршрута между администратором, пользователем и средствами защиты информации (функциями безопасности средств защиты информации)";
            window.critLabel21_5.Text = "Запрет несанкционированной удаленной активации видеокамер, микрофонов и иных периферийных устройств, которые могут активироваться удаленно, и оповещение пользователей об активации таких устройств";
            window.critLabel21_6.Text = "Передача и контроль целостности атрибутов безопасности (меток безопасности), связанных с информацией, при обмене информацией с иными информационными системами";
            window.critLabel21_7.Text = "Контроль санкционированного и исключение несанкционированного использования технологий мобильного кода, в том числе регистрация событий, связанных с использованием технологии мобильного кода, их анализ и реагирование на нарушения, связанные с использованием технологии мобильного кода";
            window.critLabel21_8.Text = "Контроль санкционированного и исключение несанкционированного использования технологий передачи речи, в том числе регистрация событий, связанных с использованием технологий передачи речи, их анализ и реагирование на нарушения, связанные с использованием технологий передачи речи";
            window.critLabel21_9.Text = "Контроль санкционированной и исключение несанкционированной передачи видеоинформации, в том числе регистрация событий, связанных с передачей видеоинформации, их анализ и реагирование на нарушения, связанные с передачей видеоинформации";
            window.critLabel21_10.Text = "Подтверждение происхождения источника информации, получаемой в процессе определения сетевых адресов по сетевым именам или определения сетевых имен по сетевым адресам";
            window.critLabel21_11.Text = "Обеспечение подлинности сетевых соединений (сеансов взаимодействия), в том числе для защиты от подмены сетевых устройств и сервисов";
            window.critLabel21_12.Text = "Исключение возможности отрицания пользователем факта отправки информации другому пользователю";
            window.critLabel21_13.Text = "Исключение возможности отрицания пользователем факта получения информации от другого пользователя";
            window.critLabel21_14.Text = "Использование устройств терминального доступа для обработки информации";
            window.critLabel21_15.Text = "Защита архивных файлов, параметров настройки средств защиты информации и программного обеспечения и иных данных, не подлежащих изменению в процессе обработки информации";
            window.critLabel21_16.Text = "Выявление, анализ и блокирование в информационной системе скрытых каналов передачи информации в обход реализованных мер защиты информации или внутри разрешенных сетевых протоколов";
            window.critLabel21_17.Text = "Обеспечение загрузки и исполнения программного обеспечения с машинных носителей информации, доступных только для чтения, и контроль целостности данного программного обеспечения";
            window.critLabel21_18.Text = "Защита беспроводных соединений, применяемых в информационной системе";
            window.critLabel21_19.Text = "Исключение доступа пользователя к информации, возникшей в результате действий предыдущего пользователя через реестры, оперативную память, внешние запоминающие устройства и иные общие для пользователей ресурсы информационной системы";
            window.critLabel21_20.Text = "Защита информационной системы от угроз безопасности информации, направленных на отказ в обслуживании этой информационной системы";
            window.critLabel21_21.Text = "Защита периметра (физических и (или) логических границ) информационной системы при ее взаимодействии с иными информационными системами и информационно-телекоммуникационными сетями";
            window.critLabel21_22.Text = "Прекращение сетевых соединений по их завершении или по истечении заданного оператором временного интервала неактивности сетевого соединения";
            window.critLabel21_23.Text = "Использование в информационной системе или ее сегментах различных типов общесистемного, прикладного и специального программного обеспечения (создание гетерогенной среды)";
            window.critLabel21_24.Text = "Использование прикладного и специального программного обеспечения, имеющего возможность функционирования на различных типах операционных систем";
            window.critLabel21_25.Text = "Создание (эмуляция) ложных информационных систем или их компонентов, предназначенных для обнаружения, регистрации и анализа действий нарушителей в процессе реализации угроз безопасности информации";
            window.critLabel21_26.Text = "Воспроизведение ложных и (или) скрытие истинных отдельных информационных технологий и (или) структурно-функциональных характеристик информационной системы или ее сегментов, обеспечивающее навязывание у нарушителя ложного представления об истинных информационных технологиях и (или) структурно-функциональных характеристиках информационной системы";
            window.critLabel21_27.Text = "Перевод информационной системы или ее устройств (компонентов) в заранее определенную конфигурацию, обеспечивающую защиту информации, в случае возникновения отказов (сбоев) в системе защиты информации информационной системы";
            window.critLabel21_28.Text = "Защита мобильных технических средств, применяемых в информационной системе";

            window.labelSectionName22.Content = "Раздел 22. Защита машинных носителей информации";
            window.critLabel22_1.Text = "Учет машинных носителей информации";
            window.critLabel22_2.Text = "Управление доступом к машинным носителям информации";
            window.critLabel22_3.Text = "Контроль перемещения машинных носителей информации за пределы контролируемой зоны";
            window.critLabel22_4.Text = "Исключение возможности несанкционированного ознакомления с содержанием информации, хранящейся на машинных носителях, и (или) использования носителей информации в иных информационных системах";
            window.critLabel22_5.Text = "Контроль использования интерфейсов ввода (вывода) информации на машинные носители информации";
            window.critLabel22_6.Text = "Контроль ввода (вывода) информации на машинные носители информации";
            window.critLabel22_7.Text = "Контроль подключения машинных носителей информации";
            window.critLabel22_8.Text = "Уничтожение (стирание) информации на машинных носителях при их передаче между пользователями, в сторонние организации для ремонта или утилизации, а также контроль уничтожения (стирания)";

            window.labelSectionName23.Content = "Раздел 23. Защита виртуальной инфраструктуры";
            window.critLabel23_1.Text = "Идентификация и аутентификация субъектов и объектов доступа в виртуальной инфраструктуре";
            window.critLabel23_2.Text = "Администрирование учетных записей в виртуальной инфраструктуре";
            window.critLabel23_3.Text = "Регистрация событий безопасности в виртуальной инфраструктуре";
            window.critLabel23_4.Text = "Контроль целостности виртуальной инфраструктуры и ее конфигураций";
            window.critLabel23_5.Text = "Реализация антивирусной защиты в виртуальной инфраструктуре";
            window.critLabel23_6.Text = "Администрирование антивирусной защиты в виртуальной инфраструктуре";
            window.critLabel23_7.Text = "Резервное копирование данных виртуальной инфраструктуры";

            window.labelSectionName24.Content = "Раздел 24. Защита физических носителей информации";
            window.critLabel24_1.Text = "Защита учтенных внешних устройств";
            window.critLabel24_2.Text = "Управление доступом к учтенным внешним устройствам";
            window.critLabel24_3.Text = "Контроль перемещения машинных носителей информации за пределы контролируемой зоны";
            window.critLabel24_4.Text = "Исключение возможности несанкционированного ознакомления с содержанием информации, хранящейся на машинных носителях, и (или) использования носителей информации в иных информационных системах";
            window.critLabel24_5.Text = "Контроль подключения машинных носителей информации";
            window.critLabel24_6.Text = "Затирание остаточной информации";

            window.labelSectionName25.Content = "Раздел 25. Идентификация и аутентификация в ИС";
            window.critLabel25_1.Text = "Идентификация и аутентификация пользователей";
            window.critLabel25_2.Text = "Идентификация и аутентификация процессов";
            window.critLabel25_3.Text = "Идентификация и аутентификация устройств";
            window.critLabel25_4.Text = "Идентификация и аутентификация программного обеспечения";
            window.critLabel25_5.Text = "Идентификация и аутентификация объектов файловой системы";
            window.critLabel25_6.Text = "Идентификация и аутентификация СУБД";
            window.critLabel25_7.Text = "Идентификация и аутентификация к объектам доступа";
            window.critLabel25_8.Text = "Многофакторная аутентификация";
            window.critLabel25_9.Text = "Защита аутентификационной информации";

            window.labelSectionName26.Content = "Раздел 26. Управление доступом";
            window.critLabel26_1.Text = "Администрирование учетных записей (активация, блокирование, уничтожение, назначение прав)";
            window.critLabel26_2.Text = "Защита аутентификационной информации";
            window.critLabel26_3.Text = "Управление потоками информации (фильтрация, маршрутизация, контроль соединение, однонаправленная передача)";
            window.critLabel26_4.Text = "Реализация защищенного удаленного доступа";
            window.critLabel26_5.Text = "Ограничение точек доступа при организации удаленного доступа";
            window.critLabel26_6.Text = "Контроль использования ВУ в информационной системе";
            window.critLabel26_7.Text = "Реализация технологии DPI (глубокое исследование пакетов)";
            window.critLabel26_8.Text = "Минимизация прав пользователей и администраторов";
            window.critLabel26_9.Text = "Ограничение неуспешных попыток входа в информационную систему";
            window.critLabel26_10.Text = "Установка ограничений на количество одновременных сессий";

            window.labelSectionName27.Content = "Раздел 27. Ограничение программной среды";
            window.critLabel27_1.Text = "Управление запуском (обращениями) компонентов программного обеспечения, в том числе определение запускаемых компонентов, настройка параметров запуска компонентов, контроль за запуском компонентов программного обеспечения";
            window.critLabel27_2.Text = "Управление установкой (инсталляцией) компонентов программного обеспечения, в том числе определение компонентов, подлежащих установке, настройка параметров установки компонентов, контроль за установкой компонентов программного обеспечения";
            window.critLabel27_3.Text = "Установка (инсталляция) только разрешенного к использованию программного обеспечения и (или) его компонентов";
            window.critLabel27_4.Text = "Управление временными файлами, в том числе запрет, разрешение, перенаправление записи, удаление временных файлов";

            window.labelSectionName28.Content = "Раздел 28. Обеспечение целостности информационной системы и информации";
            window.critLabel28_1.Text = "Контроль целостности программного обеспечения, включая программное обеспечение средств защиты информации";
            window.critLabel28_2.Text = "Контроль целостности информации, содержащейся в базах данных информационной системы";
            window.critLabel28_3.Text = "Обеспечение возможности восстановления программного обеспечения, включая программное обеспечение средств защиты информации, при возникновении нештатных ситуаций";
            window.critLabel28_4.Text = "Обнаружение и реагирование на поступление в информационную систему незапрашиваемых электронных сообщений (писем, документов) и иной информации, не относящихся к функционированию информационной системы (защита от спама)";
            window.critLabel28_5.Text = "Контроль содержания информации, передаваемой из информационной системы (контейнерный, основанный на свойствах объекта доступа, и контентный, основанный на поиске запрещенной к передаче информации с использованием сигнатур, масок и иных методов), и исключение неправомерной передачи информации из информационной системы";
            window.critLabel28_6.Text = "Ограничение прав пользователей по вводу информации в информационную систему";
            window.critLabel28_7.Text = "Контроль точности, полноты и правильности данных, вводимых в информационную систему";
            window.critLabel28_8.Text = "Контроль ошибочных действий пользователей по вводу и (или) передаче информации и предупреждение пользователей об ошибочных действиях";

            window.labelSectionName29.Content = "Раздел 29. Обеспечение доступности информации";
            window.critLabel29_1.Text = "Использование отказоустойчивых технических средств";
            window.critLabel29_2.Text = "Резервирование технических средств, программного обеспечения, каналов передачи информации, средств обеспечения функционирования информационной системы";
            window.critLabel29_3.Text = "Контроль безотказного функционирования технических средств, обнаружение и локализация отказов функционирования, принятие мер по восстановлению отказавших средств и их тестирование";
            window.critLabel29_4.Text = "Периодическое резервное копирование информации на резервные машинные носители информации";
            window.critLabel29_5.Text = "Обеспечение возможности восстановления информации с резервных машинных носителей информации (резервных копий) в течение установленного временного интервала";
            window.critLabel29_6.Text = "Кластеризация информационной системы и (или) ее сегментов";
            window.critLabel29_7.Text = "Контроль состояния и качества предоставления уполномоченным лицом вычислительных ресурсов (мощностей), в том числе по передаче информации";

            window.labelSectionName30.Content = "Раздел 30. Анализ защищенности информации";
            window.critLabel30_1.Text = "Выявление, анализ и устранение уязвимостей информационной системы";
            window.critLabel30_2.Text = "Контроль установки обновлений программного обеспечения, включая обновление программного обеспечения средств защиты информации";
            window.critLabel30_3.Text = "Контроль работоспособности, параметров настройки и правильности функционирования программного обеспечения и средств защиты информации";
            window.critLabel30_4.Text = "Контроль состава технических средств, программного обеспечения и средств защиты информации";
            window.critLabel30_5.Text = "Контроль правил генерации и смены паролей пользователей, заведения и удаления учетных записей пользователей, реализации правил разграничения доступом, полномочий пользователей в информационной системе";

            window.labelSectionName31.Content = "Раздел 31. Регистрация событий безопасности";
            window.critLabel31_1.Text = "Определение событий безопасности, подлежащих регистрации, и сроков их хранения";
            window.critLabel31_2.Text = "Определение состава и содержания информации о событиях безопасности, подлежащих регистрации";
            window.critLabel31_3.Text = "Сбор, запись и хранение информации о событиях безопасности в течение установленного времени хранения";
            window.critLabel31_4.Text = "Реагирование на сбои при регистрации событий безопасности, в том числе аппаратные и программные ошибки, сбои в механизмах сбора информации и достижение предела или переполнения объема (емкости) памяти";
            window.critLabel31_5.Text = "Мониторинг (просмотр, анализ) результатов регистрации событий безопасности и реагирование на них";
            window.critLabel31_6.Text = "Генерирование временных меток и (или) синхронизация системного времени в информационной системе";
            window.critLabel31_7.Text = "Защита информации о событиях безопасности";
            window.critLabel31_8.Text = "Обеспечение возможности просмотра и анализа информации о действиях отдельных пользователей в информационной системе";

        }

        private void FillTimes()
        {
            dateClassSolutionCreation = new DateTime();
            dateClassSolutionCreation = DateTime.Now;
            string ram = dateClassSolutionCreation.ToString();
            dateSolutionCreation = "" + ram[0] + ram[1] + ram[2] + ram[3] + ram[4] + ram[5] + ram[6] + ram[7] + ram[8] + ram[9];
            if (ram[12] == ':') { timeSolutionCreation = "" + "0" + ram[11] + ram[12] + ram[13] + ram[14] + ram[15] + ram[16] + ram[17]; }
            else { timeSolutionCreation = "" + ram[11] + ram[12] + ram[13] + ram[14] + ram[15] + ram[16] + ram[17] + ram[18]; }
        }

        private void FillProjectUserInfo()
        {
            System.Management.ManagementObjectSearcher searcher = new System.Management.ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            System.Management.ManagementObjectCollection collection = searcher.Get();
            usernameSolutionCreation = (string)collection.Cast<System.Management.ManagementBaseObject>().First()["UserName"];
        }

        public void FillInfoSection()
        {
            string appendedName = StringLibrary.AppendChar(name);

            window.labelSolutionTitle.Content = "Проект \"" + appendedName + "\"";
            window.labelNameSolution.Content = appendedName;
            window.labelPathSolution.Content = path;
            window.labelFullPathSolution.Content = StringLibrary.AppendChar(fullFileName);
            window.labelUserSolution.Content = usernameSolutionCreation;
            window.labelDateSolution.Content = dateSolutionCreation;
            window.labelTimeSolution.Content = timeSolutionCreation;
            window.labelCommentSolution.Content = comment;
        }

        //////////////////////////////////////////////

        public void CreateSolution()
        {
            FillLists();
            FillCriteriasComboboxes();
            FillCriteriasPseudonims();
            FillTimes();
            FillProjectUserInfo();
        }

        public void OpenProject(string path)
        {
            FillLists();
            FillCriteriasGridsTitles();
            FillCriteriasComboboxes();
            FillCriteriasPseudonims();
            SetFullProjectPath(path);
            ReadSolutionFile();
            CountSecurityAssessmentVar1();
        }

        public void FinishToCreateProject()
        {
            string ram;

            ram = window.textBoxProjectName.Text.ToString();
            ram = ram.Trim();
            name = ram;

            ram = window.labelCommentSolutionCreate.Content.ToString();
            ram = ram.Trim();
            comment = ram;
            if (comment.Length == 0) comment = "<пусто>";

            path = window.labelPathSolutionCreate.Content.ToString();
            fullFileName = path + @"\" + name + ".xsf";

            CreateSolutionFile();
            FillCriteriasGridsTitles();
        }

        //////////////////////////////////////////////

        private void CreateSolutionFile()
        {
            using (StreamWriter file = new StreamWriter(fullFileName, false, System.Text.Encoding.Default))
            {
                file.WriteLine("Solution.Name" + " " + name);
                file.WriteLine("Solution.Path" + " " + path);
                file.WriteLine("Solution.Date" + " " + dateSolutionCreation);
                file.WriteLine("Solution.Time" + " " + timeSolutionCreation);
                file.WriteLine("Solution.User" + " " + usernameSolutionCreation);
                file.WriteLine("Solution.Comment" + " " + comment);
                file.WriteLine();

                for (int i = 1; i < criteriasPseudonims.Count; i++)
                {
                    for (int j = 0; j < criteriasPseudonims[i].Count; j++)
                    {
                        file.WriteLine(criteriasPseudonims[i][j] + " " + comboboxList[i][j].SelectedIndex);
                    }
                    file.WriteLine();
                }

            }
        }

        public void UpdateSolutionFile()
        {
            CreateSolutionFile();
        }

        private void ReadSolutionFile()
        {
            if (File.Exists(fullFileName))
            {
                StreamReader file = new StreamReader(fullFileName, System.Text.Encoding.Default);
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    StringLibrary.SplitLine(line, out string parameter, out string value);

                    if (parameter == "Solution.Name")
                    {
                        name = value;
                    }
                    if (parameter == "Solution.Path")
                    {
                        path = value;
                    }
                    if (parameter == "Solution.Date")
                    {
                        dateSolutionCreation = value;
                    }
                    if (parameter == "Solution.Time")
                    {
                        timeSolutionCreation = value;
                    }
                    if (parameter == "Solution.User")
                    {
                        usernameSolutionCreation = value;
                    }
                    if (parameter == "Solution.Comment")
                    {
                        comment = value;
                    }

                    if (parameter == "InformationSecurity.Criteria_1.1") comboboxList[1][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_1.2") comboboxList[1][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_1.3") comboboxList[1][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_1.4") comboboxList[1][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_1.5") comboboxList[1][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_1.6") comboboxList[1][5].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_2.1") comboboxList[2][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_2.2") comboboxList[2][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_2.3") comboboxList[2][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_3.1") comboboxList[3][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.2") comboboxList[3][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.3") comboboxList[3][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.4") comboboxList[3][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.5") comboboxList[3][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.6") comboboxList[3][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.7") comboboxList[3][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_3.8") comboboxList[3][7].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_4.1") comboboxList[4][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_4.2") comboboxList[4][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_4.3") comboboxList[4][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_4.4") comboboxList[4][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_5.1") comboboxList[5][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_5.2") comboboxList[5][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_5.3") comboboxList[5][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_5.4") comboboxList[5][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_5.5") comboboxList[5][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_5.6") comboboxList[5][5].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_6.1") comboboxList[6][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_6.2") comboboxList[6][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_6.3") comboboxList[6][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_6.4") comboboxList[6][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_6.5") comboboxList[6][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_6.6") comboboxList[6][5].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_7.1") comboboxList[7][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_7.2") comboboxList[7][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_7.3") comboboxList[7][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_7.4") comboboxList[7][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_8.1") comboboxList[8][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_8.2") comboboxList[8][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_8.3") comboboxList[8][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_9.1") comboboxList[9][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_9.2") comboboxList[9][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_9.3") comboboxList[9][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_9.4") comboboxList[9][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_9.5") comboboxList[9][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_9.6") comboboxList[9][5].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_10.1") comboboxList[10][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_10.2") comboboxList[10][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_10.3") comboboxList[10][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_10.4") comboboxList[10][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_11.1") comboboxList[11][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_11.2") comboboxList[11][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_11.3") comboboxList[11][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_11.4") comboboxList[11][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_12.1") comboboxList[12][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_12.2") comboboxList[12][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_12.3") comboboxList[12][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_12.4") comboboxList[12][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_12.5") comboboxList[12][4].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_13.1") comboboxList[13][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_13.2") comboboxList[13][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_13.3") comboboxList[13][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_13.4") comboboxList[13][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_14.1") comboboxList[14][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_14.2") comboboxList[14][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_14.3") comboboxList[14][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_15.1") comboboxList[15][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_15.2") comboboxList[15][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_15.3") comboboxList[15][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_16.1") comboboxList[16][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_16.2") comboboxList[16][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_16.3") comboboxList[16][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_17.1") comboboxList[17][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_17.2") comboboxList[17][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_17.3") comboboxList[17][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_18.1") comboboxList[18][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_18.2") comboboxList[18][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_18.3") comboboxList[18][2].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_19.1") comboboxList[19][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_19.2") comboboxList[19][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_19.3") comboboxList[19][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_19.4") comboboxList[19][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_19.5") comboboxList[19][4].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_20.1") comboboxList[20][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.2") comboboxList[20][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.3") comboboxList[20][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.4") comboboxList[20][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.5") comboboxList[20][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.6") comboboxList[20][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.7") comboboxList[20][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.8") comboboxList[20][7].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.9") comboboxList[20][8].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_20.10") comboboxList[20][9].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_21.1") comboboxList[21][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.2") comboboxList[21][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.3") comboboxList[21][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.4") comboboxList[21][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.5") comboboxList[21][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.6") comboboxList[21][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.7") comboboxList[21][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.8") comboboxList[21][7].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.9") comboboxList[21][8].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.10") comboboxList[21][9].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.11") comboboxList[21][10].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.12") comboboxList[21][11].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.13") comboboxList[21][12].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.14") comboboxList[21][13].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.15") comboboxList[21][14].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.16") comboboxList[21][15].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.17") comboboxList[21][16].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.18") comboboxList[21][17].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.19") comboboxList[21][18].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.20") comboboxList[21][19].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.21") comboboxList[21][20].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.22") comboboxList[21][21].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.23") comboboxList[21][22].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.24") comboboxList[21][23].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.25") comboboxList[21][24].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.26") comboboxList[21][25].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.27") comboboxList[21][26].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_21.28") comboboxList[21][27].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_22.1") comboboxList[22][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.2") comboboxList[22][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.3") comboboxList[22][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.4") comboboxList[22][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.5") comboboxList[22][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.6") comboboxList[22][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.7") comboboxList[22][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_22.8") comboboxList[22][7].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_23.1") comboboxList[23][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.2") comboboxList[23][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.3") comboboxList[23][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.4") comboboxList[23][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.5") comboboxList[23][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.6") comboboxList[23][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_23.7") comboboxList[23][6].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_24.1") comboboxList[24][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_24.2") comboboxList[24][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_24.3") comboboxList[24][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_24.4") comboboxList[24][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_24.5") comboboxList[24][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_24.6") comboboxList[24][5].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_25.1") comboboxList[25][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.2") comboboxList[25][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.3") comboboxList[25][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.4") comboboxList[25][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.5") comboboxList[25][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.6") comboboxList[25][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.7") comboboxList[25][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.8") comboboxList[25][7].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_25.9") comboboxList[25][8].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_26.1") comboboxList[26][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.2") comboboxList[26][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.3") comboboxList[26][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.4") comboboxList[26][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.5") comboboxList[26][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.6") comboboxList[26][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.7") comboboxList[26][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.8") comboboxList[26][7].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.9") comboboxList[26][8].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_26.10") comboboxList[26][9].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_27.1") comboboxList[27][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_27.2") comboboxList[27][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_27.3") comboboxList[27][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_27.4") comboboxList[27][3].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_28.1") comboboxList[28][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.2") comboboxList[28][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.3") comboboxList[28][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.4") comboboxList[28][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.5") comboboxList[28][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.6") comboboxList[28][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.7") comboboxList[28][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_28.8") comboboxList[28][7].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_29.1") comboboxList[29][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.2") comboboxList[29][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.3") comboboxList[29][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.4") comboboxList[29][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.5") comboboxList[29][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.6") comboboxList[29][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_29.7") comboboxList[29][6].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_30.1") comboboxList[30][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_30.2") comboboxList[30][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_30.3") comboboxList[30][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_30.4") comboboxList[30][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_30.5") comboboxList[30][4].SelectedIndex = short.Parse(value);

                    if (parameter == "InformationSecurity.Criteria_31.1") comboboxList[31][0].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.2") comboboxList[31][1].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.3") comboboxList[31][2].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.4") comboboxList[31][3].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.5") comboboxList[31][4].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.6") comboboxList[31][5].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.7") comboboxList[31][6].SelectedIndex = short.Parse(value);
                    if (parameter == "InformationSecurity.Criteria_31.8") comboboxList[31][7].SelectedIndex = short.Parse(value);

                }

                file.Close();
            }
        }

        //////////////////////////////////////////////

        public void CountSecurityAssessmentVar1()
        {

            // CRITERIAS ASSESSMENTS
            criteriaAssessment = new List<List<double>>();
            criteriaAssessment.Add(new List<double>()); // NULL


            for (int i = 1; i < comboboxList.Count; i++)
            {
                criteriaAssessment.Add(new List<double>());

                for (int j = 0; j < comboboxList[i].Count; j++)
                {
                    criteriaAssessment[i].Add(comboboxList[i][j].SelectedIndex / (double)(comboboxList[i][j].Items.Count - 1));
                }
                
            }

            // SECTIONS ASSESSMENTS
            sectionAssessment = new List<double>();
            sectionAssessment.Add(0); // NULL

            for (int i = 1; i < comboboxList.Count; i++)
            {
                sectionAssessment.Add(Math.Round(MathLibrary.GetListElementsSum(criteriaAssessment[i]) / criteriaAssessment[i].Count, 3));
            }

            // PART ASSESSMENTS
            partAssessment = new List<double>();
            partAssessment.Add(0); // NULL
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 1, 6) / (6 - 1 + 1));      //  PART 1
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 7, 13) / (13 - 7 + 1));    //  PART 2
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 14, 31) / (31 - 14 + 1));  //  PART 3

            // INFORMATION SYSTEM ASSESSMENT
            generalAssessment = MathLibrary.GetListElementsSum(sectionAssessment) / (sectionAssessment.Count - 1);
            //generalAssessment = MathLibrary.GetListElementsSum(partAssessment) / 3;

        }

        public void CountSecurityAssessmentVar2()
        {

            // CRITERIAS ASSESSMENTS


            // SECTIONS ASSESSMENTS
            sectionAssessment = new List<double>();
            sectionAssessment.Add(0); // NULL
            for (int i = 1; i < comboboxList.Count; i++)
            {
                double est = 0;
                for (int j = 0; j < comboboxList[i].Count; j++)
                {
                    est += comboboxList[i][j].SelectedIndex / (comboboxList[i][j].Items.Count - 1);
                }
                sectionAssessment.Add(Math.Round((est / comboboxList[i].Count), 3));
            }

            // PART ASSESSMENTS
            partAssessment = new List<double>();
            partAssessment.Add(0); // NULL
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 1, 6) / (6 - 1 + 1)); //  PART 1
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 7, 13) / (13 - 7 + 1)); //  PART 2
            partAssessment.Add(MathLibrary.GetListElementsSum(sectionAssessment, 18, 31) / (31 - 18 + 1)); //  PART 3

            // INFORMATION SYSTEM ASSESSMENT
            //generalAssessment = MathLibrary.GetListElementsSum(partAssessment) / 3;
            generalAssessment = MathLibrary.GetListElementsSum(sectionAssessment) / (sectionAssessment.Count - 1);
        }


        public void CreateLabelReport()
        {

            // CLEAR REPORT LABELS
            window.labelReport1.Content = "";
            window.labelReport2.Content = "";
            window.labelReportValue1.Content = "";
            window.labelReportValue2.Content = "";


            // PRINT HEADER
            window.labelReport1.Content += "Количественная оценка защищенности информационной системы" + "\n";
            window.labelReport1.Content += "\n";

            window.labelReportValue1.Content += MathLibrary.RoundTo3(generalAssessment) + "\n";
            window.labelReportValue1.Content += "\n";


            // PRINT PARTS HEADERS
            window.labelReport1.Content += partNames.ElementAt(0) + "\n";
            window.labelReport1.Content += partNames.ElementAt(1) + "\n";
            window.labelReport1.Content += partNames.ElementAt(2) + "\n";
            window.labelReport1.Content += "\n";

            window.labelReportValue1.Content += MathLibrary.RoundTo3(partAssessment[1]) + "\n";
            window.labelReportValue1.Content += MathLibrary.RoundTo3(partAssessment[2]) + "\n";
            window.labelReportValue1.Content += MathLibrary.RoundTo3(partAssessment[3]) + "\n";
            window.labelReportValue1.Content += "\n";


            // PRINT SECTIONS PART
            for (int i = 1; i <= 6; i++) { window.labelReport1.Content += sectionNames.ElementAt(i) + "\n"; }
            window.labelReport1.Content += "\n";
            for (int i = 1; i <= 6; i++) { window.labelReportValue1.Content += sectionAssessment.ElementAt(i) + "\n"; }
            window.labelReportValue1.Content += "\n";

            for (int i = 7; i <= 13; i++) { window.labelReport1.Content += sectionNames.ElementAt(i) + "\n"; }
            window.labelReport1.Content += "\n";
            for (int i = 7; i <= 13; i++) { window.labelReportValue1.Content += sectionAssessment.ElementAt(i) + "\n"; }
            window.labelReportValue1.Content += "\n";

            for (int i = 14; i <= 17; i++) { window.labelReport1.Content += sectionNames.ElementAt(i) + "\n"; }
            window.labelReport1.Content += "\n";
            for (int i = 14; i <= 17; i++) { window.labelReportValue1.Content += sectionAssessment.ElementAt(i) + "\n"; }
            window.labelReportValue1.Content += "\n";

            for (int i = 18; i <= 31; i++) { window.labelReport2.Content += sectionNames.ElementAt(i) + "\n"; }
            window.labelReport2.Content += "\n";
            for (int i = 18; i <= 31; i++) { window.labelReportValue2.Content += sectionAssessment.ElementAt(i) + "\n"; }
            window.labelReportValue2.Content += "\n";

        }

        public void CreateTxtReport(string path)
        {
            if (Directory.Exists(path))
            {
                string fullPath = path + "/Отчет " + name + " " + ".txt";
                using (StreamWriter file = new StreamWriter(fullPath, false, System.Text.Encoding.Default))
                {

                    int maxStringLength = StringLibrary.GetMaxLength(sectionNames);

                    void writeLines(short amount)
                    {
                        for (short i = 0; i < amount; i++) file.WriteLine();
                    }
                    void printCriterias()
                    {
                        maxStringLength = 125;

                        file.WriteLine(window.labelSectionName1.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.1 " + window.critLabel1_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.2 " + window.critLabel1_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.3 " + window.critLabel1_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.4 " + window.critLabel1_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.5 " + window.critLabel1_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("1.6 " + window.critLabel1_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit1_6.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName2.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("2.1 " + window.critLabel2_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit2_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("2.2 " + window.critLabel2_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit2_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("2.3 " + window.critLabel2_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit2_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName3.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.1 " + window.critLabel3_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.2 " + window.critLabel3_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.3 " + window.critLabel3_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.4 " + window.critLabel3_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.5 " + window.critLabel3_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.6 " + window.critLabel3_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.7 " + window.critLabel3_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("3.8 " + window.critLabel3_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit3_8.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName4.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("4.1 " + window.critLabel4_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit4_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("4.2 " + window.critLabel4_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit4_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("4.3 " + window.critLabel4_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit4_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("4.4 " + window.critLabel4_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit4_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName5.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.1 " + window.critLabel5_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.2 " + window.critLabel5_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.3 " + window.critLabel5_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.4 " + window.critLabel5_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.5 " + window.critLabel5_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("5.6 " + window.critLabel5_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit5_6.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName6.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.1 " + window.critLabel6_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.2 " + window.critLabel6_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.3 " + window.critLabel6_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.4 " + window.critLabel6_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.5 " + window.critLabel6_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("6.6 " + window.critLabel6_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit6_6.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName7.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("7.1 " + window.critLabel7_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit7_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("7.2 " + window.critLabel7_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit7_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("7.3 " + window.critLabel7_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit7_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("7.4 " + window.critLabel7_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit7_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName8.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("8.1 " + window.critLabel8_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit8_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("8.2 " + window.critLabel8_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit8_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("8.3 " + window.critLabel8_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit8_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName9.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.1 " + window.critLabel9_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.2 " + window.critLabel9_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.3 " + window.critLabel9_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.4 " + window.critLabel9_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.5 " + window.critLabel9_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("9.6 " + window.critLabel9_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit9_6.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName10.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("10.1 " + window.critLabel10_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit10_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("10.2 " + window.critLabel10_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit10_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("10.3 " + window.critLabel10_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit10_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("10.4 " + window.critLabel10_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit10_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName11.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("11.1 " + window.critLabel11_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit11_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("11.2 " + window.critLabel11_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit11_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("11.3 " + window.critLabel11_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit11_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("11.4 " + window.critLabel11_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit11_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName12.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("12.1 " + window.critLabel12_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit12_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("12.2 " + window.critLabel12_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit12_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("12.3 " + window.critLabel12_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit12_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("12.4 " + window.critLabel12_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit12_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("12.5 " + window.critLabel12_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit12_5.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName13.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("13.1 " + window.critLabel13_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit13_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("13.2 " + window.critLabel13_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit13_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("13.3 " + window.critLabel13_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit13_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("13.4 " + window.critLabel13_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit13_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName14.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("14.1 " + window.critLabel14_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit14_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("14.2 " + window.critLabel14_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit14_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("14.3 " + window.critLabel14_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit14_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName15.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("15.1 " + window.critLabel15_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit15_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("15.2 " + window.critLabel15_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit15_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("15.3 " + window.critLabel15_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit15_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName16.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("16.1 " + window.critLabel16_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit16_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("16.2 " + window.critLabel16_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit16_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("16.3 " + window.critLabel16_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit16_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName17.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("17.1 " + window.critLabel17_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit17_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("17.2 " + window.critLabel17_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit17_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("17.3 " + window.critLabel17_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit17_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName18.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("18.1 " + window.critLabel18_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit18_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("18.2 " + window.critLabel18_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit18_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("18.3 " + window.critLabel18_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit18_3.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName19.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("19.1 " + window.critLabel19_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit19_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("19.2 " + window.critLabel19_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit19_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("19.3 " + window.critLabel19_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit19_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("19.4 " + window.critLabel19_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit19_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("19.5 " + window.critLabel19_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit19_5.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName20.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.1 " + window.critLabel20_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.2 " + window.critLabel20_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.3 " + window.critLabel20_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.4 " + window.critLabel20_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.5 " + window.critLabel20_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.6 " + window.critLabel20_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.7 " + window.critLabel20_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.8 " + window.critLabel20_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_8.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.9 " + window.critLabel20_9.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_9.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("20.10 " + window.critLabel20_10.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit20_10.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName21.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.1 " + window.critLabel21_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.2 " + window.critLabel21_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.3 " + window.critLabel21_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.4 " + window.critLabel21_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.5 " + window.critLabel21_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.6 " + window.critLabel21_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.7 " + window.critLabel21_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.8 " + window.critLabel21_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_8.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.9 " + window.critLabel21_9.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_9.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.10 " + window.critLabel21_10.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_10.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.11 " + window.critLabel21_11.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_11.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.12 " + window.critLabel21_12.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_12.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.13 " + window.critLabel21_13.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_13.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.14 " + window.critLabel21_14.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_14.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.15 " + window.critLabel21_15.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_15.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.16 " + window.critLabel21_16.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_16.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.17 " + window.critLabel21_17.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_17.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.18 " + window.critLabel21_18.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_18.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.19 " + window.critLabel21_19.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_19.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.20 " + window.critLabel21_20.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_20.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.21 " + window.critLabel21_21.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_21.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.22 " + window.critLabel21_22.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_22.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.23 " + window.critLabel21_23.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_23.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.24 " + window.critLabel21_24.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_24.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.25 " + window.critLabel21_25.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_25.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.26 " + window.critLabel21_26.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_26.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.27 " + window.critLabel21_27.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_27.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("21.28 " + window.critLabel21_28.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit21_28.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName22.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.1 " + window.critLabel22_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.2 " + window.critLabel22_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.3 " + window.critLabel22_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.4 " + window.critLabel22_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.5 " + window.critLabel22_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.6 " + window.critLabel22_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.7 " + window.critLabel22_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("22.8 " + window.critLabel22_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit22_8.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName23.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.1 " + window.critLabel23_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.2 " + window.critLabel23_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.3 " + window.critLabel23_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.4 " + window.critLabel23_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.5 " + window.critLabel23_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.6 " + window.critLabel23_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("23.7 " + window.critLabel23_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit23_7.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName24.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.1 " + window.critLabel24_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.2 " + window.critLabel24_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.3 " + window.critLabel24_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.4 " + window.critLabel24_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.5 " + window.critLabel24_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("24.6 " + window.critLabel24_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit24_6.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName25.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.1 " + window.critLabel25_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.2 " + window.critLabel25_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.3 " + window.critLabel25_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.4 " + window.critLabel25_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.5 " + window.critLabel25_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.6 " + window.critLabel25_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.7 " + window.critLabel25_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.8 " + window.critLabel25_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_8.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("25.9 " + window.critLabel25_9.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit25_9.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName26.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.1 " + window.critLabel26_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.2 " + window.critLabel26_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.3 " + window.critLabel26_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.4 " + window.critLabel26_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.5 " + window.critLabel26_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.6 " + window.critLabel26_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.7 " + window.critLabel26_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.8 " + window.critLabel26_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_8.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.9 " + window.critLabel26_9.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_9.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("26.10 " + window.critLabel26_10.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit26_10.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName27.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("27.1 " + window.critLabel27_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit27_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("27.2 " + window.critLabel27_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit27_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("27.3 " + window.critLabel27_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit27_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("27.4 " + window.critLabel27_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit27_4.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName28.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.1 " + window.critLabel28_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.2 " + window.critLabel28_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.3 " + window.critLabel28_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.4 " + window.critLabel28_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.5 " + window.critLabel28_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.6 " + window.critLabel28_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.7 " + window.critLabel28_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("28.8 " + window.critLabel28_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit28_8.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName29.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.1 " + window.critLabel29_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.2 " + window.critLabel29_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.3 " + window.critLabel29_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.4 " + window.critLabel29_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.5 " + window.critLabel29_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.6 " + window.critLabel29_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("29.7 " + window.critLabel29_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit29_7.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName30.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("30.1 " + window.critLabel30_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit30_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("30.2 " + window.critLabel30_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit30_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("30.3 " + window.critLabel30_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit30_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("30.4 " + window.critLabel30_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit30_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("30.5 " + window.critLabel30_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit30_5.SelectedItem).Text);
                        file.WriteLine();

                        file.WriteLine(window.labelSectionName31.Content.ToString());
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.1 " + window.critLabel31_1.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_1.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.2 " + window.critLabel31_2.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_2.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.3 " + window.critLabel31_3.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_3.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.4 " + window.critLabel31_4.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_4.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.5 " + window.critLabel31_5.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_5.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.6 " + window.critLabel31_6.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_6.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.7 " + window.critLabel31_7.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_7.SelectedItem).Text);
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(("31.8 " + window.critLabel31_8.Text.ToString()), maxStringLength) + "	" + ((TextBlock)window.comboboxCrit31_8.SelectedItem).Text);
                        file.WriteLine();
                    }

                    // TITLE
                    file.WriteLine("ОТЧЁТ ПО ОЦЕНКЕ ЗАЩИЩЕННОСТИ ИНФОРМАЦИОННОЙ СИСТЕМЫ");
                    writeLines(2);

                    // PRINT COMMON PROKECT INFO
                    file.WriteLine(StringLibrary.IncreaseStringToMaxLength("Название проекта", maxStringLength) + "\t" + name);
                    file.WriteLine(StringLibrary.IncreaseStringToMaxLength("Дата и время создания", maxStringLength) + "\t" + dateSolutionCreation + " " + timeSolutionCreation);
                    file.WriteLine(StringLibrary.IncreaseStringToMaxLength("Комментарий", maxStringLength) + "\t" + comment);
                    writeLines(2);

                    // PRINT HEADER
                    file.WriteLine(StringLibrary.IncreaseStringToMaxLength("Количественная оценка защищенности информационной системы", maxStringLength) + "\t" + MathLibrary.RoundTo3(generalAssessment));
                    writeLines(2);

                    // PRINT PARTS HEADERS
                    for (int i = 0; i < 3; i++) file.WriteLine(StringLibrary.IncreaseStringToMaxLength(partNames.ElementAt(i), maxStringLength) + "\t" + MathLibrary.RoundTo3(partAssessment[i + 1]));
                    writeLines(2);

                    // PRINT SECTIONS PART
                    for (int i = 1; i <= 31; i++)
                    {
                        file.WriteLine(StringLibrary.IncreaseStringToMaxLength(sectionNames.ElementAt(i), maxStringLength) + "\t" + sectionAssessment.ElementAt(i));
                        if (i == 6 || i == 13 || i == 17 || i == 31) file.WriteLine();
                    }
                    writeLines(2);

                    printCriterias();

                    // OPEN CREATED TXT DOCUMENT
                    System.Diagnostics.Process txtDocument = new System.Diagnostics.Process();
                    txtDocument.StartInfo.FileName = "notepad.exe";
                    txtDocument.StartInfo.Arguments = fullPath;
                    txtDocument.Start();

                }
            }
        }

        public void CreateDocxReport(string path)
        {

            if (Directory.Exists(path))
            {

                DateTime date = new DateTime();
                date = DateTime.Now;
                string documentName = "/Отчет " + name + " " + ".docx";
                string documentFullPath = path + documentName;

                Word.Application application = new Word.Application();
                Word.Document document = application.Documents.Add();

                void WriteTitle(string text)
                {
                    Word.Paragraph localParagraph;
                    localParagraph = application.ActiveDocument.Paragraphs.Add();
                    localParagraph.Range.Font.Name = "Times New Roman";
                    localParagraph.Range.Font.Size = 16;
                    localParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    localParagraph.Range.Text += text;
                }

                void WriteParagraph(string text)
                {
                    string str = text + ";";
                    Word.Paragraph localParagraph;
                    localParagraph = application.ActiveDocument.Paragraphs.Add();
                    localParagraph.Range.Font.Name = "Times New Roman";
                    localParagraph.Range.Font.Size = 12;
                    localParagraph.Range.Text += str;
                }

                void Test()
                {
                    Word.Paragraph localParagraph;
                    localParagraph = application.ActiveDocument.Paragraphs.Add();
                    localParagraph.Range.Font.Name = "Times New Roman";
                    localParagraph.Range.Font.Size = 12;
                    localParagraph.Range.Text = "";
                }

                void WriteLine()
                {
                    Word.Paragraph localParagraph;
                    localParagraph = application.ActiveDocument.Paragraphs.Add();
                    localParagraph.Range.Font.Name = "Times New Roman";
                    localParagraph.Range.Font.Size = 12;
                    localParagraph.Range.Text += "\n";
                }

                void WriteLines(byte linesAmount)
                {
                    for (byte i = 0; i < linesAmount; i++) WriteParagraph("\n");
                }

                ///////////////////////////////////////////////////

                WriteTitle("ОТЧЁТ ПО ОЦЕНКЕ ЗАЩИЩЕННОСТИ ИНФОРМАЦИОННОЙ СИСТЕМЫ");
                WriteLine();

                WriteParagraph("Название проекта: " + name);
                WriteParagraph("Дата и время создания: " + dateSolutionCreation + " " + timeSolutionCreation);
                WriteParagraph("Комментарий: " + comment);
                WriteLine();

                WriteParagraph("Количественная оценка защищенности информационной системы: " + MathLibrary.RoundTo3(generalAssessment));
                for (int i = 0; i < 3; i++) WriteParagraph(partNames.ElementAt(i) + "\t" + MathLibrary.RoundTo3(partAssessment[i + 1]));
                WriteLine();

                for (int i = 1; i <= 31; i++)
                {
                    WriteParagraph(sectionNames.ElementAt(i) + "\t" + sectionAssessment.ElementAt(i));
                    if (i == 6 || i == 13 || i == 17 || i == 31) WriteLine();
                }

                document.SaveAs2(documentFullPath);
                application.Documents.Open(documentFullPath);
                //document.Close();
                //application.Quit();

            }
        }

        //////////////////////////////////////////////

        public void UpdateAsTestFill()
        {

            Random random = new Random();

            int max;
            max = 2;

            window.comboboxCrit1_1.SelectedIndex = random.Next(max);
            window.comboboxCrit1_2.SelectedIndex = random.Next(max);
            window.comboboxCrit1_3.SelectedIndex = random.Next(max);
            window.comboboxCrit1_4.SelectedIndex = random.Next(max);
            window.comboboxCrit1_5.SelectedIndex = random.Next(max);
            window.comboboxCrit1_6.SelectedIndex = random.Next(max);

            max = 3;

            window.comboboxCrit2_1.SelectedIndex = random.Next(max);
            window.comboboxCrit2_2.SelectedIndex = random.Next(max);
            window.comboboxCrit2_3.SelectedIndex = random.Next(max);

            window.comboboxCrit3_1.SelectedIndex = random.Next(max);
            window.comboboxCrit3_2.SelectedIndex = random.Next(max);
            window.comboboxCrit3_3.SelectedIndex = random.Next(max);
            window.comboboxCrit3_4.SelectedIndex = random.Next(max);
            window.comboboxCrit3_5.SelectedIndex = random.Next(max);
            window.comboboxCrit3_6.SelectedIndex = random.Next(max);
            window.comboboxCrit3_7.SelectedIndex = random.Next(max);
            window.comboboxCrit3_8.SelectedIndex = random.Next(max);

            window.comboboxCrit4_1.SelectedIndex = random.Next(max);
            window.comboboxCrit4_2.SelectedIndex = random.Next(max);
            window.comboboxCrit4_3.SelectedIndex = random.Next(max);
            window.comboboxCrit4_4.SelectedIndex = random.Next(max);

            window.comboboxCrit5_1.SelectedIndex = random.Next(max);
            window.comboboxCrit5_2.SelectedIndex = random.Next(max);
            window.comboboxCrit5_3.SelectedIndex = random.Next(max);
            window.comboboxCrit5_4.SelectedIndex = random.Next(max);
            window.comboboxCrit5_5.SelectedIndex = random.Next(max);
            window.comboboxCrit5_6.SelectedIndex = random.Next(max);

            window.comboboxCrit6_1.SelectedIndex = random.Next(max);
            window.comboboxCrit6_2.SelectedIndex = random.Next(max);
            window.comboboxCrit6_3.SelectedIndex = random.Next(max);
            window.comboboxCrit6_4.SelectedIndex = random.Next(max);
            window.comboboxCrit6_5.SelectedIndex = random.Next(max);
            window.comboboxCrit6_6.SelectedIndex = random.Next(max);

            window.comboboxCrit7_1.SelectedIndex = random.Next(max);
            window.comboboxCrit7_2.SelectedIndex = random.Next(max);
            window.comboboxCrit7_3.SelectedIndex = random.Next(max);
            window.comboboxCrit7_4.SelectedIndex = random.Next(max);

            window.comboboxCrit8_1.SelectedIndex = random.Next(max);
            window.comboboxCrit8_2.SelectedIndex = random.Next(max);
            window.comboboxCrit8_3.SelectedIndex = random.Next(max);

            window.comboboxCrit9_1.SelectedIndex = random.Next(max);
            window.comboboxCrit9_2.SelectedIndex = random.Next(max);
            window.comboboxCrit9_3.SelectedIndex = random.Next(max);
            window.comboboxCrit9_4.SelectedIndex = random.Next(max);
            window.comboboxCrit9_5.SelectedIndex = random.Next(max);
            window.comboboxCrit9_6.SelectedIndex = random.Next(max);

            window.comboboxCrit10_1.SelectedIndex = random.Next(max);
            window.comboboxCrit10_2.SelectedIndex = random.Next(max);
            window.comboboxCrit10_3.SelectedIndex = random.Next(max);
            window.comboboxCrit10_4.SelectedIndex = random.Next(max);

            window.comboboxCrit11_1.SelectedIndex = random.Next(max);
            window.comboboxCrit11_2.SelectedIndex = random.Next(max);
            window.comboboxCrit11_3.SelectedIndex = random.Next(max);
            window.comboboxCrit11_4.SelectedIndex = random.Next(max);

            window.comboboxCrit12_1.SelectedIndex = random.Next(max);
            window.comboboxCrit12_2.SelectedIndex = random.Next(max);
            window.comboboxCrit12_3.SelectedIndex = random.Next(max);
            window.comboboxCrit12_4.SelectedIndex = random.Next(max);
            window.comboboxCrit12_5.SelectedIndex = random.Next(max);

            window.comboboxCrit13_1.SelectedIndex = random.Next(max);
            window.comboboxCrit13_2.SelectedIndex = random.Next(max);
            window.comboboxCrit13_3.SelectedIndex = random.Next(max);
            window.comboboxCrit13_4.SelectedIndex = random.Next(max);

            window.comboboxCrit14_1.SelectedIndex = random.Next(max);
            window.comboboxCrit14_2.SelectedIndex = random.Next(max);
            window.comboboxCrit14_3.SelectedIndex = random.Next(max);

            window.comboboxCrit15_1.SelectedIndex = random.Next(max);
            window.comboboxCrit15_2.SelectedIndex = random.Next(max);
            window.comboboxCrit15_3.SelectedIndex = random.Next(max);

            window.comboboxCrit16_1.SelectedIndex = random.Next(max);
            window.comboboxCrit16_2.SelectedIndex = random.Next(max);
            window.comboboxCrit16_3.SelectedIndex = random.Next(max);

            window.comboboxCrit17_1.SelectedIndex = random.Next(max);
            window.comboboxCrit17_2.SelectedIndex = random.Next(max);
            window.comboboxCrit17_3.SelectedIndex = random.Next(max);

            window.comboboxCrit18_1.SelectedIndex = random.Next(max);
            window.comboboxCrit18_2.SelectedIndex = random.Next(max);
            window.comboboxCrit18_3.SelectedIndex = random.Next(max);

            window.comboboxCrit19_1.SelectedIndex = random.Next(max);
            window.comboboxCrit19_2.SelectedIndex = random.Next(max);
            window.comboboxCrit19_3.SelectedIndex = random.Next(max);
            window.comboboxCrit19_4.SelectedIndex = random.Next(max);
            window.comboboxCrit19_5.SelectedIndex = random.Next(max);

            window.comboboxCrit20_1.SelectedIndex = random.Next(max);
            window.comboboxCrit20_2.SelectedIndex = random.Next(max);
            window.comboboxCrit20_3.SelectedIndex = random.Next(max);
            window.comboboxCrit20_4.SelectedIndex = random.Next(max);
            window.comboboxCrit20_5.SelectedIndex = random.Next(max);
            window.comboboxCrit20_6.SelectedIndex = random.Next(max);
            window.comboboxCrit20_7.SelectedIndex = random.Next(max);
            window.comboboxCrit20_8.SelectedIndex = random.Next(max);
            window.comboboxCrit20_9.SelectedIndex = random.Next(max);
            window.comboboxCrit20_10.SelectedIndex = random.Next(max);

            window.comboboxCrit21_1.SelectedIndex = random.Next(max);
            window.comboboxCrit21_2.SelectedIndex = random.Next(max);
            window.comboboxCrit21_3.SelectedIndex = random.Next(max);
            window.comboboxCrit21_4.SelectedIndex = random.Next(max);
            window.comboboxCrit21_5.SelectedIndex = random.Next(max);
            window.comboboxCrit21_6.SelectedIndex = random.Next(max);
            window.comboboxCrit21_7.SelectedIndex = random.Next(max);
            window.comboboxCrit21_8.SelectedIndex = random.Next(max);
            window.comboboxCrit21_9.SelectedIndex = random.Next(max);
            window.comboboxCrit21_10.SelectedIndex = random.Next(max);
            window.comboboxCrit21_11.SelectedIndex = random.Next(max);
            window.comboboxCrit21_12.SelectedIndex = random.Next(max);
            window.comboboxCrit21_13.SelectedIndex = random.Next(max);
            window.comboboxCrit21_14.SelectedIndex = random.Next(max);
            window.comboboxCrit21_15.SelectedIndex = random.Next(max);
            window.comboboxCrit21_16.SelectedIndex = random.Next(max);
            window.comboboxCrit21_17.SelectedIndex = random.Next(max);
            window.comboboxCrit21_18.SelectedIndex = random.Next(max);
            window.comboboxCrit21_19.SelectedIndex = random.Next(max);
            window.comboboxCrit21_20.SelectedIndex = random.Next(max);
            window.comboboxCrit21_21.SelectedIndex = random.Next(max);
            window.comboboxCrit21_22.SelectedIndex = random.Next(max);
            window.comboboxCrit21_23.SelectedIndex = random.Next(max);
            window.comboboxCrit21_24.SelectedIndex = random.Next(max);
            window.comboboxCrit21_25.SelectedIndex = random.Next(max);
            window.comboboxCrit21_26.SelectedIndex = random.Next(max);
            window.comboboxCrit21_27.SelectedIndex = random.Next(max);
            window.comboboxCrit21_28.SelectedIndex = random.Next(max);

            window.comboboxCrit22_1.SelectedIndex = random.Next(max);
            window.comboboxCrit22_2.SelectedIndex = random.Next(max);
            window.comboboxCrit22_3.SelectedIndex = random.Next(max);
            window.comboboxCrit22_4.SelectedIndex = random.Next(max);
            window.comboboxCrit22_5.SelectedIndex = random.Next(max);
            window.comboboxCrit22_6.SelectedIndex = random.Next(max);
            window.comboboxCrit22_7.SelectedIndex = random.Next(max);
            window.comboboxCrit22_8.SelectedIndex = random.Next(max);

            window.comboboxCrit23_1.SelectedIndex = random.Next(max);
            window.comboboxCrit23_2.SelectedIndex = random.Next(max);
            window.comboboxCrit23_3.SelectedIndex = random.Next(max);
            window.comboboxCrit23_4.SelectedIndex = random.Next(max);
            window.comboboxCrit23_5.SelectedIndex = random.Next(max);
            window.comboboxCrit23_6.SelectedIndex = random.Next(max);
            window.comboboxCrit23_7.SelectedIndex = random.Next(max);

            window.comboboxCrit24_1.SelectedIndex = random.Next(max);
            window.comboboxCrit24_2.SelectedIndex = random.Next(max);
            window.comboboxCrit24_3.SelectedIndex = random.Next(max);
            window.comboboxCrit24_4.SelectedIndex = random.Next(max);
            window.comboboxCrit24_5.SelectedIndex = random.Next(max);
            window.comboboxCrit24_6.SelectedIndex = random.Next(max);

            window.comboboxCrit25_1.SelectedIndex = random.Next(max);
            window.comboboxCrit25_2.SelectedIndex = random.Next(max);
            window.comboboxCrit25_3.SelectedIndex = random.Next(max);
            window.comboboxCrit25_4.SelectedIndex = random.Next(max);
            window.comboboxCrit25_5.SelectedIndex = random.Next(max);
            window.comboboxCrit25_6.SelectedIndex = random.Next(max);
            window.comboboxCrit25_7.SelectedIndex = random.Next(max);
            window.comboboxCrit25_8.SelectedIndex = random.Next(max);
            window.comboboxCrit25_9.SelectedIndex = random.Next(max);

            window.comboboxCrit26_1.SelectedIndex = random.Next(max);
            window.comboboxCrit26_2.SelectedIndex = random.Next(max);
            window.comboboxCrit26_3.SelectedIndex = random.Next(max);
            window.comboboxCrit26_4.SelectedIndex = random.Next(max);
            window.comboboxCrit26_5.SelectedIndex = random.Next(max);
            window.comboboxCrit26_6.SelectedIndex = random.Next(max);
            window.comboboxCrit26_7.SelectedIndex = random.Next(max);
            window.comboboxCrit26_8.SelectedIndex = random.Next(max);
            window.comboboxCrit26_9.SelectedIndex = random.Next(max);
            window.comboboxCrit26_10.SelectedIndex = random.Next(max);

            window.comboboxCrit27_1.SelectedIndex = random.Next(max);
            window.comboboxCrit27_2.SelectedIndex = random.Next(max);
            window.comboboxCrit27_3.SelectedIndex = random.Next(max);
            window.comboboxCrit27_4.SelectedIndex = random.Next(max);

            window.comboboxCrit28_1.SelectedIndex = random.Next(max);
            window.comboboxCrit28_2.SelectedIndex = random.Next(max);
            window.comboboxCrit28_3.SelectedIndex = random.Next(max);
            window.comboboxCrit28_4.SelectedIndex = random.Next(max);
            window.comboboxCrit28_5.SelectedIndex = random.Next(max);
            window.comboboxCrit28_6.SelectedIndex = random.Next(max);
            window.comboboxCrit28_7.SelectedIndex = random.Next(max);
            window.comboboxCrit28_8.SelectedIndex = random.Next(max);

            window.comboboxCrit29_1.SelectedIndex = random.Next(max);
            window.comboboxCrit29_2.SelectedIndex = random.Next(max);
            window.comboboxCrit29_3.SelectedIndex = random.Next(max);
            window.comboboxCrit29_4.SelectedIndex = random.Next(max);
            window.comboboxCrit29_5.SelectedIndex = random.Next(max);
            window.comboboxCrit29_6.SelectedIndex = random.Next(max);
            window.comboboxCrit29_7.SelectedIndex = random.Next(max);

            window.comboboxCrit30_1.SelectedIndex = random.Next(max);
            window.comboboxCrit30_2.SelectedIndex = random.Next(max);
            window.comboboxCrit30_3.SelectedIndex = random.Next(max);
            window.comboboxCrit30_4.SelectedIndex = random.Next(max);
            window.comboboxCrit30_5.SelectedIndex = random.Next(max);

            window.comboboxCrit31_1.SelectedIndex = random.Next(max);
            window.comboboxCrit31_2.SelectedIndex = random.Next(max);
            window.comboboxCrit31_3.SelectedIndex = random.Next(max);
            window.comboboxCrit31_4.SelectedIndex = random.Next(max);
            window.comboboxCrit31_5.SelectedIndex = random.Next(max);
            window.comboboxCrit31_6.SelectedIndex = random.Next(max);
            window.comboboxCrit31_7.SelectedIndex = random.Next(max);
            window.comboboxCrit31_8.SelectedIndex = random.Next(max);

        }

        public void ClearCriteriasComboboxes()
        {
            for (int i = 1; i < comboboxList.Count; i++)
            {
                for (int j = 0; j < comboboxList[i].Count; j++)
                {
                    comboboxList[i][j].SelectedIndex = -1;
                }
            }

        }

        public bool IsCriteriasComboboxesSelected()
        {
            for (int i = 1; i < comboboxList.Count; i++)
            {
                for (int j = 0; j < comboboxList[i].Count; j++)
                {
                    if (comboboxList[i][j].SelectedIndex == -1) return false;
                }

            }
            return true;
        }

    }

}