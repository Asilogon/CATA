using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;


namespace CATA
{
    public partial class Form2 : Form
    {
        private string main_project_folder()
        {
            /*
             Определяет тип счётчика. 
             Возвращяет имя основной папки в соостветсвии с типом счётчика.
             Для счётчиков Типа "РиМ 384.0X/2" изменяет символ "/" на "-" в названии.
             */

            string form_name = "Не верный тип ПУ";
            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                string comboBox1_wey = "";
                if (comboBox1.Text == "РиМ 384.01/2") { comboBox1_wey = "РиМ 384.01-2"; }
                else if (comboBox1.Text == "РиМ 384.02/2") { comboBox1_wey = "РиМ 384.02-2"; }

                form_name = @"АКТ " + textBox1.Text + " " + comboBox1_wey + " №№ " + textBox3.Text + ", " + textBox4.Text;
            }
            else
            {
                form_name = @"АКТ " + textBox1.Text + " " + comboBox1.Text + " № " + textBox2.Text;
            }

            return form_name;
        }
        private void creat_seva_fail(string wey_save)
        {
            /*
                Создаёт фаел сохранения xlm-seve
            */

            XDocument xlm = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XComment("Параметры сохранения проекта "),
                new XElement("settings",
                new XElement("param1", textBox1.Text),  // Номер АКТа
                new XElement("param2", comboBox1.Text), // Тип ПУ
                new XElement("param3", textBox2.Text),  // Номер ПУ
                new XElement("param4", textBox3.Text),  // Номер 1-го ДИЭ
                new XElement("param5", textBox4.Text),  // Номер 2-го ДИЭ
                new XElement("param6", null),           // Данные в поле "Отправитель"
                new XElement("param7", null),           // Данные в поле "Дата изгот/пост"
                new XElement("param8", null),           // Данные в поле "Содержание обращения пользователя"
                new XElement("param9", null),           // Данные в поле "Комплектность"
                new XElement("param10", null),          // Данные в поле "Результат осмотра"
                new XElement("param11", null),          // Данные в поле "Проверка функций"
                new XElement("param12", null),          // Данные в поле "Проверка погрешностей для ПУ"
                new XElement("param13", null),          // Данные в поле "Проверка погрешностей для 1-го ДИЭ"
                new XElement("param14", null),          // Данные в поле "Проверка погрешностей для 2-го ДИЭ"
                new XElement("param15", null),          // Данные в поле "Вывод"
                new XElement("param16", null),          // Данные в поле "Вид ремонта"
                new XElement("param17", null),          // Данные в поле "Примечание"
                new XElement("param18", false),          // Чек бокс "Проверка погрешностей" принемает True или False
                new XElement("param19", false),          // Чек бокс "Вид ремонта" принемает True или False
                new XElement("param20", false),          // Чек бокс "Примечание" принемает True или False
                new XElement("param21", null)));        // Путь к "Приложение1"
            
            xlm.Save(wey_save);
        }
        private void new_record_excel(string wey_save)
        {
            string wey_TA = Properties.Settings.Default.wey_TA_settings;
            string wet_excel = wey_TA +@"\Stack.xlsx";

            Excel.Application ex = new Excel.Application();             //Объявляем приложение
            ex.Workbooks.Open(wet_excel,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);


            ex.Visible = false;            //Отобразить Excel
            
            ex.DisplayAlerts = false;            //Отключить отображение окон с сообщениями

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);             //Получаем первый лист документа (счет начинается с 1)

            int sheet_rows = sheet.UsedRange.Rows.Count;                                    //получает количество строк на листе

            sheet.Cells[sheet_rows + 1, 1] = String.Format(textBox1.Text);                  // Номер АКТа

            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                string comboBox1_wey = "";
                if (comboBox1.Text == "РиМ 384.01/2") { comboBox1_wey = "РиМ 384.01-2"; }
                else if (comboBox1.Text == "РиМ 384.02/2") { comboBox1_wey = "РиМ 384.02-2"; }

                sheet.Cells[sheet_rows + 1, 2] = String.Format(comboBox1_wey); // Тип ПУ
                sheet.Cells[sheet_rows + 1, 3] = String.Format(textBox3.Text); // Номер ПУ1
                sheet.Cells[sheet_rows + 1, 4] = String.Format(textBox4.Text); // Номер ПУ2
            }
            else
            {
                sheet.Cells[sheet_rows + 1, 2] = String.Format(comboBox1.Text); // Тип ПУ
                sheet.Cells[sheet_rows + 1, 3] = String.Format(textBox2.Text); // Номер ПУ1
            }

            sheet.Cells[sheet_rows + 1, 7] = String.Format(wey_save); // Путь сохранения

            DateTime dateOnly = DateTime.Now;

            sheet.Cells[sheet_rows + 1, 5] = String.Format(dateOnly.ToString("dd.MM.yyyy")); // Дата изменений
            sheet.Cells[sheet_rows + 1, 6] = String.Format(dateOnly.ToShortTimeString()); // Время изменений


            ex.Application.ActiveWorkbook.SaveAs(wet_excel, Type.Missing,
                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                  Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            ex.Quit();

        }
        public void set_f3(string form_name, string name_wey_save)
        {
            Program.f1.Visible = false;
            Form3 f3 = new Form3();
            f3.form_text = form_name;
            f3.namber_akt = textBox1.Text;
            f3.combobox = comboBox1.Text;
            f3.namber_py = textBox2.Text;
            f3.namber_py_dia_1 = textBox3.Text;
            f3.namber_py_dia_2 = textBox4.Text;
            f3.wey_save = name_wey_save;
            f3.Show();
            Close();

        }

        public Form2()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string wey_TA = Properties.Settings.Default.wey_TA_settings;        // Получает путь к корневой папки
            string name_project_folder = main_project_folder();                 // получает имя проекта
            string main_wey_project = wey_TA + @"\" + name_project_folder;      // Переменная с путём основной папки проекта
            string name_wey_seva_fail = main_wey_project + @"\Save_" + name_project_folder + ".xlm"; // Имя файла сохранения

            DirectoryInfo drInfo = new DirectoryInfo(main_wey_project);         // Проверяет существует ли папка

            if (drInfo.Exists) //Если существует
            {
                if (MessageBox.Show(
                    "Проект уже существует:\n" + drInfo.FullName + "\nИспользовать существующий?",
                    "Создание проекта",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning) == DialogResult.Yes) // Если пользователь нажал да
                {
                    Program.f1.Visible = false;
                    Form3 f3 = new Form3();
                    f3.set_seve = name_wey_seva_fail; //Загружает сохранение
                    f3.wey_save = name_wey_seva_fail;
                    f3.Show();
                    this.Close();
                }

            }
            else //Если не существует
            {
                Directory.CreateDirectory(main_wey_project);                        // Создаёт основную папку проекта
                Directory.CreateDirectory(main_wey_project + @"\Фото");             // Создаёт папку в нутри основной
                Directory.CreateDirectory(main_wey_project + @"\Сопровод");
                Directory.CreateDirectory(main_wey_project + @"\Журналы");

                creat_seva_fail(name_wey_seva_fail);             // Создаёт Save-fail

                new_record_excel(name_wey_seva_fail); //Добавляет в Stack запись с указанным путём сохранения

                set_f3(name_project_folder, name_wey_seva_fail); //Запускает Forme3 с пред настройками  
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                textBox3.Enabled = true;
                textBox4.Enabled = true;
                textBox2.Enabled = false;

            }
            else
            {
                textBox3.Enabled = false;
                textBox4.Enabled = false;
                textBox2.Enabled = true;

            }
        }
    }
}
