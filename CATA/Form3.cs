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
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Xceed.Words.NET;

namespace CATA
{
    public partial class Form3 : Form
    {

        public string form_text{set { Text = Text +"  |  "+value; }}
        public string combobox {set {comboBox1.Text = value;}}
        public string namber_akt { set { textBox1.Text = value; } }
        public string namber_py { set { textBox4.Text = value; }}
        public string namber_py_dia_1 { set { textBox2.Text = value; } }
        public string namber_py_dia_2 { set { textBox3.Text = value; } }
        public string wey_save { get; set; }
        public string set_seve 
        {
            set
            {
                XDocument doc = XDocument.Load(value);
                textBox1.Text = doc.Element("settings").Element("param1").Value;
                comboBox1.Text = doc.Element("settings").Element("param2").Value;
                textBox2.Text = doc.Element("settings").Element("param4").Value;
                textBox3.Text = doc.Element("settings").Element("param5").Value;
                textBox4.Text = doc.Element("settings").Element("param3").Value;
                textBox5.Text = doc.Element("settings").Element("param6").Value;
                textBox6.Text = doc.Element("settings").Element("param7").Value;
                textBox7.Text = doc.Element("settings").Element("param8").Value;
                textBox8.Text = doc.Element("settings").Element("param9").Value;
                textBox9.Text = doc.Element("settings").Element("param10").Value;
                textBox10.Text = doc.Element("settings").Element("param11").Value;
                textBox11.Text = doc.Element("settings").Element("param12").Value;
                textBox12.Text = doc.Element("settings").Element("param13").Value;
                textBox13.Text = doc.Element("settings").Element("param14").Value;
                textBox14.Text = doc.Element("settings").Element("param15").Value;
                textBox15.Text = doc.Element("settings").Element("param16").Value;
                textBox16.Text = doc.Element("settings").Element("param17").Value;
                checkBox1.Checked = Boolean.Parse(doc.Element("settings").Element("param18").Value);
                checkBox2.Checked = Boolean.Parse(doc.Element("settings").Element("param19").Value);
                checkBox3.Checked = Boolean.Parse(doc.Element("settings").Element("param20").Value);
                pril = doc.Element("settings").Element("param21").Value;

                


                if (pril != "")
                {
                    pictureBox1.Image = System.Drawing.Image.FromFile(wey_True_icon);
                }

                if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
                {
                    //textBox2.Enabled = true;
                    //textBox3.Enabled = true;
                    //textBox4.Enabled = false;

                    textBox2.Visible = true;
                    textBox3.Visible = true;
                    textBox4.Visible = false;

                    label5.Visible = false;
                    label3.Visible = true;
                    label4.Visible = true;

                    this.Text = this.Text + " | АКТ " + textBox1.Text + " | " + comboBox1.Text + " | №№ " + textBox2.Text + " , " + textBox3.Text;

                    if (checkBox1.Checked == true)
                    {
                        if (textBox12.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                        else if (textBox12.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                        if (textBox13.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                        else if (textBox13.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                        textBox12.Enabled = true;
                        textBox13.Enabled = true;
                        textBox11.Enabled = false;
                    }
                }
                else
                {
                    //textBox2.Enabled = false;
                    //textBox3.Enabled = false;
                    //textBox4.Enabled = true;

                    textBox2.Visible = false;
                    textBox3.Visible = false;
                    textBox4.Visible = true;

                    label5.Visible = true;
                    label3.Visible = false;
                    label4.Visible = false;

                    this.Text = this.Text + " | АКТ " + textBox1.Text + " | " + comboBox1.Text + " | № " + textBox4.Text;


                    if (checkBox1.Checked == true)
                    {
                        if (textBox11.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                        else if (textBox11.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                        textBox12.Enabled = false;
                        textBox13.Enabled = false;
                        textBox11.Enabled = true;
                    }
                }
            }

        }
        public string get_clear_tip_py(string s)
        {
            if ((s == "РиМ 384.01/2") || (s == "РиМ 384.02/2"))
            {
                string comboBox1_wey = "";
                if (s == "РиМ 384.01/2") { comboBox1_wey = "РиМ 384.01-2"; }
                else if (s == "РиМ 384.02/2") { comboBox1_wey = "РиМ 384.02-2"; }
                return comboBox1_wey;
            }
            else
            {
                return s;
            }
        }





        string pril = "";
                
        string wey_False_icon = @"Image\False.png";
        string wey_True_icon = @"Image\True.png";
        string wey_Title_image = @"Image\Title.jpg";


        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string wey_TA = Properties.Settings.Default.wey_TA_settings;
        }


        private void comboBox1_TextChanged(Object sender, EventArgs e)
        {

            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                //textBox2.Enabled = true;
                //textBox3.Enabled = true;
                //textBox4.Enabled = false;

                textBox2.Visible = true;
                textBox3.Visible = true;
                textBox4.Visible = false;

                label5.Visible = false;
                label3.Visible = true;
                label4.Visible = true;

                if (checkBox1.Checked == true)
                {
                    if (textBox12.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                    else if (textBox12.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                    if (textBox13.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                    else if (textBox13.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                    textBox12.Enabled = true;
                    textBox13.Enabled = true;
                    textBox11.Enabled = false;
                }
            }
            else
            {
                //textBox2.Enabled = false;
                //textBox3.Enabled = false;
                //textBox4.Enabled = true;

                textBox2.Visible = false;
                textBox3.Visible = false;
                textBox4.Visible = true;

                label5.Visible = true;
                label3.Visible = false;
                label4.Visible = false;

                if (checkBox1.Checked == true)
                {
                    if (textBox11.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                    else if (textBox11.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                    textBox12.Enabled = false;
                    textBox13.Enabled = false;
                    textBox11.Enabled = true;
                }

            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            // путь к документу

            string pathDocument = "";
            string wey_part_TA = "";
            string wey_TA = Properties.Settings.Default.wey_TA_settings;

            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                string comboBox1_wey = "";
                if (comboBox1.Text == "РиМ 384.01/2") { comboBox1_wey = "РиМ 384.01-2"; }
                else if (comboBox1.Text == "РиМ 384.02/2") { comboBox1_wey = "РиМ 384.02-2"; }

                if (wey_part_TA == "")
                {
                    wey_part_TA = wey_TA + @"\АКТ " + textBox1.Text + " " + comboBox1_wey + " №№ " + textBox2.Text + ", " + textBox3.Text;
                }

                DirectoryInfo drInfo = new DirectoryInfo(wey_part_TA);

                if (drInfo.Exists)
                {
                    pathDocument = wey_part_TA + @"\АКТ " + textBox1.Text + " " + comboBox1_wey + " №№ " + textBox2.Text + ", " + textBox3.Text + ".docx";

                }
                else
                {
                    pathDocument = wey_TA + @"\АКТ " + textBox1.Text + " " + comboBox1_wey + " №№ " + textBox2.Text + ", " + textBox3.Text + ".docx";

                }

            }
            else
            {

                if (wey_part_TA == "")
                {
                    wey_part_TA = wey_TA + @"\АКТ " + textBox1.Text + " " + comboBox1.Text + " № " + textBox4.Text;
                }

                DirectoryInfo drInfo = new DirectoryInfo(wey_part_TA);

                if (drInfo.Exists)
                {
                    pathDocument = wey_part_TA + @"\АКТ " + textBox1.Text + " " + comboBox1.Text + " № " + textBox4.Text + ".docx";
                }
                else
                {
                    pathDocument = wey_TA + @"\АКТ " + textBox1.Text + " " + comboBox1.Text + " № " + textBox4.Text + ".docx";
                }

            }


            /*  Создаёт документ и таблюцу в нём. 
                pathDocument - Путь к месту создания документа 
                 
                Внутренние  функции:
                    table_high_voltage - создаёт таблицу для ВЫСОКОВОЛЬТНЫХ счётчиков
                    table_low_voltage - создаёт таблицу для НИЗКОВОЛЬТНЫХ счётчиков
            */
            Creattable(pathDocument);


        }

        private void splitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox1.Checked == true)
            {

                pictureBox8.Visible = true;
                if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
                {
                    textBox12.Enabled = true;
                    textBox13.Enabled = true;
                    textBox11.Enabled = false;

                }
                else if (comboBox1.Text == "")
                {

                }
                else
                {
                    textBox12.Enabled = false;
                    textBox13.Enabled = false;
                    textBox11.Enabled = true;

                }
            }
            else
            {
                pictureBox8.Visible = false;
                textBox12.Enabled = false;
                textBox13.Enabled = false;
                textBox11.Enabled = false;
            }
        }


        private void splitContainer9_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                if (textBox12.Text != "")
                {
                    pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon);
                }
                else if (textBox12.Text == "")
                {
                    pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon);
                }
            }

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

            if (checkBox2.Checked == true)
            {
                pictureBox10.Visible = true;
                textBox15.Enabled = true;
            }
            else if (checkBox2.Checked == false)
            {
                pictureBox10.Visible = false;
                textBox15.Enabled = false;
            }
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            if (textBox16.Text != "") { pictureBox11.Image = System.Drawing.Image.FromFile(wey_True_icon); }
            else if (textBox16.Text == "") { pictureBox11.Image = System.Drawing.Image.FromFile(wey_False_icon); }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            if (OPF.ShowDialog() == DialogResult.OK)
            {
                pril = OPF.FileName;
                pictureBox1.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
        }
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                pictureBox11.Visible = true;
                textBox16.Enabled = true;
            }
            else if (checkBox3.Checked == false)
            {
                pictureBox11.Visible = false;
                textBox16.Enabled = false;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click_1(object sender, EventArgs e)
        {

        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string wey_TA = Properties.Settings.Default.wey_TA_settings;
        }

        private void folderBrowserDialog1_HelpRequest_1(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click_2(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text != "")
            {
                pictureBox2.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox5.Text == "")
            {
                pictureBox2.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text != "")
            {
                pictureBox3.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox6.Text == "")
            {
                pictureBox3.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (textBox8.Text != "")
            {
                pictureBox4.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox8.Text == "")
            {
                pictureBox4.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Text != "")
            {
                pictureBox5.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox7.Text == "")
            {
                pictureBox5.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            if (textBox9.Text != "")
            {
                pictureBox6.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox9.Text == "")
            {
                pictureBox6.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            if (textBox10.Text != "")
            {
                pictureBox7.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox10.Text == "")
            {
                pictureBox7.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if (textBox11.Text != "")
            {
                pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon);
            }
            else if (textBox11.Text == "")
            {
                pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon);
            }
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                if (textBox13.Text != "")
                {
                    pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon);
                }
                else if (textBox13.Text == "")
                {
                    pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon);
                }
            }
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            if (textBox14.Text != "") { pictureBox9.Image = System.Drawing.Image.FromFile(wey_True_icon); }
            else if (textBox14.Text == "") { pictureBox9.Image = System.Drawing.Image.FromFile(wey_False_icon); }
        }

        private void textBox15_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox15.Text != "") { pictureBox10.Image = System.Drawing.Image.FromFile(wey_True_icon); }
            else if (textBox15.Text == "") { pictureBox10.Image = System.Drawing.Image.FromFile(wey_False_icon); }
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {

        }


        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();


        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            toolStripButton3.Enabled = false;
            string wey_TA = Properties.Settings.Default.wey_TA_settings;
            XDocument doc = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XComment("Параметры сохранения проекта"),
                new XElement("settings",
                new XElement("param1", textBox1.Text),
                new XElement("param2", comboBox1.Text),
                new XElement("param3", textBox4.Text),
                new XElement("param4", textBox2.Text),
                new XElement("param5", textBox3.Text),
                new XElement("param6", textBox5.Text),
                new XElement("param7", textBox6.Text),
                new XElement("param8", textBox7.Text),
                new XElement("param9", textBox8.Text),
                new XElement("param10", textBox9.Text),
                new XElement("param11", textBox10.Text),
                new XElement("param12", textBox11.Text),
                new XElement("param13", textBox12.Text),
                new XElement("param14", textBox13.Text),
                new XElement("param15", textBox14.Text),
                new XElement("param16", textBox15.Text),
                new XElement("param17", textBox16.Text),
                new XElement("param18", checkBox1.Checked),
                new XElement("param19", checkBox2.Checked),
                new XElement("param20", checkBox3.Checked),
                new XElement("param21", pril)));
            doc.Save(wey_save);

                       
            string wet_excel = wey_TA + @"\Stack.xlsx";

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

            for (int i = 1;i <= sheet_rows; i++)
            {

                Excel.Range forYach = sheet.Cells[i, 1] as Excel.Range; // Номер АКТа
                

                if (textBox1.Text==forYach.Value2.ToString())
                {
                    Excel.Range forYach1 = sheet.Cells[i, 2] as Excel.Range; // Тип ПУ
                    if (get_clear_tip_py(comboBox1.Text) == forYach1.Value2.ToString())
                    {
                        if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
                        {
                            Excel.Range forYach2 = sheet.Cells[i, 3] as Excel.Range; // Номер ПУ1
                            Excel.Range forYach3 = sheet.Cells[i, 4] as Excel.Range; // Номер ПУ2
                            if (textBox2.Text == forYach2.Value2.ToString())
                            {
                                if (textBox3.Text == forYach3.Value2.ToString())
                                {
                                    DateTime dateOnly = DateTime.Now;
                                    sheet.Cells[i, 5] = String.Format(dateOnly.ToString("dd.MM.yyyy")); // Дата изменений
                                    sheet.Cells[i, 6] = String.Format(dateOnly.ToShortTimeString()); // Время изменений
                                    ex.Application.ActiveWorkbook.SaveAs(wet_excel, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            
                            Excel.Range forYach2 = sheet.Cells[i, 3] as Excel.Range; // Номер ПУ1
                            if (textBox4.Text == forYach2.Value2.ToString())
                            {
                                DateTime dateOnly = DateTime.Now;
                                sheet.Cells[i, 5] = String.Format(dateOnly.ToString("dd.MM.yyyy")); // Дата изменений
                                sheet.Cells[i, 6] = String.Format(dateOnly.ToShortTimeString()); // Время изменений
                                ex.Application.ActiveWorkbook.SaveAs(wet_excel, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                break;
                            }
                        }
                    }

                }
                
            }

            ex.Quit();
                       

            MessageBox.Show(
                "Проект сохранён:\n" + wey_save,
                "Сообщение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            toolStripButton3.Enabled = true;
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

            OpenFileDialog Xl = new OpenFileDialog();
            Xl.Filter = "xlm files (*.xlm)|*.xlm";
            if (Xl.ShowDialog() == DialogResult.OK)
            {
                XDocument doc = XDocument.Load(Xl.FileName);
                textBox1.Text = doc.Element("settings").Element("param1").Value;
                comboBox1.Text = doc.Element("settings").Element("param2").Value;
                textBox2.Text = doc.Element("settings").Element("param3").Value;
                textBox3.Text = doc.Element("settings").Element("param4").Value;
                textBox4.Text = doc.Element("settings").Element("param5").Value;
                textBox5.Text = doc.Element("settings").Element("param6").Value;
                textBox6.Text = doc.Element("settings").Element("param7").Value;
                textBox7.Text = doc.Element("settings").Element("param8").Value;
                textBox8.Text = doc.Element("settings").Element("param9").Value;
                textBox9.Text = doc.Element("settings").Element("param10").Value;
                textBox10.Text = doc.Element("settings").Element("param11").Value;
                textBox11.Text = doc.Element("settings").Element("param12").Value;
                textBox12.Text = doc.Element("settings").Element("param13").Value;
                textBox13.Text = doc.Element("settings").Element("param14").Value;
                textBox14.Text = doc.Element("settings").Element("param15").Value;
                textBox15.Text = doc.Element("settings").Element("param16").Value;
                textBox16.Text = doc.Element("settings").Element("param17").Value;
                checkBox1.Checked = Boolean.Parse(doc.Element("settings").Element("param18").Value);
                checkBox2.Checked = Boolean.Parse(doc.Element("settings").Element("param19").Value);
                checkBox3.Checked = Boolean.Parse(doc.Element("settings").Element("param20").Value);
                pril = doc.Element("settings").Element("param21").Value;


                if (pril != "")
                {
                    pictureBox1.Image = System.Drawing.Image.FromFile(wey_True_icon);
                }

                if ((comboBox1.Text != "РиМ 384.01/2") || (comboBox1.Text != "РиМ 384.02/2"))
                {
                    textBox2.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = true;

                    textBox2.Visible = false;
                    textBox3.Visible = false;
                    textBox4.Visible = true;

                    label5.Visible = true;
                    label3.Visible = false;
                    label4.Visible = false;

                    if (checkBox1.Checked == true)
                    {
                        if (textBox11.Text != "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_True_icon); }
                        else if (textBox11.Text == "") { pictureBox8.Image = System.Drawing.Image.FromFile(wey_False_icon); }

                        textBox12.Enabled = false;
                        textBox13.Enabled = false;
                        textBox11.Enabled = true;
                    }
                }
            }
        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void toolStripSeparator1_Click(object sender, EventArgs e)
        {
                    }
    }
}
