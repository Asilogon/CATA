using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CATA
{
    public partial class Form1 : Form
    {

        string wey_new = "";

        private IEnumerable<Control> GetAllTextBoxControls(Control container)
        {
            List<Control> controlList = new List<Control>();
            foreach (Control c in container.Controls)
            {
                controlList.AddRange(GetAllTextBoxControls(c));
                if (c is TextBox)
                    controlList.Add(c);
            }
            return controlList;
        }

        public void start_excel()

        {
            string wet_excel = @"D:\Технический анализ\Stack.xlsx";

            Excel.Application ex = new Excel.Application();


            ex.Workbooks.Open(wet_excel,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing);

            //Отобразить Excel
            ex.Visible = false;
            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = false;

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            int sheet_rows = sheet.UsedRange.Rows.Count;
           
            Button[] buttins = splitContainer1.Panel1.Controls.OfType<Button>().ToArray();

            for (int i = 1; i < sheet_rows; i++)
            {
                if (i > 6) { break; };
                //Получение одной ячейки как ранга
                Excel.Range forYach = sheet.Cells[i+1, 1] as Excel.Range;
                Excel.Range forYach1 = sheet.Cells[i + 1, 2] as Excel.Range;
                Excel.Range forYach2 = sheet.Cells[i + 1, 3] as Excel.Range;
                Excel.Range forYach3 = sheet.Cells[i + 1, 4] as Excel.Range;
                Excel.Range forYach4 = sheet.Cells[i + 1, 5] as Excel.Range;
                Excel.Range forYach5 = sheet.Cells[i + 1, 6] as Excel.Range;
                Excel.Range forYach6 = sheet.Cells[i + 1, 7] as Excel.Range;
                //Получаем значение из ячейки и преобразуем в строку                
                string bat_text1 = "";
                string bat_text2 = "";
                string bat_text3 = "";

                bat_text1 = "АКТ " + forYach.Value2.ToString();

                if ((forYach1.Value2.ToString() == "РиМ 384.01-2") || (forYach1.Value2.ToString() == "РиМ 384.02-2"))
                { bat_text2 = forYach1.Value2.ToString() + " №№ " + forYach2.Value2.ToString() + ", " + forYach3.Value2.ToString(); }
                else
                { bat_text2 = forYach1.Value2.ToString() + " №" + forYach2.Value2.ToString(); }

                bat_text3 = forYach5.Value2.ToString() +"  "+ forYach4.Value2.ToString();

                wey_new = forYach6.Value2.ToString();
                buttins[6-i].Visible = true;
                buttins[6-i].Text = bat_text1 + "\n" + bat_text2 + "\n" + bat_text3;
                buttins[6-i].Click += new EventHandler(button_Click);
                buttins[6-i].Tag = wey_new;
            }
                                   
            ex.Quit();

        }

        public Form1()
        {
            Program.f1 = this; // теперь f1 будет ссылкой на форму Form1
            InitializeComponent();
            start_excel();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog Xl = new OpenFileDialog();
            Xl.Filter = "xlm files (*.xlm)|*.xlm";
            if (Xl.ShowDialog() == DialogResult.OK)
            {
                Program.f1.Visible = false;
                Form3 f3 = new Form3();
                f3.set_seve = Xl.FileName;
                f3.wey_save = Xl.FileName;
                f3.Show();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
    
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {



        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void button_Click(object sender, EventArgs e)
        {

            var button = sender as Button;
            

            Program.f1.Visible = false;
            Form3 f3 = new Form3();
            f3.set_seve = button.Tag.ToString();
            f3.wey_save = button.Tag.ToString();
            f3.Show();
        }
        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
