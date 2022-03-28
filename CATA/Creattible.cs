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

using Xceed.Words.NET;

//Выделите фрагмент кода, который хотите закоментировать и нажмите Ctrl + K, C (удерживая Ctrl нажать K, затем C). Для снятия комментариев нужно выделить закомментированный кусок и нажать Ctrl + K, U.


namespace CATA
{
    public partial class Form3 : Form
    {
        public void Creattable(string pathDocument)
        {
            // создаём документ
            DocX document = DocX.Create(pathDocument);

            // Изменение размера полей. Пропорция полей (100*X)/3.53 , Где Х необходимое количество СМ в word
            document.MarginTop = 24F;
            document.MarginLeft = 42.5F;
            document.MarginRight = 28.32F;
            document.MarginBottom = 35.41F;

            // Нумирация страниц



            document.AddFooters();
            document.Footers.First.PageNumbers = true;
            document.DifferentOddAndEvenPages = false;

            Paragraph paragraph_colun = document.Footers.First
                .InsertParagraph("Лист ")
                .Font("Times New Roman")
                .FontSize(12);
            paragraph_colun.AppendPageNumber(PageNumberFormat.normal);

            paragraph_colun
                .Append(" Листов ")
                .Font("Times New Roman")
                .FontSize(12);

            paragraph_colun.AppendPageCount(PageNumberFormat.normal);
            paragraph_colun.Alignment = Alignment.right;


            document.Footers.Even.InsertParagraph(paragraph_colun);
            document.Footers.Odd.InsertParagraph(paragraph_colun);



            // загрузка изображения

            Xceed.Words.NET.Image image = document.AddImage(wey_Title_image);
            Picture im = image.CreatePicture();
            im.Width = 681;
            im.Height = 121;

            // создание параграфа
            Paragraph paragraph1 = document.InsertParagraph();
            // вставка изображения в параграф
            paragraph1.AppendPicture(im);
            paragraph1.Append("\n");

            // выравнивание параграфа по центру
            paragraph1.Alignment = Alignment.center;

            Paragraph paragraph2 = document.InsertParagraph("УТВЕРЖДАЮ")
                .Font("Times New Roman")
                .FontSize(14);
            paragraph2.IndentationBefore = 10.70F;
            paragraph2.Alignment = Alignment.center;

            Paragraph paragraph3 = document.InsertParagraph("Заместитель генерального директора по качеству АО «РиМ»\n________________Вологдин А.В.\n«___» _________________ 2022г.")
                .Font("Times New Roman")
                .FontSize(12);
            paragraph3.IndentationBefore = 10.94F;

            if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
            {
                table_high_voltage(document);
            }
            else
            {
                table_low_voltage(document);
            }

            if (checkBox3.Checked == true)
            {
                Paragraph paragraph6 = document
                    .InsertParagraph("\nПримечание: ")
                    .Font("Times New Roman")
                    .FontSize(12)
                    .Bold();
                paragraph6.Alignment = Alignment.left;
                paragraph6.Append(textBox16.Text).Font("Times New Roman").FontSize(12);

            }

            // Создание второй таблице
            Table table1 = document.AddTable(3, 2);// (строки,столбцы)
            table1.Alignment = Alignment.center;     // располагаем таблицу по центру
            table1.Design = TableDesign.TableGrid;   // меняем стандартный дизайн таблицы

            table1.Rows[0].Cells[0].Paragraphs[0]
                .Append("Члены комиссии:")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.left;
            table1.Rows[1].Cells[0].Paragraphs[0]
                .Append("Главный конструктор направления АИИСКУЭ")
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.left;
            table1.Rows[2].Cells[0].Paragraphs[0]
                .Append("Главный метролог")
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.left;
            table1.Rows[1].Cells[1].Paragraphs[0]
                .Append("А.В. Лапчук")
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.right;
            table1.Rows[2].Cells[1].Paragraphs[0]
                .Append("П.С. Утовка")
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.right;

            table1.Rows[0].Cells[1].Width = 300;
            table1.Rows[0].Cells[0].Width = 500;

            document.InsertParagraph().InsertTableAfterSelf(table1);

            document.InsertSection();

            document.MarginTop = 24F;
            document.MarginLeft = 42.5F;
            document.MarginRight = 28.32F;
            document.MarginBottom = 35.41F;

            if (pril != "")
            {
                Paragraph paragraph_p1 = document.InsertParagraph("Приложение 1\n")
                    .Font("Times New Roman")
                    .Bold()
                    .FontSize(12);
                paragraph_p1.Alignment = Alignment.center;


                // вставляет изображение
                Xceed.Words.NET.Image image1 = document.AddImage(pril);
                Picture im1 = image1.CreatePicture();
                im1.Width = 700;
                im1.Height = 950;


                // создание параграфа
                Paragraph paragraph10 = document.InsertParagraph();
                // вставка изображения в параграф
                paragraph10.AppendPicture(im1);

                Paragraph paragraph_p2 = document.InsertParagraph("Рисунок 1 – Сопроводительное письмо")
                    .Font("Times New Roman")
                    .FontSize(12);
                paragraph_p2.Alignment = Alignment.center;

            }

            // сохраняем документ
            document.Save();
        }



        public void table_high_voltage(DocX document)
        {

            //Создание шапки для ИПУЭ 
            Paragraph paragraph4 = document.InsertParagraph("\nАКТ № " + textBox1.Text + "\nТехнический анализ ИПУЭ " + comboBox1.Text + " ДИЭ зав. №№" + textBox2.Text + ", " + textBox3.Text)
                    .Font("Times New Roman")
                    .FontSize(12);
            paragraph4.Alignment = Alignment.center;


            //Создание таблицы для ИПУЭ.     
            Table table = document.AddTable(10, 3);// (строки,столбцы)
            table.Alignment = Alignment.center;     // располагаем таблицу по центру
            table.Design = TableDesign.TableGrid;   // меняем стандартный дизайн таблицы

            // 1 строка
            table.Rows[0].Cells[0].Paragraphs[0]
                .Append("Наименование")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;


            table.Rows[0].MergeCells(1, 2); //Объеденяет 1 и 2 столбец в 0 строке

            table.Rows[0].Cells[1].Paragraphs[0]
                .Append("Содержание")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[0].Cells[1].Width = 500;

            // 2 строка
            table.Rows[1].MergeCells(1, 2); //Объеденяет 1 и 2 столбец в 0 строке

            table.Rows[1].Cells[0].Paragraphs[0]
                .Append("Отправитель")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[1].Cells[1].Paragraphs[0]
                .Append(textBox5.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            // 3 строка
            table.Rows[2].MergeCells(1, 2); //Объеденяет 1 и 2 столбец в 0 строке

            table.Rows[2].Cells[0].Paragraphs[0]
                .Append("Дата изгот./пост.")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[2].Cells[1].Paragraphs[0]
                .Append(textBox6.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            // 4 строка
            table.Rows[3].MergeCells(1, 2); //Объеденяет 1 и 2 столбец в 0 строке

            table.Rows[3].Cells[0].Paragraphs[0]
                .Append("Комплектность")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[3].Cells[1].Paragraphs[0]
                .Append(textBox8.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            // 5 строка
            table.Rows[4].MergeCells(1, 2); //Объеденяет 1 и 2 столбец в 0 строке

            table.Rows[4].Cells[0].Paragraphs[0]
                .Append("Содержание обращения потребителя")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[4].Cells[1].Paragraphs[0]
                .Append(textBox7.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table.Rows[4].Cells[0].Width = 250;

            // 6-8 строка                      
            table.Rows[5].Cells[1].Paragraphs[0]
                .Append("Зав. № ДИЭ")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;


            table.Rows[5].Cells[0].Paragraphs[0]
                .Append("Результаты внешнего осмотра")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[6].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                .Append(textBox2.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[7].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                .Append(textBox3.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[5].Cells[2].Paragraphs[0]
                .Append(textBox9.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table.Rows[5].Cells[2].Width = 800;

            table.MergeCellsInColumn(0, 5, 7); // объединяет ячейке 5,7 в колонке 0 
            table.MergeCellsInColumn(2, 5, 7); // объединяет ячейке 5,7 в колонке 2 

            // 9-10 строка  
            table.Rows[8].Cells[0].Paragraphs[0]
                .Append("Результат проверки функционирования и обследования")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[8].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                .Append(textBox2.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[9].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                .Append(textBox3.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[8].Cells[2].Paragraphs[0]
                .Append(textBox10.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table.MergeCellsInColumn(0, 8, 9); // объединяет ячейке 8,9 в колонке 0 
            table.MergeCellsInColumn(2, 8, 9); // объединяет ячейке 8,9 в колонке 2 



            if (checkBox1.Checked == true) // Создание строки Проверка погрешности
            {
                table.InsertRow();
                table.InsertRow();
                int row = table.RowCount;

                table.Rows[row - 2].Cells[0].Paragraphs[0]
                    .Append("Проверка погрешности")
                    .Bold()
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;

                table.Rows[row - 2].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                    .Append(textBox2.Text)
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;

                table.Rows[row - 1].Cells[1].Paragraphs[0] // Вставляет номер счётчика
                    .Append(textBox3.Text)
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;

                table.MergeCellsInColumn(0, row - 2, row - 1);

                if ((comboBox1.Text == "РиМ 384.01/2") || (comboBox1.Text == "РиМ 384.02/2"))
                {
                    table.Rows[row - 2].Cells[2].Paragraphs[0]
                        .Append(textBox12.Text)
                        .Font("Times New Roman")
                        .FontSize(11)
                        .Alignment = Alignment.both;

                    table.Rows[row - 1].Cells[2].Paragraphs[0]
                        .Append(textBox13.Text)
                        .Font("Times New Roman")
                        .FontSize(11)
                        .Alignment = Alignment.both;
                }
            }
            // Строка Вывода
            table.InsertRow();
            int row1 = table.RowCount;

            table.Rows[row1 - 1].Cells[0].Paragraphs[0]
                .Append("Вывод")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table.Rows[row1 - 1].Cells[1].Paragraphs[0]
                .Append(textBox14.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table.Rows[row1 - 1].MergeCells(1, 2);

            // Строка Вид ремонта

            if (checkBox2.Checked == true)
            {
                table.InsertRow();
                int row2 = table.RowCount;

                table.Rows[row2 - 1].Cells[0].Paragraphs[0]
                    .Append("Вид ремонта")
                    .Bold()
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;

                table.Rows[row2 - 1].Cells[1].Paragraphs[0]
                    .Append(textBox15.Text)
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.both;

                table.Rows[row2 - 1].MergeCells(1, 2);
            }


            document.InsertParagraph().InsertTableAfterSelf(table);     // создаём параграф и вставляем таблицу
                                                                        // Строка примечания
        }
        public void table_low_voltage(DocX document)
        {
            //Создание шапки для ПУ 
            Paragraph paragraph4 = document.InsertParagraph("\nАКТ № " + textBox1.Text + "\nТехнический анализ счётчика " + comboBox1.Text + " зав. № " + textBox4.Text)
                .Font("Times New Roman")
                .FontSize(14);
            paragraph4.Alignment = Alignment.center;

            //Создание таблицы для ПУ.     
            Table table2 = document.AddTable(7, 2);// (строки,столбцы)
            table2.Alignment = Alignment.center;     // располагаем таблицу по центру
            table2.Design = TableDesign.TableGrid;   // меняем стандартный дизайн таблицы

            // 1 строка Наименование Содержание
            table2.Rows[0].Cells[0].Paragraphs[0]
                .Append("Наименование")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;


            table2.Rows[0].Cells[1].Paragraphs[0]
                .Append("Содержание")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[0].Cells[1].Width = 500;

            // 2 строка Отправитель        
            table2.Rows[1].Cells[0].Paragraphs[0]
                .Append("Отправитель")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[1].Cells[1].Paragraphs[0]
                .Append(textBox5.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            // 3 строка Дата изгот./пост.
            table2.Rows[2].Cells[0].Paragraphs[0]
                .Append("Дата изгот./пост.")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[2].Cells[1].Paragraphs[0]
                .Append(textBox6.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            // 4 строка Комплектность
            table2.Rows[3].Cells[0].Paragraphs[0]
                .Append("Комплектность")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[3].Cells[1].Paragraphs[0]
                .Append(textBox8.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            // 5 строка Содержание обращения потребителя
            table2.Rows[4].Cells[0].Paragraphs[0]
                .Append("Содержание обращения потребителя")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[4].Cells[1].Paragraphs[0]
                .Append(textBox7.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table2.Rows[4].Cells[0].Width = 250;

            // 6 строка Результаты внешнего осмотра
            table2.Rows[5].Cells[0].Paragraphs[0]
                .Append("Результаты внешнего осмотра")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[5].Cells[1].Paragraphs[0]
                .Append(textBox9.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            table2.Rows[5].Cells[1].Width = 800;

            // 7 строка Результат проверки функционирования и обследования
            table2.Rows[6].Cells[0].Paragraphs[0]
                .Append("Результат проверки функционирования и обследования")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[6].Cells[1].Paragraphs[0]
                .Append(textBox10.Text.ToString())
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            if (checkBox1.Checked == true) // Создание строки Проверка погрешности
            {
                table2.InsertRow();
                int row = table2.RowCount;

                table2.Rows[row - 1].Cells[0].Paragraphs[0]
                    .Append("Проверка погрешности")
                    .Bold()
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;



                if ((comboBox1.Text != "РиМ 384.01/2") || (comboBox1.Text != "РиМ 384.02/2"))
                {
                    table2.Rows[row - 1].Cells[1].Paragraphs[0]
                        .Append(textBox11.Text)
                        .Font("Times New Roman")
                        .FontSize(11)
                        .Alignment = Alignment.both;
                }
            }

            // Строка Вывода
            table2.InsertRow();
            int row1 = table2.RowCount;

            table2.Rows[row1 - 1].Cells[0].Paragraphs[0]
                .Append("Вывод")
                .Bold()
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.center;

            table2.Rows[row1 - 1].Cells[1].Paragraphs[0]
                .Append(textBox14.Text)
                .Font("Times New Roman")
                .FontSize(11)
                .Alignment = Alignment.both;

            // Строка Вид ремонта

            if (checkBox2.Checked == true)
            {
                table2.InsertRow();
                int row2 = table2.RowCount;

                table2.Rows[row2 - 1].Cells[0].Paragraphs[0]
                    .Append("Вид ремонта")
                    .Bold()
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.center;

                table2.Rows[row2 - 1].Cells[1].Paragraphs[0]
                    .Append(textBox15.Text)
                    .Font("Times New Roman")
                    .FontSize(11)
                    .Alignment = Alignment.both;

            }

            document.InsertParagraph().InsertTableAfterSelf(table2);
        }

    }
}






