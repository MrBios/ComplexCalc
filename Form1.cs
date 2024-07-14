using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using GemBox.Spreadsheet;


namespace ComplexCalc
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        List<complex> complexList = new List<complex>();

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = getNum(textBox1.Text).ToString();
            textBox2.Text = getNum(textBox2.Text).ToString();
            if (test())
            {
                if (complexList.Count == 0)
                {
                    richTextBox1.Text = "";
                    complexList.Add(complexadd(0, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                else
                {
                    complexList.Add(complexadd(1, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        bool test()
        {
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                return true;
            }
            else
            {
                MessageBox.Show("Ошибка заполнения чисел");
                return false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = getNum(textBox1.Text).ToString();
            textBox2.Text = getNum(textBox2.Text).ToString();
            if (test())
            {
                if (complexList.Count == 0)
                {
                    richTextBox1.Text = "";
                    complexList.Add(complexadd(0, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                else
                {
                    complexList.Add(complexadd(2, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = getNum(textBox1.Text).ToString();
            textBox2.Text = getNum(textBox2.Text).ToString();
            if (test())
            {
                if (Double.Parse(textBox1.Text) == 0 && Double.Parse(textBox2.Text) == 0)
                {
                    MessageBox.Show("Вы пытаетесь делить на ноль");
                    return;
                }
                if (complexList.Count == 0)
                {
                    richTextBox1.Text = "";
                    complexList.Add(complexadd(0, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                else
                {
                    complexList.Add(complexadd(3, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = getNum(textBox1.Text).ToString();
            textBox2.Text = getNum(textBox2.Text).ToString();
            if (test())
            {
                if (complexList.Count == 0)
                {
                    richTextBox1.Text = "";
                    complexList.Add(complexadd(0, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                else
                {
                    complexList.Add(complexadd(4, Convert.ToDouble(textBox1.Text), Convert.ToDouble(textBox2.Text)));
                }
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        complex complexadd(int type, double a, double b)
        {
            switch (type)
            {
                case 0:
                    break;
                case 1:
                    richTextBox1.Text += "+\n";
                    break;
                case 2:
                    richTextBox1.Text += "-\n";
                    break;
                case 3:
                    richTextBox1.Text += ":\n";
                    break;
                case 4:
                    richTextBox1.Text += "*\n";
                    break;
            }
            richTextBox1.Text += "z = " + a.ToString() + " + " + b.ToString() + "i\n";
            return (new complex(type, a, b));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            complexList.Clear();
            richTextBox1.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
        }

        Tuple<complex, List<string>> getResult()
        {

            List<string> text = new List<string>();
            text.Add("Решение примера");
            string g = "";
            foreach (var i in complexList)
            {
                switch (i.type)
                {
                    case 0:
                        break;
                    case 1:
                        g += " + ";
                        break;
                    case 2:
                        g += " - ";
                        break;
                    case 3:
                        g += " : ";
                        break;
                    case 4:
                        g += " * ";
                        break;
                }
                g += "(z = " + i.a.ToString() + " + " + i.b.ToString() + "i)";
            }
            text.Add(g);
            text.Add("-----");
            while (ishere(3, 4))
            {
                int ind = -1;
                for (int i = 0; i < complexList.Count; i++)
                {
                    if (complexList[i].type == 3 || complexList[i].type == 4)
                    {
                        ind = i;
                        break;
                    }
                }
                if (ind != -1)
                {
                    Tuple<complex, List<string>> tup;
                    if (complexList[ind].type == 3)
                    {
                        tup = del(complexList[ind - 1], complexList[ind]);
                    }
                    else if (complexList[ind].type == 4)
                    {
                        tup = mult(complexList[ind - 1], complexList[ind]);
                    }
                    else continue;
                    complexList.RemoveAt(ind);
                    complexList[ind - 1] = tup.Item1;
                    foreach (var i in tup.Item2)
                    {
                        text.Add(i);
                    }
                    text.Add("Промежуточный результат расчетов");
                    g = "";
                    foreach (var i in complexList)
                    {
                        switch (i.type)
                        {
                            case 0:
                                break;
                            case 1:
                                g += " + ";
                                break;
                            case 2:
                                g += " - ";
                                break;
                            case 3:
                                g += " : ";
                                break;
                            case 4:
                                g += " * ";
                                break;
                        }
                        g += "(z = " + i.a.ToString() + " + " + i.b.ToString() + "i)";
                    }
                    text.Add(g);
                    text.Add("-----");
                }
            }

            while (ishere(1, 2))
            {
                int ind = -1;
                for (int i = 0; i < complexList.Count; i++)
                {
                    if (complexList[i].type == 1 || complexList[i].type == 2)
                    {
                        ind = i;
                        break;
                    }
                }
                if (ind != -1)
                {
                    Tuple<complex, List<string>> tup;
                    if (complexList[ind].type == 1)
                    {
                        tup = sum(complexList[ind - 1], complexList[ind]);
                    }
                    else if (complexList[ind].type == 2)
                    {
                        tup = minus(complexList[ind - 1], complexList[ind]);
                    }
                    else continue;
                    complexList.RemoveAt(ind);
                    complexList[ind - 1] = tup.Item1;
                    foreach (var i in tup.Item2)
                    {
                        text.Add(i);
                    }
                    text.Add("Промежуточный результат расчетов");
                    g = "";
                    foreach (var i in complexList)
                    {
                        switch (i.type)
                        {
                            case 0:
                                break;
                            case 1:
                                g += " + ";
                                break;
                            case 2:
                                g += " - ";
                                break;
                            case 3:
                                g += " : ";
                                break;
                            case 4:
                                g += " * ";
                                break;
                        }
                        g += "(z = " + i.a.ToString() + " + " + i.b.ToString() + "i)";
                    }
                    text.Add(g);
                    text.Add("-----");
                }
            }
            text.Add("Результат вычислений : z = " + complexList[0].a.ToString() + " + " + complexList[0].b.ToString() + "i");
            complex rrr = complexList[0];
            complexList.Clear();
            return new Tuple<complex, List<string>>(rrr, text);
        }

        [STAThread]
        public void CreateWordDocument(List<string> textList)
        {
            System.Windows.Forms.SaveFileDialog ofd = new System.Windows.Forms.SaveFileDialog();
            ofd.DefaultExt = ".doc";
            ofd.AddExtension = true;
            ofd.Filter = "Word file|*.doc";
            string filePath;

            var re = ofd.ShowDialog();
            if (re == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                MessageBox.Show("Ошибка выбора пути");
                return;
            }

            // Create a new Word document
            using (DocX document = DocX.Create(filePath))
            {
                string all = "";
                // Loop through each element in the list and insert it into the document
                foreach (string text in textList)
                {
                    all += text + "\n";
                    // Insert the text and a new line

                }
                document.InsertParagraph(all).InsertPageBreakAfterSelf();
                // Save the document
                document.Save();
            }

            MessageBox.Show("Файл успешно сохранен");

        }

        Tuple<complex, List<string>> sum(complex n1, complex n2)
        {
            List<string> r = new List<string>();
            r.Add("(z1 = " + n1.a.ToString() + " + " + n1.b.ToString() + "i) + (z2 = " + n2.a.ToString() + " + " + n2.b.ToString() + "i)");
            r.Add("Посчитаем сумму по формуле z = (a1 + a2) + (b1 + b2)i");
            r.Add("z = (" + n1.a.ToString() + " + " + n2.a.ToString() + ") + (" + n1.b.ToString() + " + " + n2.b.ToString() + ")i = " + (n1.a + n2.a).ToString() + " + " + (n1.b + n2.b).ToString() + "i");
            r.Add("-----");
            return new Tuple<complex, List<string>>(new complex(n1.type, n1.a + n2.a, n1.b + n2.b), r);
        }

        Tuple<complex, List<string>> minus(complex n1, complex n2)
        {
            List<string> r = new List<string>();
            r.Add("(z1 = " + n1.a.ToString() + " + " + n1.b.ToString() + "i) - (z2 = " + n2.a.ToString() + " + " + n2.b.ToString() + "i)");
            r.Add("Посчитаем разницу по формуле z = (a1 - a2) + (b1 - b2)i");
            r.Add("z = (" + n1.a.ToString() + " - " + n2.a.ToString() + ") + (" + n1.b.ToString() + " - " + n2.b.ToString() + ")i = " + (n1.a - n2.a).ToString() + " + " + (n1.b - n2.b).ToString() + "i");
            r.Add("-----");
            return new Tuple<complex, List<string>>(new complex(n1.type, n1.a - n2.a, n1.b - n2.b), r);
        }

        Tuple<complex, List<string>> mult(complex n1, complex n2)
        {
            List<string> r = new List<string>();
            r.Add("(z1 = " + n1.a.ToString() + " + " + n1.b.ToString() + "i) * (z2 = " + n2.a.ToString() + " + " + n2.b.ToString() + "i)");
            r.Add("Посчитаем произведение по формуле z = (a1*a2 - b1*b2) + (a1*b2 + a2*b1)i");
            double a = (n1.a * n2.a) - (n1.b * n2.b);
            double b = (n1.a * n2.b) + (n2.a * n1.b);
            r.Add(String.Format("z = ({0}*{1} - {2}*{3}) + ({4}*{5} + {6}*{7})i = {8} + {9}i", n1.a, n2.a, n1.b, n2.b, n1.a, n2.b, n2.a, n1.b, a, b));
            r.Add("-----");
            return new Tuple<complex, List<string>>(new complex(n1.type, a, b), r);
        }

        Tuple<complex, List<string>> del(complex n1, complex n2)
        {
            List<string> r = new List<string>();
            r.Add("(z1 = " + n1.a.ToString() + " + " + n1.b.ToString() + "i) : (z2 = " + n2.a.ToString() + " + " + n2.b.ToString() + "i)");
            r.Add("Посчитаем деление по формуле z = ((a1*a2 + b1*b2) + (a2*b1 - a1*b2))/(a2^2 + b2^2)");
            double a = ((n1.a * n2.a) + (n1.b * n2.b)) / ((n2.a * n2.a) + (n2.b * n2.b));
            double b = ((n2.a * n1.b) - (n1.a * n2.b)) / ((n2.a * n2.a) + (n2.b * n2.b));
            r.Add(String.Format("z = (({0}*{1} + {2}*{3}) + ({1}*{2} - {0}*{3}))/({1}^2 + {3}^2) = {4} + {5}i", n1.a, n2.a, n1.b, n2.b, a, b));
            r.Add("-----");
            return new Tuple<complex, List<string>>(new complex(n1.type, a, b), r);
        }

        bool ishere(int type1, int type2)
        {
            bool result = false;
            foreach (var i in complexList)
            {
                if (i.type == type1 || i.type == type2) result = true;
            }
            return result;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (complexList.Count < 2)
            {
                MessageBox.Show("Пример не полный");
                return;
            }
            var resu = getResult();
            complexList.Clear();
            richTextBox1.Text = "";
            foreach (var i in resu.Item2)
            {
                richTextBox1.Text += i + "\n";
            }
            CreateExcelDocument(resu.Item2);
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (complexList.Count < 2)
            {
                MessageBox.Show("Пример не полный");
                return;
            }
            var resu = getResult();
            complexList.Clear();
            richTextBox1.Text = "";
            foreach (var i in resu.Item2)
            {
                richTextBox1.Text += i + "\n";
            }
            CreateWordDocument(resu.Item2);
        }

        public void CreateExcelDocument(List<string> elements)
        {
            System.Windows.Forms.SaveFileDialog ofd = new System.Windows.Forms.SaveFileDialog();
            ofd.DefaultExt = ".xlsx";
            ofd.AddExtension = true;
            ofd.Filter = "Exel file|*.xlsx";
            string filePath;

            var re = ofd.ShowDialog();
            if (re == DialogResult.OK)
            {
                filePath = ofd.FileName;
            }
            else
            {
                MessageBox.Show("Ошибка выбора пути");
                return;
            }
            // Set the license key for GemBox.Spreadsheet
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            // Create a new Excel workbook
            var workbook = new ExcelFile();

            // Add a new worksheet to the workbook
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // Insert the elements into the cells of the first row
            for (int i = 0; i < elements.Count; i++)
            {
                worksheet.Cells[i, 0].Value = elements[i].Replace("-----", "");
            }

            // Save the workbook to a file
            workbook.Save(filePath);

            MessageBox.Show("Файл успешно сохранен");
        }

        double getNum(string input)
        {
            try
            {
                double num = 0;
                List<string> list = new List<string>();
                list = input.Split(' ').ToList();
                num = Convert.ToDouble(list[0]);
                for (int i = 1; i < list.Count; i += 2)
                {
                    switch (list[i])
                    {
                        case "+":
                            num += Convert.ToDouble(list[i + 1]);
                            break;
                        case "-":
                            num -= Convert.ToDouble(list[i + 1]);
                            break;
                        case "*":
                            num *= Convert.ToDouble(list[i + 1]);
                            break;
                        case "/":
                            if (num == 0 && Convert.ToDouble(list[i + 1]) == 0)
                            {
                                num = 0;
                            }
                            else
                            {
                                if (Convert.ToDouble(list[i + 1]) == 0)
                                {
                                    MessageBox.Show("Вы пытаетесь делить на 0.\nЭто невозможная математическая операция\nрезультат деления будет приравнен нулю");
                                    num = 0;
                                }
                                else
                                {
                                    num /= Convert.ToDouble(list[i + 1]);
                                }
                            }
                            break;
                        case "\\":
                            if (num == 0 && Convert.ToDouble(list[i + 1]) == 0)
                            {
                                num = 0;
                            }
                            else
                            {
                                num /= Convert.ToDouble(list[i + 1]);
                            }
                            break;
                        case ":":
                            if (num == 0 && Convert.ToDouble(list[i + 1]) == 0)
                            {
                                num = 0;
                            }
                            else
                            {
                                num /= Convert.ToDouble(list[i + 1]);
                            }
                            break;
                    }
                }
                return num;
            }
            catch
            {
                MessageBox.Show("Ошибка в определении значений");
                return 0;
            }
        }

        class complex
        {
            //1 - + 2 - - 3 - : 4 - * 0 - nothing
            public int type;
            public double a;
            public double b;
            public complex(int type, double a, double b)
            {
                this.type = type;
                this.a = a;
                this.b = b;
            }
        }
    }
}
