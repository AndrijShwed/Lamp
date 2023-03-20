using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Заміна_ламп_вул.осв
{
   
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        

        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            string village = comboBox1.Text;
            switch (village)
            {
                case "Бережниця":
                  comboBox2.Items.Add("Шевченка");
                  comboBox2.Items.Add("Дорошенка");
                  comboBox2.Items.Add("Надбережна");
                  comboBox2.Items.Add("Бандери С.");
                  comboBox2.Items.Add("Рогізнянська");
                  comboBox2.Items.Add("І.Франка");
                  comboBox2.Items.Add("Космонавтів");
                  comboBox2.Items.Add("Нова");
                  comboBox2.Items.Add("Молодіжна");
                  comboBox2.Items.Add("Лісна");
                  comboBox2.Items.Add("Садова");
                  comboBox2.Items.Add("Зелена");
                  comboBox2.Items.Add("Алея Неб.Сотні");
                  comboBox2.Items.Add("Церква-цвинтар");
                    break;
                case "Рогізно":
                    comboBox2.Items.Add("Шевченка");
                    comboBox2.Items.Add("І.Франка");
                    comboBox2.Items.Add("Лесі Українки");
                    comboBox2.Items.Add("Зелена");
                    comboBox2.Items.Add("Садова");
                    comboBox2.Items.Add("Вузька");
                    break;
                case "Заболотівці":
                    comboBox2.Items.Add("Миру");
                    comboBox2.Items.Add("Шевченка");
                    comboBox2.Items.Add("Космонавтів");
                    comboBox2.Items.Add("Героїв України");
                    break;
                case "Журавків":
                    comboBox2.Items.Add("Шевченка");
                    comboBox2.Items.Add("Довбуша-Миру");
                    comboBox2.Items.Add("Б.Хмельницького");
                    comboBox2.Items.Add("Чорновола В.");
                    comboBox2.Items.Add("Лісна");
                    break;
                case "Загурщина":
                    comboBox2.Items.Add("Шевченка");
                    comboBox2.Items.Add("Зелена");
                    break;
                default:
                    return;

            }
        
        }
        class Lamp
        {
            public string Village;
            public string Street;
            public string SupportNumb;
            public string Producer;
            public DateTime Date;
            public int Warranty;
                                  
        }

        class CountLamp
        {
            public string Status;
            public DateTime DateCount;

        }
        List<Lamp> lamps = new List<Lamp>();

        List<CountLamp> countLamps = new List<CountLamp>();
        int pNum = 1;
       
        private void button1_Click(object sender, EventArgs e)
        {
            InputFileLamp();
            InputFileCountLamp();

            if ((textBox1.TextLength == 0) || (textBox3.TextLength == 0) || (textBox4.TextLength == 0)  ||
                (comboBox1.SelectedItem == null) || (comboBox2.SelectedItem == null))
            {
                return;
            }
            
            else
            {
                
                string newSupportNumb = textBox1.Text;
                string newVillage = comboBox1.Text;
                string newStreet = comboBox2.Text;
                string newProducer = textBox4.Text;
                int newWarranty;
                if(Int32.TryParse((textBox3.Text), out int j))
                {
                    newWarranty = j;
                }
                else
                {
                    return;
                }
                int n = 0;
                int k = 1;

                DateTime DateCh = DateTime.Now;

                if (lamps.Count == 0)
                {
                    lamps.Add(new Lamp()
                    {
                        Village = newVillage,
                        Street = newStreet,
                        SupportNumb = newSupportNumb,
                        Date = DateCh,
                        Producer = newProducer,
                        Warranty = newWarranty

                    }) ;
                    countLamps.Add(new CountLamp()
                    {
                        DateCount = DateCh,
                        Status = "First"

                    });

                    listBox1.Items.Add((pNum) + ") с." + comboBox1.Text + ", вул." + comboBox2.Text +
                        ", опора N " + textBox1.Text + "  дата заміни - " + DateCh.ToShortDateString()+
                        ". Виробник - " + newProducer+", гарантія -" + newWarranty + " р.");
                    listBox1.Items.Add("     Лампа замінена вперше.");
                    listBox1.Items.Add(new string('-', 85));
                    pNum++;
                }
                else
                {
                    foreach (var item in lamps)
                    {
                        

                        DateTime dateOld = item.Date;

                        DateTime dateOldWarranty = dateOld.AddYears(item.Warranty);
                        
                        TimeSpan diff = dateOldWarranty - DateCh;

                        int days = Convert.ToInt32(diff.TotalDays);


                        if ((item.Village == newVillage) && (item.Street == newStreet) && (item.SupportNumb == newSupportNumb))
                        {

                            if (item.Date.ToShortDateString() == DateCh.ToShortDateString())
                            {
                                listBox1.Items.Add("На цій опорі сьогогдні вже замінено лампу");
                                return;
                            }
                            
                            if (days > 0)
                            {
                              
                               lamps.Remove(item);
                                lamps.Add(new Lamp()
                                {
                                    Village = newVillage,
                                    Street = newStreet,
                                    SupportNumb = newSupportNumb,
                                    Producer = newProducer,
                                    Date = DateCh,
                                    Warranty = newWarranty

                                });

                                countLamps.Add(new CountLamp()
                                {
                                    DateCount = DateCh,
                                    Status = "Warranty"
                                });

                                listBox1.Items.Add((pNum) + ") с." + comboBox1.Text + ", вул." +
                                    comboBox2.Text + ", опора N " + textBox1.Text + " дата заміни - " +
                                    DateCh.ToShortDateString()+ ". Виробник - " + newProducer + ", гарантія -" + newWarranty + " р.");
                                listBox1.Items.Add("     Лампа, яку замінили, на гарантії, залишилось - " + days +
                                    " днів, встановлена - " + dateOld.ToShortDateString() +
                                    " р." );
                                listBox1.Items.Add(new string('-', 85));
                                pNum++; 
                                break;
                            }
                            else
                            {
                                lamps.Remove(item);
                                lamps.Add(new Lamp()
                                {
                                    Village = newVillage,
                                    Street = newStreet,
                                    SupportNumb = newSupportNumb,
                                    Producer = newProducer,
                                    Date = DateCh,
                                    Warranty = newWarranty
                                });

                                countLamps.Add(new CountLamp()
                                {
                                    DateCount = DateCh,
                                    Status = "WarrantyEnd"
                                });

                                listBox1.Items.Add((pNum) + ") с." + comboBox1.Text + ", вул." +
                                    comboBox2.Text + ", опора N " + textBox1.Text + ", дата заміни - " +
                                    DateCh.ToShortDateString()+ ". Виробник - " + newProducer + ", гарантія -" + newWarranty + " р.");
                                listBox1.Items.Add("     Гарантія у заміненої лампи закінчилась");
                                listBox1.Items.Add(new string('-', 85));
                                pNum++;
                               break;
                            }
                            
                        }
                        if (k == lamps.Count)
                        {             

                            listBox1.Items.Add((pNum) + ") с." + comboBox1.Text + ", вул." +
                                comboBox2.Text + ", опора N " + textBox1.Text + ", дата заміни - " +
                               DateCh.ToShortDateString()+ ". Виробник - " + newProducer + ", гарантія -" + newWarranty + " р.");
                            listBox1.Items.Add("     Лампа замінена вперше.");
                            listBox1.Items.Add(new string('-', 85));
                            n = 1;
                            pNum++;
                        }

                        k++;

                    }


                }
                if (n == 1)
                {
                    lamps.Add(new Lamp()
                    {
                        Village = newVillage,
                        Street = newStreet,
                        SupportNumb = newSupportNumb,
                        Producer = newProducer,
                        Date = DateCh,
                        Warranty = newWarranty
                    });
                    n = 0;
                    countLamps.Add(new CountLamp()
                    {
                        DateCount = DateCh,
                        Status = "First"
                    });
                }


                    comboBox2.Text = null;
                textBox1.Text = null;
            }
            string path = "D:\\Lamp\\Список замінених ламп.txt";
            using (StreamWriter sw = new StreamWriter(path, false))
            {

                foreach (var item in lamps)
                {
                    sw.WriteLine("");
                    sw.WriteLine(item.Village);
                    sw.WriteLine(item.Street);
                    sw.WriteLine(item.SupportNumb);
                    sw.WriteLine(item.Date);
                    sw.WriteLine(item.Warranty);
                    sw.WriteLine(item.Producer);

                }

            }
            string path1 = "D:\\Lamp\\Кількість замінених ламп.txt";
            using (StreamWriter sw = new StreamWriter(path1, false))
            {

                foreach (var item in countLamps)
                {
                    sw.WriteLine("");
                    sw.WriteLine(item.DateCount);
                    sw.WriteLine(item.Status);

                }

            }

        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            button2.BackColor = Color.Red;
        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.MediumTurquoise;
        }

        private void button2_MouseClick(object sender, MouseEventArgs e)
        {
            this.Close();
        }

        Point LastPoint;

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            LastPoint = new Point(e.X, e.Y);
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - LastPoint.X;
                this.Top += e.Y - LastPoint.Y;
            }
        }

        //Print
        private int Lines = 0;
        private string[] result;
        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.Text == null)
            {
                return;
            }
            int countLines = listBox1.Items.Count+15;
            result = new string[countLines];

            // задаємо текст для друку
            result[Lines] = "               Протокол проведення заміни ламп вуличного освітлення ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "Бережницький старостинський округ";
            Lines++;
            result[Lines] = "  "; 
            Lines++;
            result[Lines] = "Дата проведення "+ DateTime.Now.ToShortDateString() + " р.";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            foreach (var item in listBox1.Items)
            {
                result[Lines] = item.ToString();
                Lines++;
            }
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "    Електрик _________________      підпис _____________";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "  ";
            Lines++;
            result[Lines] = "    Староста _________________      підпис _____________";
            Lines ++;

            //об'єкт для друку
            PrintDocument printDocument = new PrintDocument();

            //обробник друку
            printDocument.PrintPage += PrintPageHandler;

            // діалог налаштування друку
            PrintDialog printDialog = new PrintDialog();

            // встановлення об'єкта друку для його налаштувань
            printDialog.Document = printDocument;

            // якщо в діалозі була натиснута клавіша ОК
            if (printDialog.ShowDialog() == DialogResult.OK)
                printDialog.Document.Print(); // друкуємо


        }
        
        int counter = 0;//лічильник строк друку
        int curPage = 1;// поточна сторінка
        // обробник події друку
        void PrintPageHandler(object sender, PrintPageEventArgs e)
        {
            float leftMargin = e.MarginBounds.Left; // відступ зліва у документі
            float topMargin = e.MarginBounds.Top; // відступ зверху в документі
            float yPos = 0; // поточна позиція Y для виведення рядка

            int i = 0;// лічильник рядків на сторінці

            Font myFont = new Font("Arial", 16, FontStyle.Regular, GraphicsUnit.Pixel);
            Font myFont1 = new Font("Arial Italic", 20, FontStyle.Regular, GraphicsUnit.Pixel);

            int nLines = (int)((e.MarginBounds.Height - myFont1.GetHeight(e.Graphics)) / myFont.GetHeight(e.Graphics));
            int nPages = ((Lines - 1) / nLines) + 1;
            if (counter == 0)
            {
                yPos = topMargin;
                e.Graphics.DrawString(result[0], myFont1, Brushes.Black, leftMargin, yPos, new StringFormat());
                counter++;
            }

            while ((i < nLines) && (counter < Lines))
            {
                yPos = topMargin + myFont1.GetHeight(e.Graphics) + i * myFont.GetHeight(e.Graphics);
                // друк рядка result
                e.Graphics.DrawString(result[counter], myFont, Brushes.Black, leftMargin, yPos, new StringFormat());
                i++;
                counter++;
            }
            // визначення чи потрібна ще одна сторінка
            e.HasMorePages = false;

            if (curPage < nPages)
            {
                // надання ще однієї сторінки
                e.HasMorePages = true;
                curPage++;

            }
        }

        int f;
        int w;
        int we;
        private void button5_Click(object sender, EventArgs e)
        {
            f = w = we = 0;

            InputFileCountLamp();

            if ((textBox2.TextLength == 0) || (comboBox3.SelectedItem == null))
            {
                return;
            }
            int Year = Convert.ToInt32(textBox2.Text);
            int Month = Convert.ToInt32(comboBox3.Text);
            
            if ((Year > DateTime.Now.Year )||(Year < 2020))
            {
                listBox2.Items.Add("Некоректно введено рік ");
                return;
            }
            if((Year == DateTime.Now.Year)&&(Month > DateTime.Now.Month))
            {
                listBox2.Items.Add("Некоректно введена дата ");
                return;
            }
            foreach (var item in countLamps)
            {
                if (Convert.ToInt32(comboBox3.Text) == 0)
                {
                    if (item.DateCount.Year == Year)
                    {
                        CountStatistic(item.Status);
                    }
                }
                else if ((item.DateCount.Year == Year) && (item.DateCount.Month == Month))
                {
                    CountStatistic(item.Status);
                   
                }
            }
            
            int all = we + f + w;

            if (Convert.ToInt32(comboBox3.Text) == 0)
            {
                listBox2.Items.Add("Всього за " + Year + " рік замінено " + all +" ламп вуличного освітлення.");
                listBox2.Items.Add(new string('-', 85));
                listBox2.Items.Add("В тому числі : 1) по гарантії - " + w +" шт.");
                listBox2.Items.Add("                       2) гарантія закінчилась у - " + we +" шт.");
                listBox2.Items.Add("                       3) замінено вперше - " + f +" шт.");
            }
            else
            {
                listBox2.Items.Add("Всього за " + Month + "-й місяць " + Year +
                                        " року замінено " + all +" ламп.");
                listBox2.Items.Add(new string('-', 85));
                listBox2.Items.Add("В тому числі : 1) по гарантії - " + w +" шт.");
                listBox2.Items.Add("                       2) гарантія закінчилась у - " + we +" шт.");
                listBox2.Items.Add("                       3) замінено вперше - " + f +" шт.");

            }

        }
        private void CountStatistic(string k)
        {
            switch (k)
            {
                case "First":
                    f++;
                    break;
                case "Warranty":
                    w++;
                    break;
                case "WarrantyEnd":
                    we++;
                    break;
                default:
                    break;
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
        }

        //Print
        private int Lines1 = 0;
        private string[] result1;
        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox2.Text == null)
            {
                return;
            }
            string Year = textBox2.Text;
            string Month = comboBox3.Text;
            if (Convert.ToInt32(comboBox3.Text) == 0)
            {
                Month = " весь ";
            }
            int countLines1 = listBox2.Items.Count + 19;
            result1 = new string[countLines1];
            // задаємо текст для друку
            result1[Lines1] = "            Витяг з бази даних по заміні ламп вуличного освітлення ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "    За період  "+ Month + " . "+ Year +" рік.";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "Бережницький старостинський округ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "Дата " + DateTime.Now.ToShortDateString() + " р.";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            foreach (var item in listBox2.Items)
            {
                result1[Lines1] = item.ToString();
                Lines1++;
            }
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "    Електрик _________________      підпис _____________";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "  ";
            Lines1++;
            result1[Lines1] = "    Староста _________________      підпис _____________";
            Lines1++;

            //объект для печати
            PrintDocument printDocument = new PrintDocument();

            //обработчик события печати
            printDocument.PrintPage += PrintPageHandler1;

            // диалог настройки печати
            PrintDialog printDialog = new PrintDialog();

            // установка объекта печати для его настройки
            printDialog.Document = printDocument;

            // если в диалоге было нажато ОК
            if (printDialog.ShowDialog() == DialogResult.OK)
                printDialog.Document.Print(); // печатаем


        }

        int counter1 = 0;//лічильник строк друку
        int curPage1 = 1;// поточна сторінка
        // обробник події друку
        void PrintPageHandler1(object sender, PrintPageEventArgs e)
        {
            float leftMargin = e.MarginBounds.Left; // відступ зліва у документі
            float topMargin = e.MarginBounds.Top; // відступ зверху в документі
            float yPos = 0; // поточна позиція Y для виведення рядка

            int i = 0;// лічильник рядків на сторінці

            Font myFont = new Font("Arial", 16, FontStyle.Regular, GraphicsUnit.Pixel);
            Font myFont1 = new Font("Arial Italic", 20, FontStyle.Regular, GraphicsUnit.Pixel);

            int nLines = (int)((e.MarginBounds.Height-myFont1.GetHeight(e.Graphics)) / myFont.GetHeight(e.Graphics));
            int nPages = ((Lines1 - 1) / nLines) + 1;

            if (counter1 == 0)
            {
                yPos = topMargin;
                e.Graphics.DrawString(result1[0], myFont1, Brushes.Black, leftMargin, yPos, new StringFormat());
                counter1++;
            }

            while ((i < nLines) && (counter1 < Lines1))
            {

                yPos = topMargin + myFont1.GetHeight(e.Graphics) + i * myFont.GetHeight(e.Graphics);
                // друк рядка result
                e.Graphics.DrawString(result1[counter1], myFont, Brushes.Black, leftMargin, yPos, new StringFormat());
                i++;
                counter1++;

            }
            // визначення чи потрібна ще одна сторінка
            e.HasMorePages = false;

            if (curPage1 < nPages)
            {
                // надання ще однієї сторінки
                e.HasMorePages = true;
                curPage1++;

            }
        }

        private void InputFileLamp()
        {
            lamps.Clear();
 
            string path = "D:\\Lamp\\Список замінених ламп.txt";
            using (FileStream file = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    string line;

                    while (sr.ReadLine() != null)
                    {
                        line = sr.ReadLine();
                        string newVillage1 = line;
                        line = sr.ReadLine();
                        string newStreet1 = line;
                        line = sr.ReadLine();
                        string newSupportNumb1 = line;
                        line = sr.ReadLine();
                        DateTime newDate1 = Convert.ToDateTime(line);
                        line = sr.ReadLine();
                        int Warranty1 = int.Parse(line);
                        line = sr.ReadLine();
                        string newProducer1 = line;
                   


                        lamps.Add(new Lamp()
                        {
                            Village = newVillage1,
                            Street = newStreet1,
                            SupportNumb = newSupportNumb1,
                            Date = newDate1,
                            Warranty = Warranty1,
                            Producer = newProducer1

                        });

                    }

                }
            }
        }

        private void InputFileCountLamp()
        {
            countLamps.Clear();

            string path1 = "D:\\Lamp\\Кількість замінених ламп.txt";
            using (FileStream file = new FileStream(path1, FileMode.OpenOrCreate))
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    string line;

                    while (sr.ReadLine() != null)
                    {
                       
                        line = sr.ReadLine();
                        DateTime newDate1Count = Convert.ToDateTime(line);
                        line = sr.ReadLine();
                        string newcountStatus1 = line;


                        countLamps.Add(new CountLamp()
                        { 
                            DateCount = newDate1Count,
                            Status = newcountStatus1

                        });

                    }

                }
            }
        }
    }
    
}
