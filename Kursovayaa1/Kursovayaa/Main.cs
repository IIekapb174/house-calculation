using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Kursovayaa
{
    public partial class Main : Form
    {
        int k = 0; //листать картинки
        double ploshad;

        public Main()
        {
            InitializeComponent();
            LoadData();
        }


        private void LoadData() // загрузить бд в прайс лист
        {
            string ConnectString = "Data Source=.\\SQLEXPRESS;Initial Catalog=Kursovaya;" +
            "Integrated Security=true;"; //подключение к бд

            SqlConnection myConnection = new SqlConnection(ConnectString);
            myConnection.Open();

            string query = "SELECT * FROM products ORDER BY products_category, products_price ";

            SqlCommand command = new SqlCommand(query, myConnection); //запрос к бд
            SqlDataReader reader = command.ExecuteReader(); // чтение данных 

            List<string[]> data = new List<string[]>(); // загружать данные в список 
            while (reader.Read()) // построчно считывать данные создавая новый элемент
            {
                data.Add(new string[4]);

                data[data.Count - 1][0] = reader[0].ToString();
                data[data.Count - 1][1] = reader[1].ToString();
                data[data.Count - 1][2] = reader[2].ToString();
                data[data.Count - 1][3] = reader[3].ToString();
            }

            reader.Close();
            myConnection.Close();

            foreach (string[] s in data) // добавить новые строки в dataGridView
                dataGridView1.Rows.Add(s);

            TableDesign(dataGridView1);
        }

        void SaveTable(DataGridView What_Save) // Сохранить в excel
        {

            if (System.IO.File.Exists(@"Save.xlsx"))
            {
                System.IO.File.Delete(@"Save.xlsx");
            }

            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + "Save.xlsx";

            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            for (int i = 1; i < What_Save.RowCount + 1; i++)
            {
                for (int j = 1; j < What_Save.ColumnCount + 1; j++)
                {
                    worksheet.Rows[i].Columns[j] = What_Save.Rows[i - 1].Cells[j - 1].Value;
                }
            }

            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();

        }
        void TableDesign(DataGridView dgv) // дизайн таблицы
        {
            dgv.BorderStyle = BorderStyle.None;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238, 239, 249);
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.DefaultCellStyle.SelectionBackColor = Color.DarkTurquoise;
            dgv.DefaultCellStyle.SelectionForeColor = Color.WhiteSmoke;
            dgv.BackgroundColor = Color.White;

            dgv.EnableHeadersVisualStyles = false;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(20, 25, 72);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        }

        private void autor_Click(object sender, EventArgs e)
        {
            Form autor = new Autor();
            autor.Show();
        }

        private void info_Click(object sender, EventArgs e)
        {
            Form info = new Info();
            info.Show();
        }

        private void Main_Load(object sender, EventArgs e)
        {
          
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tolshinaFU_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void otdelka_Click(object sender, EventArgs e)
        {

        }

        private void Lentochniy_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e) // листать картинки
        {
            k++;
            if ( k == 0)
            {
                gotovie.Visible = true;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = false;
            }
            if (k == 1 )
            {
                gotovie.Visible = false;
                gotovie1.Visible = true;
                gotovie2.Visible = false;
                gotovie3.Visible = false;
            }
            if (k == 2)
            {
                gotovie.Visible = false;
                gotovie1.Visible = false;
                gotovie2.Visible = true;
                gotovie3.Visible = false;
            }
            if (k == 3)
            {
                gotovie.Visible = false;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = true;
            }
            if (k > 3)
            {
                k = 0;
                gotovie.Visible = true;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = false;
            }



        }

        private void back_Click(object sender, EventArgs e)
        {
            k--;
            if (k == 0)
            {
                gotovie.Visible = true;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = false;
            }
            if (k == 1)
            {
                gotovie.Visible = false;
                gotovie1.Visible = true;
                gotovie2.Visible = false;
                gotovie3.Visible = false;
            }
            if (k == 2)
            {
                gotovie.Visible = false;
                gotovie1.Visible = false;
                gotovie2.Visible = true;
                gotovie3.Visible = false;
            }
            if (k == 3)
            {
                gotovie.Visible = false;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = true;
            }
            if ( k < 0)
            {
                k = 3;
                gotovie.Visible = false;
                gotovie1.Visible = false;
                gotovie2.Visible = false;
                gotovie3.Visible = true;
            }

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void PictureBox4_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show(
                "• Cоставление проектной документации и расчет мощности;" + '\n' +
                "• Проверка электролиний, подбор и комплектация электрооборудования;" + '\n' +
                "• Монтаж электропроводки распределительных и групповых линий;" + '\n' +
                "• Установка электрического щитового оборудования под ключ;" + '\n' +
                "• Монтаж освещения.",
                "Об услуге",
                MessageBoxButtons.OK,
                MessageBoxIcon.None,
                MessageBoxDefaultButton.Button1
                  );
        }

        private void otoplenie_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
              "• Выбор котла и топлива;" + '\n' +
              "• Устройство котельной и монтаж ее составных частей;" + '\n' +
              "• Установка радиаторов;" + '\n' +
              "• Пусконаладочные работы ;" + '\n' +
              "• Составление плана разводки магистралей;" + '\n' +
              "• Составление сметной документации.", 
              "Об услуге",
              MessageBoxButtons.OK,
              MessageBoxIcon.None,
              MessageBoxDefaultButton.Button1
              );
        }

        private void voda_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
             "• Разработка пакета технической документации;" + '\n' + 
             "• Рабочие эскизы и схемы коммуникаций;" + '\n' +
             "• Рабочие эскизы и схемы коммуникаций;" + '\n' +
             "• Рабочие эскизы и схемы коммуникаций;" + '\n' +
             "• Установку труб водоснабжения.",
             "Об услуге",
             MessageBoxButtons.OK,
             MessageBoxIcon.None,
             MessageBoxDefaultButton.Button1
             );
        }

        private void textBoxOkna_TextChanged(object sender, EventArgs e)
        {
            okna.MaxLength = 2;
         
        }

        private void textBoxDveri_TextChanged(object sender, EventArgs e)
        {
            dveri.MaxLength = 2;
        }

        private void Result_Click(object sender, EventArgs e)
        {
           
            if (textBoxPloshad.Text == "" || textBoxPloshad.Text == "0")
            {
                MessageBox.Show(
                    "Укажите площадь дома",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1
                    );
              
            }
            else
            {
                ploshad = double.Parse(textBoxPloshad.Text.Replace(".", ","));
                textBoxData.Text = DateTime.Now.ToString("dd.MM.yyyy");

                kek.Visible = false;
                panel1.Visible = true;
                label1.Visible = false;
                labelitog.Visible = true;

                textBoxS.Text = textBoxPloshad.Text;

                if (comboBoxetazi.SelectedItem == null || comboBoxetazi.SelectedItem =="1 этаж")
                {
                    textBoxEtazi.Text = "1 этаж";
                }
                else
                {
                    textBoxEtazi.Text = comboBoxetazi.SelectedItem.ToString();
                    ploshad = ploshad / 2;
                    textBoxPloshad.Text = Convert.ToString(ploshad);
                    
                }

                if(okna.Text == "" || okna.Text =="0")
                {
                    okna.Text = "1";
                }
                if (dveri.Text == "" || dveri.Text =="0")
                {
                    dveri.Text = "1";
                }
                okon.Text = okna.Text;
                dverei.Text = dveri.Text;


                ploshad = ploshad - (Convert.ToDouble(okna.Text) * 1.8);
                textBoxPloshad.Text = Convert.ToString(ploshad);

                ploshad = ploshad - (Convert.ToDouble(dveri.Text) * 2.1);
                textBoxPloshad.Text = Convert.ToString(ploshad);

                InsertTable();
            }

        }

        public void InsertTable () // вставить в итоговый результат
        {
            SqlConnection con = new SqlConnection("Data Source =.\\SQLEXPRESS; Initial Catalog = Kursovaya; " +
                        "Integrated Security=true;");
            DataTable dt = new DataTable();

            if (fundament.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_id, products_category, products_name, products_price * " + textBoxS.Text +
                "FROM products " +
                "WHERE products_name LIKE '%" + fundament.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (steni.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_id, products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM products " +
                "WHERE products_name LIKE '%" + steni.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (perecritiya.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_id, products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM products " +
                "WHERE products_name LIKE '%" + perecritiya.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (crovlya.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_id, products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM products " +
                "WHERE products_name LIKE '%" + crovlya.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (truba.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT * " +
                "FROM products " +
                "WHERE products_name LIKE '%" + truba.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (fasad.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_id, products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM products " +
                "WHERE products_name LIKE '%" + fasad.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }
            if (lestnica.SelectedIndex > -1)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT * " +
                "FROM products " +
                "WHERE products_name LIKE '%" + lestnica.SelectedItem.ToString() + "' ", con);
                sda.Fill(dt);
            }


            if (checkBoxEconom.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM otdelka " +
                "WHERE products_name LIKE 'Эконом' ", con);
                sda.Fill(dt);
            }
            if (checkBoxStandart.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM otdelka " +
                "WHERE products_name LIKE 'Стандарт' ", con);
                sda.Fill(dt);
            }
            if (checkBoxPremium.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                "SELECT products_category, products_name, products_price * " + textBoxPloshad.Text +
                "FROM otdelka " +
                "WHERE products_name LIKE 'Премиум' ", con);
                sda.Fill(dt);
            }

            if (checkBoxElectro.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                 "SELECT products_category, products_name, products_price " +
                 "FROM inzeneria " +
                 "WHERE products_name LIKE 'Электрика' ", con);
                sda.Fill(dt);
            }
            if (checkBoxOtoplenie.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                 "SELECT products_category, products_name, products_price " +
                 "FROM inzeneria " +
                "WHERE products_name LIKE 'Отопление' ", con);
                sda.Fill(dt);
            }
            if (checkBoxVoda.Checked)
            {
                SqlDataAdapter sda = new SqlDataAdapter(
                 "SELECT products_category, products_name, products_price " +
                 "FROM inzeneria " +
                 "WHERE products_name LIKE 'Водопровод' ", con);
                sda.Fill(dt);
            }

            DataRow MyRow5 = dt.NewRow();
            MyRow5[1] = "";
            dt.Rows.Add(MyRow5);
            dt.AcceptChanges();

            DataRow MyRow4 = dt.NewRow();
            MyRow4[1] = "Дата";
            MyRow4[2] = textBoxData.Text;
            dt.Rows.Add(MyRow4);
            dt.AcceptChanges();

            DataRow MyRow = dt.NewRow();
            MyRow[1] = "Площадь дома";
            MyRow[2] = textBoxS.Text;
            dt.Rows.Add(MyRow);
            dt.AcceptChanges();

            DataRow MyRow1 = dt.NewRow();
            MyRow1[1] = "Количество этажей";
            MyRow1[2] = textBoxEtazi.Text;
            dt.Rows.Add(MyRow1);
            dt.AcceptChanges();

            DataRow MyRow2 = dt.NewRow();
            MyRow2[1] = "Количество окон";
            MyRow2[2] = okon.Text;
            dt.Rows.Add(MyRow2);
            dt.AcceptChanges();

            DataRow MyRow3 = dt.NewRow();
            MyRow3[1] = "Количество дверей";
            MyRow3[2] =dverei.Text;
            dt.Rows.Add(MyRow3);
            dt.AcceptChanges();


            dataGridView2.DataSource = dt;


            dataGridView2.Columns[0].HeaderText = "ID товара";
            dataGridView2.Columns[1].HeaderText = "Конструкция";
            dataGridView2.Columns[2].HeaderText = "Наименование";
            dataGridView2.Columns[3].HeaderText = "Стоимость";
            dataGridView2.Columns.Remove("Column2");
            dataGridView2.Columns[4].HeaderText = "Стоимость компонента";

            TableDesign(dataGridView2);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBoxPloshad.MaxLength = 5;          
        }

        private void fundament_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Lestnica_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBoxPloshad_KeyPress(object sender, KeyPressEventArgs e) // Чтение только цифр и точки
        { 
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8 && number != '.') 
            {
                e.Handled = true;
            } 
        }

        private void textBoxS_TextChanged(object sender, EventArgs e)
        {
 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveTable(dataGridView2);
            Application.Exit();
        }

        private void comboBoxetazi_SelectedIndexChanged(object sender, EventArgs e)
        {
            cokol();
        }
        public void cokol()
        {
            if (comboBoxetazi.Text == "1 этаж" && fundament.SelectedItem != "Цокольный этаж (2,5м)" || comboBoxetazi.SelectedItem == null && fundament.SelectedItem != "Цокольный этаж (2,5м)")
            {
                lestnica.Enabled = false;
                lestnica.SelectedItem = null;

            }
            else
            {
                lestnica.Enabled = true;
            }
            
        }

        private void okna_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void dveri_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void textBoxEtazi_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxEconom_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxStandart.Checked = false;
            checkBoxPremium.Checked = false;
        }
        private void checkBoxStandart_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxEconom.Checked = false;
            checkBoxPremium.Checked = false;
        }

        private void checkBoxPremium_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxEconom.Checked = false;
            checkBoxStandart.Checked = false;
        }
    }
}
