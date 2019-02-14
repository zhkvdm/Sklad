using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace DataAcc
{
    public partial class FormAddProduct : Form
    {
        OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;
        DataTable dataTable3 = new DataTable();

        string ProductName;

        Form1 form1;


        public FormAddProduct()
        {
            InitializeComponent();
        }

        private void FormAddProduct_Load(object sender, EventArgs e)
        {
            form1 = this.Owner as Form1;
            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);
            
            //dbCon.Open();
            /*
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Продукция.КодПродукции, Продукция.Название 
                FROM Продукция;", dbCon);

            DataTable dataTableProduction = new DataTable();
            dbAdapter.Fill(dataTableProduction);

            foreach (DataRow RowProduction in dataTableProduction.Rows)
            {
                comboBox1.Items.Add(RowProduction[1].ToString());
            }
            */
            //Выбор имени продукта из БД
            //Происходит при выборе создания нового продукта на основе выбранного
            if (form1.OrderIdToAdd != "")
            {
                DataTable dataTable2 = new DataTable();

                dbAdapter = new OleDbDataAdapter(
                    @"SELECT Продукция.Название
                    FROM Продукция INNER JOIN Заказы ON Продукция.КодПродукции = Заказы.КодПродукции
                    WHERE Заказы.КодЗаказа = " + form1.OrderIdToAdd + ";", dbCon);
                dbAdapter.Fill(dataTable2);

                foreach (DataRow Row in dataTable2.Rows)
                {
                    ProductName = Row[0].ToString();
                }

                textBox1.Text = ProductName + " - 1";
            }

            //dbCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductName = textBox1.Text.ToString();
            string MaxProductNum = "";

            if (ProductName == "")
            {
                MessageBox.Show("Необходимо ввести название продукта", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                //dbCon.Open();

                DataTable dataTable1 = new DataTable();

                dbAdapter = new OleDbDataAdapter(
                    @"SELECT Продукция.КодПродукции 
                    FROM Продукция 
                    WHERE Продукция.Название = '" + ProductName + "';", dbCon);
                dbAdapter.Fill(dataTable1);

                foreach (DataRow Row in dataTable1.Rows)
                {
                    if (Row[0].ToString() != "")
                    {
                        MessageBox.Show("Продукт с таким именем уже существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //dbCon.Close();
                        return;
                    }
                }

                //dbCon.Close();
            }

            string 
                OrderNum = textBox2.Text.ToString(),
                Note = textBox3.Text.ToString();

            if (OrderNum == "")
            {
                //OrderNum = "0";
                MessageBox.Show("Необходимо ввести номер заказа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //dbCon.Open();

            OleDbCommand myOleDbCommand;

            myOleDbCommand = new OleDbCommand(
                @"INSERT INTO Продукция (Название) 
                VALUES ('" + ProductName + "');", dbCon);

            SendOleDbCommand(myOleDbCommand);

            
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Max(Продукция.КодПродукции) AS Выражение1 
                FROM Продукция;", dbCon);
            dbAdapter.Fill(dataTable3);

            foreach (DataRow Row in dataTable3.Rows)
            {
                MaxProductNum = Row[0].ToString();
            }

            if (MaxProductNum == "")
            {
                //dbCon.Close();
                return;
            }

            if(Note != "")
                myOleDbCommand = new OleDbCommand(
                    @"INSERT INTO Заказы (НомерЗаказа, КодПродукции, Прим) 
                    VALUES ('" + OrderNum + "'," + MaxProductNum + ",'" + Note + "')", dbCon);
            else
                myOleDbCommand = new OleDbCommand(
                    @"INSERT INTO Заказы (НомерЗаказа, КодПродукции) 
                    VALUES ('" + OrderNum + "'," + MaxProductNum + ")", dbCon);

            SendOleDbCommand(myOleDbCommand);
            

            

            form1.ClearDataTables();
            form1.ProductsList_Load();
            form1.comboBox1.SelectedIndex = (form1.comboBox1.Items.Count - 1);
            if (form1.OrderIdToAdd != "")
            {
                form1.dataGridView1_Fill(form1.OrderIdToAdd);
                form1.dataGridView2_Fill(form1.OrderIdToAdd);
            }

            //dbCon.Close();

            form1.SaveProduct(form1.comboBox1.SelectedIndex);//comboBox1.SelectedIndex

            //form1.ProductSaveFlag = false;

            //MessageBox.Show("Новый продукт добавлен", "Добавление продукта", MessageBoxButtons.OK, MessageBoxIcon.Information);

            form1.OrderIdToAdd = "";

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Отправка SQL-команды
        private void SendOleDbCommand(OleDbCommand oleDbCommand)
        {
            try
            {
                oleDbCommand.ExecuteNonQuery();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
