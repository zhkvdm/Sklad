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
    public partial class FormAddLimitCard : Form
    {
        OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;

        Form1 form1;

        string id, OrderNum;

        public FormAddLimitCard()
        {
            InitializeComponent();
        }

        private void FormAddLimitCard_Load(object sender, EventArgs e)
        {
            form1 = this.Owner as Form1;

            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);

            string str = form1.comboBox1.SelectedItem.ToString();
            id = "";
            OrderNum = "";

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                OrderNum = OrderNum + str[i];

            //dbCon.Open();
            /*
            DataTable dataTable1 = new DataTable();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Заказы.КодЗаказа 
                FROM Продукция INNER JOIN Заказы ON Продукция.КодПродукции = Заказы.КодПродукции 
                WHERE Продукция.КодПродукции = " + id + ";", dbCon);

            dbAdapter.Fill(dataTable1);

            foreach (DataRow Row in dataTable1.Rows)
            {
                OrderNum = Row[0].ToString();
            }
            */

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string LimitCardName = "", Numbers = "", Volume = "", CurrentPlanNum = "";
            
            LimitCardName = textBox1.Text.ToString();
            Numbers = textBox2.Text.ToString();
            //Volume = textBox3.Text.ToString();
            Volume = numericUpDown1.Value.ToString();

            if (LimitCardName == "" || Numbers == "" || Volume == "")
            {
                MessageBox.Show(
                    "Заполнены не все поля", "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            OleDbCommand myOleDbCommand;

            myOleDbCommand = new OleDbCommand(
                @"INSERT INTO План (КодЗаказа, Объем, ЗавНомера) 
                VALUES (" + OrderNum + ", " + Volume + ", '" + Numbers + "');", dbCon);

            SendOleDbCommand(myOleDbCommand);

            DataTable dataTable2 = new DataTable();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Max(План.КодПлана) AS Выражение1 
                FROM План;", dbCon);
            dbAdapter.Fill(dataTable2);

            foreach (DataRow Row in dataTable2.Rows)
            {
                CurrentPlanNum = Row[0].ToString();
            }

            myOleDbCommand = new OleDbCommand(
                @"INSERT INTO Лимитки (НомерЛимитки, КодПлана, Объем) 
                VALUES ('" + LimitCardName + "', " + CurrentPlanNum + ", " + Volume + ");", dbCon);

            SendOleDbCommand(myOleDbCommand);

            //dbCon.Close();

            form1.dataGridView2_Fill(OrderNum);

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
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
