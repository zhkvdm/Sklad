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
    public partial class FormAddType : Form
    {
        OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;

        DataTable dataTable = new DataTable();

        Form1 form1;

        public FormAddType()
        {
            InitializeComponent();
        }

        private void FormAddType_Load(object sender, EventArgs e)
        {
            form1 = this.Owner as Form1;

            textBox2.Text = "1000";

            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);
            //dbCon.Open();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT ТипыТипов.КодТипаТипов, ТипыТипов.Название 
                FROM ТипыТипов 
                ORDER BY ТипыТипов.КодТипаТипов;", dbCon);

            dbAdapter.Fill(dataTable);

            foreach (DataRow Row in dataTable.Rows)
            {
                comboBox1.Items.Add("[" + Row[0].ToString() + "] " + Row[1].ToString());
            }

            //dbCon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string TypeName = textBox1.Text.ToString();

            if (TypeName == "")
            {
                MessageBox.Show("Необходимо ввести название типа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string str, idOfBaseType = "", order = textBox2.Text.ToString();
            int s, f;

            if (comboBox1.SelectedIndex != -1)
            {
                str = comboBox1.SelectedItem.ToString();

                s = str.IndexOf("[");
                f = str.IndexOf("]");
                for (int i = s + 1; i < f; i++)
                    idOfBaseType = idOfBaseType + str[i];
            }
            else
            {
                MessageBox.Show("Необходимо выбрать базовый тип элемента", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (order == "")
            {
                order = "1000";
                textBox2.Text = order;
            }

            //dbCon.Open();

            OleDbCommand myOleDbCommand;

            myOleDbCommand = new OleDbCommand(
                @"INSERT INTO ТипыЭлементов (Название, Порядок, Принадлежность) 
                VALUES ('" + TypeName + "', " + order + "," + idOfBaseType + ")", dbCon);

            try
            {
                myOleDbCommand.ExecuteNonQuery();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //dbCon.Close();

            form1.ElementsList_Load();

            MessageBox.Show("Тип элементов добавлен", "Добавление типа", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
