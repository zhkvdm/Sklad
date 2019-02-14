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
    public partial class FormAddBaseType : Form
    {
        OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;

        Form1 form1;

        public FormAddBaseType()
        {
            InitializeComponent();
        }

        private void FormAddBaseType_Load(object sender, EventArgs e)
        {
            form1 = this.Owner as Form1;
            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string BaseTypeName = textBox1.Text.ToString();

            if (BaseTypeName == "")
            {
                MessageBox.Show("Необходимо ввести название типа", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string Order = textBox2.Text.ToString(), Note = textBox3.Text.ToString();

            if (Order == "")
            {
                Order = "0";
            }

            //dbCon.Open();

            OleDbCommand myOleDbCommand;

            myOleDbCommand = new OleDbCommand(
                @"INSERT INTO ТипыТипов (Название, Порядок, Прим) 
                VALUES ('" +
                BaseTypeName + "', " +
                Order + ",'" +
                Note + "')",
                dbCon);

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

            MessageBox.Show("Базавый тип добавлен", "Добавление базового типа", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
