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
    public partial class FormAddProd : Form
    {
        OleDbConnection dbCon;

        Form2 form2;

        //string myConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\admin\\Documents\\Sklad.accdb";
        public string myConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Sklad.accdb";

        public FormAddProd()
        {
            InitializeComponent();
        }

        private void FormAddProd_Load(object sender, EventArgs e)
        {
            form2 = this.Owner as Form2;

            dbCon = form2.dbCon;// = new OleDbConnection(myConnString);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string
                ProdName = textBox1.Text.ToString(),
                OwnerForm = textBox2.Text.ToString();

            if (ProdName == "")
            {
                MessageBox.Show("Необходимо ввести название производителя", "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            else
            {
                //dbCon.Open();

                OleDbCommand myOleDbCommand;

                if (OwnerForm != "")
                    myOleDbCommand = new OleDbCommand(
                        @"INSERT INTO Производители (Название, ФормаСобственности) 
                        VALUES ('" + ProdName + "','" + OwnerForm + "')", dbCon);
                else
                    myOleDbCommand = new OleDbCommand(
                        @"INSERT INTO Производители (Название) 
                        VALUES ('" + ProdName + "')", dbCon);

                try
                {
                    myOleDbCommand.ExecuteNonQuery();
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //dbCon.Close();

                form2.ProdList_Load();

                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
