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
    public partial class Form2 : Form
    {
        public OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;

        DataTable
            dataTable = new DataTable(),
            dataTableLimitCardsList = new DataTable(),
            dataTableLimitCardsItems = new DataTable();

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form1 form1 = this.Owner as Form1;

            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);

            ProdList_Load();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormAddProd formAddProd = new FormAddProd();
            formAddProd.ShowDialog(this);
        }

        //Загрузка списка производителей
        public void ProdList_Load()
        {
            listBox1.Items.Clear();

            //dbCon.Open();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Производители.КодПроизводителя, Производители.Название, Производители.ФормаСобственности 
                FROM Производители 
                ORDER BY Производители.КодПроизводителя;", dbCon);

            DataTable dataTable = new DataTable();
            dbAdapter.Fill(dataTable);

            foreach (DataRow Row in dataTable.Rows)
            {
                listBox1.Items.Add("[" +
                    Row[0].ToString() + "]  " +
                    Row[2].ToString() + " " +
                    Row[1].ToString());
            }

            //dbCon.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            { }
            else
            {
                return;
            }

            string str = listBox1.SelectedItem.ToString(), id = "";

            DialogResult dialogResult = MessageBox.Show(
                    "Удалить производителя '" + str.Remove(0,str.IndexOf(" ") + 2) + "'?", "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            { }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }
            
            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                id = id + str[i];

            //dbCon.Open();

            OleDbCommand myOleDbCommand;
            myOleDbCommand = new OleDbCommand(@"DELETE FROM Производители WHERE Производители.КодПроизводителя = " + id + ";", dbCon);

            try
            {
                myOleDbCommand.ExecuteNonQuery();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //dbCon.Close();

            ProdList_Load();
        }

    }
}
