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
    public partial class FormAddElement : Form
    {
        OleDbConnection dbCon;

        OleDbDataAdapter dbAdapter;

        DataTable
            dataTable = new DataTable(),
            dataTableTypes = new DataTable(),
            dataTableProductions = new DataTable();

        Form1 form1;

        public FormAddElement()
        {
            InitializeComponent();
        }

        private void FormAddElement_Load(object sender, EventArgs e)
        {
            form1 = this.Owner as Form1;

            dbCon = form1.dbCon;// = new OleDbConnection(form1.myConnString);
            //dbCon.Open();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT ТипыЭлементов.КодТипа, ТипыЭлементов.Название 
                FROM ТипыЭлементов 
                ORDER BY ТипыЭлементов.КодТипа;", dbCon);

            dbAdapter.Fill(dataTableTypes);

            foreach (DataRow Row in dataTableTypes.Rows)
            {
                comboBox1.Items.Add("[" + Row[0].ToString() + "] " + Row[1].ToString());
            }

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Производители.КодПроизводителя, Производители.Название, Производители.ФормаСобственности 
                FROM Производители 
                ORDER BY Производители.КодПроизводителя;", dbCon);

            dbAdapter.Fill(dataTableProductions);

            foreach (DataRow Row in dataTableProductions.Rows)
            {
                comboBox2.Items.Add("[" + Row[0].ToString() + "] " + Row[1].ToString());
            }

            //dbCon.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ElementName = textBox1.Text.ToString();

            if (ElementName == "")
            {
                MessageBox.Show("Необходимо ввести название элемента", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string str, idOfType = "", idOfProd = "";
            int s, f;

            if (comboBox1.SelectedIndex != -1)
            {
                str = comboBox1.SelectedItem.ToString();

                s = str.IndexOf("[");
                f = str.IndexOf("]");
                for (int i = s + 1; i < f; i++)
                    idOfType = idOfType + str[i];
            }
            else 
            {
                MessageBox.Show("Необходимо выбрать тип элемента", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (comboBox2.SelectedIndex != -1)
            {
                str = comboBox2.SelectedItem.ToString();

                s = str.IndexOf("[");
                f = str.IndexOf("]");
                for (int i = s + 1; i < f; i++)
                    idOfProd = idOfProd + str[i];
            }
            
            //dbCon.Open();

            OleDbCommand myOleDbCommand;
            
            if(idOfProd != "")
                myOleDbCommand = new OleDbCommand(
                    @"INSERT INTO Элементы (ОбозначениепоКД, КодТипа, КодПроизводителя) 
                    VALUES ('" + ElementName + "', " + idOfType + "," + idOfProd + ")", dbCon);
            else
                myOleDbCommand = new OleDbCommand(
                    @"INSERT INTO Элементы (ОбозначениепоКД, КодТипа) 
                    VALUES ('" + ElementName + "', " + idOfType + ")", dbCon);

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

            MessageBox.Show("Элемент добавлен", "Добавление элемента", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
