using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace DataAcc
{
    public partial class Form1 : Form
    {
        private Word.Application WordApp;

        public OleDbConnection dbCon;
        OleDbDataAdapter dbAdapter;
        OleDbCommand myOleDbCommand;

        int comboBox1LastSelection;
        bool CancelFlag, ExceptionFlag, Cell3Flag;

        public DataTable dataTable = new DataTable();
        DataTable
            dataTableLimitCardsList = new DataTable(),
            dataTableLimitCardsItems = new DataTable();

        TreeNode treeNode;

        //public string myConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\admin\\Documents\\Sklad.accdb";
        public string myConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Sklad.accdb";

        public Form1()
        {
            InitializeComponent();
        }


        #region Обработка событий формы

        //Событие "Загрузка формы"
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1LastSelection = -1;
            OrderIdToAdd = "";
            ProductSaveFlag = true;
            ExceptionFlag = false;
            CancelFlag = false;
            Cell3Flag = false;

            dbCon = new OleDbConnection(myConnString);

            try
            {
                dbCon.Open();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ExceptionFlag = true;
                this.Close();
            }
            //dbCon.Close();

            ElementsList_Load();
            ProductsList_Load();

            comboBox1.SelectedIndex = 0;
            
            //dataGridView1.ClearSelection();
        }

        //Кнопка формы "->"
        private void AddButton_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null)
            {
                MessageBox.Show("Необходимо выбрать элемент для добавления\nиз списка элементов", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string str = treeView1.SelectedNode.Text, id = "";
            int dataGridView1Count;

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                id = id + str[i];
            if (id.Length < 4)
                return;

            dataGridView1Count = dataGridView1.Rows.Count;

            for (int i = 0; i < dataGridView1Count; i++)
                if (dataGridView1[0, i].Value.ToString() == id)
                {
                    MessageBox.Show("Данный элемент уже добавлен", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                    return;
                }

            //dbCon.Open();

            dbAdapter = new OleDbDataAdapter(
                "SELECT Элементы.КодЭлемента  AS [Код эл-та], Элементы.ОбозначениепоКД AS Элемент, ТипыЭлементов.Название AS [Тип элемента], '1' AS Количество " +
                "FROM ТипыЭлементов INNER JOIN Элементы ON ТипыЭлементов.КодТипа = Элементы.КодТипа " +
                "WHERE Элементы.КодЭлемента = " + id + " " +
                "ORDER BY Элементы.КодЭлемента;"
                , dbCon);

            dbAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;

            dataGridView_Form(1);

            //dbCon.Close();

            dataGridView1.ClearSelection();
            treeView1.SelectedNode = null;

            ProductSaveFlag = false;
        }

        //Кнопка формы "<-"
        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.CurrentRow != null && dataGridView1.SelectedRows.Count != 0)
                dataGridView1.Rows.Remove(dataGridView1.SelectedRows[0]);

            dataGridView1.ClearSelection();

            ProductSaveFlag = false;
        }

        //Обработка события "Изменен выбранный продукт в списке продуктов"
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Проверка сохранен ли текущий продукт и имеет ли он состав для сохранения
            if (!ProductSaveFlag && dataGridView1.Rows.Count > 0 && comboBox1LastSelection != comboBox1.SelectedIndex)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "Текущий продукт не сохранен.\nСохранить продукт?", "Предупреждение",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    SaveProduct(comboBox1LastSelection);
                    ProductSaveFlag = true;
                }
                else if (dialogResult == DialogResult.No)
                {
                    ProductSaveFlag = true;
                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    CancelFlag = true;
                    comboBox1.SelectedIndex = comboBox1LastSelection;
                    return;
                }
            }

            //ProductSaveFlag = true;
            string str = comboBox1.SelectedItem.ToString(), id = "";

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                id = id + str[i];

            if (!CancelFlag && comboBox1LastSelection != comboBox1.SelectedIndex)
            {
                dataGridView1_Fill(id);
                dataGridView2_Fill(id);
                ProductSaveFlag = true;
            }

            comboBox1LastSelection = comboBox1.SelectedIndex;

            CancelFlag = false;

            dataGridView1.Select();

            dataGridView1.ClearSelection();
        }

        //Обработка события "Редактирование ячейки" в dataGridView1
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 3)
            {
                TextBox tb = (TextBox)e.Control;
                tb.KeyPress += new KeyPressEventHandler(tb_KeyPress);
            }
        }

        void tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && !(e.KeyChar == ','))
                if (e.KeyChar != (char)Keys.Back)
                    e.Handled = true;
        }

        //Получение фокуса ячейкой
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell.ColumnIndex == 3)
            {
                Cell3Flag = true;
            }
            else
                Cell3Flag = false;
        }

        //Проверено значения ячейки
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (Cell3Flag)
            {
                string str = comboBox1.SelectedItem.ToString(), id = "";

                int s = str.IndexOf("["),
                    f = str.IndexOf("]");
                for (int i = s + 1; i < f; i++)
                    id = id + str[i];

                //MessageBox.Show(dataGridView1[dataGridView1.CurrentCell.ColumnIndex, dataGridView1.CurrentRow.Index].Value.ToString());
                myOleDbCommand = new OleDbCommand(
                    @"UPDATE Состав SET Объем = " + 
                    dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString()+ " " +
                    "WHERE (((Состав.КодЗаказа)=" + id +
                    ") AND ((Состав.Элемент)=" + dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString() + "));",
                    dbCon);

                //    @"INSERT INTO Состав (КодЗаказа, Элемент, Объем) 
                //    VALUES (" +
                //    OrderNum/*Код заказа*/ + ", " +
                //    dataGridView1[0, i].Value.ToString()/*Код элта*/ + ", " +
                //    Count/*Количество*/ + ");", dbCon);
                
                SendOleDbCommand(myOleDbCommand);
                dataGridView1.ClearSelection();
            }
        }

        //Обработка события "Щелчок по ячейке" в dataGridView2
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(dataGridView2.Rows.Count == 0)
                return;

            string id = dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString();

            dataTableLimitCardsItems.Clear();

            //dbCon.Open();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Элементы.КодЭлемента AS [Код эл-та], Элементы.ОбозначениепоКД AS Элемент, ТипыЭлементов.Название AS [Тип элемента], Состав.Объем AS [Количество], ROUND(Состав.Объем * Лимитки.Объем, 2) AS [Всего] 
                FROM ТипыЭлементов INNER JOIN (Элементы INNER JOIN ((Заказы INNER JOIN (План INNER JOIN Лимитки ON План.КодПлана = Лимитки.КодПлана) ON Заказы.КодЗаказа = План.КодЗаказа) INNER JOIN Состав ON Заказы.КодЗаказа = Состав.КодЗаказа) ON Элементы.КодЭлемента = Состав.Элемент) ON ТипыЭлементов.КодТипа = Элементы.КодТипа 
                WHERE Лимитки.КодЛимитки = " + id + @" 
                ORDER BY Элементы.КодЭлемента;", dbCon);
            dbAdapter.Fill(dataTableLimitCardsItems);
            dataGridView3.DataSource = dataTableLimitCardsItems;

            dataGridView3.ClearSelection();

            dataGridView_Form(3);

            //dbCon.Close();
        }

        //Переключение вкладок tabControl1
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();
            dataTableLimitCardsItems.Clear();
        }

        //Кнопка меню "Производители"
        private void производителиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog(this);
        }

        //Кнопка меню "Создать лимитку"
        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = Лимитка;
            FormAddLimitCard formAddLimitCard = new FormAddLimitCard();
            formAddLimitCard.ShowDialog(this);
        }

        //Кнопка меню "Печать лимитной карты"
        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = Лимитка;

            if (dataGridView3.RowCount == 0)
            {
                MessageBox.Show("Лимитная карта пуста", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string CountOfLimitCard = "1";
            DataTable dataTable1 = new DataTable();

            object oMissing = System.Reflection.Missing.Value;

            Word.Documents WordDocuments;
            Word.Document WordDocument;

            try
            {
                //Создаем объект Word - равносильно запуску Word
                WordApp = new Word.Application();
                //Делаем его видимым
                WordApp.Visible = true;

                Object template = Type.Missing;
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;

                Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

                WordDocument = WordApp.Documents.Add(
                    ref template, ref newTemplate, ref documentType, ref visible);

                string ProductName, LimitCardName = "", LimitCardDate = "";
                ProductName = comboBox1.SelectedItem.ToString();
                ProductName = ProductName.Remove(0, ProductName.IndexOf(" ") + 2);
                ProductName = ProductName.Remove(ProductName.IndexOf("[") - 1);

                dataTable1.Clear();

                dbAdapter = new OleDbDataAdapter(
                    @"SELECT Лимитки.Объем, Лимитки.НомерЛимитки, Лимитки.ДатаСоставления 
                        FROM  Лимитки 
                        WHERE Лимитки.КодЛимитки = " + dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString() /*[0, i - 2].Value.ToString()*/ + " ;",
                    dbCon);
                dbAdapter.Fill(dataTable1);

                foreach (DataRow Row in dataTable1.Rows)
                {
                    CountOfLimitCard = Row[0].ToString();
                    LimitCardName = Row[1].ToString();
                    LimitCardDate = Row[2].ToString();
                }
                LimitCardDate = LimitCardDate.Remove(LimitCardDate.IndexOf(" "));
                LimitCardName = LimitCardName.Replace("/", ".");
                LimitCardName = LimitCardName.Replace("\\", ".");
                LimitCardName = LimitCardName.Replace("|", ".");
                LimitCardName = LimitCardName.Replace("?", ".");
                LimitCardName = LimitCardName.Replace(":", ".");
                LimitCardName = LimitCardName.Replace("*", ".");
                LimitCardName = LimitCardName.Replace("<", ".");
                LimitCardName = LimitCardName.Replace(">", ".");
                LimitCardName = LimitCardName.Replace("\"", "'");

                ProductName = ProductName.Replace("/", ".");
                ProductName = ProductName.Replace("\\", ".");
                ProductName = ProductName.Replace("|", ".");
                ProductName = ProductName.Replace("?", ".");
                ProductName = ProductName.Replace(":", ".");
                ProductName = ProductName.Replace("*", ".");
                ProductName = ProductName.Replace("<", ".");
                ProductName = ProductName.Replace(">", ".");
                ProductName = ProductName.Replace("\"", "'");

                string FN = ProductName.ToString() + @" лим " + LimitCardName.ToString() + @" " + CountOfLimitCard.ToString() + @".docx";
                object fileName = FN;
                WordDocument.SaveAs(ref fileName, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                  ref oMissing, ref oMissing);

                WordApp.Selection.TypeParagraph();
                //wordrange.Font.Color = Word.WdColor.wdColorRed;
                WordApp.Selection.Font.Name = "Times New Roman";
                WordApp.Selection.Font.Size = 14;
                WordApp.Selection.Font.Bold = 1;
                WordApp.Selection.TypeText(ProductName + " Лимитка " + LimitCardName + " " + CountOfLimitCard + " - шт. от " + LimitCardDate);
                WordApp.Selection.TypeParagraph();
                WordApp.Selection.Font.Size = 12;
                WordApp.Selection.Font.Bold = 0;
                WordApp.Selection.TypeText("Лимитная ведомость");
                WordApp.Selection.TypeParagraph();
                WordApp.Selection.Font.Size = 10;

                int rowCount = dataGridView3.RowCount,
                    NumRowCount = 1;
                bool CycleFlag = true;
                
                //Добавляем таблицу и получаем объект wordtable 
                Word.Table WordTable = WordDocument.Tables.Add(WordApp.Selection.Range, 1/*rowCount*/, 8,
                                  ref defaultTableBehavior, ref autoFitBehavior);

                WordApp.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                WordApp.Application.Selection.PageSetup.LeftMargin = 20f;

                //Форматирование документа
                //
                //Настройка ширины столбцов таблицы
                /*
                 * ПЕРЕНЕСЕНО ПОСЛЕ ЗАПОЛНЕНИЯ ТАБЛИЦЫ ДАННЫМИ
                WordTable.Columns[1].SetWidth(15.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[2].SetWidth(142.2f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[3].SetWidth(233.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[4].SetWidth(108, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[5].SetWidth(54, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[6].SetWidth(39, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[7].SetWidth(108, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[8].SetWidth(83.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                */

                //dbCon.Open();

                Word.Range WordCellRange;

                //WordCellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                
                //Заполнение "шапки" таблицы
                WordCellRange = WordTable.Cell(1, 1).Range;
                WordCellRange.Text = "№";

                WordCellRange = WordTable.Cell(1, 2).Range;
                WordCellRange.Text = "Наименование";

                WordCellRange = WordTable.Cell(1, 3).Range;
                WordCellRange.Text = "Обозначение";

                WordCellRange = WordTable.Cell(1, 4).Range;
                WordCellRange.Text = "Допустимая замена";

                WordCellRange = WordTable.Cell(1, 5).Range;
                WordCellRange.Text = "Прим.";

                WordCellRange = WordTable.Cell(1, 6).Range;
                WordCellRange.Text = "*" + CountOfLimitCard;

                WordCellRange = WordTable.Cell(1, 7).Range;
                WordCellRange.Text = "Кол/цена";

                WordCellRange = WordTable.Cell(1, 8).Range;
                WordCellRange.Text = "Дата";
                //
                //
                //WordTable.Rows[1].Range.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth225pt;

                //Заполнение таблицы данными
                while (CycleFlag)
                {
                    WordTable.Rows.Add(ref oMissing);//Добавление новой строки в таблицу

                    //WordTable.Rows[WordTable.Rows.Count].Range.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth225pt;

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 1).Range;
                    WordCellRange.Text = NumRowCount.ToString();

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 2).Range;
                    //Наименование
                    WordCellRange.Text = dataGridView3[2, NumRowCount - 1].Value.ToString();

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 3).Range;
                    //Обозначение
                    WordCellRange.Text = dataGridView3[1, NumRowCount - 1].Value.ToString();

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 4).Range;
                    //Допустимая замена
                    //WordCellRange.Text = Change;

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 5).Range;
                    //WordCellRange.Text = "Прим.";
                    WordCellRange.Text = dataGridView3[3, NumRowCount - 1].Value.ToString();

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 6).Range;
                    //WordCellRange.Text = "*";
                    WordCellRange.Text = dataGridView3[4, NumRowCount - 1].Value.ToString();

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 7).Range;
                    //WordCellRange.Text = "Кол/цена";

                    WordCellRange = WordTable.Cell(WordTable.Rows.Count, 8).Range;
                    //WordCellRange.Text = "Дата";

                    dataTable1 = new DataTable();
                    //dataTable1.Clear();

                        dbAdapter = new OleDbDataAdapter(
                            @"SELECT СоставСборки.Элемент, СоставСборки.Объем, ROUND(СоставСборки.Объем * " + CountOfLimitCard + @",2) AS Объем 
                        FROM Элементы INNER JOIN СоставСборки ON Элементы.КодЭлемента = СоставСборки.КодЭлемента 
                        WHERE Элементы.КодЭлемента = " + dataGridView3[0, NumRowCount - 1].Value.ToString() + ";", dbCon);

                    dbAdapter.Fill(dataTable1);

                    foreach (DataRow Row in dataTable1.Rows)
                    {
                        DataTable dataTable2 = new DataTable();
                        dbAdapter = new OleDbDataAdapter(
                        @"SELECT Элементы.ОбозначениепоКД, ТипыЭлементов.Название
                        FROM ТипыЭлементов INNER JOIN Элементы ON ТипыЭлементов.КодТипа = Элементы.КодТипа 
                        WHERE Элементы.КодЭлемента = " + Row[0].ToString() + ";",dbCon);
                        dbAdapter.Fill(dataTable2);

                        foreach (DataRow Row2 in dataTable2.Rows)
                        {
                            WordTable.Rows.Add(ref oMissing);//Добавление новой строки в таблицу

                            WordCellRange = WordTable.Cell(WordTable.Rows.Count, 2).Range;
                            //Наименование
                            WordCellRange.Text = Row2[1].ToString();

                            WordCellRange = WordTable.Cell(WordTable.Rows.Count, 3).Range;
                            //Обозначение
                            WordCellRange.Text = Row2[0].ToString();

                            WordCellRange = WordTable.Cell(WordTable.Rows.Count, 5).Range;
                            //WordCellRange.Text = "Прим.";
                            WordCellRange.Text = Row[1].ToString();

                            WordCellRange = WordTable.Cell(WordTable.Rows.Count, 6).Range;
                            //WordCellRange.Text = "*";
                            WordCellRange.Text = Row[2].ToString();
                        }
                    }

                    NumRowCount++;
                    if (NumRowCount == rowCount + 1)//Проверка на выход из цикла
                        CycleFlag = false;
                    
                }


                //Настройка ширины столбцов таблицы
                WordTable.Columns[1].SetWidth(15.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[2].SetWidth(142.2f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[3].SetWidth(233.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[4].SetWidth(108, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[5].SetWidth(54, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[6].SetWidth(39, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[7].SetWidth(108, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                WordTable.Columns[8].SetWidth(83.9f, Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);

                object rStart = 0;
                object rEnd = WordDocument.Content.End;
                WordDocument.Range(ref rStart, ref rEnd).InsertAfter("\n0 - расход материалов – “по потребности”");

                //dbCon.Close();

            }
            catch (Exception exception)
            {
                //MessageBox.Show("Невозможно создать документ Word", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(exception.Message + "\nМетод:\n" + exception.TargetSite.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Кнопка меню "Создать новый продукт"
        private void новыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверка сохранен ли текущий продукт и имеет ли он состав для сохранения
            if (!ProductSaveFlag && dataGridView1.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "Текущий продукт не сохранен.\nСохранить продукт и создать новый?", "Предупреждение",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    SaveProduct(comboBox1.SelectedIndex);
                }
                else if (dialogResult == DialogResult.No)
                {

                }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }
            }

            FormAddProduct formAddProduct = new FormAddProduct();
            formAddProduct.ShowDialog(this);
        }

        //Кнопка меню "Создать новый продукт на основе имеющегося"
        private void наОсновеИмеющегосяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверка сохранен ли текущий продукт и имеет ли он состав для сохранения
            if (!ProductSaveFlag && dataGridView1.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "Текущий продукт не сохранен.\nСохранить продукт и создать новый?", "Предупреждение",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    SaveProduct(comboBox1.SelectedIndex);
                }
                else if (dialogResult == DialogResult.No)
                { }
                else if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }
            }

            string str = comboBox1.SelectedItem.ToString(), id = "";

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                id = id + str[i];

            OrderIdToAdd = id;

            FormAddProduct formAddProduct = new FormAddProduct();
            formAddProduct.ShowDialog(this);
        }

        //Кнопка меню "Добавить элемент"
        private void элементToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAddElement formAddElement = new FormAddElement();
            formAddElement.ShowDialog(this);
        }

        //Кнопка меню "Обновить список элементов"
        private void обновитьСписокЭлементовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ElementsList_Load();
        }

        //Кнопка меню "Удалить элемент"
        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode == null)
                return;

            string str;

            if (treeView1.SelectedNode.Parent == null) // Тип типа
                str = "базовый тип '";
            else if (treeView1.SelectedNode.Parent.Parent == null) // Тип
                str = "тип элементов '";
            else // Элемент
                str = "элемент '";

            str += treeView1.SelectedNode.Text;

            DialogResult dialogResult = MessageBox.Show(
                    "Удалить " + str + "' ?", "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes) { }
            else if (dialogResult == DialogResult.No)
                return;

            Delete();
        }

        //Кнопка меню "Добавить тип"
        private void типToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAddType formAddType = new FormAddType();
            formAddType.ShowDialog(this);
        }

        //Кнопка меню "Создать базовый тип"
        private void базовыйТипToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAddBaseType formAddBaseType = new FormAddBaseType();
            formAddBaseType.ShowDialog(this);
        }

        //Кнопка меню "Сохранить продукт"
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(!ProductSaveFlag)
                SaveProduct(comboBox1.SelectedIndex);
        }

        //Кнопка меню "Удалить продукт"
        private void уToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string str = comboBox1.SelectedItem.ToString();

            DialogResult dialogResult = MessageBox.Show(
                    "Удалить текущий продукт?\n" + str, "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes){}
            else if (dialogResult == DialogResult.No)
            {
                return;
            }

            string id = "", OrderNum = "", PlanNum = "";

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                OrderNum = OrderNum + str[i];

            //dbCon.Open();

            DataTable dataTable1 = new DataTable();

            //+
            dbAdapter = new OleDbDataAdapter(
                @"SELECT План.КодПлана 
                FROM Заказы INNER JOIN (План INNER JOIN Лимитки ON План.КодПлана = Лимитки.КодПлана) ON Заказы.КодЗаказа = План.КодЗаказа 
                WHERE Заказы.КодЗаказа=" + OrderNum + ";", dbCon);

            dbAdapter.Fill(dataTable1);

            foreach (DataRow Row in dataTable1.Rows)
            {
                //OrderNum = Row[0].ToString();
                PlanNum = Row[0].ToString();
            }

            //OleDbCommand myOleDbCommand;

            //++++
            try
            {
                if (PlanNum != "")//(OrderNum != "")
                {
                    myOleDbCommand = new OleDbCommand(@"DELETE FROM Лимитки WHERE Лимитки.КодПлана = " + PlanNum + ";",
                        dbCon);
                    SendOleDbCommand(myOleDbCommand);

                    myOleDbCommand = new OleDbCommand(@"DELETE FROM План WHERE План.КодЗаказа = " + OrderNum + ";",
                        dbCon);
                    SendOleDbCommand(myOleDbCommand);

                    myOleDbCommand = new OleDbCommand(@"DELETE FROM Состав WHERE Состав.КодЗаказа = " + OrderNum + ";",
                        dbCon);
                    SendOleDbCommand(myOleDbCommand);

                }
                //+++
                myOleDbCommand = new OleDbCommand(@"DELETE FROM Заказы WHERE Заказы.КодЗаказа = " + OrderNum + ";",
                        dbCon);
                SendOleDbCommand(myOleDbCommand);

                /*
                myOleDbCommand = new OleDbCommand(@"DELETE FROM Продукция WHERE Продукция.КодПродукции = " + id + ";",
                    dbCon);

                SendOleDbCommand(myOleDbCommand);
                */
                //dbCon.Close();

                ProductsList_Load();
            }
            catch (Exception exception)
            {
                //MessageBox.Show("Невозможно создать документ Word", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(exception.Message, exception.TargetSite.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Кнопка меню "Обновить список продуктов"
        private void обновитьСписокПродуктовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProductsList_Load();
        }

        //Кнопка меню "Справка"
        private void справкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                Process SysInfo = new Process();
                SysInfo.StartInfo.ErrorDialog = true;
                SysInfo.StartInfo.FileName = "Help.chm";
                SysInfo.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Кнопка меню "О программе"
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAbout formAbout = new FormAbout();
            formAbout.ShowDialog(this);
        }

        //Событие "Изменение свойства SelectedIndex" у comboBox1
        private void comboBox1_Enter(object sender, EventArgs e)
        {
            comboBox1LastSelection = comboBox1.SelectedIndex;
        }

        //Событие "Ошибка" у dataGridView1
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show(
                "Некорректный формат данных", "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

            ExceptionFlag = true;

            dataGridView1.CancelEdit();
            dataGridView1.EndEdit();
        }

        //Событие "Выход"
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Проверка сохранен ли текущий продукт и имеет ли он состав для сохранения
            if (!ProductSaveFlag && dataGridView1.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "Текущий продукт не сохранен.\nСохранить продукт и создать новый?", "Предупреждение",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Warning);
                if (dialogResult == DialogResult.Yes)
                {
                    SaveProduct(comboBox1.SelectedIndex);
                }
                else if (dialogResult == DialogResult.No)
                { }
                else if (dialogResult == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
            if (!ExceptionFlag)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "Выйти из программы?", "Предупреждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                { }
                else if (dialogResult == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            else
            {
                ExceptionFlag = false;
            }

            dbCon.Close();
        }

        #endregion


        
        #region Загрузка списков (Элементы, Продукты)

        //Загрузка списка элементов
        public void ElementsList_Load()
        {
            treeView1.Nodes.Clear();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT ТипыТипов.КодТипаТипов, ТипыТипов.Название
                FROM ТипыТипов;", dbCon);

            DataTable dataTableTypeOfTypes = new DataTable();
            dbAdapter.Fill(dataTableTypeOfTypes);

            foreach (DataRow RowTypeOfTypes in dataTableTypeOfTypes.Rows)
            {
                dbAdapter = new OleDbDataAdapter(
                    @"SELECT ТипыЭлементов.КодТипа, ТипыЭлементов.Название 
                    FROM ТипыТипов INNER JOIN ТипыЭлементов ON ТипыТипов.КодТипаТипов = ТипыЭлементов.Принадлежность 
                    WHERE ТипыТипов.КодТипаТипов = " + RowTypeOfTypes[0].ToString() + ";", dbCon);

                DataTable dataTableElementTypes = new DataTable();
                dbAdapter.Fill(dataTableElementTypes);

                treeNode = new TreeNode(RowTypeOfTypes[1].ToString() + " [" + RowTypeOfTypes[0].ToString() + "]");

                treeView1.Nodes.Add(treeNode);

                foreach (DataRow RowElementTypes in dataTableElementTypes.Rows)
                {
                    dbAdapter = new OleDbDataAdapter(
                        @"SELECT Элементы.КодЭлемента, Элементы.ОбозначениепоКД 
                        FROM ТипыЭлементов INNER JOIN Элементы ON ТипыЭлементов.КодТипа = Элементы.КодТипа 
                        WHERE (([Элементы].[КодТипа]=" +
                        RowElementTypes[0].ToString() + "));", dbCon);

                    DataTable dataTableElements = new DataTable();
                    dbAdapter.Fill(dataTableElements);

                    treeNode = new TreeNode(RowElementTypes[1].ToString() + " [" + RowElementTypes[0].ToString() + "]");

                    treeView1.Nodes[treeView1.Nodes.Count - 1].
                        Nodes.Add(treeNode);

                    foreach (DataRow RowElements in dataTableElements.Rows)
                    {
                        treeNode = new TreeNode(
                        "[" +
                        RowElements[0].ToString() + "] " +
                        RowElements[1].ToString());

                        treeView1.Nodes[treeView1.Nodes.Count - 1].
                            Nodes[treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Count - 1].
                            Nodes.Add(treeNode);
                    }
                }
            }
        }

        //Загрузка списка продуктов
        public void ProductsList_Load()
        {
            comboBox1.Items.Clear();

            //+++++
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Заказы.КодЗаказа, Продукция.Название, Заказы.Прим, Заказы.НомерЗаказа 
                FROM Продукция INNER JOIN Заказы ON Продукция.КодПродукции = Заказы.КодПродукции ORDER BY Заказы.КодЗаказа;", dbCon);

            DataTable dataTableProduction = new DataTable();
            dbAdapter.Fill(dataTableProduction);

            foreach (DataRow RowProduction in dataTableProduction.Rows)
            {
                comboBox1.Items.Add("зак№" +
                    RowProduction[3].ToString() + "  " +
                    RowProduction[1].ToString() + " [" +
                    RowProduction[0].ToString() + "] ");
            }

            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

        }

        #endregion

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

        //Задает ширину столбцов в таблицах dataGridView
        void dataGridView_Form(int code)
        {
            switch (code)
            {
                case 1:
                    //dataGridView1.Columns[0].Width = 90;
                    //dataGridView1.Columns[1].Width = 251;
                    //dataGridView1.Columns[2].Width = 213;
                    //dataGridView1.Columns[3].Width = 80;
                    dataGridView1.Columns[0].Visible = false;
                    dataGridView1.Columns[0].ReadOnly = true;
                    dataGridView1.Columns[1].ReadOnly = true;
                    dataGridView1.Columns[2].ReadOnly = true;
                    break;
                case 2:
                    dataGridView2.Columns[0].Visible = false;
                    //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    /*
                    dataGridView2.Columns[0].Width = 80;
                    dataGridView2.Columns[1].Width = 113;
                    dataGridView2.Columns[2].Width = 140;
                    dataGridView2.Columns[3].Width = 80;
                    dataGridView2.Columns[4].Width = 120;
                    */
                    dataGridView2.Sort(dataGridView2.Columns[4], ListSortDirection.Descending);
                    break;
                case 3:
                    dataGridView3.Columns[0].Visible = false;
                    //dataGridView3.Columns[0].Width = 80;
                    break;
            }
        }


        //Заполнение таблицы dataGridView1 - таблица состава выбранного продукта
        public void dataGridView1_Fill(string id)
        {
            dataTable.Clear();

            //dbCon.Open();
            //+++
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Заказы.Прим
                FROM Заказы
                WHERE Заказы.КодЗаказа = " + id + ";"
                , dbCon);

            DataTable Dt = new DataTable();
            dbAdapter.Fill(Dt);

            foreach (DataRow Row in Dt.Rows)
            {
                LabelNote.Text = Row[0].ToString();
            }
            //+++
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Элементы.КодЭлемента AS [Код эл-та], Элементы.ОбозначениепоКД AS Элемент, ТипыЭлементов.Название AS [Тип элемента], Состав.Объем AS Количество
                FROM (ТипыЭлементов INNER JOIN Элементы ON ТипыЭлементов.КодТипа = Элементы.КодТипа) INNER JOIN (Заказы INNER JOIN Состав ON Заказы.КодЗаказа = Состав.КодЗаказа) ON Элементы.КодЭлемента = Состав.Элемент
                WHERE Заказы.КодЗаказа = " + id + @" 
                ORDER BY Элементы.КодЭлемента;"
                , dbCon);

            dbAdapter.Fill(dataTable);
            dataGridView1.DataSource = dataTable;

            dataGridView_Form(1);

            dataGridView1.ClearSelection();

            //dbCon.Close();
        }

        //Заполнение таблицы dataGridView2 - таблица состава выбранного продукта
        public void dataGridView2_Fill(string id)
        {
            dataTableLimitCardsList.Clear();
            dataTableLimitCardsItems.Clear();
            //+++
            dbAdapter = new OleDbDataAdapter(
                @"SELECT Лимитки.КодЛимитки AS [Код лим], Лимитки.НомерЛимитки AS [Номер лим], План.ЗавНомера AS [Серийные номера], Лимитки.Объем AS Количество, Лимитки.ДатаСоставления AS [Дата создания]
                FROM Заказы INNER JOIN (План INNER JOIN Лимитки ON План.КодПлана = Лимитки.КодПлана) ON Заказы.КодЗаказа = План.КодЗаказа
                WHERE Заказы.КодЗаказа = " + id + @" 
                ORDER BY Лимитки.КодЛимитки;", dbCon);

            dbAdapter.Fill(dataTableLimitCardsList);
            dataGridView2.DataSource = dataTableLimitCardsList;

            dataGridView_Form(2);

            //dbCon.Close();
        }

        //Очищает таблицы dataGridView
        public void ClearDataTables()
        {
            dataTable.Clear();
            dataTableLimitCardsList.Clear();
            dataTableLimitCardsItems.Clear();
            comboBox1.ResetText();
        }

        //Сохранение текущего продукта
        public void SaveProduct(int index)
        {
            string str = comboBox1.Items[index].ToString(), id = "", OrderNum = "";
            int dataGridView1RowsCount = dataGridView1.Rows.Count;

            if (dataGridView1RowsCount == 0)
            {
                MessageBox.Show(
                    "Текущий продукт не имеет состава", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                OrderNum = OrderNum + str[i];

            //dbCon.Open();

            DataTable dataTable1 = new DataTable();
            /*
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
            if (OrderNum == "")
            {
                //dbCon.Close();
                return;
            }

            myOleDbCommand = new OleDbCommand(@"DELETE FROM Состав WHERE Состав.КодЗаказа = " + OrderNum + ";", dbCon);

            SendOleDbCommand(myOleDbCommand);

            for (int i = 0; i < dataGridView1RowsCount; i++)
            {
                string Count = dataGridView1[3, i].Value.ToString();
                if (Count.IndexOf(",") != -1)
                    Count = Count.Replace(",", ".");

                myOleDbCommand = new OleDbCommand(
                    @"INSERT INTO Состав (КодЗаказа, Элемент, Объем) 
                    VALUES (" +
                    OrderNum/*Код заказа*/ + ", " +
                    dataGridView1[0, i].Value.ToString()/*Код элта*/ + ", " +
                    Count/*Количество*/ + ");", dbCon);

                SendOleDbCommand(myOleDbCommand);
            }

            MessageBox.Show(
                "Текущий продукт сохранен", "Информация",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            ProductSaveFlag = true;

            //dbCon.Close();
        }


        #region Уаление элементов/типов/типов типов

        //Удаляет выбранный элемент/тип/тип типа
        private void Delete()
        {
            //dbCon.Open();

            string str, id;

            str = treeView1.SelectedNode.Text;
            id = "";

            int s = str.IndexOf("["),
                f = str.IndexOf("]");
            for (int i = s + 1; i < f; i++)
                id = id + str[i];

            if (treeView1.SelectedNode.Parent == null) // Тип типа
            {
                str = treeView1.SelectedNode.Text;
                DeleteGeneralType(id);
            }
            else if (treeView1.SelectedNode.Parent.Parent == null) // Тип
            {
                str = treeView1.SelectedNode.Text;
                DeleteType(id);
            }
            else // Элемент
            {
                if (!DeleteElement(id))
                    return;
            }

            //dbCon.Close();

            ElementsList_Load();

            MessageBox.Show("Элемент удален", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool DeleteElement(string id)
        {
            //OleDbCommand myOleDbCommand;
            myOleDbCommand = new OleDbCommand(@"DELETE FROM Элементы WHERE Элементы.КодЭлемента = " + id + ";", dbCon);
            try
            {
                myOleDbCommand.ExecuteNonQuery();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private void DeleteType(string id)
        {
            DataTable dataTableDeleteType = new DataTable();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT Элементы.КодЭлемента 
                FROM ТипыЭлементов INNER JOIN Элементы ON ТипыЭлементов.КодТипа = Элементы.КодТипа 
                WHERE ТипыЭлементов.КодТипа = " + id + ";", dbCon);
            dbAdapter.Fill(dataTableDeleteType);

            foreach (DataRow RowDeleteType in dataTableDeleteType.Rows)
            {
                DeleteElement(RowDeleteType[0].ToString());
            }

            //OleDbCommand myOleDbCommand;
            myOleDbCommand = new OleDbCommand(@"DELETE FROM ТипыЭлементов WHERE ТипыЭлементов.КодТипа = " + id + ";", dbCon);

            SendOleDbCommand(myOleDbCommand);
        }

        private void DeleteGeneralType(string id)
        {
            DataTable dataTableDeleteGeneralType = new DataTable();

            dbAdapter = new OleDbDataAdapter(
                @"SELECT ТипыЭлементов.КодТипа 
                FROM ТипыТипов INNER JOIN ТипыЭлементов ON ТипыТипов.КодТипаТипов = ТипыЭлементов.Принадлежность 
                WHERE ТипыТипов.КодТипаТипов = " + id + ";", dbCon);
            dbAdapter.Fill(dataTableDeleteGeneralType);

            foreach (DataRow RowDeleteGeneralType in dataTableDeleteGeneralType.Rows)
            {
                DeleteType(RowDeleteGeneralType[0].ToString());
            }

            //OleDbCommand myOleDbCommand;
            myOleDbCommand = new OleDbCommand(@"DELETE FROM ТипыТипов WHERE ТипыТипов.КодТипаТипов = " + id + ";", dbCon);

            SendOleDbCommand(myOleDbCommand);
        }

        #endregion

    }
}
