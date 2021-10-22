using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApp6
{

    public partial class Form1 : Form
    {
        private Panel buttonPanel = new Panel();
        private DataGridView songsDataGridView = new DataGridView();
        private Button addNewRowButton = new Button();
        private Button deleteRowButton = new Button();
        private Button Button1 = new Button();
        private Button Button2 = new Button();
        private Button Buttontxt = new Button();

        public Form1()
        {
            this.Load += new EventHandler(Form1_Load);

        }
        private void Form1_Load(System.Object sender, System.EventArgs e)
        {
            SetupLayout();
            SetupDataGridView();
            PopulateDataGridView();
        }
        private void songsDataGridView_CellFormatting(object sender,
        System.Windows.Forms.DataGridViewCellFormattingEventArgs e)
        {
            if (e != null)
            {
                if (this.songsDataGridView.Columns[e.ColumnIndex].Name ==
               "Release Date")
                {
                    if (e.Value != null)
                    {
                        try
                        {
                            e.Value = DateTime.Parse(e.Value.ToString())
                            .ToLongDateString();
                            e.FormattingApplied = true;
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine("{0} is not a valid date.",
                           e.Value.ToString());
                        }
                    }
                }
            }
        }
        private void addNewRowButton_Click(object sender, EventArgs e)
        {
            this.songsDataGridView.Rows.Add();
        }
        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            if (this.songsDataGridView.SelectedRows.Count > 0 &&
            this.songsDataGridView.SelectedRows[0].Index !=
            this.songsDataGridView.Rows.Count - 1)
            {
                this.songsDataGridView.Rows.RemoveAt(
                this.songsDataGridView.SelectedRows[0].Index);
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            songsDataGridView.Sort(new RowComparer(SortOrder.Ascending));
        }
        private void Buttontxt_Click(object sender, EventArgs e)
        {
            FileStream fs = new FileStream("база.txt", FileMode.Open);
            StreamWriter streamWriter = new StreamWriter(fs);
            try
            {
                for (int j = 0; j < songsDataGridView.Rows.Count; j++)
                {
                    for (int i = 0; i < songsDataGridView.Rows[j].Cells.Count; i++)
                    {
                        streamWriter.Write(songsDataGridView.Rows[j].Cells[i].Value+"  ||  ");
                    }

                    streamWriter.WriteLine();
                }

                streamWriter.Close();
                fs.Close();

                MessageBox.Show("Файл успешно сохранен");
            }
            catch
            {
                MessageBox.Show("Ошибка при сохранении файла!");
            }

        }
    
    
        private void Button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            for (int i = 0; i < songsDataGridView.Columns.Count; i++)
            {
                ExcelApp.Cells[i + 1].Characters.Font.Bold = true;
                ExcelApp.Cells[i + 1] = songsDataGridView.Columns[i].HeaderText;
                


            }
                for (int i = 1; i < songsDataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < songsDataGridView.ColumnCount; j++)
                {
                 
                    ExcelApp.Cells[i + 1, j + 1] = songsDataGridView.Rows[i-1].Cells[j].Value;
                }
            }        
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }


        private void SetupLayout()
        {
            this.Size = new Size(600, 500);
            addNewRowButton.Text = "Add Row";
            addNewRowButton.Location = new Point(10, 10);
            addNewRowButton.Click += new EventHandler(addNewRowButton_Click);
            deleteRowButton.Text = "Delete Row";
            deleteRowButton.Location = new Point(100, 10);
            deleteRowButton.Click += new EventHandler(deleteRowButton_Click);
            Button1.Text = "Sort";
            Button1.Location = new Point(200, 10);
            Button1.Click += new EventHandler(Button1_Click);
            Button2.Text = "Excel";
            Button2.Location = new Point(300, 10);
            Button2.Click += new EventHandler(Button2_Click);
            buttonPanel.Controls.Add(Button2);
            Buttontxt.Text = "Inport in txt";
            Buttontxt.Location = new Point(400, 10);
            Buttontxt.Click += new EventHandler(Buttontxt_Click);
            buttonPanel.Controls.Add(Buttontxt);
            buttonPanel.Controls.Add(addNewRowButton);
            buttonPanel.Controls.Add(deleteRowButton);
            buttonPanel.Controls.Add(Button1);
            buttonPanel.Height = 50;
            buttonPanel.Dock = DockStyle.Bottom;
            this.Controls.Add(this.buttonPanel);
        }
        private class RowComparer : System.Collections.IComparer
        {
            private static int sortOrderModifier = 1;

            public RowComparer(SortOrder sortOrder)
            {
                if (sortOrder == SortOrder.Descending)
                {
                    sortOrderModifier = -1;
                }
                else if (sortOrder == SortOrder.Ascending)
                {
                    sortOrderModifier = 1;
                }
            }

            public int Compare(object x, object y)
            {
                DataGridViewRow DataGridViewRow1 = (DataGridViewRow)x;
                DataGridViewRow DataGridViewRow2 = (DataGridViewRow)y;

                // Try to sort based on the Last Name column.
                int CompareResult = System.String.Compare(
                    DataGridViewRow1.Cells[1].Value.ToString(),
                    DataGridViewRow2.Cells[1].Value.ToString());

                // If the Last Names are equal, sort based on the First Name.
                if (CompareResult == 0)
                {
                    CompareResult = System.String.Compare(
                        DataGridViewRow1.Cells[0].Value.ToString(),
                        DataGridViewRow2.Cells[0].Value.ToString());
                }
                return CompareResult * sortOrderModifier;
            }
        }
        private void SetupDataGridView()
        {
            this.Controls.Add(songsDataGridView);
            songsDataGridView.ColumnCount = 5;
            songsDataGridView.ColumnHeadersDefaultCellStyle.BackColor =
           Color.Navy;
            songsDataGridView.ColumnHeadersDefaultCellStyle.ForeColor =
           Color.White;
            songsDataGridView.ColumnHeadersDefaultCellStyle.Font =
            new Font(songsDataGridView.Font, FontStyle.Bold);
            songsDataGridView.Name = "songsDataGridView";
            songsDataGridView.Location = new Point(8, 8);
            songsDataGridView.Size = new Size(500, 250);          
            songsDataGridView.AutoSizeRowsMode =
            DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            songsDataGridView.ColumnHeadersBorderStyle =
            DataGridViewHeaderBorderStyle.Single;
            songsDataGridView.CellBorderStyle =
           DataGridViewCellBorderStyle.Single;
            songsDataGridView.GridColor = Color.Black;
            songsDataGridView.RowHeadersVisible = false;
            songsDataGridView.Columns[3].Name = "Release Date";
            songsDataGridView.Columns[4].Name = "Track";
            songsDataGridView.Columns[0].Name = "Title";
            songsDataGridView.Columns[1].Name = "Artist";
            songsDataGridView.Columns[2].Name = "Album";
            songsDataGridView.Columns[4].DefaultCellStyle.Font =
            new Font(songsDataGridView.DefaultCellStyle.Font,
           FontStyle.Italic);
            songsDataGridView.DefaultCellStyle.WrapMode =
    DataGridViewTriState.True;
            songsDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            songsDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells; ;
            songsDataGridView.SelectionMode =
            DataGridViewSelectionMode.FullRowSelect;
            songsDataGridView.MultiSelect = false;
            songsDataGridView.Dock = DockStyle.Fill;
            songsDataGridView.Sort(new RowComparer(SortOrder.Ascending));
            songsDataGridView.CellFormatting += new
            DataGridViewCellFormattingEventHandler(
            songsDataGridView_CellFormatting);
        }
        private void PopulateDataGridView()
        {
            string[] row0 = { "Revolution 9",
 "Beatles", "The Beatles [White Album]", "11/22/1968", "29" };
            string[] row1 = { "Fools Rush In",
 "Frank Sinatra", "Nice 'N' Easy","1960", "6" };
            string[] row2 = {  "One of These Days",
 "Pink Floyd", "Meddle","11/11/1971", "1" };
            string[] row3 = {  "Where Is My Mind?",
 "Pixies", "Surfer Rosa","1988", "7" };
            string[] row4 = {  "Can't Find My Mind",
 "Cramps", "Psychedelic Jungle","5/1981", "9" };
            string[] row5 = {  "Scatterbrain. (As Dead As Leaves.)",
 "Radiohead", "Hail to the Thief","6/10/2003", "13" };
            string[] row6 = {  "Dress", "P J Harvey", "Dry", "6/30/1992", "3" };
            songsDataGridView.Rows.Add(row0);
            songsDataGridView.Rows.Add(row1);
            songsDataGridView.Rows.Add(row2);
            songsDataGridView.Rows.Add(row3);
            songsDataGridView.Rows.Add(row4);
            songsDataGridView.Rows.Add(row5);
            songsDataGridView.Rows.Add(row6);

        }
       

    }
}
