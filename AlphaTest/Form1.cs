using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;



namespace AlphaTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet ds = new DataSet("Channel");

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Width = 150;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 150; 

            btnExport.Enabled = true;
            btnFilter.Enabled = true;
      
        }

        private void button2_Click(object sender, EventArgs e)
        {
            btnSort.Enabled = true;
            string text;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i=0; i< dataGridView1.Rows.Count; i++)
                {
                    text = dataGridView1[3, i].Value.ToString();

                    if (!text.Contains("Политика"))
                    {
                        dataGridView1.Rows.RemoveAt(dataGridView1.CurrentRow.Index);
                    }
                }
            }
            dataGridView1.DataSource = ds.Tables[0];   
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ds.ReadXml(System.Windows.Forms.Application.StartupPath + @"\data.xml");
        }

        bool flagSort = false;
        private void btnSort_Click(object sender, EventArgs e)
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-CA");
            dataGridView1.Columns[4].DefaultCellStyle.Format = "ddd, dd MMM yyyy HH:mm:ss";

            System.Data.DataTable dt = ds.Tables[0];
            System.Data.DataTable dtCloned = dt.Clone();
            dtCloned.Columns[4].DataType = typeof(DateTime);

            foreach (DataRow row in dt.Rows)
            {
                dtCloned.ImportRow(row);
            }
            dataGridView1.DataSource = dtCloned;

            if (!flagSort)
            {
                this.dataGridView1.Sort(this.dataGridView1.Columns["pubDate"], ListSortDirection.Ascending);
                flagSort = true;
            }
            else
            {
                this.dataGridView1.Sort(this.dataGridView1.Columns["pubDate"], ListSortDirection.Descending);
                flagSort = false;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            int columns = dataGridView1.Columns.Count;
            int rows = dataGridView1.Rows.Count;
            string[,] array = new string[rows, columns];

            for (int i = 0; i <= dataGridView1.Rows.Count-1; i++)
                for (int j = 0; j <= dataGridView1.Columns.Count-1; j++)
                {
                    array[i, j] = dataGridView1[j, i].Value.ToString();
                }

            Word.Application application = new Word.Application();
            Object missing = Type.Missing;
            application.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Word.Document document = application.ActiveDocument;
            Word.Range range = application.Selection.Range;
            Object behiavor = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehiavor = Word.WdAutoFitBehavior.wdAutoFitFixed;
            document.Tables.Add(range, rows, columns, ref behiavor, ref autoFitBehiavor);
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    document.Tables[1].Cell(i + 1, j + 1).Range.Text = array[i, j].ToString();
            application.Visible = true;
        }
    }
}

