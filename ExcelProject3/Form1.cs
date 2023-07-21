using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;

namespace ExcelProject3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strFilePath = string.Empty;

            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.FileName = "*.xlsx";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                strFilePath = openFileDialog1.FileName;
                WorkBook workBook = WorkBook.Load(strFilePath);
                WorkSheet workSheet = workBook.WorkSheets[0];
                WorkSheet firstSheet = workBook.DefaultWorkSheet;
                int rowCount = firstSheet.RowCount;
                int colCount = firstSheet.ColumnCount;
                //List<string> list = new List<string>();
                foreach (var cell in workSheet["A1:D6"])
                {
                    //list.Add(cell.ToString());
                    listBox1.Items.Add(cell.ToString());
                }
            }
            //listBox1.Items.Add(list);
            //for(int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 0; j < colCount; j++)
            //    {
            //        cell = workSheet[ij];
            //        listBox1.Items.Add(cell.Value.ToString());
            //    }
            //}

        }
    }
}
