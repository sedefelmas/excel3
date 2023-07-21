﻿using System;
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
            string filePath = string.Empty;

            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.FileName = "*.xlsx";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet workSheet = workBook.WorkSheets[0];
                Range usedRange = workSheet;
                int rowCount = workSheet.RowCount;
                int colCount = workSheet.ColumnCount;
                string cellValue;
                ICollection<string> cellRange = workSheet.GetNamedRanges();
                //List<string> list = new List<string>();
                foreach (var cell in workSheet["A1:D7"])
                {
                    if(cell.ToString()!= "" && cell != null) 
                    {
                    //list.Add(cell.ToString());
                    listBox1.Items.Add(cell.ToString());
                    } 
                }
                //foreach(var cell in cellRange)
                //{
                //    cell = 
                //    listBox1.Items.Add(cell);
                //}                
                //for (int i = 1; i <= rowCount; i++)
                //{
                //    for (int j = 1; j < colCount; j++)
                //    {
                //        if (workSheet.GetCellAt(i, j).Value.ToString() != null)
                //        {
                //            cellValue = workSheet.GetCellAt(i, j).StringValue;
                //            listBox1.Items.Add(cellValue);
                //        }
                //    }
                //}
            }
            //listBox1.Items.Add(list);


        }
    }
}
