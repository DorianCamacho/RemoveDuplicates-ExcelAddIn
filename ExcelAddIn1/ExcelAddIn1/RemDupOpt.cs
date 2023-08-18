using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelAddIn1.ColumnSelection;
using System.Diagnostics;

namespace ExcelAddIn1
{
    public partial class RemDupOpt : Form
    {
        string RowF = RowFlag.MinRow;
        public RemDupOpt()
        {
            
            InitializeComponent();
        }

        public static class SharedData
        {
            public static string DataToShare {get;set;}
            public static string DataHeader {get;set;}
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void RemoveButton_Click(object sender, EventArgs e)
        { 
            if (radioExpSel.Checked)
            {
                SharedData.DataToShare = "ExpandSelection";
                if (DataHeadersBox.Checked) { SharedData.DataHeader = "Checked"; }
                else { SharedData.DataHeader = "Unchecked"; }
                ColumnSelection columnSelection = new ColumnSelection();
                if (RowFlag.MinRow == "true") { this.Close(); }
                else
                {
                    columnSelection.ShowDialog();
                    this.Close();
                }
            }
            else if (radioContSel.Checked)
            {
                SharedData.DataToShare = "ContinueSelection";
                if (DataHeadersBox.Checked) { SharedData.DataHeader = "Checked"; }
                else { SharedData.DataHeader = "Unchecked"; }
                ColumnSelection columnSelection = new ColumnSelection();
                if (RowFlag.MinRow == "true") { this.Close(); }
                else
                {
                    columnSelection.ShowDialog();
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Choose an option");
            }

        }

        private void Cancelbutton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        void radioExpSel_CheckedChanged(object sender, EventArgs e)
        {

        }

        void radioContSel_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void DataHeadersBox_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
