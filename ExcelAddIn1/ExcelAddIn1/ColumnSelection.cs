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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using static ExcelAddIn1.RemDupOpt;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ColumnSelection : Form
    {
        private List<int> selectedColumns;
        string RemDupOptData = SharedData.DataToShare;
        string RemData = SharedData.DataHeader;

        public ColumnSelection()
        {   
            InitializeComponent();
            if (RemDupOptData == "ContinueSelection")
            {
                SelectedBox();
            }
            else
            {
                UsedBox();
            }
        }

        public static class RowFlag
        {
            public static string MinRow { get; set; }
        }

        private void ColumnSelection_Load(object sender, EventArgs e)
        {
            
        }

        public void SelectedBox()
        {
            RowFlag.MinRow = "false";
            Excel.Application activeApplication = Globals.ThisAddIn.Application;
            Range selectedRange = activeApplication.Selection;
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanel1.WrapContents = false;
            flowLayoutPanel1.AutoScroll = true;
            int cols = selectedRange.Columns.Count;
            int rows = selectedRange.Rows.Count;

            if (rows == 1)
            {
                MessageBox.Show("No duplicates to eliminate");
                RowFlag.MinRow = "true";
                //System.Windows.Forms.Application.ExitThread();
            }
            else
            {
                Excel.Range firstRow = selectedRange.Rows[1] as Excel.Range;
                for (int colIndex = 1; colIndex <= cols; colIndex++)
                {
                    Excel.Range cell = firstRow.Cells[1, colIndex] as Excel.Range;
                    string columnName = cell.Value != null ? cell.Value.ToString() : string.Empty;

                    System.Windows.Forms.CheckBox checkBox = new System.Windows.Forms.CheckBox();
                    if (RemData == "Checked") { checkBox.Text = columnName; }
                    else { checkBox.Text = colIndex.ToString(); }
                    checkBox.Tag = colIndex;
                    checkBox.AutoSize = true;
                    checkBox.Anchor = AnchorStyles.Left;
                    flowLayoutPanel1.Controls.Add(checkBox);
                }
            }
        }

        public void UsedBox()
        {
            RowFlag.MinRow = "false";
            Excel.Application activeApplication = Globals.ThisAddIn.Application;
            Excel.Worksheet worksheet = activeApplication.ActiveSheet;
            Excel.Range usedRange = worksheet.UsedRange;
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanel1.WrapContents = false;
            flowLayoutPanel1.AutoScroll = true;
            int cols = usedRange.Columns.Count;
            int rows = usedRange.Rows.Count;

            if (rows == 1)
            {
                RowFlag.MinRow = "true";
                MessageBox.Show("No duplicates to eliminate");
            }
            else
            {
                Excel.Range firstRow = usedRange.Rows[1] as Excel.Range;
                for (int colIndex = 1; colIndex <= cols; colIndex++)
                {
                    Excel.Range cell = firstRow.Cells[1, colIndex] as Excel.Range;
                    string columnName = cell.Value != null ? cell.Value.ToString() : string.Empty;

                    System.Windows.Forms.CheckBox checkBox = new System.Windows.Forms.CheckBox();
                    if (RemData == "Checked") { checkBox.Text = columnName; }
                    else { checkBox.Text = colIndex.ToString(); }
                    checkBox.Tag = colIndex;
                    checkBox.AutoSize = true;
                    checkBox.Anchor = AnchorStyles.Left;
                    flowLayoutPanel1.Controls.Add(checkBox);
                }
            }
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SelectAllButton_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.CheckBox checkBox in flowLayoutPanel1.Controls)
            {
                checkBox.Checked = true;
            }
        }

        private void UnselctAllButton_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.CheckBox checkBox in flowLayoutPanel1.Controls)
            {
                checkBox.Checked = false;
            }
        }

        private void DataHeaderCheckBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            selectedColumns = new List<int>();
            int z = 0;
            foreach (System.Windows.Forms.CheckBox checkBox in flowLayoutPanel1.Controls.OfType<System.Windows.Forms.CheckBox>())
            {
                if (checkBox.Checked)
                {
                    int columnIndex = (int)checkBox.Tag;
                    selectedColumns.Add(columnIndex);
                }
                else
                {
                    selectedColumns.Add(z);
                }
            }

            string RemDupOptData = SharedData.DataToShare;
            
            if (RemDupOptData == "ContinueSelection")
            {
                SelectedRangeDuplicates();
            }
            else
            {
                UsedRangeDuplicates();
            }
            this.Close();
        }

        public void SelectedRangeDuplicates()
        {
            Excel.Application activeApplication = Globals.ThisAddIn.Application;
            Range selectedRange = activeApplication.Selection;
            Excel.Worksheet worksheet = activeApplication.ActiveSheet;

            int rows = selectedRange.Rows.Count;
            int cols = selectedRange.Columns.Count;
            const string sep = "<sep>";
            string[] rowData = new string[rows];

            if (rows == 1)
            {
                MessageBox.Show("No duplicates to eliminate");
            }
            else
            {
                int i;

                if (RemData == "Checked")
                {
                    i = 1;
                }
                else
                {
                    i = 0;
                }
                for (i += 1; i <= rows; i++)
                {
                    int f = 0;
                    string ConcatRows = string.Empty;
                    for (int j = 1; j <= cols; j++)
                    {
                        if (selectedColumns[f] != 0  && f <= selectedColumns.Count-1)
                        {
                            Excel.Range cell = selectedRange.Cells[i, j] as Excel.Range;
                            if (cell.Value == null)
                            {
                                ConcatRows += sep;
                            }
                            else
                            {
                                ConcatRows += cell.Value.ToString() + sep;
                            }
                        }
                        f++;
                    }
                    if (rowData.Contains(ConcatRows))
                    {
                        Excel.Range deleteRow = selectedRange.Rows[i] as Excel.Range;
                        deleteRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        rows--;
                        i--;
                    }
                    else
                    {
                        rowData[i - 1] = ConcatRows;
                    }
                }
            }
        }

        public void UsedRangeDuplicates()
        {
            Excel.Application activeApplication = Globals.ThisAddIn.Application;
            Excel.Worksheet worksheet = activeApplication.ActiveSheet;
            Excel.Range usedRange = worksheet.UsedRange;

            int rows = usedRange.Rows.Count;
            int cols = usedRange.Columns.Count;
            const string sep = "<sep>";
            string[] rowData = new string[rows];

            if (rows == 1)
            {
                MessageBox.Show("No duplicates to eliminate");
            }
            else
            {
                int i;

                if (RemData == "Checked")
                {
                    i = 1;
                }
                else
                {
                    i = 0;
                }
                for (i += 1; i <= rows; i++)
                {
                    int f = 0;
                    string ConcatRows = string.Empty;
                    for (int j = 1; j <= cols; j++)
                    {
                        if (selectedColumns[f] != 0 && f <= selectedColumns.Count - 1)
                        {
                            Excel.Range cell = usedRange.Cells[i, j] as Excel.Range;
                            if (cell.Value == null)
                            {
                                ConcatRows += sep;
                            }
                            else
                            {
                                ConcatRows += cell.Value.ToString() + sep;
                            }
                        }
                        f++;
                    }
                    if (rowData.Contains(ConcatRows))
                    {
                        Excel.Range deleteRow = usedRange.Rows[i] as Excel.Range;
                        deleteRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        rows--;
                        i--;
                    }
                    else
                    {
                        rowData[i - 1] = ConcatRows;
                    }
                }
            }
        } 
    }
}


