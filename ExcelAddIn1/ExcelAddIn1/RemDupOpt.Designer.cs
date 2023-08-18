namespace ExcelAddIn1
{
    partial class RemDupOpt
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.radioExpSel = new System.Windows.Forms.RadioButton();
            this.radioContSel = new System.Windows.Forms.RadioButton();
            this.Cancelbutton = new System.Windows.Forms.Button();
            this.RemoveButton = new System.Windows.Forms.Button();
            this.DataHeadersBox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Control;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(6, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(389, 35);
            this.textBox1.TabIndex = 1;
            this.textBox1.TabStop = false;
            this.textBox1.Text = "Microsoft Excel found data next to your selection. Because you have not selected " +
    "this data, it will not be removed.";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Control;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(6, 52);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(160, 16);
            this.textBox2.TabIndex = 2;
            this.textBox2.TabStop = false;
            this.textBox2.Text = "What do you want to do?";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // radioExpSel
            // 
            this.radioExpSel.AutoSize = true;
            this.radioExpSel.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioExpSel.Location = new System.Drawing.Point(12, 74);
            this.radioExpSel.Name = "radioExpSel";
            this.radioExpSel.Size = new System.Drawing.Size(133, 20);
            this.radioExpSel.TabIndex = 4;
            this.radioExpSel.TabStop = true;
            this.radioExpSel.Text = "Expand the selection";
            this.radioExpSel.UseVisualStyleBackColor = true;
            this.radioExpSel.CheckedChanged += new System.EventHandler(this.radioExpSel_CheckedChanged);
            // 
            // radioContSel
            // 
            this.radioContSel.AutoSize = true;
            this.radioContSel.Font = new System.Drawing.Font("Microsoft Tai Le", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioContSel.Location = new System.Drawing.Point(12, 99);
            this.radioContSel.Name = "radioContSel";
            this.radioContSel.Size = new System.Drawing.Size(211, 20);
            this.radioContSel.TabIndex = 5;
            this.radioContSel.TabStop = true;
            this.radioContSel.Text = "Continue with the current selection";
            this.radioContSel.UseVisualStyleBackColor = true;
            this.radioContSel.CheckedChanged += new System.EventHandler(this.radioContSel_CheckedChanged);
            // 
            // Cancelbutton
            // 
            this.Cancelbutton.Location = new System.Drawing.Point(312, 147);
            this.Cancelbutton.Name = "Cancelbutton";
            this.Cancelbutton.Size = new System.Drawing.Size(83, 26);
            this.Cancelbutton.TabIndex = 6;
            this.Cancelbutton.Text = "Cancel";
            this.Cancelbutton.UseVisualStyleBackColor = true;
            this.Cancelbutton.Click += new System.EventHandler(this.Cancelbutton_Click);
            // 
            // RemoveButton
            // 
            this.RemoveButton.Location = new System.Drawing.Point(174, 147);
            this.RemoveButton.Name = "RemoveButton";
            this.RemoveButton.Size = new System.Drawing.Size(123, 25);
            this.RemoveButton.TabIndex = 7;
            this.RemoveButton.Text = "Remove Duplicates...";
            this.RemoveButton.UseVisualStyleBackColor = true;
            this.RemoveButton.Click += new System.EventHandler(this.RemoveButton_Click);
            // 
            // DataHeadersBox
            // 
            this.DataHeadersBox.AutoSize = true;
            this.DataHeadersBox.Location = new System.Drawing.Point(12, 125);
            this.DataHeadersBox.Name = "DataHeadersBox";
            this.DataHeadersBox.Size = new System.Drawing.Size(125, 17);
            this.DataHeadersBox.TabIndex = 8;
            this.DataHeadersBox.Text = "My data has headers";
            this.DataHeadersBox.UseVisualStyleBackColor = true;
            this.DataHeadersBox.CheckedChanged += new System.EventHandler(this.DataHeadersBox_CheckedChanged);
            // 
            // RemDupOpt
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(407, 185);
            this.Controls.Add(this.DataHeadersBox);
            this.Controls.Add(this.RemoveButton);
            this.Controls.Add(this.Cancelbutton);
            this.Controls.Add(this.radioContSel);
            this.Controls.Add(this.radioExpSel);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Name = "RemDupOpt";
            this.Text = "Remove Duplicates (Case Sensitive)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.RadioButton radioExpSel;
        private System.Windows.Forms.RadioButton radioContSel;
        private System.Windows.Forms.Button Cancelbutton;
        private System.Windows.Forms.Button RemoveButton;
        private System.Windows.Forms.CheckBox DataHeadersBox;
    }
}