namespace Excel_Upload_To_DataGridView_SQL_Server
{
    partial class Form1
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
            this.txt_Url = new System.Windows.Forms.TextBox();
            this.btn_format = new System.Windows.Forms.Button();
            this.btn_ExcelUpload = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_save = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // txt_Url
            // 
            this.txt_Url.Location = new System.Drawing.Point(25, 13);
            this.txt_Url.Name = "txt_Url";
            this.txt_Url.Size = new System.Drawing.Size(756, 22);
            this.txt_Url.TabIndex = 0;
            this.txt_Url.Click += new System.EventHandler(this.txt_Url_Click);
            // 
            // btn_format
            // 
            this.btn_format.Location = new System.Drawing.Point(787, 12);
            this.btn_format.Name = "btn_format";
            this.btn_format.Size = new System.Drawing.Size(125, 23);
            this.btn_format.TabIndex = 1;
            this.btn_format.Text = "Format Excel";
            this.btn_format.UseVisualStyleBackColor = true;
            this.btn_format.Click += new System.EventHandler(this.btn_format_Click);
            // 
            // btn_ExcelUpload
            // 
            this.btn_ExcelUpload.Location = new System.Drawing.Point(918, 13);
            this.btn_ExcelUpload.Name = "btn_ExcelUpload";
            this.btn_ExcelUpload.Size = new System.Drawing.Size(127, 23);
            this.btn_ExcelUpload.TabIndex = 2;
            this.btn_ExcelUpload.Text = "Upload Excel";
            this.btn_ExcelUpload.UseVisualStyleBackColor = true;
            this.btn_ExcelUpload.Click += new System.EventHandler(this.btn_ExcelUpload_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(25, 64);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1020, 432);
            this.dataGridView1.TabIndex = 3;
            // 
            // btn_save
            // 
            this.btn_save.Location = new System.Drawing.Point(441, 502);
            this.btn_save.Name = "btn_save";
            this.btn_save.Size = new System.Drawing.Size(125, 40);
            this.btn_save.TabIndex = 4;
            this.btn_save.Text = "Save";
            this.btn_save.UseVisualStyleBackColor = true;
            this.btn_save.Click += new System.EventHandler(this.btn_save_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.btn_save);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btn_ExcelUpload);
            this.Controls.Add(this.btn_format);
            this.Controls.Add(this.txt_Url);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Form1";
            this.Text = "Excel Data Upload";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_Url;
        private System.Windows.Forms.Button btn_format;
        private System.Windows.Forms.Button btn_ExcelUpload;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_save;
    }
}

