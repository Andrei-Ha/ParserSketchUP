
namespace ParserSketchUP
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label_info = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label_result = new System.Windows.Forms.Label();
            this.label_result2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(597, 37);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Открыть .csv файл отчета";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(30, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(52, 15);
            this.label1.TabIndex = 1;
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Text = "path_file";
            // 
            // label_info
            // 
            this.label_info.AutoSize = true;
            this.label_info.Location = new System.Drawing.Point(1, 1);
            this.label_info.Name = "label_info";
            this.label_info.Size = new System.Drawing.Size(52, 15);
            this.label_info.TabIndex = 1;
            this.label_info.ForeColor = System.Drawing.Color.Red;
            this.label_info.Text = "Для правильной генерации данных при создании отчёта в SketchUP необходимо:\n"+
                " из \"Группировать по:\" убрать DefenitionName \n" +
                "а в \"Атрибутах отчета\" должно быть лишь: Name, Description";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(597, 80);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(150, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Сгенерировать .xls";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label_result
            // 
            this.label_result.AutoSize = true;
            this.label_result.Location = new System.Drawing.Point(100, 110);
            this.label_result.Name = "label_result";
            this.label_result.Size = new System.Drawing.Size(0, 15);
            this.label_result.TabIndex = 3;
            // 
            // label_result2
            // 
            this.label_result2.AutoSize = true;
            this.label_result2.Location = new System.Drawing.Point(400, 110);
            this.label_result2.Name = "label_result2";
            this.label_result2.Size = new System.Drawing.Size(0, 15);
            this.label_result2.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label_result2);
            this.Controls.Add(this.label_result);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label_info);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Parser";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label_info;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label_result;
        private System.Windows.Forms.Label label_result2;
    }
}

