namespace Lab5MSOffice
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.launchMSWordButton = new System.Windows.Forms.Button();
            this.launchMSExcelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(23, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(223, 66);
            this.label1.TabIndex = 1;
            this.label1.Text = "MS Word\r\nОткрытие существующего документа D:/Lab5.docx. Добавление текстовой стро" +
    "ки. Стиль вводимого теста Полужирный.";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(298, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(223, 46);
            this.label2.TabIndex = 2;
            this.label2.Text = "MS Excel\r\nОбъединение содержимого 2-х документов.";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 136);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Текст:";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(75, 133);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(149, 20);
            this.textBox2.TabIndex = 6;
            // 
            // launchMSWordButton
            // 
            this.launchMSWordButton.Location = new System.Drawing.Point(75, 181);
            this.launchMSWordButton.Name = "launchMSWordButton";
            this.launchMSWordButton.Size = new System.Drawing.Size(75, 23);
            this.launchMSWordButton.TabIndex = 9;
            this.launchMSWordButton.Text = "Выполнить";
            this.launchMSWordButton.UseVisualStyleBackColor = true;
            this.launchMSWordButton.Click += new System.EventHandler(this.launchMSWordButton_Click);
            // 
            // launchMSExcelButton
            // 
            this.launchMSExcelButton.Location = new System.Drawing.Point(345, 112);
            this.launchMSExcelButton.Name = "launchMSExcelButton";
            this.launchMSExcelButton.Size = new System.Drawing.Size(75, 23);
            this.launchMSExcelButton.TabIndex = 10;
            this.launchMSExcelButton.Text = "Выполнить";
            this.launchMSExcelButton.UseVisualStyleBackColor = true;
            this.launchMSExcelButton.Click += new System.EventHandler(this.launchMSExcelButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(533, 289);
            this.Controls.Add(this.launchMSExcelButton);
            this.Controls.Add(this.launchMSWordButton);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button launchMSWordButton;
        private System.Windows.Forms.Button launchMSExcelButton;
    }
}

