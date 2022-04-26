
namespace DomPecheyKP
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
            this.ProductType = new System.Windows.Forms.GroupBox();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.ClientName = new System.Windows.Forms.TextBox();
            this.ManagerName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ChimneyElements = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Manufacturer = new System.Windows.Forms.GroupBox();
            this.radioButton17 = new System.Windows.Forms.RadioButton();
            this.radioButton5 = new System.Windows.Forms.RadioButton();
            this.radioButton6 = new System.Windows.Forms.RadioButton();
            this.radioButton7 = new System.Windows.Forms.RadioButton();
            this.radioButton8 = new System.Windows.Forms.RadioButton();
            this.Diameter = new System.Windows.Forms.GroupBox();
            this.radioButton16 = new System.Windows.Forms.RadioButton();
            this.radioButton15 = new System.Windows.Forms.RadioButton();
            this.radioButton14 = new System.Windows.Forms.RadioButton();
            this.radioButton13 = new System.Windows.Forms.RadioButton();
            this.radioButton9 = new System.Windows.Forms.RadioButton();
            this.radioButton10 = new System.Windows.Forms.RadioButton();
            this.radioButton11 = new System.Windows.Forms.RadioButton();
            this.radioButton12 = new System.Windows.Forms.RadioButton();
            this.MetalThickness = new System.Windows.Forms.GroupBox();
            this.radioButton19 = new System.Windows.Forms.RadioButton();
            this.radioButton20 = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.Number = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NameElement = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CountElement = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PriceElement = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SumElement = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ProductType.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ChimneyElements)).BeginInit();
            this.Manufacturer.SuspendLayout();
            this.Diameter.SuspendLayout();
            this.MetalThickness.SuspendLayout();
            this.SuspendLayout();
            // 
            // ProductType
            // 
            this.ProductType.Controls.Add(this.radioButton4);
            this.ProductType.Controls.Add(this.radioButton3);
            this.ProductType.Controls.Add(this.radioButton2);
            this.ProductType.Controls.Add(this.radioButton1);
            this.ProductType.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ProductType.Location = new System.Drawing.Point(12, 51);
            this.ProductType.Name = "ProductType";
            this.ProductType.Size = new System.Drawing.Size(724, 63);
            this.ProductType.TabIndex = 0;
            this.ProductType.TabStop = false;
            this.ProductType.Text = "Тип товара";
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton4.Location = new System.Drawing.Point(561, 26);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(119, 23);
            this.radioButton4.TabIndex = 3;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "Банная печь";
            this.radioButton4.UseVisualStyleBackColor = true;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton3.Location = new System.Drawing.Point(387, 26);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(117, 23);
            this.radioButton3.TabIndex = 2;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "Печь-камин";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton2.Location = new System.Drawing.Point(6, 26);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(149, 23);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Каминная топка";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton1.Location = new System.Drawing.Point(169, 26);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(173, 23);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Отопительная печь";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "Имя клиента";
            // 
            // ClientName
            // 
            this.ClientName.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ClientName.Location = new System.Drawing.Point(117, 17);
            this.ClientName.Name = "ClientName";
            this.ClientName.Size = new System.Drawing.Size(182, 26);
            this.ClientName.TabIndex = 2;
            // 
            // ManagerName
            // 
            this.ManagerName.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ManagerName.Location = new System.Drawing.Point(412, 17);
            this.ManagerName.Name = "ManagerName";
            this.ManagerName.Size = new System.Drawing.Size(182, 26);
            this.ManagerName.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(325, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "Менеджер";
            // 
            // ChimneyElements
            // 
            this.ChimneyElements.AllowUserToAddRows = false;
            this.ChimneyElements.BackgroundColor = System.Drawing.Color.Wheat;
            this.ChimneyElements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ChimneyElements.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Number,
            this.NameElement,
            this.CountElement,
            this.PriceElement,
            this.SumElement});
            this.ChimneyElements.Location = new System.Drawing.Point(12, 358);
            this.ChimneyElements.Name = "ChimneyElements";
            this.ChimneyElements.RowTemplate.Height = 25;
            this.ChimneyElements.Size = new System.Drawing.Size(724, 150);
            this.ChimneyElements.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(12, 337);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(150, 18);
            this.label3.TabIndex = 6;
            this.label3.Text = "Элементы дымохода";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(12, 124);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 18);
            this.label4.TabIndex = 8;
            this.label4.Text = "Название";
            // 
            // Manufacturer
            // 
            this.Manufacturer.Controls.Add(this.radioButton17);
            this.Manufacturer.Controls.Add(this.radioButton5);
            this.Manufacturer.Controls.Add(this.radioButton6);
            this.Manufacturer.Controls.Add(this.radioButton7);
            this.Manufacturer.Controls.Add(this.radioButton8);
            this.Manufacturer.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.Manufacturer.Location = new System.Drawing.Point(12, 203);
            this.Manufacturer.Name = "Manufacturer";
            this.Manufacturer.Size = new System.Drawing.Size(724, 63);
            this.Manufacturer.TabIndex = 4;
            this.Manufacturer.TabStop = false;
            this.Manufacturer.Text = "Производитель";
            // 
            // radioButton17
            // 
            this.radioButton17.AutoSize = true;
            this.radioButton17.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton17.Location = new System.Drawing.Point(593, 26);
            this.radioButton17.Name = "radioButton17";
            this.radioButton17.Size = new System.Drawing.Size(84, 23);
            this.radioButton17.TabIndex = 4;
            this.radioButton17.TabStop = true;
            this.radioButton17.Text = "Permetr";
            this.radioButton17.UseVisualStyleBackColor = true;
            // 
            // radioButton5
            // 
            this.radioButton5.AutoSize = true;
            this.radioButton5.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton5.Location = new System.Drawing.Point(467, 28);
            this.radioButton5.Name = "radioButton5";
            this.radioButton5.Size = new System.Drawing.Size(62, 23);
            this.radioButton5.TabIndex = 3;
            this.radioButton5.TabStop = true;
            this.radioButton5.Text = "Craft";
            this.radioButton5.UseVisualStyleBackColor = true;
            // 
            // radioButton6
            // 
            this.radioButton6.AutoSize = true;
            this.radioButton6.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton6.Location = new System.Drawing.Point(325, 26);
            this.radioButton6.Name = "radioButton6";
            this.radioButton6.Size = new System.Drawing.Size(78, 23);
            this.radioButton6.TabIndex = 2;
            this.radioButton6.TabStop = true;
            this.radioButton6.Text = "Ferrum";
            this.radioButton6.UseVisualStyleBackColor = true;
            // 
            // radioButton7
            // 
            this.radioButton7.AutoSize = true;
            this.radioButton7.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton7.Location = new System.Drawing.Point(173, 26);
            this.radioButton7.Name = "radioButton7";
            this.radioButton7.Size = new System.Drawing.Size(88, 23);
            this.radioButton7.TabIndex = 1;
            this.radioButton7.TabStop = true;
            this.radioButton7.Text = "Везувий";
            this.radioButton7.UseVisualStyleBackColor = true;
            // 
            // radioButton8
            // 
            this.radioButton8.AutoSize = true;
            this.radioButton8.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton8.Location = new System.Drawing.Point(31, 26);
            this.radioButton8.Name = "radioButton8";
            this.radioButton8.Size = new System.Drawing.Size(78, 23);
            this.radioButton8.TabIndex = 0;
            this.radioButton8.TabStop = true;
            this.radioButton8.Text = "Дымок";
            this.radioButton8.UseVisualStyleBackColor = true;
            // 
            // Diameter
            // 
            this.Diameter.Controls.Add(this.radioButton16);
            this.Diameter.Controls.Add(this.radioButton15);
            this.Diameter.Controls.Add(this.radioButton14);
            this.Diameter.Controls.Add(this.radioButton13);
            this.Diameter.Controls.Add(this.radioButton9);
            this.Diameter.Controls.Add(this.radioButton10);
            this.Diameter.Controls.Add(this.radioButton11);
            this.Diameter.Controls.Add(this.radioButton12);
            this.Diameter.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.Diameter.Location = new System.Drawing.Point(214, 272);
            this.Diameter.Name = "Diameter";
            this.Diameter.Size = new System.Drawing.Size(522, 63);
            this.Diameter.TabIndex = 5;
            this.Diameter.TabStop = false;
            this.Diameter.Text = "Диамметр";
            this.Diameter.Enter += new System.EventHandler(this.Diameter_Enter);
            // 
            // radioButton16
            // 
            this.radioButton16.AutoSize = true;
            this.radioButton16.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton16.Location = new System.Drawing.Point(450, 28);
            this.radioButton16.Name = "radioButton16";
            this.radioButton16.Size = new System.Drawing.Size(52, 23);
            this.radioButton16.TabIndex = 7;
            this.radioButton16.TabStop = true;
            this.radioButton16.Text = "300";
            this.radioButton16.UseVisualStyleBackColor = true;
            // 
            // radioButton15
            // 
            this.radioButton15.AutoSize = true;
            this.radioButton15.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton15.Location = new System.Drawing.Point(384, 28);
            this.radioButton15.Name = "radioButton15";
            this.radioButton15.Size = new System.Drawing.Size(52, 23);
            this.radioButton15.TabIndex = 6;
            this.radioButton15.TabStop = true;
            this.radioButton15.Text = "250";
            this.radioButton15.UseVisualStyleBackColor = true;
            this.radioButton15.CheckedChanged += new System.EventHandler(this.radioButton15_CheckedChanged);
            // 
            // radioButton14
            // 
            this.radioButton14.AutoSize = true;
            this.radioButton14.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton14.Location = new System.Drawing.Point(317, 28);
            this.radioButton14.Name = "radioButton14";
            this.radioButton14.Size = new System.Drawing.Size(53, 23);
            this.radioButton14.TabIndex = 5;
            this.radioButton14.TabStop = true;
            this.radioButton14.Text = "200";
            this.radioButton14.UseVisualStyleBackColor = true;
            // 
            // radioButton13
            // 
            this.radioButton13.AutoSize = true;
            this.radioButton13.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton13.Location = new System.Drawing.Point(253, 28);
            this.radioButton13.Name = "radioButton13";
            this.radioButton13.Size = new System.Drawing.Size(50, 23);
            this.radioButton13.TabIndex = 4;
            this.radioButton13.TabStop = true;
            this.radioButton13.Text = "180";
            this.radioButton13.UseVisualStyleBackColor = true;
            // 
            // radioButton9
            // 
            this.radioButton9.AutoSize = true;
            this.radioButton9.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton9.Location = new System.Drawing.Point(190, 26);
            this.radioButton9.Name = "radioButton9";
            this.radioButton9.Size = new System.Drawing.Size(49, 23);
            this.radioButton9.TabIndex = 3;
            this.radioButton9.TabStop = true;
            this.radioButton9.Text = "150";
            this.radioButton9.UseVisualStyleBackColor = true;
            // 
            // radioButton10
            // 
            this.radioButton10.AutoSize = true;
            this.radioButton10.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton10.Location = new System.Drawing.Point(128, 26);
            this.radioButton10.Name = "radioButton10";
            this.radioButton10.Size = new System.Drawing.Size(48, 23);
            this.radioButton10.TabIndex = 2;
            this.radioButton10.TabStop = true;
            this.radioButton10.Text = "130";
            this.radioButton10.UseVisualStyleBackColor = true;
            this.radioButton10.CheckedChanged += new System.EventHandler(this.radioButton10_CheckedChanged);
            // 
            // radioButton11
            // 
            this.radioButton11.AutoSize = true;
            this.radioButton11.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton11.Location = new System.Drawing.Point(65, 26);
            this.radioButton11.Name = "radioButton11";
            this.radioButton11.Size = new System.Drawing.Size(49, 23);
            this.radioButton11.TabIndex = 1;
            this.radioButton11.TabStop = true;
            this.radioButton11.Text = "120";
            this.radioButton11.UseVisualStyleBackColor = true;
            // 
            // radioButton12
            // 
            this.radioButton12.AutoSize = true;
            this.radioButton12.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton12.Location = new System.Drawing.Point(6, 26);
            this.radioButton12.Name = "radioButton12";
            this.radioButton12.Size = new System.Drawing.Size(45, 23);
            this.radioButton12.TabIndex = 0;
            this.radioButton12.TabStop = true;
            this.radioButton12.Text = "115";
            this.radioButton12.UseVisualStyleBackColor = true;
            // 
            // MetalThickness
            // 
            this.MetalThickness.Controls.Add(this.radioButton19);
            this.MetalThickness.Controls.Add(this.radioButton20);
            this.MetalThickness.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.MetalThickness.Location = new System.Drawing.Point(12, 272);
            this.MetalThickness.Name = "MetalThickness";
            this.MetalThickness.Size = new System.Drawing.Size(196, 63);
            this.MetalThickness.TabIndex = 5;
            this.MetalThickness.TabStop = false;
            this.MetalThickness.Text = "Толщина металла";
            // 
            // radioButton19
            // 
            this.radioButton19.AutoSize = true;
            this.radioButton19.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton19.Location = new System.Drawing.Point(105, 26);
            this.radioButton19.Name = "radioButton19";
            this.radioButton19.Size = new System.Drawing.Size(75, 23);
            this.radioButton19.TabIndex = 1;
            this.radioButton19.TabStop = true;
            this.radioButton19.Text = "0,8 мм";
            this.radioButton19.UseVisualStyleBackColor = true;
            // 
            // radioButton20
            // 
            this.radioButton20.AutoSize = true;
            this.radioButton20.Font = new System.Drawing.Font("Constantia", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.radioButton20.Location = new System.Drawing.Point(6, 28);
            this.radioButton20.Name = "radioButton20";
            this.radioButton20.Size = new System.Drawing.Size(74, 23);
            this.radioButton20.TabIndex = 0;
            this.radioButton20.TabStop = true;
            this.radioButton20.Text = "0,5 мм";
            this.radioButton20.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(661, 124);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 9;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 514);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(360, 94);
            this.checkedListBox1.TabIndex = 10;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(378, 514);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(661, 514);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 12;
            this.button3.Text = "button3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // Number
            // 
            this.Number.HeaderText = "№ п.п.";
            this.Number.Name = "Number";
            this.Number.ReadOnly = true;
            // 
            // NameElement
            // 
            this.NameElement.FillWeight = 500F;
            this.NameElement.HeaderText = "Наименование";
            this.NameElement.Name = "NameElement";
            this.NameElement.ReadOnly = true;
            // 
            // CountElement
            // 
            this.CountElement.HeaderText = " Кол-во шт.";
            this.CountElement.Name = "CountElement";
            // 
            // PriceElement
            // 
            this.PriceElement.HeaderText = "Цена за 1 шт.";
            this.PriceElement.Name = "PriceElement";
            // 
            // SumElement
            // 
            this.SumElement.HeaderText = "Цена без скидки";
            this.SumElement.Name = "SumElement";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 613);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.MetalThickness);
            this.Controls.Add(this.Diameter);
            this.Controls.Add(this.Manufacturer);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ChimneyElements);
            this.Controls.Add(this.ManagerName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ClientName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProductType);
            this.Name = "Form1";
            this.Text = "Комерческое предложение";
            this.ProductType.ResumeLayout(false);
            this.ProductType.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ChimneyElements)).EndInit();
            this.Manufacturer.ResumeLayout(false);
            this.Manufacturer.PerformLayout();
            this.Diameter.ResumeLayout(false);
            this.Diameter.PerformLayout();
            this.MetalThickness.ResumeLayout(false);
            this.MetalThickness.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox ProductType;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton4;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox ClientName;
        private System.Windows.Forms.TextBox ManagerName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView ChimneyElements;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox Manufacturer;
        private System.Windows.Forms.RadioButton radioButton5;
        private System.Windows.Forms.RadioButton radioButton6;
        private System.Windows.Forms.RadioButton radioButton7;
        private System.Windows.Forms.RadioButton radioButton8;
        private System.Windows.Forms.GroupBox Diameter;
        private System.Windows.Forms.RadioButton radioButton9;
        private System.Windows.Forms.RadioButton radioButton10;
        private System.Windows.Forms.RadioButton radioButton11;
        private System.Windows.Forms.RadioButton radioButton12;
        private System.Windows.Forms.RadioButton radioButton15;
        private System.Windows.Forms.RadioButton radioButton14;
        private System.Windows.Forms.RadioButton radioButton13;
        private System.Windows.Forms.RadioButton radioButton16;
        private System.Windows.Forms.GroupBox MetalThickness;
        private System.Windows.Forms.RadioButton radioButton19;
        private System.Windows.Forms.RadioButton radioButton20;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.RadioButton radioButton17;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Number;
        private System.Windows.Forms.DataGridViewTextBoxColumn NameElement;
        private System.Windows.Forms.DataGridViewTextBoxColumn CountElement;
        private System.Windows.Forms.DataGridViewTextBoxColumn PriceElement;
        private System.Windows.Forms.DataGridViewTextBoxColumn SumElement;
    }
}

