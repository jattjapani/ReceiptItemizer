namespace ReceiptItemizer
{
    partial class frmReceiptItemizer
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
            this.dgvItemsDisplay = new System.Windows.Forms.DataGridView();
            this.ItemDsc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ItemAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Tax = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Coupon = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AfterTax = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OneThird = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OneFourth = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtItemDesc = new System.Windows.Forms.TextBox();
            this.txtItemAmount = new System.Windows.Forms.TextBox();
            this.txtTaxRate = new System.Windows.Forms.TextBox();
            this.txtCoupon = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnClearRow = new System.Windows.Forms.Button();
            this.cmbCity = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.ckbNonTaxable = new System.Windows.Forms.CheckBox();
            this.dgvTotalDisplay = new System.Windows.Forms.DataGridView();
            this.ItemsTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TotalTax = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CouponTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Total1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OneThirdTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OneFourthTotal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnClearField = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvItemsDisplay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTotalDisplay)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvItemsDisplay
            // 
            this.dgvItemsDisplay.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvItemsDisplay.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ItemDsc,
            this.ItemAmount,
            this.Tax,
            this.Coupon,
            this.AfterTax,
            this.OneThird,
            this.OneFourth});
            this.dgvItemsDisplay.Location = new System.Drawing.Point(254, 26);
            this.dgvItemsDisplay.Name = "dgvItemsDisplay";
            this.dgvItemsDisplay.Size = new System.Drawing.Size(743, 242);
            this.dgvItemsDisplay.TabIndex = 12;
            // 
            // ItemDsc
            // 
            this.ItemDsc.HeaderText = "Item Description";
            this.ItemDsc.Name = "ItemDsc";
            // 
            // ItemAmount
            // 
            this.ItemAmount.HeaderText = "Item Amount";
            this.ItemAmount.Name = "ItemAmount";
            // 
            // Tax
            // 
            this.Tax.HeaderText = "Tax";
            this.Tax.Name = "Tax";
            // 
            // Coupon
            // 
            this.Coupon.HeaderText = "Coupon";
            this.Coupon.Name = "Coupon";
            // 
            // AfterTax
            // 
            this.AfterTax.HeaderText = "Total After Tax ";
            this.AfterTax.Name = "AfterTax";
            // 
            // OneThird
            // 
            this.OneThird.HeaderText = "One Third";
            this.OneThird.Name = "OneThird";
            // 
            // OneFourth
            // 
            this.OneFourth.HeaderText = "One Fourth";
            this.OneFourth.Name = "OneFourth";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 15;
            this.label1.Text = "Item Description";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Item Amount";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(51, 139);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Tax Rate";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(58, 192);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Coupon";
            // 
            // txtItemDesc
            // 
            this.txtItemDesc.Enabled = false;
            this.txtItemDesc.Location = new System.Drawing.Point(108, 63);
            this.txtItemDesc.Name = "txtItemDesc";
            this.txtItemDesc.Size = new System.Drawing.Size(100, 20);
            this.txtItemDesc.TabIndex = 1;
            // 
            // txtItemAmount
            // 
            this.txtItemAmount.Enabled = false;
            this.txtItemAmount.Location = new System.Drawing.Point(108, 99);
            this.txtItemAmount.Name = "txtItemAmount";
            this.txtItemAmount.Size = new System.Drawing.Size(100, 20);
            this.txtItemAmount.TabIndex = 2;
            // 
            // txtTaxRate
            // 
            this.txtTaxRate.Enabled = false;
            this.txtTaxRate.Location = new System.Drawing.Point(108, 135);
            this.txtTaxRate.Name = "txtTaxRate";
            this.txtTaxRate.Size = new System.Drawing.Size(100, 20);
            this.txtTaxRate.TabIndex = 3;
            // 
            // txtCoupon
            // 
            this.txtCoupon.Enabled = false;
            this.txtCoupon.Location = new System.Drawing.Point(108, 189);
            this.txtCoupon.Name = "txtCoupon";
            this.txtCoupon.Size = new System.Drawing.Size(100, 20);
            this.txtCoupon.TabIndex = 4;
            this.txtCoupon.Text = "0";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(115, 230);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(39, 23);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(870, 405);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(52, 23);
            this.btnUpdate.TabIndex = 6;
            this.btnUpdate.Text = "Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnClearRow
            // 
            this.btnClearRow.Location = new System.Drawing.Point(172, 230);
            this.btnClearRow.Name = "btnClearRow";
            this.btnClearRow.Size = new System.Drawing.Size(66, 23);
            this.btnClearRow.TabIndex = 10;
            this.btnClearRow.Text = "Clear Row";
            this.btnClearRow.UseVisualStyleBackColor = true;
            this.btnClearRow.Click += new System.EventHandler(this.btnClearRow_Click);
            // 
            // cmbCity
            // 
            this.cmbCity.FormattingEnabled = true;
            this.cmbCity.Items.AddRange(new object[] {
            "Clovis",
            "Fowler",
            "Fresno",
            "Hanford",
            "Selma",
            "Tulare",
            "Visalia"});
            this.cmbCity.Location = new System.Drawing.Point(108, 26);
            this.cmbCity.Name = "cmbCity";
            this.cmbCity.Size = new System.Drawing.Size(100, 21);
            this.cmbCity.Sorted = true;
            this.cmbCity.TabIndex = 0;
            this.cmbCity.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(78, 34);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(24, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "City";
            // 
            // ckbNonTaxable
            // 
            this.ckbNonTaxable.AutoSize = true;
            this.ckbNonTaxable.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.ckbNonTaxable.Location = new System.Drawing.Point(121, 166);
            this.ckbNonTaxable.Name = "ckbNonTaxable";
            this.ckbNonTaxable.Size = new System.Drawing.Size(87, 17);
            this.ckbNonTaxable.TabIndex = 11;
            this.ckbNonTaxable.Text = "Non Taxable";
            this.ckbNonTaxable.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.ckbNonTaxable.UseVisualStyleBackColor = true;
            // 
            // dgvTotalDisplay
            // 
            this.dgvTotalDisplay.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTotalDisplay.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ItemsTotal,
            this.TotalTax,
            this.CouponTotal,
            this.Total1,
            this.OneThirdTotal,
            this.OneFourthTotal});
            this.dgvTotalDisplay.Location = new System.Drawing.Point(354, 291);
            this.dgvTotalDisplay.Name = "dgvTotalDisplay";
            this.dgvTotalDisplay.Size = new System.Drawing.Size(643, 99);
            this.dgvTotalDisplay.TabIndex = 13;
            // 
            // ItemsTotal
            // 
            this.ItemsTotal.HeaderText = "Items Total";
            this.ItemsTotal.Name = "ItemsTotal";
            // 
            // TotalTax
            // 
            this.TotalTax.HeaderText = "TotalTax";
            this.TotalTax.Name = "TotalTax";
            // 
            // CouponTotal
            // 
            this.CouponTotal.HeaderText = "Coupon Total";
            this.CouponTotal.Name = "CouponTotal";
            // 
            // Total1
            // 
            this.Total1.HeaderText = "Total";
            this.Total1.Name = "Total1";
            // 
            // OneThirdTotal
            // 
            this.OneThirdTotal.HeaderText = "One Third Total";
            this.OneThirdTotal.Name = "OneThirdTotal";
            // 
            // OneFourthTotal
            // 
            this.OneFourthTotal.HeaderText = "One Fourth Total";
            this.OneFourthTotal.Name = "OneFourthTotal";
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(945, 405);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(52, 23);
            this.btnExit.TabIndex = 9;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(792, 405);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(55, 23);
            this.btnClearAll.TabIndex = 8;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.UseVisualStyleBackColor = true;
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnClearField
            // 
            this.btnClearField.Location = new System.Drawing.Point(23, 230);
            this.btnClearField.Name = "btnClearField";
            this.btnClearField.Size = new System.Drawing.Size(74, 23);
            this.btnClearField.TabIndex = 7;
            this.btnClearField.Text = "Clear Field";
            this.btnClearField.UseVisualStyleBackColor = true;
            this.btnClearField.Click += new System.EventHandler(this.btnClearField_Click);
            // 
            // frmReceiptCalculator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1035, 450);
            this.Controls.Add(this.dgvTotalDisplay);
            this.Controls.Add(this.ckbNonTaxable);
            this.Controls.Add(this.cmbCity);
            this.Controls.Add(this.btnClearAll);
            this.Controls.Add(this.btnClearRow);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnClearField);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtCoupon);
            this.Controls.Add(this.txtTaxRate);
            this.Controls.Add(this.txtItemAmount);
            this.Controls.Add(this.txtItemDesc);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgvItemsDisplay);
            this.Name = "frmReceiptCalculator";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dgvItemsDisplay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTotalDisplay)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvItemsDisplay;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtItemDesc;
        private System.Windows.Forms.TextBox txtItemAmount;
        private System.Windows.Forms.TextBox txtTaxRate;
        private System.Windows.Forms.TextBox txtCoupon;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnClearRow;
        private System.Windows.Forms.ComboBox cmbCity;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox ckbNonTaxable;
        private System.Windows.Forms.DataGridView dgvTotalDisplay;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemsTotal;
        private System.Windows.Forms.DataGridViewTextBoxColumn TotalTax;
        private System.Windows.Forms.DataGridViewTextBoxColumn CouponTotal;
        private System.Windows.Forms.DataGridViewTextBoxColumn Total1;
        private System.Windows.Forms.DataGridViewTextBoxColumn OneThirdTotal;
        private System.Windows.Forms.DataGridViewTextBoxColumn OneFourthTotal;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.Button btnClearField;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemDsc;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemAmount;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tax;
        private System.Windows.Forms.DataGridViewTextBoxColumn Coupon;
        private System.Windows.Forms.DataGridViewTextBoxColumn AfterTax;
        private System.Windows.Forms.DataGridViewTextBoxColumn OneThird;
        private System.Windows.Forms.DataGridViewTextBoxColumn OneFourth;
    }
}

