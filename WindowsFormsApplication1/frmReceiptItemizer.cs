using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReceiptItemizer
{
    public partial class frmReceiptItemizer : Form
    {
        public frmReceiptItemizer()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String City;
            decimal CurrentTax;

            City = cmbCity.Text;
            TaxRate T = new TaxRate();

            CurrentTax = T.SetTax(City);

            txtTaxRate.Text = CurrentTax.ToString();
            txtItemDesc.Enabled = true;
            txtItemAmount.Enabled = true;
            txtCoupon.Enabled = true;

        }
        private void btnClearField_Click(object sender, EventArgs e)
        {
            txtItemDesc.Text = "";
            txtItemAmount.Text = "";
            txtCoupon.Text = "0";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            decimal dItemAmount;
            decimal dTax;
            decimal dCoupon = 0;
            decimal dAfterTax;
            decimal dOneThird;
            decimal dOneFourth;
                      
            Calculate C = new Calculate();

            try
            {
                dItemAmount = Convert.ToDecimal(txtItemAmount.Text);
                dTax = C.Tax(Convert.ToDecimal(txtTaxRate.Text), dItemAmount);
                dCoupon = Convert.ToDecimal(txtCoupon.Text);

                if (ckbNonTaxable.Checked == false)
                {
                    dAfterTax = C.TotalAfterTax(Convert.ToDecimal(txtTaxRate.Text), dItemAmount);
                    dAfterTax -= Convert.ToDecimal(txtCoupon.Text);
                }
                else
                {

                    dAfterTax = Convert.ToDecimal(txtItemAmount.Text);
                    dAfterTax -= Convert.ToDecimal(txtCoupon.Text);
                    dTax = 0;
                }

                dOneThird = C.OneThird(dAfterTax);
                dOneFourth = C.OneFourth(dAfterTax);

                dgvItemsDisplay.Rows.Add(txtItemDesc.Text, dItemAmount.ToString("#.##"), dTax.ToString("#.##"), dCoupon.ToString("#.##"), 
                                          dAfterTax.ToString("#.##"), dOneThird.ToString("#.##"), dOneFourth.ToString("#.##"));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }           
            
        }

        private void btnClearRow_Click(object sender, EventArgs e)
        {
            try
            {
                dgvItemsDisplay.Rows.RemoveAt(dgvItemsDisplay.SelectedRows[0].Index);
            }
            catch (Exception)
            {
                MessageBox.Show("No Row Selected!");
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            decimal dItemAmountTotal = 0;
            decimal dTotalTax = 0;
            decimal dCouponTotal = 0;
            decimal dTotal = 0;
            decimal dOneThirdTotal = 0;
            decimal dOneFourthTotal = 0;

            dgvTotalDisplay.Rows.Clear();
           
            try
            {
              for (int i = 0; i < dgvItemsDisplay.Rows.Count; i++)
			{
                dItemAmountTotal += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[1].Value);
                if (dgvItemsDisplay.Rows[i].Cells[2].Value != "") 
                {
                   dTotalTax += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[2].Value);
                }
                
                if (dgvItemsDisplay.Rows[i].Cells[3].Value != "")
                {
                    dCouponTotal += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[3].Value);
                }
                dTotal += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[4].Value);
                dOneThirdTotal += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[5].Value);
                dOneFourthTotal += Convert.ToDecimal(dgvItemsDisplay.Rows[i].Cells[6].Value);
                
			}
              dgvTotalDisplay.Rows.Add(dItemAmountTotal.ToString("c"), dTotalTax.ToString("c"), dCouponTotal.ToString("c"), dTotal.ToString("c"),
                                       dOneThirdTotal.ToString("c"), dOneFourthTotal.ToString("c"));
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                MessageBox.Show(ex.StackTrace);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            try
            {
                dgvItemsDisplay.Rows.Clear();
                dgvTotalDisplay.Rows.Clear();
                txtItemDesc.Text = "";
                txtItemAmount.Text = "";
                txtCoupon.Text = "0";
            }
            catch (Exception)
            {
                MessageBox.Show("No Row Selected!");
            }
        }

       
    }
}
