Public Class Calculator

    Public dblItemAmount1, dblItemAmount2 As Double, dblItemAmount3, dblItemAmount4, dblItemAmount5
    Public dblItemTotal1, dblItemTotal2 As Double, dblItemTotal3, dblItemTotal4, dblItemTotal5 As Double
    Public dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5 As Double
    Public dblItemAmountTotal, dblTaxTotal, dblCouponsTotal As Double
    Public dblFinalItemTotal, dblFinalOneThirdTotal, dblFinalOneFourthTotal As Double
    Public dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5 As Double
    Public dblOneThird1, dblOneThird2 As Double, dblOneThird3, dblOneThird4, dblOneThird5 As Double
    Public dblOneFourth1, dblOneFourth2 As Double, dblOneFourth3, dblOneFourth4, dblOneFourth5 As Double
    Public dblCurrentTax As Double

    Private Sub Calculator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCurrentTax.Text = "Pick a City"
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged

        If ComboBox1.SelectedItem.ToString() = "Clovis" Then

            dblCurrentTax = 0.0823
            lblCurrentTax.Text = dblCurrentTax.ToString("p")
            txtItemDisc1.Enabled = True
            txtItemAmount1.Enabled = True
        End If
        If ComboBox1.SelectedItem.ToString() = "Fresno" Then

            dblCurrentTax = 0.0823
            lblCurrentTax.Text = dblCurrentTax.ToString("p")
            txtItemDisc1.Enabled = True
            txtItemAmount1.Enabled = True
        End If
        If ComboBox1.SelectedItem.ToString() = "Selma" Then

            dblCurrentTax = 0.0873
            lblCurrentTax.Text = dblCurrentTax.ToString("p")
            txtItemDisc1.Enabled = True
            txtItemAmount1.Enabled = True
        End If
        If ComboBox1.SelectedItem.ToString() = "Tulare" Then

            dblCurrentTax = 0.085
            lblCurrentTax.Text = dblCurrentTax.ToString("p")
            txtItemDisc1.Enabled = True
            txtItemAmount1.Enabled = True
        End If
    End Sub

    Private Sub txtItem1_TextChanged(sender As Object, e As EventArgs) Handles txtItemAmount1.TextChanged

        Dim strItemAmount1 As String
        Dim blnValidItemEntry1 As Boolean


        strItemAmount1 = CStr(txtItemAmount1.Text)

        blnValidItemEntry1 = fncValidEntry(strItemAmount1)

        If (blnValidItemEntry1 = True) Then

            dblItemAmount1 = CDbl(txtItemAmount1.Text)
            dblItemTotal1 = dblItemAmount1

            dblOneThird1 = fncFractional(dblItemTotal1, 3)
            dblOneFourth1 = fncFractional(dblItemTotal1, 4)
            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)

            fncItemDisplay(dblItemTotal1.ToString("c"), dblOneThird1.ToString("c"), dblOneFourth1.ToString("c"), _
                           lblItemTotal1.Text, lblOneThird1.Text, lblOneFourth1.Text)

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)

            fncTotalDisplay(dblItemAmountTotal.ToString("c"), dblFinalItemTotal.ToString("c"), dblFinalOneThirdTotal.ToString("c"), _
                            dblFinalOneFourthTotal.ToString("c"), lblItemsAmountTotal.Text, lblRegTotal.Text, lblOneThirdTotal.Text, _
                            lblOneFourthTotal.Text)

            cbxItem1.Enabled = True
            txtCoupon1.Enabled = True
            btnReset1.Enabled = True
            txtItemDisc2.Enabled = True
            txtItemAmount2.Enabled = True
            btnResetAll.Enabled = True

        Else
            txtItemAmount1.Text = ""

        End If

    End Sub

    Private Sub cbxItem1_CheckedChanged(sender As Object, e As EventArgs) Handles cbxItem1.CheckedChanged

        If (cbxItem1.Checked = True) Then

            dblItemTax1 = fncItemTax(dblItemAmount1, dblCurrentTax)
            lblItemTax1.Text = dblItemTax1.ToString("c")

            dblItemTotal1 = fncItemTotal(dblItemAmount1, dblItemTax1)
            lblItemTotal1.Text = dblItemTotal1.ToString("c")

            dblOneThird1 = fncFractional(dblItemTotal1, 3)
            lblOneThird1.Text = dblOneThird1.ToString("c")

            dblOneFourth1 = fncFractional(dblItemTotal1, 4)
            lblOneFourth1.Text = dblOneFourth1.ToString("c")

            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
            lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

            dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
            lblTaxTotal.Text = dblTaxTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

        End If

    End Sub

    Private Sub txtCoupon1_TextChanged(sender As Object, e As EventArgs) Handles txtCoupon1.TextChanged

        Dim strCouponAmount1 As String
        Dim blnValidCouponEntry1 As Boolean


        strCouponAmount1 = CStr(txtCoupon1.Text)

        blnValidCouponEntry1 = fncValidEntry(strCouponAmount1)

        If (blnValidCouponEntry1 = True) Then

            dblCoupon1 = CDbl(txtCoupon1.Text)

            dblItemTotal1 = fncItemTotal(dblItemAmount1, dblItemTax1)

            dblItemTotal1 = fncCoupon(dblCoupon1, dblItemTotal1)
            lblItemTotal1.Text = dblItemTotal1.ToString("c")

            dblOneThird1 = fncFractional(dblItemTotal1, 3)
            lblOneThird1.Text = dblOneThird1.ToString("c")

            dblOneFourth1 = fncFractional(dblItemTotal1, 4)
            lblOneFourth1.Text = dblOneFourth1.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

            txtItemAmount1.Update()

        Else
            txtCoupon1.Text = ""

        End If

    End Sub

    Private Sub txtItem2_TextChanged(sender As Object, e As EventArgs) Handles txtItemAmount2.TextChanged

        Dim strItemAmount2 As String
        Dim blnValidItemEntry2 As Boolean

        strItemAmount2 = CStr(txtItemAmount2.Text)

        blnValidItemEntry2 = fncValidEntry(strItemAmount2)

        If (blnValidItemEntry2 = True) Then

            dblItemAmount2 = CDbl(txtItemAmount2.Text)

            dblItemTax2 = fncItemTax(dblItemAmount2, dblCurrentTax)
            lblItemTax2.Text = dblItemTax2.ToString("C")

            dblItemTotal2 = fncItemTotal(dblItemAmount2, dblItemTax2)
            lblItemTotal2.Text = dblItemTotal2.ToString("c")

            dblOneThird2 = fncFractional(dblItemTotal2, 3)
            lblOneThird2.Text = dblOneThird2.ToString("c")

            dblOneFourth2 = fncFractional(dblItemTotal2, 4)
            lblOneFourth2.Text = dblOneFourth2.ToString("c")

            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
            lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

            dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
            lblTaxTotal.Text = dblTaxTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

            txtCoupon2.Enabled = True
            btnReset2.Enabled = True
            txtItemDisc3.Enabled = True
            txtItemAmount3.Enabled = True

        Else
            txtItemAmount2.Text = ""

        End If


    End Sub

    Private Sub txtCoupon2_TextChanged(sender As Object, e As EventArgs) Handles txtCoupon2.TextChanged

        Dim strCouponAmount2 As String
        Dim blnValidCouponEntry2 As Boolean


        strCouponAmount2 = CStr(txtCoupon2.Text)

        blnValidCouponEntry2 = fncValidEntry(strCouponAmount2)

        If (blnValidCouponEntry2 = True) Then

            dblCoupon2 = CDbl(txtCoupon2.Text)

            dblItemTotal2 = fncItemTotal(dblItemAmount2, dblItemTax2)

            dblItemTotal2 = fncCoupon(dblCoupon2, dblItemTotal2)
            lblItemTotal2.Text = dblItemTotal2.ToString("c")

            dblOneThird2 = fncFractional(dblItemTotal2, 3)
            lblOneThird2.Text = dblOneThird2.ToString("c")

            dblOneFourth2 = fncFractional(dblItemTotal2, 4)
            lblOneFourth2.Text = dblOneFourth2.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

        Else
            txtCoupon2.Text = ""

        End If

    End Sub

    Private Sub txtItem3_TextChanged(sender As Object, e As EventArgs) Handles txtItemAmount3.TextChanged

        Dim strItemAmount3 As String
        Dim blnValidItemEntry3 As Boolean


        strItemAmount3 = CStr(txtItemAmount3.Text)

        blnValidItemEntry3 = fncValidEntry(strItemAmount3)

        If (blnValidItemEntry3 = True) Then

            dblItemAmount3 = CDbl(txtItemAmount3.Text)

            dblItemTax3 = fncItemTax(dblItemAmount3, dblCurrentTax)
            lblItemTax3.Text = dblItemTax3.ToString("c")

            dblItemTotal3 = fncItemTotal(dblItemAmount3, dblItemTax3)
            lblItemTotal3.Text = dblItemTotal3.ToString("c")

            dblOneThird3 = fncFractional(dblItemTotal3, 3)
            lblOneThird3.Text = dblOneThird3.ToString("c")

            dblOneFourth3 = fncFractional(dblItemTotal3, 4)
            lblOneFourth3.Text = dblOneFourth3.ToString("c")

            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
            lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

            dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
            lblTaxTotal.Text = dblTaxTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblItemAmount5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

            txtCoupon3.Enabled = True
            btnReset3.Enabled = True
            txtItemDisc4.Enabled = True
            txtItemAmount4.Enabled = True

        Else
            txtItemAmount3.Text = ""

        End If


    End Sub

    Private Sub txtCoupon3_TextChanged(sender As Object, e As EventArgs) Handles txtCoupon3.TextChanged

        Dim strCouponAmount3 As String
        Dim blnValidCouponEntry3 As Boolean


        strCouponAmount3 = CStr(txtCoupon3.Text)

        blnValidCouponEntry3 = fncValidEntry(strCouponAmount3)

        If (blnValidCouponEntry3 = True) Then

            dblCoupon3 = CDbl(txtCoupon3.Text)

            dblItemTotal3 = fncItemTotal(dblItemAmount3, dblItemTax3)

            dblItemTotal3 = fncCoupon(dblCoupon3, dblItemTotal3)
            lblItemTotal3.Text = dblItemTotal3.ToString("c")

            dblOneThird3 = fncFractional(dblItemTotal3, 3)
            lblOneThird3.Text = dblOneThird3.ToString("c")

            dblOneFourth3 = fncFractional(dblItemTotal3, 4)
            lblOneFourth3.Text = dblOneFourth3.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")
        Else
            txtCoupon3.Text = ""

        End If
    End Sub

    Private Sub txtItem4_TextChanged(sender As Object, e As EventArgs) Handles txtItemAmount4.TextChanged

        Dim strItemAmount4 As String
        Dim blnValidItemEntry4 As Boolean


        strItemAmount4 = CStr(txtItemAmount4.Text)

        blnValidItemEntry4 = fncValidEntry(strItemAmount4)

        If (blnValidItemEntry4 = True) Then

            dblItemAmount4 = CDbl(txtItemAmount4.Text)

            dblItemTax4 = fncItemTax(dblItemAmount4, dblCurrentTax)
            lblItemTax4.Text = dblItemTax4.ToString("c")

            dblItemTotal4 = fncItemTotal(dblItemAmount4, dblItemTax4)
            lblItemTotal4.Text = dblItemTotal4.ToString("c")

            dblOneThird4 = fncFractional(dblItemTotal4, 3)
            lblOneThird4.Text = dblOneThird4.ToString("c")

            dblOneFourth4 = fncFractional(dblItemTotal4, 4)
            lblOneFourth4.Text = dblOneFourth4.ToString("c")

            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
            lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

            dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
            lblTaxTotal.Text = dblTaxTotal.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblItemAmount5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

            txtCoupon4.Enabled = True
            btnReset4.Enabled = True
            txtItemDisc5.Enabled = True
            txtItemAmount5.Enabled = True

        Else
            txtItemAmount4.Text = ""

        End If

    End Sub

    Private Sub txtCoupon4_TextChanged(sender As Object, e As EventArgs) Handles txtCoupon4.TextChanged

        Dim strCouponAmount4 As String
        Dim blnValidCouponEntry4 As Boolean


        strCouponAmount4 = CStr(txtCoupon4.Text)

        blnValidCouponEntry4 = fncValidEntry(strCouponAmount4)

        If (blnValidCouponEntry4 = True) Then

            dblCoupon4 = CDbl(txtCoupon4.Text)

            dblItemTotal4 = fncItemTotal(dblItemAmount4, dblItemTax4)

            dblItemTotal4 = fncCoupon(dblCoupon4, dblItemTotal4)
            lblItemTotal4.Text = dblItemTotal4.ToString("c")

            dblOneThird4 = fncFractional(dblItemTotal4, 3)
            lblOneThird4.Text = dblOneThird4.ToString("c")

            dblOneFourth4 = fncFractional(dblItemTotal4, 4)
            lblOneFourth4.Text = dblOneFourth4.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

        Else
            txtCoupon4.Text = ""

        End If
    End Sub

    Private Sub txtItem5_TextChanged(sender As Object, e As EventArgs) Handles txtItemAmount5.TextChanged

        Dim strItemAmount5 As String
        Dim blnValidItemEntry5 As Boolean


        strItemAmount5 = CStr(txtItemAmount5.Text)

        blnValidItemEntry5 = fncValidEntry(strItemAmount5)

        If (blnValidItemEntry5 = True) Then

            dblItemAmount5 = CDbl(txtItemAmount5.Text)

            dblItemTax5 = fncItemTax(dblItemAmount5, dblCurrentTax)
            lblItemTax5.Text = dblItemTax5.ToString("c")

            dblItemTotal5 = fncItemTotal(dblItemAmount5, dblItemTax5)
            lblItemTotal5.Text = dblItemTotal5.ToString("c")

            dblOneThird5 = fncFractional(dblItemTotal5, 3)
            lblOneThird5.Text = dblOneThird5.ToString("c")

            dblOneFourth5 = fncFractional(dblItemTotal5, 4)
            lblOneFourth5.Text = dblOneFourth5.ToString("c")

            dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
            lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

            dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
            lblTaxTotal.Text = dblTaxTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

            txtCoupon5.Enabled = True
            btnReset5.Enabled = True

        Else
            txtItemAmount5.Text = ""

        End If

    End Sub

    Private Sub txtCoupon5_TextChanged(sender As Object, e As EventArgs) Handles txtCoupon5.TextChanged

        Dim strCouponAmount5 As String
        Dim blnValidCouponEntry5 As Boolean


        strCouponAmount5 = CStr(txtCoupon5.Text)

        blnValidCouponEntry5 = fncValidEntry(strCouponAmount5)

        If (blnValidCouponEntry5 = True) Then

            dblCoupon5 = CDbl(txtCoupon5.Text)

            dblItemTotal5 = fncItemTotal(dblItemAmount5, dblItemTax5)

            dblItemTotal5 = fncCoupon(dblCoupon5, dblItemTotal5)
            lblItemTotal5.Text = dblItemTotal5.ToString("c")

            dblOneThird5 = fncFractional(dblItemTotal5, 3)
            lblOneThird5.Text = dblOneThird5.ToString("c")

            dblOneFourth5 = fncFractional(dblItemTotal5, 4)
            lblOneFourth5.Text = dblOneFourth5.ToString("c")

            dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
            lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

            dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
            lblRegTotal.Text = dblFinalItemTotal.ToString("c")

            dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
            lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

            dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
            lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

        Else
            txtCoupon5.Text = ""

        End If

    End Sub

    Private Sub btnReset1_Click(sender As Object, e As EventArgs) Handles btnReset1.Click

        fncResetDisplays(txtItemDisc1.Text, txtItemAmount1.Text, lblItemTax1.Text, txtCoupon1.Text, lblItemTotal1.Text, _
                            lblOneThird1.Text, lblOneFourth1.Text)

        fncResetValues(dblItemAmount1, dblItemTax1, dblCoupon1, dblItemTotal1, dblOneThird1, dblOneFourth1)

        txtCoupon1.Enabled = False
        btnReset1.Enabled = False

        dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
        lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

        dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
        lblTaxTotal.Text = dblTaxTotal.ToString("c")

        dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
        lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

        dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
        lblRegTotal.Text = dblFinalItemTotal.ToString("c")

        dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
        lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

        dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
        lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

    End Sub

    Private Sub btnReset2_Click(sender As Object, e As EventArgs) Handles btnReset2.Click

        fncResetDisplays(txtItemDisc2.Text, txtItemAmount2.Text, lblItemTax2.Text, txtCoupon2.Text, lblItemTotal2.Text, _
                            lblOneThird2.Text, lblOneFourth2.Text)

        fncResetValues(dblItemAmount2, dblItemTax2, dblCoupon2, dblItemTotal2, dblOneThird2, dblOneFourth2)

        txtCoupon2.Enabled = False
        btnReset2.Enabled = False

        dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
        lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

        dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
        lblTaxTotal.Text = dblTaxTotal.ToString("c")

        dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
        lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

        dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
        lblRegTotal.Text = dblFinalItemTotal.ToString("c")

        dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
        lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

        dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
        lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

    End Sub

    Private Sub btnReset3_Click(sender As Object, e As EventArgs) Handles btnReset3.Click

        fncResetDisplays(txtItemDisc3.Text, txtItemAmount3.Text, lblItemTax3.Text, txtCoupon3.Text, lblItemTotal3.Text, _
                            lblOneThird3.Text, lblOneFourth3.Text)

        fncResetValues(dblItemAmount3, dblItemTax3, dblCoupon3, dblItemTotal3, dblOneThird3, dblOneFourth3)

        txtCoupon3.Enabled = False
        btnReset3.Enabled = False
        dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
        lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

        dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
        lblTaxTotal.Text = dblTaxTotal.ToString("c")

        dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
        lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

        dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
        lblRegTotal.Text = dblFinalItemTotal.ToString("c")

        dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
        lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

        dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
        lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

    End Sub

    Private Sub btnReset4_Click(sender As Object, e As EventArgs) Handles btnReset4.Click

        fncResetDisplays(txtItemDisc4.Text, txtItemAmount4.Text, lblItemTax4.Text, txtCoupon4.Text, lblItemTotal4.Text, _
                                   lblOneThird4.Text, lblOneFourth4.Text)

        fncResetValues(dblItemAmount4, dblItemTax4, dblCoupon4, dblItemTotal4, dblOneThird4, dblOneFourth4)

        txtCoupon4.Enabled = False
        btnReset4.Enabled = False

        dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
        lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

        dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
        lblTaxTotal.Text = dblTaxTotal.ToString("c")

        dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
        lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

        dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
        lblRegTotal.Text = dblFinalItemTotal.ToString("c")

        dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
        lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

        dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
        lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

    End Sub

    Private Sub btnReset5_Click(sender As Object, e As EventArgs) Handles btnReset5.Click

        fncResetDisplays(txtItemDisc5.Text, txtItemAmount5.Text, lblItemTax5.Text, txtCoupon5.Text, lblItemTotal5.Text, _
                           lblOneThird5.Text, lblOneFourth5.Text)

        fncResetValues(dblItemAmount5, dblItemTax5, dblCoupon5, dblItemTotal5, dblOneThird5, dblOneFourth5)

        txtCoupon5.Enabled = False
        btnReset5.Enabled = False

        dblItemAmountTotal = fncFinalTotal(dblItemAmount1, dblItemAmount2, dblItemAmount3, dblItemAmount4, dblItemAmount5)
        lblItemsAmountTotal.Text = dblItemAmountTotal.ToString("c")

        dblTaxTotal = fncFinalTotal(dblItemTax1, dblItemTax2, dblItemTax3, dblItemTax4, dblItemTax5)
        lblTaxTotal.Text = dblTaxTotal.ToString("c")

        dblCouponsTotal = fncFinalTotal(dblCoupon1, dblCoupon2, dblCoupon3, dblCoupon4, dblCoupon5)
        lblCouponsTotal.Text = dblCouponsTotal.ToString("c")

        dblFinalItemTotal = fncFinalTotal(dblItemTotal1, dblItemTotal2, dblItemTotal3, dblItemTotal4, dblItemTotal5)
        lblRegTotal.Text = dblFinalItemTotal.ToString("c")

        dblFinalOneThirdTotal = fncFinalTotal(dblOneThird1, dblOneThird2, dblOneThird3, dblOneThird4, dblOneThird5)
        lblOneThirdTotal.Text = dblFinalOneThirdTotal.ToString("c")

        dblFinalOneFourthTotal = fncFinalTotal(dblOneFourth1, dblOneFourth2, dblOneFourth3, dblOneFourth4, dblOneFourth5)
        lblOneFourthTotal.Text = dblFinalOneFourthTotal.ToString("c")

    End Sub

    Private Sub btnResetAll_Click(sender As Object, e As EventArgs) Handles btnResetAll.Click

        lblCurrentTax.Text = "Pick a City"
        dblCurrentTax = 0

        btnReset1.PerformClick()
        btnReset2.PerformClick()
        btnReset3.PerformClick()
        btnReset4.PerformClick()
        btnReset5.PerformClick()

        txtItemAmount1.Enabled = False
        txtItemDisc1.Enabled = False
        txtItemAmount2.Enabled = False
        txtItemDisc2.Enabled = False
        txtItemAmount3.Enabled = False
        txtItemDisc3.Enabled = False
        txtItemAmount4.Enabled = False
        txtItemDisc4.Enabled = False
        txtItemAmount5.Enabled = False
        txtItemDisc5.Enabled = False

        fncResetDisplays(lblItemsAmountTotal.Text, lblTaxTotal.Text, lblCouponsTotal.Text, lblRegTotal.Text, _
                         lblOneThirdTotal.Text, lblOneFourthTotal.Text, lblOneFourthTotal.Text)

        fncResetValues(dblItemAmountTotal, dblTaxTotal, dblCouponsTotal, dblFinalItemTotal, dblFinalOneThirdTotal, _
                       dblFinalOneFourthTotal)

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

End Class
