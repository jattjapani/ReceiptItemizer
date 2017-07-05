Module Calculations

    Public Function fncValidEntry(ByVal ItemAmount As String) As Boolean

        Dim blnValidEntry As Boolean = False
        Dim dblEnteredItemAmount As Double

        If IsNumeric(ItemAmount) Then
            blnValidEntry = True
            dblEnteredItemAmount = CDbl(ItemAmount)
        End If

        If (dblEnteredItemAmount > 0) Then
            blnValidEntry = True
        End If

        Return blnValidEntry
    End Function

    Public Function fncItemTax(ByVal dblItemAmount As Double, ByVal dblCurrentTax As Double) As Double

        Return dblItemAmount * dblCurrentTax

    End Function

    Public Function fncItemTotal(ByVal dblItemAmount As Double, ByVal dblItemTax As Double) As Double

        Return dblItemAmount + dblItemTax

    End Function

    Public Function fncCoupon(ByVal dblCouponAmount As Double, ByVal dblRegAmount As Double) As Double

        Return dblRegAmount - dblCouponAmount

    End Function

    Public Function fncFractional(ByVal dblItemTotal As Double, ByVal intFactor As Integer) As Double

        Return dblItemTotal / intFactor

    End Function

    Public Function fncOneFourth(ByVal dblItemTotal As Double) As Double

        Return dblItemTotal / 4

    End Function

    Public Function fncFinalTotal(ByVal dblItemTotal1 As Double, ByVal dblItemTotal2 As Double, ByVal dblItemTotal3 As Double, _
                                  ByVal dblItemTotal4 As Double, ByVal dblItemTotal5 As Double) As Double

        Return dblItemTotal1 + dblItemTotal2 + dblItemTotal3 + dblItemTotal4 + dblItemTotal5

    End Function

    Public Function fncItemDisplay(ByVal Value1, ByVal Value2, ByVal Value3, ByRef Display1, _
                                   ByRef Display2, ByRef Display3) As Boolean

        Display1 = Value1
        Display2 = Value2
        Display3 = Value3

        Return True
    End Function

    Public Function fncTotalDisplay(ByVal Value1, ByVal Value2, ByVal Value3, ByVal Value4, _
                               ByRef Display1, ByRef Display2, ByRef Display3, ByRef Display4) As Boolean

        Display1 = Value1
        Display2 = Value2
        Display3 = Value3
        Display4 = Value4

        Return True
    End Function

    Public Function fncResetDisplays(ByRef Display1, ByRef Display2, ByRef Display3, ByRef Display4, _
                                     ByRef Display5, ByRef Display6, ByRef Display7) As Boolean

        Display1 = ""
        Display2 = ""
        Display3 = ""
        Display4 = ""
        Display5 = ""
        Display6 = ""
        Display7 = ""

        Return True
    End Function

    Public Function fncResetValues(ByRef Value1, ByRef Value2, ByRef Value3, _
                                       ByRef Value4, ByRef Value5, ByRef Value6) As Boolean

        Value1 = 0
        Value2 = 0
        Value3 = 0
        Value4 = 0
        Value5 = 0
        Value6 = 0

        Return True
    End Function

End Module
