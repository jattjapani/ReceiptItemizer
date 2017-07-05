namespace ReceiptItemizer
{
    public class Calculate
    {
        public decimal Tax (decimal dCurrentTax, decimal dItemAmount )
        {
            return dItemAmount * dCurrentTax;   
        }
        public decimal TotalAfterTax (decimal dCurrentTax, decimal dItemAmount)
        {
            return dItemAmount + (dItemAmount * dCurrentTax);           
        }

        public decimal OneThird (decimal dItemTotal)
        {
            return dItemTotal / 3;
        }
        public decimal OneFourth (decimal dItemTotal)
        {
            return dItemTotal / 4;
        }
    }
}
