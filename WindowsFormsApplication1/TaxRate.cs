namespace ReceiptCalculator
    {
    public class TaxRate
    {
        private decimal ClovisTax = .08225m;
        private decimal FowlerTax = .08225m;
        private decimal FresnoTax = .08225m;
        private decimal SelmaTax =  .08725m;
        private decimal VisaliaTax = .08250m;
        private decimal HanfordTax = .07500m;
        private decimal TulareTax = .08500m;

        public decimal SetTax (string city )
        {
            switch (city)
            {
                case "Clovis":
                    return ClovisTax;
                case "Fowler":
                    return FowlerTax;
                case "Fresno":
                    return FresnoTax;
                case "Selma":
                    return SelmaTax;
                case "Visalia":
                    return VisaliaTax;
                case "Hanford":
                    return HanfordTax;
                case "Tulare":
                    return TulareTax;
                default:
                    return 0;
            }

        }
        
    }
}
