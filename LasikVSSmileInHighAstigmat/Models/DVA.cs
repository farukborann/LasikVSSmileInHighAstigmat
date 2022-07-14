using System;
using System.Collections.Generic;
using System.Linq;

namespace LasikVSSmileInHighAstigmat.Models
{
    public class DVA
    {
        public double Decimal { get; set; }
        public string Snellen { get; set; }
        public double LogMar => Math.Log10(1 / Decimal);

        public DVA(double _decimal, string snellen)
        {
            Decimal = _decimal;
            Snellen = snellen;
        }

        static DVA DVA_2_0 = new(2.0, "20/10");
        static DVA DVA_1_5 = new(1.5, "20/12,5");
        static DVA DVA_1_2 = new(1.2, "20/16");
        static DVA DVA_1_0 = new(1.0, "20/20");
        static DVA DVA_0_9 = new(0.9, "20/20");
        static DVA DVA_0_8 = new(0.8, "20/25");
        static DVA DVA_0_7 = new(0.7, "20/25");
        static DVA DVA_0_6 = new(0.6, "20/32");
        static DVA DVA_0_5 = new(0.5, "20/40");
        static DVA DVA_0_4 = new(0.4, "20/50");
        static DVA DVA_0_3 = new(0.3, "20/63");
        static DVA DVA_0_2 = new(0.2, "20/100"); /* 80 *********************/
        static DVA DVA_0_1 = new(0.1, "20/200");
        static DVA DVA_0_05 = new(0.05, "20/400");

        static List<DVA> DVAModels = new() { DVA_0_05, DVA_0_1, DVA_0_2, DVA_0_3, DVA_0_4, DVA_0_5, DVA_0_6, DVA_0_7, DVA_0_8, DVA_0_9, DVA_1_0, DVA_1_2, DVA_1_5, DVA_2_0 };

        public static DVA FindDVA_Decimal(double _decimal)
        {
            return DVAModels.First(x => x.Decimal.Equals(_decimal));
        }
        
        public static DVA FindDVA_Snellen(string snellen)
        {
            return DVAModels.First(x => x.Snellen.Equals(snellen));
        }
        
        public static DVA FindDVA_LogMar(double logMar)
        {
            return DVAModels.First(x => x.LogMar.Equals(logMar));
        }
    }
}
