using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PROJECT_BI
{
    using System.Net;
    using System.IO;
    using Excel = Microsoft.Office.Interop.Excel; 
    class Company : Equity
    {

        
        private String benchId;
        private String benchmark;
        private String industry;
        private Index index;
        private double alpha3M, alpha6M, alpha1Y, alpha3Y, alpha5Y, alpha;
        private double beta3M, beta6M, beta1Y, beta3Y, beta5Y, beta;
        private List<double> relativePerf;
        private double trackingError, informationRatio;
        private String sector;
        //public Company() { }
        public Company(string bi,string b,string c,string i,string ind,string n,string s,string z) : base(bi,b,c,i,ind,n,s,z)
        {
            benchId = bi;
            benchmark = b;
            industry = ind;
            sector = s;
            relativePerf = new List<double>();
            index = new Index();
           /* country = c;
            id = i;
            name = n;
            
            zone = z;
            urlTemplate = "http://real-chart.finance.yahoo.com/table.csv?s=" + id;
           */
        }
        public Company(string bi, string b, string c, string i, string ind, string n, string s, string z, Index bench)
            : base(bi, b, c, i, ind, n, s, z)
        {
            benchId = bi;
            benchmark = b;
            industry = ind;
            sector = s;
            index = bench;
            relativePerf = new List<double>();
            /* country = c;
             id = i;
             name = n;
            
             zone = z;
             urlTemplate = "http://real-chart.finance.yahoo.com/table.csv?s=" + id;
            */
        }
        private string UrlBuilder(DateTime start, DateTime end)
        {
            /* exemples:
               * 6 janvier 2012 / 28 juin 2015, journalier
               *  http://real-chart.finance.yahoo.com/table.csv?s=EN.PA&a=00&b=6&c=2012&d=05&e=28&f=2015&g=d&ignore=.csv
               * 
               * 
              //http://real-chart.finance.yahoo.com/table.csv?s=EN.PA&d=5&e=22&f=2015&g=d&a=0&b=3&c=2000&ignore=.csv
               * 
               * 
               * */
            String url = urlTemplate;
            url += "&a=" + (start.Month - 1) + "&b=" + start.Day + "&c=" + start.Year + "&d=" + (end.Month - 1) + "&e=" + end.Day + "&f=" + end.Year + "g=d&ignore=.csv";
            Console.WriteLine(url);
            return url;
        }
        
        public virtual bool Equals(Object o)
        {
            Console.WriteLine("Calling equals");
            Company e = o as Company;
            return (e.id == id && e.name == name);
        }
        public override int GetHashCode()
        {
            Console.WriteLine("Calling HashCode");
            return benchId.GetHashCode();
        }
        public override void ComputePerformances()
        {
            Console.WriteLine("computeperformance()");
            base.ComputePerformances();
            relativePerf.Add(0);
            int nb = perf.Count();
            for (int i = 1; i < nb; i++ )
                ComputeRelativePerf(i);

        }
        private void ComputeRelativePerf(int i)
        {
            KeyValuePair<DateTime,double> p =  datesAndValuation.ElementAt(i);
            DateTime currentDate = p.Key;
            int benchIdx = index.findDateIndex(currentDate);
            double perfC = perf.ElementAt(i), perfB = index.GetPerformanceAt(benchIdx);
            relativePerf.Add((double)(perfC - perfB));

        }
        private void ComputeBeta()
        {

            Console.WriteLine("calcule beta for : " + name);
            List<double> benchPerfs = index.GetPerformances();
            List<DateTime> datesBench=index.GetDates();
            List<DateTime> dates = GetDates();

            beta3M = BetaPeriod(25, benchPerfs, datesBench, dates);
            beta6M = BetaPeriod(50, benchPerfs, datesBench, dates);
            beta1Y = BetaPeriod(100, benchPerfs, datesBench, dates);
            beta3Y = BetaPeriod(300, benchPerfs, datesBench, dates);
            beta5Y = BetaPeriod(500, benchPerfs, datesBench, dates);
            beta = BetaPeriod(0, benchPerfs, datesBench, dates);
        }
        private double BetaPeriod(int period, List<double> benchPerfs,List<DateTime> benchDates, List<DateTime>dates)
        {
            Console.WriteLine("calcul beta " + period + " for " + name);
            int offset=0;
            double sum=0;
            int nbValues = 0, nbValuesB = 0;
            double Ex, Ey,varY=0;
            double B = VALUE_IF_ERROR;
            // récupérer la dernière date de l'intervalle qui nous intéresse
            switch (period)
            {
                case 25: offset = -3;break;
                case 50: offset = -6; break;
                case 100: offset = -12;break;
                case 300 :offset = -36;break;
                case 500 : offset= -60;break;
            }
            DateTime lastOne =(offset ==0)? dates.First(): dates.Last().AddMonths(offset);
            DateTime currentB = benchDates.Last(), current = dates.Last();
            // pour chaque date de cet intervalle, calculer la moyenne pour le benchmark et l'action
            // ESPERANCE 
            Console.WriteLine("last : " + lastOne + " dates.Last() : " + dates.Last());
            if (lastOne >= dates.First())
            {
                int i = dates.Count() - 1;
                do
                {
                    sum += this.perf.ElementAt(i);
                    nbValues++;
                    i--;
                    current = dates.ElementAt(i);
                }
                while (current >= lastOne && i >0);
                Ex = sum / nbValues;
                sum = 0;
                // ESPEREANCE INDEX
                i = benchDates.Count()-1;
                do
                {
                    sum += benchPerfs.ElementAt(i);
                    nbValuesB++;
                    i--;
                    currentB = benchDates.ElementAt(i);
                }
                while (currentB >= lastOne && i > 0);

                Ey = sum / nbValuesB;

                // INDEX VARIANCE
                int nb = Math.Min(nbValues, nbValuesB);
                int at = benchPerfs.Count()-1;
                for (i = 0; i < nbValuesB; i++)
                {
                    varY += Math.Pow(benchPerfs.ElementAt(at) - Ey, 2);
                    at--;
                }
                varY /= nbValuesB;
                double cov = 0;

                // COVARIANCE
               // nbValues--;
               // nbValuesB--;
                nbValues = dates.Count - 1;
                nbValuesB = benchPerfs.Count - 1;
                for (i = 0; i < nb; i++)
                {
                    cov += ((benchPerfs.ElementAt(nbValuesB) - Ey) * (perf.ElementAt(nbValues) - Ex));
                    nbValues--;
                    nbValuesB--;
                }
                cov /= nb;
                B= cov / varY;
                
                Console.Write("Beta " + period + " for " + name + " = " + B);
            }

            return B;
        }

       
        private void ComputeAlpha()
        {
            alpha3M = AlphaPeriod(25);
            alpha6M = AlphaPeriod(50);
            alpha1Y = AlphaPeriod(100);
            alpha3Y = AlphaPeriod(300);
            alpha5Y = AlphaPeriod(500);
            alpha = AlphaPeriod(0); // 0 is for all time
        }
        //Alpha = Perf annualisée fond - B*perf anualisée (indice)
        private double AlphaPeriod(int period)
        {
            double A = VALUE_IF_ERROR, P=0,Pi=VALUE_IF_ERROR;
            double B=0;
            switch(period)
            {
                case 0: 
                    B = beta; 
                    P =perfAT;
                    Pi = index.GetATPerf();
                    break;
                case 25: B  =beta3M;
                    P =perf3M;
                    Pi = index.Get3MPerf();
                    break;
                case 50: B = beta6M; 
                    P =perf6M;
                    Pi = index.Get6MPerf();
                    break;
                case 100: B  =beta1Y;
                    P =perf1Y;
                    Pi = index.Get1YPerf();
                    break;
                case 300 : B  =beta3Y;
                    P =perf3Y;
                    Pi = index.Get3YPerf();
                    break;
                case 500 :  B  =beta5Y;
                    P =perf5Y;
                    Pi = index.Get5YPerf();
                    break;
            }
            if (B != VALUE_IF_ERROR && Pi != VALUE_IF_ERROR)
            {
                A = P - B * Pi;
            }
            return A;

        }
        public override void ComputeIndicators()
        {
           base.ComputeIndicators();
           Console.WriteLine("j'ai calculé les indicateurs " + name);

           ComputeBeta();
           Console.WriteLine("j'ai calculté le Beeta" + name);
           ComputeAlpha();
           Console.WriteLine("j'ai calculté le alpha" + name);

           ComputeTrackingError();
           ComputeInformationRatio();
        }

        private void ComputeTrackingError()
        {
            int nb = relativePerf.Count(), i;
            double sum = 0.0,avg = relativePerf.Average();
            for (i = 0; i < nb; i++)
            {
                sum += Math.Pow(relativePerf.ElementAt(i) - avg,2);
            }
            trackingError= sum /= (nb-1);

        }
        private void ComputeInformationRatio()
        {
            informationRatio = (perfAT - index.GetATPerf()) / trackingError;
        }
        public override String ToString()
        {
            String s = "STOCK : " + '\n';
            s += "benchID= " + benchId + " benchmark = " + benchmark + " country = " + country + " id = " + id + "industry = " + industry + "name = " + name + " sector = " + sector + " zone =" + zone + '\n';
            s += '\n' + '\n' + "PERFORMANCES : " + '\n' + "3 M : " + perf3M + " 6M : " + perf6M + " 1YEAR = " + perf1Y + " 3YEAR= " + perf3Y + " 5Y : " + perf5Y + " All Time : " + perfAT + '\n';
            s+=" ALPHA : " + '\n' + "3M : " + alpha3M + " 6M : " + alpha6M + " 1Y : " + alpha1Y + " 3Y : " + alpha3Y + " 5Y : " + alpha5Y + '\n';
            s+= "BETA : " + '\n' + "3M : "  + beta3M + " 6M : " + beta6M + " 1Y : " + beta1Y + " 3Y : " + beta3Y + " 5Y : " + beta5Y + '\n';
            s += "VOLAT : " + volatility + " anuual volat : " + annualVolatility + '\n';
            Console.WriteLine(s);
           // Console.ReadKey();
            String s2 = "";
          /*  s2+=  "PERFS : " + '\n';
            for(int i = 0 ; i < perf.Count(); i++)
                s2+= perf.ElementAt(i) + " ! ";
           * */
            s2 += "RELATIVE PERFS : " + '\n';
            for(int i = 0 ; i < relativePerf.Count(); i++)
                s2+= relativePerf.ElementAt(i) + " ! ";
            s2 += "MOVING AVG 2Y: " + '\n';
            for (int i = 0; i < avg2Y.Count(); i++)
                s2 += avg2Y.Values.ElementAt(i) + " ! ";
            Console.WriteLine(s2);
            return s +'\n' + s2;
        }
        public String Output()
        {
            List<DateTime>dates = datesAndValuation.Keys.ToList();
            List<double>values= datesAndValuation.Values.ToList();

            String output = "<stock benchID=\"" + benchId + '\"' + '\n' +
                         "benchmark=\"" + benchmark + '\"' + '\n' +
                         "id=\"" + id + '\"' + '\n' +
                         "industry=\"" + industry + '\"' + '\n' +
                       "name=\"" + name + '\"' + '\n' +
                      "sector=\"" + sector + '\"' + '>' + '\n' +
                    "<zone zname=\"" + zone + "\">" + '\n' +
                      "<country cname=\"" + country + "\" />" + '\n' + "</zone>"+ '\n' + 
                      "<sector sname=\"" + sector + "\" >" + '\n' +
                      "<industry iname=\"" + industry + "\" />" + '\n' +
                      "</sector>" + '\n' + "<prices>";
            for(int i = 0 ; i < datesAndValuation.Count();i++)
            {
                double mm4=-1,mm12=-1,mm24=-1;
                avg4M.TryGetValue(dates.ElementAt(i),out mm4);
                avg1Y.TryGetValue(dates.ElementAt(i),out mm12);
                avg2Y.TryGetValue(dates.ElementAt(i),out mm24);
                string d = dates.ElementAt(i).ToString().Substring(0, 10);
                output += "<obs relativePerf= \"" + Math.Round(relativePerf.ElementAt(i),2) + "\" mm24 =" + (((int)mm24 != 0) ? "\"" + Math.Round(mm24,2) : "\"") + "\" mm12 = \"" + (((int)mm12 != 0) ? Math.Round(mm12,2) + "\"" : "\"");
                output += " mm4 = \"" + (((int)mm4 != 0) ? Math.Round(mm4,2) + "\"" : "\"") + " priceBench=\"" + Math.Round(index.ValueAt(i),2) + "\" price = \"" + Math.Round(values.ElementAt(i),2) + "\"";
                output += "  date=\"" + d + "\" />";
                output +='\n';
            }
            output += "</prices>" + '\n' + "<indicators> " + '\n';
            output += "<indicator" + ((alpha3M != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha3M,2) + "\"" : "") + ((beta3M != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta3M,2) + "\"" : "") + " perf=\"" + Math.Round(perf3M,2) + "\" period =\"3M\" />";
            output += "<indicator" + ((alpha6M != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha6M, 2) + "\"" : "") + ((beta6M != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta6M, 2) + "\"" : "") + " perf=\"" + Math.Round(perf6M, 2) + "\" period =\"6M\" />";
            output += "<indicator" + ((alpha1Y != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha1Y, 2) + "\"" : "") + ((beta1Y != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta1Y, 2) + "\"" : "") + " perf=\"" + Math.Round(perf1Y, 2) + "\" period =\"1Y\" />";
            output += "<indicator" + ((alpha3Y != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha3Y, 2) + "\"" : "") + ((beta3Y != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta3Y, 2) + "\"" : "") + " perf=\"" + Math.Round(perf3Y, 2) + "\" period =\"3Y\" />";
            output += "<indicator" + ((alpha5Y != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha5Y, 2) + "\"" : "") + ((beta5Y != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta5Y, 2) + "\"" : "") + " perf=\"" + Math.Round(perf5Y, 2) + "\" period =\"5Y\" />";
            output += "<indicator" + ((alpha != VALUE_IF_ERROR) ? " alpha=\"" + Math.Round(alpha, 2) + "\"" : "") + ((beta != VALUE_IF_ERROR) ? " beta=\"" + Math.Round(beta, 2) + "\"" : "") + " perf=\"" + Math.Round(perfAT, 2) + "\" period =\"All-Time\" />";
            output += "<volat>" + '\n' +
                "<vol Vol_weekly = \"" + Math.Round(volatility,2) + "\" />" +
                "<vol Vol_annual = \"" + Math.Round(annualVolatility, 2) + "\" />" + '\n' + "</volat>" + '\n' + "</indicators>" + '\n' +
                "<trackingError te= \"" + Math.Round(trackingError,2) + "\" />" + '\n' +
                "<informationRatio ir= \"" + Math.Round(informationRatio, 2) + "\" />" + '\n' + "</stock>" + '\n';
        
            return output;
        }

    }
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    