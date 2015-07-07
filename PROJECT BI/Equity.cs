using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PROJECT_BI
{ 
    using System.Net;
    using System.IO;
    using Excel = Microsoft.Office.Interop.Excel;
    
    public class Equity
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        public static int NB_WEEKS_IN_A_YEAR = 52;
        public static double VALUE_IF_ERROR = 666.0;

        public static int DAYS_OFFSET = 7;
        
        
        protected String country;
        protected String id;

        protected String name;

        protected String zone;
        protected String urlTemplate;
        protected Dictionary<DateTime, double> datesAndValuation;
        protected List<double> realValues;
        protected List<double> perf;


        protected double perf3M, perf6M, perf1Y, perf3Y, perf5Y,perfAT;
        protected Dictionary<DateTime,double> avg4M, avg1Y, avg2Y;
        protected double volatility, annualVolatility;
        public Equity()
        {
            datesAndValuation = new Dictionary<DateTime, double>();
            realValues = new List<double>();
            perf = new List<double>();
            perf3M = 0;
            perf6M = 0;
            perf1Y = 0;
            perf3Y = 0;
            perf5Y = 0;
            avg4M = new Dictionary<DateTime,double>();
            avg1Y = new Dictionary<DateTime, double>();
            avg2Y = new Dictionary<DateTime, double>();
        }
        public Equity(string bi,string b,string c,string i,string ind,string n,string s,string z)
        {
            country = c;
            id = i;
            name = n;
            zone = z;
            urlTemplate = "http://real-chart.finance.yahoo.com/table.csv?s=" + id;
            datesAndValuation = new Dictionary<DateTime, double>();
            realValues = new List<double>();
            perf = new List<double>();
            perf3M = 0;
            perf6M = 0;
            perf1Y = 0;
            perf3Y = 0;
            perf5Y = 0;
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
        public bool Download (DateTime start, DateTime end)
        {
            datesAndValuation.Add(start,0);
            string url = UrlBuilder(start, end);
            string fileName = name + ".csv";
            WebClient webClient = new WebClient();
            webClient.DownloadFile(url, fileName);
            return File.Exists(fileName);
            //return true;
        }
        private void ReadExcelFile()
        {
            MyBook = MyApp.Workbooks.Open(Environment.CurrentDirectory  + "\\" + name + ".csv");
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int close= 7,date=1;
            int offset = DAYS_OFFSET - 2;
            System.Array HistoricalData = (System.Array)MySheet.get_Range("A2", "G" +lastRow.ToString()).Cells.Value;
            /*Console.WriteLine("contenu du fichier : " + MyBook.Name + " : ");
            for (int i = 1; i <= HistoricalData.Length; i++)
            {
                Console.Write(HistoricalData.GetValue(i, 1));
                Console.Write(" : ");
                Console.WriteLine(HistoricalData.GetValue(i, 7));
            }*/
            DateTime theoriticalDate = datesAndValuation.ElementAt(0).Key,currentDate;
            datesAndValuation.Remove(theoriticalDate);
            double value=0,Pi=0,PiMinusOne=100;
          //  List<double> realValues = new List<double>();

            int i = lastRow-1;
            do
            {
                currentDate = DateTime.Parse(HistoricalData.GetValue(i,date).ToString());
                if (i == lastRow - 1)
                    theoriticalDate = currentDate;
                while( theoriticalDate < currentDate)
                {
                       i++;
                       currentDate = DateTime.Parse(HistoricalData.GetValue(i,date).ToString());
                }
               // value= double.Parse(HistoricalData.GetValue(i, close).ToString());
                Pi = (double.Parse(HistoricalData.GetValue(i, close).ToString()));
                PiMinusOne = ( i== lastRow-1 ? 100 : realValues.ElementAt(realValues.Count()-1));
                realValues.Add(double.Parse(HistoricalData.GetValue(i, close).ToString()));
                value = (i == lastRow - 1) ? (100) : (100 * Pi / PiMinusOne);
                
                // USING THE THEORITICAL DATE
                //datesAndValuation.Add(theoriticalDate, value);
                // USING THE REAL DATE
                datesAndValuation.Add(currentDate, value);
                i-=offset;
                theoriticalDate = theoriticalDate.AddDays(DAYS_OFFSET);
            }
            while (i >=1);

            MyBook.Close();
            Console.WriteLine( name + " " +datesAndValuation.Count() + "pair of values : ");
            
        }
        public void Initialize()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            ReadExcelFile();
        }
        public virtual bool Equals(Object o)
        {
            Console.WriteLine("Calling equals");
            Equity e = o as Equity;
            return (e.id == id && e.name == name);
        }
        public override int GetHashCode()
        {
            Console.WriteLine("Calling HashCode");
            return id.GetHashCode();
        }
        public virtual void ComputePerformances()
        {
            ComputeLocalPerf();
            ComputeYearlyPerf(0.25); // 3 month
            ComputeYearlyPerf(0.5); // 6 month
            ComputeYearlyPerf(1); // 1 YEARS
            ComputeYearlyPerf(3); // 3 YEARS
            ComputeYearlyPerf(5); // 5 YEARS
            ComputeYearlyPerf(0); //ALL TIME
        }
        public virtual void ComputeIndicators()
        {
            ComputeVolatility();
            ComputeAnnualVolatility();
            ComputeMovingAverage();
        }
        public List<double> GetPerformances()
        {
            return perf;
        }
        public void ComputeMovingAverage()
        {
            avg4M = MovingAverage(-4);
            avg1Y = MovingAverage(-12);
            avg2Y = MovingAverage(-24);
            Console.WriteLine("calcule moving average for : " + name);
        }
        private Dictionary<DateTime,double> MovingAverage(int period)
        {
            int i, nb = datesAndValuation.Count(), nbValues, currentIndex, location;
            double sum;
            DateTime current;
            List<DateTime> dates = datesAndValuation.Keys.ToList();
            List<double> values = datesAndValuation.Values.ToList();
            DateTime start = datesAndValuation.Keys.First();
            DateTime stop = start.AddMonths(period);
            Dictionary<DateTime, double> movingAvg = new Dictionary<DateTime, double>();
            for( i = 0 ; i < nb ; i++)
            {
                current = dates.ElementAt(i);
                if(start <= current.AddMonths(period)) // si assez de dates disponibles
                {
                    //calculer moyenne 
                    nbValues = 0;
                    currentIndex = i;
                    sum=0;
                    location = findDateIndex(current.AddMonths(period));
                    do
                    {
                        sum += values.ElementAt(location);
                        nbValues++;
                        location++;
                    }
                    while (location < dates.Count && dates.ElementAt(location) <= current );
                    movingAvg.Add(current, sum / nbValues);
                    
                }
            }
           
           return movingAvg;

        }
        protected void ComputeVolatility()
        {
            volatility = 0;
            double avgPerf = perf.Average();
            double elt;
            int i,nb = perf.Count();
            for (i = 0; i < nb; i++)
            {
                elt = perf.ElementAt(i);
                volatility += (elt - avgPerf) * (elt - avgPerf);
            }
            volatility /= (nb - 1);
        }
        protected void ComputeAnnualVolatility()
        {

            annualVolatility = volatility * Math.Sqrt(NB_WEEKS_IN_A_YEAR);
        }
        protected void ComputeLocalPerf()
        {
            int i = 1, nb = realValues.Count();
            perf.Add(1);
            for (i = 1; i < nb; i++)
            {
                double tmp = (realValues.ElementAt(i) - realValues.First()) / realValues.First();
                perf.Add(tmp);
                
            }

          /*
            Console.WriteLine(" Local PERFS : ");
            for (i = 0; i < nb; i++)
            {
                Console.WriteLine(perf.ElementAt(i));
            }
           */
        }

     /*   public double GetBeta()
        {
            return beta;
        }*/
        public List<DateTime> GetDates()
        {
            return datesAndValuation.Keys.ToList();
        }
        protected bool ComputeYearlyPerf(double period)
        {
            DateTime start = datesAndValuation.Last().Key, end = datesAndValuation.First().Key;
            DateTime lastDate = start;
            TimeSpan t;
            int _case = 0;
            int indexDate;
            double perf=0.0;
            string msg="";
            bool sufficientInterval = false;
            if (period == 0.25)
            {
                lastDate = start.AddMonths(-3);
                t= start - end;
                sufficientInterval = (t.Days >=90);
                msg = "3 MONTHS";
                _case = 1;
            }
            else if (period == 0.5)
            {
                lastDate = start.AddMonths(-6);
                t= start - end;
                sufficientInterval = (t.Days >= 180);
                msg = "6 MONTHS";
                _case = 2;
            }
            else if (period == 1)
            {
                lastDate = start.AddYears(-1);
                t= start - end;
                sufficientInterval = (t.Days >= 365);
                msg = "1 YEAR";
                _case = 3;
            }
            else if (period == 3)
            {
                lastDate =start.AddYears(-3);
                t= start - end;
                sufficientInterval = (t.Days >= (365*3));
                msg = "3 YEARS";
                _case = 4;
            }
            else if (period == 5)
            {
                lastDate = start.AddYears(-5);
                t= start - end;
                sufficientInterval = (t.Days >= (365*5));
                msg = "5 YEARS";
                _case = 5;
            }
            else
            {
                lastDate = end;
                t = start - end;
                msg = "ALL TIME";
                _case = 6;
                sufficientInterval = true;
            }
            if (sufficientInterval)
            {
                indexDate = findDateIndex(lastDate);
                perf = datesAndValuation.Last().Value / datesAndValuation.ElementAt(indexDate).Value;
               // t = datesAndValuation.ElementAt(indexDate).Key - start;
                t = start - datesAndValuation.ElementAt(indexDate).Key;
                perf = Math.Pow(perf, (365.0 / (double)(t.Days))) - 1;
                Console.WriteLine(name + "PERF FOR :" + msg + " = " + perf);
                switch (_case)
                {
                    case 1: perf3M = perf; break;
                    case 2: perf6M = perf; break;
                    case 3: perf1Y = perf; break;
                    case 4: perf3Y = perf; break;
                    case 5: perf5Y = perf; break;
                    case 6: perfAT = perf; break;
                }
            }
            else
            {
                Console.WriteLine(name + " : Not enough data in order to compute performance for " + msg);
                switch (_case)
                {
                    case 1: perf3M = VALUE_IF_ERROR; break;
                    case 2: perf6M = VALUE_IF_ERROR; break;
                    case 3: perf1Y = VALUE_IF_ERROR; break;
                    case 4: perf3Y = VALUE_IF_ERROR; break;
                    case 5: perf5Y = VALUE_IF_ERROR; break;
                    case 6: perfAT = VALUE_IF_ERROR; break;
                }
            }
            return sufficientInterval;
        }
        public int findDateIndex(DateTime firstAcceptableDate)
        { 
            
            int i=0;
            DateTime tmp;
            do
            {
                i++;
                tmp = datesAndValuation.ElementAt(i).Key;
            }
            while ( tmp < firstAcceptableDate);
            return i;
        }
        public double Get3MPerf()
        {
            return perf3M;
        }
        public double Get6MPerf()
        {
            return perf6M;
        }
        public double Get1YPerf()
        {
            return perf1Y;
        }
        public double Get3YPerf()
        {
            return perf3Y;
        }
        public double Get5YPerf()
        {
            return perf5Y;
        }
        public double GetATPerf()
        {
            return perfAT;
        }

    }


}
