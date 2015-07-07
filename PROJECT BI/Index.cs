using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PROJECT_BI
{
    using System.Net;
    using System.IO;
    using Excel = Microsoft.Office.Interop.Excel; 
    class Index : Equity
    {
        
        public Index()
        {

        }
        
        public Index(string c ,string i,string n,string z)
        {
            
            country = c;
            id = i;
            name = n;
            zone = z;
            urlTemplate = "http://real-chart.finance.yahoo.com/table.csv?s=" + id;
           
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
            Index e = o as Index;
            return (e.id == id && e.name == name);
        }
        public override int GetHashCode()
        {
            Console.WriteLine("Calling HashCode");
            return id.GetHashCode();
        }
        public double GetPerformanceAt(int index)
        {
            return perf.ElementAt(index);
        }
        public double ValueAt(int i)
        {
            return datesAndValuation.Values.ElementAt(i);
        }
        public override string ToString()
        {
            String s = "STOCK : " + '\n';
            s +=  " country = " + country + " id = " + id +  "name = " + name + " zone =" + zone + '\n';
            s += '\n' + '\n' + "PERFORMANCES : " + '\n' + "3 M : " + perf3M + " 6M : " + perf6M + " 1YEAR = " + perf1Y + " 3YEAR= " + perf3Y + " 5Y : " + perf5Y + " All Time : " + perfAT + '\n';
            s += "VOLAT : " + volatility + " anuual volat : " + annualVolatility + '\n';
            Console.WriteLine(s);
           // Console.ReadKey();
            String s2 = "";
            s2 += "PERFS : " + '\n';
            for (int i = 0; i < perf.Count(); i++)
                s2 += perf.ElementAt(i) + " ! ";
            s2 += "MOVING AVG 2Y: " + '\n';
            for (int i = 0; i < avg2Y.Count(); i++)
                s2 += avg2Y.Values.ElementAt(i) + " ! ";
            Console.WriteLine(s2);
            return s + '\n' + s2;
        }

    }
}
