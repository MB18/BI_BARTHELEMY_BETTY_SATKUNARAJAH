using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PROJECT_BI
{
    class Program
    {

        // A RENDRE LE 7 AU SOIR


        public static string INPUT_FILE_NAME="input.xml";
        public static string OUTPUT_FILE_NAME="output.xml";
        private List<Equity> equities;

        private string input;
        private DateTime start, end;

        public Program(string fileName)
        {
            string line;
            int counter = 0;
           // input = System.IO.File.ReadAllText(fileName);
            System.IO.StreamReader file = new System.IO.StreamReader(fileName);
            while ((line = file.ReadLine()) != null)
            {
                if (counter > 0)
                    input = input +  line + '\n';
                counter++;
            }
            equities = new List<Equity>();
            
        }

        public void parsing()
        {
           //Console.WriteLine("test : " + input);
            bool lastInfo = false;
            string buffer="";
            string startdate;
            string benchId = "", benchmark = "", country = "", id = "", industry = "", name = "", sector = "", zone = "";
            int i = 0, j=0;
            char caracter;
            //Console.WriteLine("Debut du parsing : ");
            do
            {
                caracter = input.ElementAt(i);
                //Console.Write(caracter);
                if( caracter !=' ' & caracter !='\n')
                {
                    buffer += caracter;
                    if (buffer == "<input")
                    {
                        buffer ="";
                    }
                    else if (buffer == "startdate")
                    {
                        buffer = "";
                        i += 3;
                        do
                        {
                            buffer += input.ElementAt(i);
                            i++;
                            j++;
                        }
                        while (j < 10);
                        startdate = buffer;
                        initDates(startdate);
                        i += 2;
                        buffer = "";
                    }
                    else if (buffer == "<stock")
                    {
                        buffer = "";
                    }
                    else if (buffer == "benchID=")
                    {
                        i += 2;
                        do
                        {
                            benchId += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";

                    }

                    else if (buffer == "benchmark=")
                    {
                        i += 2;
                        do
                        {
                            benchmark += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";

                    }
                    else if (buffer == "country=")
                    {
                        i += 2;
                        do
                        {
                            country += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                    }
                    else if (buffer == "id=")
                    {
                        i += 2;
                        do
                        {
                            id += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                    }
                    else if (buffer == "industry=")
                    {
                        i += 2;
                        do
                        {
                            industry += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                    }
                    else if (buffer == "name=")
                    {
                        i += 2;
                        do
                        {
                            name+= input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                    }
                    else if (buffer == "sector=")
                    {
                        i += 2;
                        do
                        {
                            sector += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                    }
                    else if (buffer == "zone=")
                    {
                        i += 2;
                        do
                        {
                            zone += input.ElementAt(i);
                            i++;
                        }
                        while (input.ElementAt(i) != '\"');
                        buffer = "";
                        lastInfo = true;
                    }
                    if (lastInfo)
                    {
                        i++;
                        do
                        {
                            caracter = input.ElementAt(i);
                            if (caracter != ' ')
                                buffer += caracter;
                            i++;
                        }
                        while (buffer != "/>");
                        CreateStock(benchId, benchmark, country, id, industry, name, sector, zone);
                        benchId = "";
                        benchmark = "";
                        country = "";
                        id = "";
                        industry = "";
                        name = "";
                        sector="";
                        zone = "";
                        lastInfo = false;
                        buffer = "";
                    }

                }
                i++;
            }
            while (i < input.Length);
        }
        void CreateStock(string benchId, string benchmark, string country, string id, string industry, string name, string sector, string zone)
        {
            Console.WriteLine("bench ID : " + benchId);
            Console.WriteLine("benchmark: " + benchmark);
            Console.WriteLine("country : " + country);
            Console.WriteLine("id : " + id);
            Console.WriteLine("industry : " + industry);
            Console.WriteLine("name: " + name);
            Console.WriteLine("sector : " + sector);
            Console.WriteLine("zone : " + zone);
            
            bool newStock = true;
            int i = 0;
            Index index = CreateBenchmark(country, benchId, benchmark, zone);
            

          
            Equity e = new Company(benchId, benchmark, country, id, industry, name, sector, zone, index);
            while (newStock && i < equities.Count())
            {
                if( equities.ElementAt(i) is Company)
                    newStock = !(equities.ElementAt(i).Equals((Object)e));
                i++;
            }
            if (newStock) 
                if (e.Download(start,end))
                    equities.Add(e);
            
        }
        Index CreateBenchmark(string country,string benchId, string benchmark,string zone)
        {
            Index index = new Index(country, benchId, benchmark, zone);
            bool newStock=true;
            int i = 0;
            while (newStock && i < equities.Count())
            {
                if( equities.ElementAt(i) is Index)
                    newStock = !(equities.ElementAt(i).Equals((Object)index));
                i++;
            }
            if (newStock)
            {
                if (index.Download(start, end))
                    equities.Add(index);
            }
            else
                i--;
            return (Index)equities.ElementAt(i);
        }
        public void initializeEquities()
        {
            int nb = equities.Count;
            if (nb >= 1)
            {
                int i;
                for (i = 0; i < nb; i++)
                {
                   equities.ElementAt(i).Initialize();
                }
            }
        }
        public void ComputePerfsAndIndicators()
        {
            Equity e;
            for (int i = 0; i < equities.Count(); i++)
            {
                e = equities.ElementAt(i);
                e.ComputePerformances();
                e.ComputeIndicators();
            }
            
        }

        
        public void CreateOutput()
        {
            String output="<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + '\n'  + "<data>" + '\n';
            Equity e;
            for (int i = 0; i < equities.Count(); i++)
            {
                e = equities.ElementAt(i);
                if (e is Company)
                    output += ((Company)e).Output();
            }
            
            output += '\n' + "</data>";
            String output2 = output.Replace(',', '.');
            System.IO.File.WriteAllText(OUTPUT_FILE_NAME, output);
        }
        
        private void initDates(string startdate)
        {
            int year, month, day;
            year = int.Parse((startdate.Substring(0, 4)));
            month = int.Parse(startdate.Substring(5, 2));
            day = int.Parse(startdate.Substring(5, 2));
            start = new DateTime(year, month, day);
            end = DateTime.Now;

            
        }

        public override string ToString()
        {
            Equity e;
            string s="";
            for (int i = 0; i < equities.Count(); i++)
            {
                e = equities.ElementAt(i);
                //s += e.ToString();
                //s += '\n' + '\n';
            }
            return s;
        }
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            
            //lecture du fichier
            Program p = new Program("input.xml");
            string text = INPUT_FILE_NAME;
            text = System.IO.File.ReadAllText("input.xml");
            
            //Console.WriteLine(text);
            p.parsing();
            p.initializeEquities();
            p.ComputePerfsAndIndicators();
            Console.WriteLine(p);
            p.CreateOutput();
         
           
        }

       
    }


}
