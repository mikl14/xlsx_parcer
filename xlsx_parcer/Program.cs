using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace xlsx_parcer
{
    class Program
    {

        struct field
        {
            public string name;
            public string INN;
            public string type_nalog;
            public string nalog;

            public field(string Name, string inn, string Type_nalog, string Nalog)
            {
                name = Name;
                INN = inn;
                type_nalog = Type_nalog;
                nalog = Nalog;
            }
        }

        static void Main(string[] args)
        {

            List<field> fields = new List<field>();


            string paths = @".\files";
            string[] files = Directory.GetFiles(paths);
          
           
            try
            {
                Workbook wb = new Workbook();

                // Когда вы создаете новую книгу, в книгу добавляется по умолчанию «Лист1».
                Worksheet sheet = wb.Worksheets[0];

                int count_cells = 0;
                foreach (string file in files)
                {
                    string path = file;

                    for (int j = 0; j < Find_docs(path).Length; j++)
                    {
                        //Console.Write("*");
                        for (int i = 1; i < Count_substr(Find_docs(path)[j], "СумУплНал=") + 1; i++)
                        {
                            Cell cell = sheet.Cells[count_cells, 0];
                            cell.PutValue(Find_INN(Find_docs(path)[j]));

                            cell = sheet.Cells[count_cells, 2];
                            cell.PutValue(Find_name(Find_docs(path)[j]));


                            // Console.WriteLine(Find_name(Find_docs(path)[j]) + " " + Find_INN(Find_docs(path)[j]) + " " + Find_type_nalog(Find_docs(path)[j])[i] + " " + Find_nalog(Find_docs(path)[j])[i]);
                            if (Find_type_nalog(Find_docs(path)[j])[i] == "УСНО")
                            {
                                cell = sheet.Cells[count_cells, 3];
                                cell.PutValue(Find_type_nalog(Find_docs(path)[j])[i]);
                                cell = sheet.Cells[count_cells, 4];
                                cell.PutValue(Find_type_nalog(Find_docs(path)[j])[i]);
                            }
                            else
                            {
                                cell = sheet.Cells[count_cells, 4];
                                cell.PutValue(Find_type_nalog(Find_docs(path)[j])[i]);
                            }
                            cell = sheet.Cells[count_cells, 5];
                            cell.PutValue(Find_nalog(Find_docs(path)[j])[i]);

                            count_cells++;

                        }

                    }
                }
                wb.Save("Excel.xlsx", SaveFormat.Xlsx);

            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
        }

        public static int Count_substr(String str,String substr)
        {
            int count = 0;
            int index = str.IndexOf(substr);
            while (index != -1)
            {
                count++;
                index = str.IndexOf(substr, index + 1);
            }
            return count;
        }
        public static int Count_Str(string str, char ch)
        {
            int count = 0;

            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ch)
                {
                    count++;
                }
            }

            return count;
        }

        public static string Find_name(string str)
        { 
           try
              {
                str = str.Substring(str.IndexOf("НаимОрг=") + 8, str.IndexOf("ИНН")-10);
                str =  str.Replace("&quot;", "\"");    
              }
           catch
              {
                            
              }                     
            return str;
  
        }

        public static string Find_INN(string str)
        {

            try
            {

                str = str.Substring(str.IndexOf("ИННЮЛ=") + 6);
                str = str.Substring(0,str.IndexOf("\"/>")+1);
                str = str.Replace("&quot;", "\"");
            }
            catch
            {
        
            }
            return str;
        }

        public static string[] Find_nalog(string str)
        {

                    string[] words = str.Split("СумУплНал=", StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 1; i < words.Length; i++)
                    {
                        try
                        {
                            words[i] = words[i].Substring(1, words[i].IndexOf("\"/>")-1);
                            
                        }
                        catch
                        {
                        }

                    }
                    return words;
        }

        public static string[] Find_type_nalog(string str)
        {

            string[] words = str.Split("НаимНалог=\"", StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i < words.Length; i++)
            {
                try
                {

                    words[i] = words[i].Substring(0, words[i].IndexOf("\""));

                    if (words[i] == "Налог, взимаемый в связи с  применением упрощенной  системы налогообложения")
                    {
                        words[i] = "УСНО";
                    }
                    else
                    {
                        words[i] = "";
                    }
                 
                }
                catch
                {
                 
                }
            }
            return words;
        }

        public static string[] Find_docs(string path)
        {

            using (StreamReader sr = new StreamReader(path))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {

                    string[] words = line.Split("СведНП", StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 1; i < words.Length; i++)
                    {
                        try
                        {

                            words[i] = words[i].Substring(0, words[i].IndexOf("ДатаСост="));

                            
                            // Console.WriteLine(words[i]);
                        }
                        catch
                        {
                            // Console.WriteLine("каво");
                        }

                    }
                    return words;

                }
            }


            return null;

            // Console.WriteLine("Hello World!");
        }


    
}
}

