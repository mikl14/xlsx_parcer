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


            string path = @"21.xml";

            int last_i = 0;
            try
            {
                // Создайте экземпляр объекта Workbook, который представляет файл Excel.
                Workbook wb = new Workbook();

                // Когда вы создаете новую книгу, в книгу добавляется по умолчанию «Лист1».
                Worksheet sheet = wb.Worksheets[0];

                for (int i = 1; i < Find_name(path).Length;i++)
                {
                    Cell cell = sheet.Cells[i+last_i, 0];
                    cell.PutValue(Find_INN(path)[i]);

                    cell = sheet.Cells[i+last_i, 2];
                    cell.PutValue(Find_name(path)[i]);

                    if (Find_type_nalog(path)[i]=="УСНО") 
                    {
                        cell = sheet.Cells[i+last_i, 3];
                        cell.PutValue(Find_type_nalog(path)[i]);
                        cell = sheet.Cells[i+last_i, 4];
                        cell.PutValue(Find_type_nalog(path)[i]);
                    }
                    else
                    {
                        cell = sheet.Cells[i+last_i, 4];
                        cell.PutValue(Find_type_nalog(path)[i]);
                    }
                    cell = sheet.Cells[i+last_i, 5];
                    cell.PutValue(Find_nalog(path)[i]);
                }
                last_i+=Find_name(path).Length;
                // Сохраните Excel как файл .xlsx.
                wb.Save("Excel.xlsx", SaveFormat.Xlsx);

            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
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

        public static string[] Find_name(string path)
        {
           
                using (StreamReader sr = new StreamReader(path))
                {
                    string line;
                   
                    while ((line = sr.ReadLine()) != null)
                    {

                        string[] words = line.Split("НаимОрг=", StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 1; i < words.Length;i++)
                        {
                            try
                            {

                                words[i] = words[i].Substring(1, words[i].IndexOf("&quot;\"") +5);
                                words[i] =  words[i].Replace("&quot;", "\"");
                                if(Count_Str(words[i],'\"') % 2 != 0)
                                {
                                    words[i] += '\"';
                                }
                              //  Console.WriteLine(words[i]);
                            }
                            catch
                            {
                             //   Console.WriteLine("каво");
                            }
                           
                        }
                        return words;

                    }
                }


            return null;

            // Console.WriteLine("Hello World!");
        }

        public static string[] Find_INN(string path)
        {

            using (StreamReader sr = new StreamReader(path))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {

                    string[] words = line.Split("ИННЮЛ=", StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 1; i < words.Length; i++)
                    {
                        try
                        {

                            words[i] = words[i].Substring(0, words[i].IndexOf("\"/>")+1);
                           
                   
                          //  Console.WriteLine(words[i]);
                        }
                        catch
                        {
                          //  Console.WriteLine("каво");
                        }

                    }
                    return words;

                }
            }


            return null;

            // Console.WriteLine("Hello World!");
        }

        public static string[] Find_nalog(string path)
        {

            using (StreamReader sr = new StreamReader(path))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {

                    string[] words = line.Split("СумУплНал=", StringSplitOptions.RemoveEmptyEntries);

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
            }


            return null;

            // Console.WriteLine("Hello World!");
        }

        public static string[] Find_type_nalog(string path)
        {

            using (StreamReader sr = new StreamReader(path))
            {
                string line;

                while ((line = sr.ReadLine()) != null)
                {

                    string[] words = line.Split("НаимНалог=\"", StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 1; i < words.Length; i++)
                    {
                        try
                        {

                            words[i] = words[i].Substring(0, words[i].IndexOf("\""));

                            if(words[i] == "Налог, взимаемый в связи с  применением упрощенной  системы налогообложения")
                            {
                                words[i] = "УСНО";
                            }
                            else
                            {
                                words[i] = "";
                            }
                          //  Console.WriteLine(words[i]);
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

        public static void WriteExcel()
        {
            // Создайте экземпляр объекта Workbook, который представляет файл Excel.
            Workbook wb = new Workbook();

            // Когда вы создаете новую книгу, в книгу добавляется по умолчанию «Лист1».
            Worksheet sheet = wb.Worksheets[0];

            // Получите доступ к ячейке «A1» на листе.
            Cell cell = sheet.Cells[0,1];

            // Введите «Привет, мир!» текст в ячейку «А1».
            cell.PutValue("Hello World!");

            // Сохраните Excel как файл .xlsx.
            wb.Save("Excel.xlsx", SaveFormat.Xlsx);
        }
    
}
}

