using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace Exceldata
{
    class Program
    {
        static void Main(string[] args)
        {
                List<string> mnfNumbers = new List<string>();
                List<string> mnfPartQuantity = new List<string>();
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string filePath = @"C:\Users\tekorius\Desktop\Book1.xlsx";
                int lookForMnf = 0;
                int lookForMnfQuantity = 7;

            mnfNumbers = GetPartNumber(filePath, lookForMnf);
           mnfPartQuantity = GetNumber(filePath, lookForMnfQuantity);

            foreach (var item in mnfNumbers)
            {
                Console.WriteLine(item);
            }
            foreach (var item in mnfPartQuantity)
            {
                Console.WriteLine(item);
            }

            Console.ReadKey();
        }
   
        public static List<string> GetPartNumber(string filePath, int lookForMnf)
        {
            List<string> mnfNumbers = new List<string>();
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                string mnf = Convert.ToString(reader.GetValue(lookForMnf));

                                if (mnf != "Buh. Nr.")
                                {
                                    Console.WriteLine("Blogas pavadinimas arba stulpelio vieta \"Buh, Nr.\"!");
                                    reader.Close();
                                }
                                else if (mnf == "Buh. Nr.")
                                {
                                    do
                                    {
                                        if (mnf != "Buh. Nr.")
                                        {
                                            if (mnf != string.Empty)
                                            {
                                                mnfNumbers.Add(mnf);
                                            }
                                        }
                                        mnf = Convert.ToString(reader.GetValue(lookForMnf));
                                    } while (reader.Read());
                                }
                            }

                        } while (reader.NextResult());
                        var result = reader.AsDataSet();
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Dokumentas yra atidarytas, uzdarykite ");
            }
            return mnfNumbers;
        }
        public static List<string> GetNumber(string filePath, int lookForMnfQuantity)
        {
            List<string> mnfNumbers = new List<string>();
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                string mnf = Convert.ToString(reader.GetValue(lookForMnfQuantity));

                                if (mnf != "QTY")
                                {
                                    Console.WriteLine("Blogas pavadinimas arba stulpelio vieta \"QTY\"!");
                                    reader.Close();
                                }
                                else if (mnf == "QTY")
                                {
                                    do
                                    {
                                        if (mnf != "QTY")
                                        {
                                            if (mnf != string.Empty)
                                            {
                                                mnfNumbers.Add(mnf);
                                            }
                                        }
                                        mnf = Convert.ToString(reader.GetValue(lookForMnfQuantity));
                                    } while (reader.Read());
                                }
                            }

                        } while (reader.NextResult());
                        var result = reader.AsDataSet();
                    }
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Dokumentas yra atidarytas, uzdarykite ");
            }
            return mnfNumbers;
        }
    }
}
