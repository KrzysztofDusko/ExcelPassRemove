using System;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelDataReader;


namespace ExcelPass
{
    class Program
    {
        static int Main(string[] args)
        {
            Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var conf = new ExcelReaderConfiguration { Password = "123456" };
            NumberFormatInfo nfi = new NumberFormatInfo
            {
                NumberDecimalSeparator = "."
            };

            try
            {
                ExcelTest(@"C:\Users\dusko\Desktop\secretExcelFile.xlsx", @"C:\Users\dusko\Desktop\wellKnownSecrets.csv", ';', conf, nfi);
            }
            catch (ExcelDataReader.Exceptions.InvalidPasswordException e)
            {
                Console.WriteLine(e.Message);
                return 1;
            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.Message);
                return 2;
            }


            return 0;
        }

        static void ExcelTest(string xlsxPath, string csvPath, char d, ExcelReaderConfiguration conf, NumberFormatInfo nfi)
        {
            using (var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream, conf))
                {
                    int wNum = 0;
                    do
                    {
                        int kolNum = reader.FieldCount;
                        int rowNum = reader.RowCount;

                        StreamWriter sw = new StreamWriter($"{csvPath[0..(csvPath.Length - 4)]}{++wNum}.csv", false, Encoding.UTF8);
                        Type cellType;

                        while (reader.Read())
                        {
                            for (int i = 0; i < kolNum; i++)
                            {
                                cellType = reader.GetFieldType(i);

                                if (cellType == typeof(string))
                                {
                                    sw.Write(reader.GetString(i));
                                }
                                else if (cellType == typeof(double) || cellType == typeof(decimal))
                                {
                                    sw.Write(reader.GetDouble(i).ToString(nfi));
                                }
                                else if (cellType == typeof(DateTime))
                                {
                                    sw.Write(reader.GetDateTime(i).ToString("yyyy-MM-dd HH:mm:ss"));
                                }
                                else
                                {
                                    sw.Write(reader.GetValue(i));
                                }
                                if (i < kolNum - 1)
                                {
                                    sw.Write(d);
                                }
                            }
                            sw.WriteLine();
                        }
                        sw.Close();
                    } while (reader.NextResult());


                }
            }
        }


    }
}
