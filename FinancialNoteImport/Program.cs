using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using OfficeFormatUtility;
using System.IO;

namespace FinancialNoteImport
{
    class Program
    {
        private static readonly int expectedLenght;

        static void Main(string[] args)
        {
            try
            { 
            List<String> arguments = new List<string>(args);
            //obsolete_importFromtxt();
            //importwithOFU();
            checkArgs(arguments);
            DataContext dc = DataContext.CreateInstance();

                if (arguments.Contains("IC")) //caricamento Nexi RAW
                    dc.importByXLSX("C:\\Temp\\movimenti.xlsx");

                if (arguments.Contains("IB"))
                dc.importByCSV("C:\\Temp\\ListaMovimenti.CSV");

                

                if (arguments.Contains("C"))
                dc.consolidate();

            if (arguments.Contains("F"))
                dc.fillWallets();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured: " + ex.ToString());
            }
        }

        private static void checkArgs(List<String> arguments)
        {

            if (arguments.Count < 1)
                throw new ArgumentException("No argument supplied");
        }

        static void importwithOFU()
        {
            LegacyExcelManager excelManager = new LegacyExcelManager();
            DataTable dt = excelManager.GetFirstSheet("C:\\Temp\\ListaMovimenti.xls");
        }

        static void importFromOpenXML()
        {
            SmlDocument smldoc = new SmlDocument("C:\\Users\\fabio\\Desktop\\Book1.xlsx");

            var rng = SmlDataRetriever.RetrieveRange(smldoc, "Sheet1", 2, 2, 8, 8);



            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName("C:\\Users\\fabio\\Desktop\\Book1.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    int startRow;
                    int startCol;

                    findStartData(doc, out startCol, out startRow, 10, 10);



                }
                //streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "FormulasUpdated.xlsx"));
            }
        }

        static private void findStartData(SpreadsheetDocument doc, out int col, out int row, int maxCol, int maxRow)
        {
            var wsheet = WorksheetAccessor.GetWorksheet(doc, "Sheet1");
            col = row = 1;
            Object cv = null;

            do
            {
                cv = WorksheetAccessor.GetCellValue(doc, wsheet, col, row);
                if (cv != null)
                {
                    do
                    {
                        col -= 1;
                        cv = WorksheetAccessor.GetCellValue(doc, wsheet, col, row);
                    } while (cv != null);
                    col += 1;
                    return;
                }
                else
                {
                    col += 1;
                    row += 1;
                }
            } while (cv == null && row <= maxRow && col <= maxCol);
        }

        private static void obsolete_importFromtxt()
        {
            SqlConnectionStringBuilder sqlStrBld = new SqlConnectionStringBuilder(@"Server=localhost\SQLEXPRESS;Database=mainDB;Trusted_Connection=True;");
            SqlConnection sqlConn = new SqlConnection(sqlStrBld.ConnectionString);
            sqlConn.Open();

            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("insert into dbo.Raw_Finance (amount, valueDate, description, feeType, appliedTaxes, appliedWelfare, appliedSavings, appliedAi, appliedFree) values");
            strBld.AppendLine("(@amount, @date, @description, @feeType, @appliedTaxes, @appliedWelfare, @appliedSavings, @appliedAi, @appliedFree)");

            string fileLocation = "C:\\Users\\fabio\\Desktop\\importFN.txt";
            var lineCounter = 0;
            using (var fr = new System.IO.StreamReader(fileLocation))
            {
                string line = fr.ReadLine();
                while (line != null)
                {
                    if (lineCounter > 0)
                    {
                        string[] lineSplit = line.Split(Convert.ToChar("\t"));
                        SqlParameter[] prm = { };
                        try
                        {
                            prm = mapFieldToParam(lineSplit);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                        if (prm.Length == 9)
                        {
                            using (var cmd = sqlConn.CreateCommand())
                            {
                                cmd.CommandType = System.Data.CommandType.Text;
                                cmd.CommandText = strBld.ToString();
                                cmd.Parameters.AddRange(prm);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    lineCounter++;
                    if (lineCounter % 10 == 0)
                    {
                        Console.WriteLine("Processed {0} lines ", lineCounter.ToString());
                    }

                    line = fr.ReadLine();
                }
                Console.ReadLine();
            }
        }

        private static SqlParameter[] mapFieldToParam(string[] lineSplit)
        {
            List<SqlParameter> lstParam = new List<SqlParameter>();
            if (lineSplit.Length == expectedLenght)
            {
                string[] dateSplit = lineSplit[1].Split(Convert.ToChar("-"));
                lstParam.Add(new SqlParameter("@amount", fixDecimal(lineSplit[0])));
                lstParam.Add(new SqlParameter("@date", new DateTime(Convert.ToInt32(dateSplit[0]), Convert.ToInt32(dateSplit[1]), Convert.ToInt32(dateSplit[2]))));
                lstParam.Add(new SqlParameter("@description", lineSplit[2]));
                lstParam.Add(new SqlParameter("@feeType", lineSplit[3]));
                lstParam.Add(new SqlParameter("@appliedTaxes", fixDecimal(lineSplit[4])));
                lstParam.Add(new SqlParameter("@appliedWelfare", fixDecimal(lineSplit[5])));
                lstParam.Add(new SqlParameter("@appliedSavings", fixDecimal(lineSplit[6])));
                lstParam.Add(new SqlParameter("@appliedAi", fixDecimal(lineSplit[7])));
                lstParam.Add(new SqlParameter("@appliedFree", fixDecimal(lineSplit[8])));
            }
            else
                Console.WriteLine("Error parsing line: " + String.Join<string>(",", lineSplit));

            return lstParam.ToArray();
        }

        private static Decimal fixDecimal(string v)
        {
            return Convert.ToDecimal(String.IsNullOrWhiteSpace(v) ? "0" : v.Replace(",", "").Replace("\"", ""));
        }
    }
}
