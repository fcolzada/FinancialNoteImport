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

        static int expectedLenght = 9;

        static void Main(string[] args)
        {
            //obsolete_importFromtxt();
            //importwithOFU();
            importByCSV();
        }

        private static void importByCSV()
        {
            List<Dictionary<String, String>> parsedCSVBanca = new List<Dictionary<string, string>>();
            List<String> header = new List<string>();
            int lineCounter = 0;
            string curLine = String.Empty;
            using (StreamReader sr = new StreamReader("C:\\Temp\\ListaMovimenti.CSV"))
            {
                curLine = sr.ReadLine();
                while (!string.IsNullOrEmpty(curLine))
                {
                    List<String> stripLine = curLine.Split(";".ToCharArray()).ToList();
                    //Header management
                    if (lineCounter < 1)
                    {
                        foreach (string k in stripLine)
                        {
                            header.Add(k.Trim("\"".ToArray()));
                        }
                    }
                    //Body management
                    else
                    {
                        Dictionary<String, String> csvLine = new Dictionary<string, string>();
                        if (header.Count != stripLine.Count())
                            throw new ApplicationException("Wrong element number on line " + (lineCounter+1).ToString() + ": " + curLine);
                        for(int curPos = 0; curPos < header.Count(); curPos++)
                        {
                            csvLine.Add(header[curPos], stripLine[curPos]);    
                        }
                        parsedCSVBanca.Add(csvLine);
                    }
                    lineCounter++;
                    curLine = sr.ReadLine();

                }

            }

            foreach (var l in parsedCSVBanca)
            { writeEntryToDB(l); }
        }

        private static void writeEntryToDB(Dictionary<String, String> dictLine)
        {
            var LAST_RECORDED_DATE = new DateTime(2018, 4, 28);
            SqlConnectionStringBuilder sqlStrBld = new SqlConnectionStringBuilder(@"Server=localhost\SQLEXPRESS;Database=mainDB;Trusted_Connection=True;");
            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("insert into dbo.Raw_Finance (amount, valueDate, description, notes) values");
            strBld.AppendLine("(@amount, @date, @description, @notes)");

            using (SqlConnection sqlConn = new SqlConnection(sqlStrBld.ConnectionString))
            {
                sqlConn.Open();
                // mapping sui parametri
                List<SqlParameter> prm = new List<SqlParameter>();
                foreach(string k in dictLine.Keys)
                {
                    switch (k.ToLower())
                    {
                        //Data movimento;Data valuta;Descrizione operazione;Causale;Importo;Divisa *;Saldo progressivo;Nota personale; 
                        case "data valuta":
                            var dateSplit = dictLine[k].Split('/');
                            DateTime itemDate = new DateTime(int.Parse(dateSplit[2]), int.Parse(dateSplit[1]), int.Parse(dateSplit[0]));
                            if (itemDate <= LAST_RECORDED_DATE)
                                return;
                            prm.Add(new SqlParameter("@date", itemDate));
                            break;
                        case "descrizione operazione":
                            string descrizione = dictLine[k].Substring(0, dictLine[k].Length > 200 ? 200-1 : dictLine[k].Length);
                            prm.Add(new SqlParameter("@description", descrizione));
                            break;
                        case "causale":
                            string causale = dictLine[k].Substring(0, dictLine[k].Length > 200 ? 200 - 1 : dictLine[k].Length);
                            prm.Add(new SqlParameter("@notes", causale));
                            break;
                        case "importo":
                            Decimal importo = Decimal.Parse(dictLine[k].Replace(',', '.'));
                            prm.Add(new SqlParameter("@amount", importo));
                            break;
                    }

                }

                using (var cmd = sqlConn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    cmd.Parameters.AddRange(prm.ToArray());
                    cmd.ExecuteNonQuery();
                }
            }
                
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
