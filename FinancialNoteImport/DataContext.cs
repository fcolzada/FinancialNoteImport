using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialNoteImport
{
    enum EnumWallet { 
        Taxes,
        Welfare,
        Savings,
        Free,
        Emergency
    }

    class DataContext
    {
        static private int expectedLenght = 9;
        static private SqlConnectionStringBuilder sqlStrBld = new SqlConnectionStringBuilder(@"Server=localhost\SQLEXPRESS;Database=mainDB;Trusted_Connection=True;");
        private DateTime LAST_RECORDED_DATE;
        
        public static DataContext CreateInstance()
        {
            DataContext dt = new DataContext();
            dt.intializeContext();
            return dt;
        }

        private DataContext()
        {
           
        }

        internal void importByXLSX(string v)
        {
            DataTable dt = new DataTable();
            List<Dictionary<String, String>> parsedExcelBanca = new List<Dictionary<string, string>>();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@"C:\Temp\movimenti.xlsx", false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = dt.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dt.Rows.Add(tempRow);
                }

            }
            dt.Rows.RemoveAt(0); //...so i'm taking it out here.

            foreach(DataRow r in dt.Rows)
            {
                //Data valuta;Descrizione operazione;Causale;Importo
                Dictionary<String, String> excelLine = new Dictionary<string, string>();

                String dd = r["Data"].ToString();
                String[] splitDate = dd.Split(Char.Parse("/"));
                DateTime dtdate = new DateTime(int.Parse(splitDate[2]), int.Parse(splitDate[1]), int.Parse(splitDate[0]));
                String importo = (Decimal.Parse(r[@"Importo (€)"].ToString()) * -1).ToString();

                excelLine.Add("data valuta", dtdate.ToString("yyyyMMdd"));
                excelLine.Add("descrizione operazione", r["descrizione"].ToString());
                excelLine.Add("causale", "NEXI IMPORT");
                excelLine.Add("importo", importo);

                parsedExcelBanca.Add(excelLine);
            }

            using (FinanceDB dbctx = new FinanceDB())
            {
                dbctx.Configuration.ValidateOnSaveEnabled = false;

                DateTime baseline = DateTime.ParseExact(
                    dbctx.AppConfiguration.Where(x => x._namespace == "FinancialNoteImport" && x.name == "baselineNexi").SingleOrDefault().value,
                    "yyyy-MM-dd", CultureInfo.InvariantCulture);

                foreach (var l in parsedExcelBanca)
                { writeEntryToDB(l, baseline); }

            }

            //StringBuilder strBld = new StringBuilder();
            
            //using (FinanceDB dbctx = new FinanceDB())
            //{
            //    dbctx.Configuration.ValidateOnSaveEnabled = false;

            //    dbctx.Database.Log = s => System.Diagnostics.Debug.WriteLine(s);
            //    DateTime baseline = DateTime.ParseExact(
            //        dbctx.AppConfiguration.Where(x => x._namespace == "FinancialNoteImport" && x.name == "baselineNexi").SingleOrDefault().value,
            //        "yyyy-MM-dd", CultureInfo.InvariantCulture);

            //    if (dt.Rows.Count > 0)
            //    {
            //        strBld.AppendLine("INSERT [dbo].[Raw_Finance]([amount], [valueDate], [description], [notes], [appliedTaxes], [appliedWelfare], [appliedSavings], [appliedAi], [appliedFree], [feeType], [n_sequence], [loadDate])");
            //        strBld.AppendLine("VALUES");
            //    }

            //    foreach (System.Data.DataRow r in dt.Rows)
            //    {
            //        string values = "(@amount, @valueDate, @description, @notes, NULL, NULL, NULL, NULL, NULL, NULL, NULL, GETDATE()),";
            //        values = values.Replace("@amount", "'" + (Decimal.Parse(r[@"Importo (€)"].ToString()) * -1).ToString() + "'");
            //        String dd = r["Data"].ToString();
            //        String[] splitDate = dd.Split(Char.Parse("/"));
            //        DateTime dtdate = new DateTime(int.Parse(splitDate[2]), int.Parse(splitDate[1]), int.Parse(splitDate[0]));
            //        values = values.Replace("@valueDate", "'" + dtdate.ToString("yyyyMMdd") + "'");
            //        values = values.Replace("@description", "'" + r["Descrizione"].ToString() + "'");
            //        values = values.Replace("@notes", "'NEXI IMPORT'");

            //        if (dtdate >= baseline)
            //            strBld.AppendLine(values);
            //    }
            //}

            //using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            //{
            //    sconn.Open();
            //    using (var cmd = sconn.CreateCommand())
            //    {
            //        cmd.CommandType = System.Data.CommandType.Text;
            //        cmd.CommandText = strBld.ToString().Trim().TrimEnd(Char.Parse(","));
            //        int affected = cmd.ExecuteNonQuery();
            //        Console.WriteLine("Imported {0} records from stage table.", affected);
            //    }

            //}


        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
{
    try
    {
        if (cell.CellValue == null)
            return "";

        string value = cell.CellValue.InnerXml;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
        }
        else
            return value;
    }
    catch (Exception ex)
    {
        return "";
    }
}


        private void intializeContext()
        {
            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                StringBuilder strBld = new StringBuilder();
                strBld.AppendLine("select CONVERT(DATE, value) from conf.AppConfiguration");
                strBld.AppendLine("where namespace = 'FinancialNoteImport' and name = 'baselineGenerali'");

                sconn.Open();

                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    SqlDataReader reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        LAST_RECORDED_DATE = reader.GetDateTime(0);
                    }
                    else
                    {
                        LAST_RECORDED_DATE = DateTime.Now;
                    }
                }
            }
            
        }

        public void consolidate()
        {
            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("select R.valueDate as d_transfer, CASE WHEN r.description like '%sdd%nexi%' THEN -2 ELSE R.amount END as m_amount, r.n_sequence as n_sequence, r.description as description, r.notes as note, CASE WHEN r.notes = 'nexi import' THEN 'CC' WHEN r.description like '%pagobancomat%' THEN 'PB' ELSE 'BA' END as idTransferType");
            strBld.AppendLine("into #RF from Raw_Finance R");

            strBld.AppendLine("insert into Transfer (d_transfer, m_amount, n_sequence, description, note, idTransferType)");
            strBld.AppendLine("select R.d_transfer, R.m_amount, R.n_sequence, R.description, R.note, R.idTransferType from #RF R");
            strBld.AppendLine("LEFT JOIN Transfer T");
            strBld.AppendLine("on R.m_amount = T.m_amount");
            strBld.AppendLine("and R.d_transfer = T.d_transfer");
            strBld.AppendLine("and R.n_sequence = T.n_sequence");
            strBld.AppendLine("where T.m_amount is null");
            strBld.AppendLine("and (R.d_transfer > '2018-07-23' OR R.note = 'nexi import')");

            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                sconn.Open();
                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    int affected = cmd.ExecuteNonQuery();
                    Console.WriteLine("Imported {0} records from stage table.", affected);
                }

            }
        }

        public void fillWallets()
        {
            StringBuilder strBld = new StringBuilder();

            strBld.AppendLine("INSERT INTO TransferWalletAllocation(d_transfer, m_amount, n_sequence, idWallet, m_allocated_amount)");
            strBld.AppendLine("SELECT T.d_transfer, T.m_amount, T.n_sequence, W.pkWallet, T.m_amount* W.p_rate");
            strBld.AppendLine("FROM Transfer T");
            strBld.AppendLine("LEFT JOIN TransferWalletAllocation TWA");
            strBld.AppendLine("on T.d_transfer = TWA.d_transfer");
            strBld.AppendLine("and T.m_amount = TWA.m_amount");
            strBld.AppendLine("and T.n_sequence = TWA.n_sequence");
            strBld.AppendLine("JOIN Wallet W");
            strBld.AppendLine("ON W.pkWallet <> 4");
            strBld.AppendLine("where TWA.d_transfer is null");
            strBld.AppendLine("and T.m_amount > 0");
            strBld.AppendLine("and T.description like '%cluster reply%'");

            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                sconn.Open();
                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    int affected = cmd.ExecuteNonQuery();
                    Console.WriteLine("Allocated salaries: {0}.", affected);
                }

            }

            strBld.Clear();



            strBld.AppendLine("select T.* from Transfer T");
            strBld.AppendLine("LEFT JOIN TransferWalletAllocation TWA");
            strBld.AppendLine("on T.d_transfer = TWA.d_transfer");
            strBld.AppendLine("and T.m_amount = TWA.m_amount");
            strBld.AppendLine("and T.n_sequence = TWA.n_sequence");
            strBld.AppendLine("where TWA.d_transfer is null");
            
            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                sconn.Open();
                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    SqlDataReader reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        decimal totalAmount = reader.GetDecimal(reader.GetOrdinal("m_amount"));
                        DateTime transferDate = reader.GetDateTime(reader.GetOrdinal("d_transfer"));
                        int sequence = reader.GetInt32(reader.GetOrdinal("n_sequence"));
                        string description = reader.GetString(reader.GetOrdinal("description")).ToLower();

                        if (totalAmount>0)
                        {
                            // Possible refund
                            allocateOnWallet(transferDate, totalAmount, sequence, EnumWallet.Taxes);
                        }
                        else
                        {
                            if (description.Contains("carrefour") || 
                                description.Contains("in s mercato") || 
                                description.Contains("atm point") || 
                                description.Contains("eurospin") ||
                                description.Contains("farmacia") || 
                                description.Contains("esselunga") || 
                                description.Contains("pam") || 
                                description.Contains("mc donald") || 
                                description.Contains("tari") || 
                                description.Contains("a2a") || 
                                description.Contains("telethon") || 
                                description.Contains("grocery") || 
                                description.Contains("fee"))
                                allocateOnWallet(transferDate, totalAmount, sequence, EnumWallet.Taxes);
                            else allocateOnWallet(transferDate, totalAmount, sequence, EnumWallet.Free);
                        }
                    }
                }

            }
        }

        private void allocateOnWallet(DateTime dateTransfer, Decimal amount, int sequence, EnumWallet walletID)
        {
            
            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("INSERT INTO TransferWalletAllocation(d_transfer, m_amount, n_sequence, idWallet, m_allocated_amount)");
            strBld.AppendLine("SELECT T.d_transfer, T.m_amount, T.n_sequence, @wallet, T.m_amount");
            strBld.AppendLine("FROM Transfer T");
            strBld.AppendLine("WHERE T.d_transfer = @dt");
            strBld.AppendLine("AND T.m_amount = @amount");
            strBld.AppendLine("AND T.n_sequence = @seq");

            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@dt", dateTransfer));
            pars.Add(new SqlParameter("@amount", amount));
            pars.Add(new SqlParameter("@seq", sequence));
            pars.Add(new SqlParameter("@wallet", walletID));

            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                sconn.Open();
                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    cmd.Parameters.AddRange(pars.ToArray());
                    int affected = cmd.ExecuteNonQuery();
                }

            }






        }

        public void importByCSV(String filePath)
        {
            List<Dictionary<String, String>> parsedCSVBanca = new List<Dictionary<string, string>>();
            List<String> header = new List<string>();
            int lineCounter = 0;
            string curLine = String.Empty;
            using (StreamReader sr = new StreamReader(filePath))
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
                            throw new ApplicationException("Wrong element number on line " + (lineCounter + 1).ToString() + ": " + curLine);
                        for (int curPos = 0; curPos < header.Count(); curPos++)
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
            { writeEntryToDB(l, LAST_RECORDED_DATE); }
        }

        private void writeEntryToDB(Dictionary<String, String> dictLine, DateTime BaseLineDate)
        {
            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("insert into dbo.Raw_Finance (amount, valueDate, description, notes, n_sequence, loadDate) ");
            strBld.AppendLine("select @amount, @date, @description, @notes, ISNULL(MAX(R.n_sequence), -1)+1, Getdate() ");
            strBld.AppendLine(" FROM dbo.Raw_Finance R where R.amount = @amount");
            strBld.AppendLine("AND R.valueDate = @date and R.description = @description and R.notes = @notes ");

            using (SqlConnection sqlConn = new SqlConnection(sqlStrBld.ConnectionString))
            {
                sqlConn.Open();
                // mapping sui parametri
                List<SqlParameter> prm = new List<SqlParameter>();
                foreach (string k in dictLine.Keys)
                {
                    switch (k.ToLower())
                    {
                        //Data movimento;Data valuta;Descrizione operazione;Causale;Importo;Divisa *;Saldo progressivo;Nota personale; 
                        case "data valuta":
                            var dateSplit = dictLine[k].Split('/');
                            DateTime itemDate = new DateTime(int.Parse(dateSplit[2]), int.Parse(dateSplit[1]), int.Parse(dateSplit[0]));
                            if (itemDate <= BaseLineDate)
                                return;
                            prm.Add(new SqlParameter("@date", itemDate));
                            break;
                        case "descrizione operazione":
                            string descrizione = dictLine[k].Substring(0, dictLine[k].Length > 200 ? 200 - 1 : dictLine[k].Length);
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

    }
}
