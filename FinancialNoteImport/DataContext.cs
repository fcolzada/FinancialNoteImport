using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FinancialNoteImport
{
    class DataContext
    {
        static private int expectedLenght = 9;
        private int lastPk;
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

        private void intializeContext()
        {
            using (SqlConnection sconn = new SqlConnection(sqlStrBld.ToString()))
            {
                StringBuilder strBld = new StringBuilder();
                strBld.AppendLine("select pkRaw_HouseFinance, valueDate ");
                strBld.AppendLine("from Raw_Finance");
                strBld.AppendLine("where pkRaw_HouseFinance = (select max(pkRaw_HouseFinance) from Raw_Finance)");

                using (var cmd = sconn.CreateCommand())
                {
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = strBld.ToString();
                    SqlDataReader reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        lastPk = reader.GetInt32(0);
                        LAST_RECORDED_DATE = reader.GetDateTime(1);
                    }
                    else
                    {
                        lastPk = 0;
                        LAST_RECORDED_DATE = DateTime.Now;
                    }
                }
            }
            throw new NotImplementedException();
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
            { writeEntryToDB(l); }
        }

        private void writeEntryToDB(Dictionary<String, String> dictLine)
        {
            StringBuilder strBld = new StringBuilder();
            strBld.AppendLine("insert into dbo.Raw_Finance (amount, valueDate, description, notes, n_sequence, loadDate) ");
            strBld.AppendLine("select @amount, @date, @description, @notes ");
            strBld.AppendLine("where not exists ");
            strBld.AppendLine("(select * from dbo.Raw_Finance where amount = @amount");
            strBld.AppendLine("AND valueDate = @date and description = @description and notes = @notes) ");

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
                            if (itemDate <= LAST_RECORDED_DATE)
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
