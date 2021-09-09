using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace MSAccessConverter
{
    class Program
    {
        public static DataTable dtMain;

        public static string inputFile = "";
        public static string table = "";
        public static string outputFile = "";

        static void Main(string[] args)
        {
            if (!Directory.Exists("./output")) Directory.CreateDirectory("./output");
            if (!Directory.Exists("./input")) Directory.CreateDirectory("./input");

            if (args.Length > 0)
            {
                if (args[0] == "help")
                {
                    Console.WriteLine($"Usage:\r\n \tMSAccessConverter.exe example.accdb Table_1 example.csv");
                    return;
                }

                inputFile = args[0];
                table = args[1];
                outputFile = args[2];
            }
            else
            {
                return;
            }

            string dbPath = @$"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=./input/{inputFile};";
            StringBuilder sb = new StringBuilder();

            try
            {
                using (OleDbConnection conn = new OleDbConnection(dbPath))
                {
                    conn.Open();

                    OleDbDataAdapter adapter = new OleDbDataAdapter($"SELECT * FROM [{table}]", conn);

                    new OleDbCommandBuilder(adapter);

                    dtMain = new DataTable();
                    adapter.Fill(dtMain);

                    var columns = dtMain.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                    sb.AppendLine(string.Join(",", columns));

                    foreach (DataRow row in dtMain.Rows)
                    {
                        var fields = row.ItemArray.Select(x => x.ToString()).ToArray();
                        sb.AppendLine(string.Join(",", fields));
                    }

                    File.WriteAllText($"./output/{outputFile}", sb.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return;
            }

            Console.WriteLine("Convert successful!");
        }
    }
}
