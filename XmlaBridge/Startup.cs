using System;
using System.Data;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.AnalysisServices.AdomdClient;

public class Startup
{
    public async Task<object> Invoke(dynamic input)
    {
        try
        {
            string server = (string)input.server;
            string database = (string)input.database;
            string query = (string)input.query;

            // Get timeout from input, default to 30 seconds
            int timeout = 30;
            if (input.timeout != null)
            {
                timeout = (int)input.timeout;
            }

            // Build connection string
            string connectionString = $"Data Source={server};Initial Catalog={database};Integrated Security=SSPI;";

            Console.WriteLine($"[ADOMD.NET] Connecting to: {server}");
            Console.WriteLine($"[ADOMD.NET] Database: {database}");
            Console.WriteLine($"[ADOMD.NET] Query: {query.Substring(0, Math.Min(200, query.Length))}...");
            Console.WriteLine($"[ADOMD.NET] Timeout: {timeout} seconds");

            using (AdomdConnection conn = new AdomdConnection(connectionString))
            {
                conn.Open();
                Console.WriteLine("[ADOMD.NET] Connection opened successfully");

                using (AdomdCommand cmd = new AdomdCommand(query, conn))
                {
                    cmd.CommandTimeout = timeout;

                    using (AdomdDataReader reader = cmd.ExecuteReader())
                    {
                        var results = new List<Dictionary<string, object>>();

                        while (reader.Read())
                        {
                            var row = new Dictionary<string, object>();

                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string columnName = reader.GetName(i);
                                object value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                                row[columnName] = value;
                            }

                            results.Add(row);
                        }

                        Console.WriteLine($"[ADOMD.NET] Query returned {results.Count} rows");

                        return new
                        {
                            success = true,
                            rowCount = results.Count,
                            data = results
                        };
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ADOMD.NET] Error: {ex.Message}");
            Console.WriteLine($"[ADOMD.NET] Stack: {ex.StackTrace}");

            return new
            {
                success = false,
                error = ex.Message,
                errorType = ex.GetType().Name,
                stackTrace = ex.StackTrace
            };
        }
    }
}
