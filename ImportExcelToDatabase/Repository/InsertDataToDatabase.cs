using Dapper;
using System.Data;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using ImportExcelToDatabase.Models;
using Npgsql;
using Oracle.ManagedDataAccess.Client;

namespace ImportExcelToDatabase.Repository
{
    public class InsertDataToDatabase
    {
        public async Task<ResponseModel> InsertAysnc(List<Dictionary<string, object>> dataList, ImportDataModel model)
        {
			try
			{
                IDbConnection connection;

                // must be upper case in each case
                switch (model.DatabaseType)
                {
                    case "MYSQL":
                        connection = new MySqlConnection(model.ConnectionString);
                        break;
                    case "POSTGRESQL":
                        connection = new NpgsqlConnection(model.ConnectionString);
                        break;
                    case "ORACLE":
                        connection = new OracleConnection(model.ConnectionString);
                        break;
                    case "MSSQL":
                    default:
                        connection = new SqlConnection(model.ConnectionString);
                        break;
                }

                using (connection)
                {
                    connection.Open();

                    // Build the SQL query dynamically based on the headers
                    var columns = new StringBuilder();
                    var values = new StringBuilder();
                    int rowIndex = 0;

                    foreach (var column in dataList[0].Keys)
                    {
                        columns.Append($"[{column}],");
                    }

                    // Remove the trailing comma
                    columns.Length--;

                    foreach (var row in dataList)
                    {
                        values.Append("(");

                        foreach (var column in row.Keys)
                        {
                            values.Append($"@{column}{rowIndex},");
                        }

                        // Remove the trailing comma and add a closing parenthesis
                        values.Length--;
                        values.Append("),");
                        rowIndex++;
                    }

                    // Remove the trailing comma
                    values.Length--;

                    string query = $"INSERT INTO {model.TableName} ({columns}) VALUES {values};";

                    // Combine all rows into a single DynamicParameters object
                    DynamicParameters parameters = new DynamicParameters();
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        foreach (var column in dataList[i].Keys)
                        {
                            parameters.Add($"{column}{i}", dataList[i][column]);
                        }
                    }

                    // Execute the query using Dapper
                    await connection.ExecuteAsync(query, parameters);

                    return new ResponseModel()
                    {
                        Success = true,
                        Data = dataList,
                        Message = $"Completed insert {dataList.Count} record(s) into table {model.TableName}"
                    };
                }
            }
			catch (Exception e)
			{
                return new ResponseModel()
                { 
                    Success = false,
                    Message = $"Exception at InsertDataToDatabase.InsertAysnc: {e.Message}" 
                };
			}
        }
    }
}
