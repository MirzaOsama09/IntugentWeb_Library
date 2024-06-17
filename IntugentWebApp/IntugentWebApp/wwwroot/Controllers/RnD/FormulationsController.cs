using IntugentWebApp.Models;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using System.Reflection.Metadata;

namespace IntugentWebApp.Controllers.RnD
{
    public class FormulationsController
    {
        private static IConfiguration _configuration { get; set; }

        public static void Initialize(IConfiguration configuration)
        {
            _configuration = configuration;
        }


        public static async Task<RNDSearchResult> GetSearchResults(string id)
        {
            using (SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("Default")))
            {
                try
                {
                    connection.Open();

                    string query = "Select TOP(20) RN.ID,DateDSCreated, RN.[Study Name], RN.[Product ID], [Product Matrix].[Product Description], R1.Employees as 'Operator', R2.Employees as 'Chemist' from[RNDDatasets] as RN " +
                                    "Left JOIN Roster AS R1 ON RN.Operator = R1.ID " +
                                    "Left JOIN Roster AS R2 ON RN.Chemist = R2.ID " +
                                    "Left Join[Product Matrix] on RN.[Product ID] = [Product Code]" +
                                    "Where RN.ID = @id";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@id", id);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            RNDSearchResult result = new RNDSearchResult();
                            if (reader.Read())
                            {
                                result.ID = reader.GetInt32(reader.GetOrdinal("ID"));
                                result.DateDSCreated = reader.GetDateTime(reader.GetOrdinal("DateDSCreated"));
                                result.StudyName = reader.IsDBNull(reader.GetOrdinal("Study Name")) ? null : reader.GetString(reader.GetOrdinal("Study Name"));
                                result.ProductID = reader.IsDBNull(reader.GetOrdinal("Product ID")) ? null : reader.GetString(reader.GetOrdinal("Product ID"));
                                result.ProductDescription = reader.IsDBNull(reader.GetOrdinal("Product Description")) ? null : reader.GetString(reader.GetOrdinal("Product Description"));
                                result.Operator = reader.IsDBNull(reader.GetOrdinal("Operator")) ? null : reader.GetString(reader.GetOrdinal("Operator"));
                                result.Chemist = reader.IsDBNull(reader.GetOrdinal("Chemist")) ? null : reader.GetString(reader.GetOrdinal("Chemist"));
                            }
                            return await Task.FromResult(result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message + ex.StackTrace);
                }
                finally
                {
                    connection.Close();
                }
            }
        }

    }
}
