using IntugentWebApp.Models;
using System.Data.SqlClient;
using System.Reflection.Metadata;

namespace IntugentWebApp.Controllers.RnD
{
    public class RndHomeController
    {
        private static IConfiguration _configuration { get; set; }

        public static void Initialize(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public static async Task<List<RNDSearchResult>> GetSearchResults()
        {
            List<RNDSearchResult> results = new List<RNDSearchResult>();

            using (SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("Default")))
            {
                try
                {
                    connection.Open();

                    string query = "Select TOP(20) RN.ID,DateDSCreated, RN.[Study Name], RN.[Product ID], [Product Matrix].[Product Description], R1.Employees as 'Operator', R2.Employees as 'Chemist' from[RNDDatasets] as RN " +
                                    "Left JOIN Roster AS R1 ON RN.Operator = R1.ID " +
                                    "Left JOIN Roster AS R2 ON RN.Chemist = R2.ID " +
                                    "Left Join[Product Matrix] on RN.[Product ID] = [Product Code]";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                RNDSearchResult result = new RNDSearchResult
                                {
                                    ID = reader.GetInt32(reader.GetOrdinal("ID")),
                                    DateDSCreated = reader.GetDateTime(reader.GetOrdinal("DateDSCreated")),
                                    StudyName = reader.IsDBNull(reader.GetOrdinal("Study Name")) ? null : reader.GetString(reader.GetOrdinal("Study Name")),
                                    ProductID = reader.IsDBNull(reader.GetOrdinal("Product ID")) ? null : reader.GetString(reader.GetOrdinal("Product ID")),
                                    ProductDescription = reader.IsDBNull(reader.GetOrdinal("Product Description")) ? null : reader.GetString(reader.GetOrdinal("Product Description")),
                                    Operator = reader.IsDBNull(reader.GetOrdinal("Operator")) ? null : reader.GetString(reader.GetOrdinal("Operator")),
                                    Chemist = reader.IsDBNull(reader.GetOrdinal("Chemist")) ? null : reader.GetString(reader.GetOrdinal("Chemist")),
                                };
                                results.Add(result);
                            }
                            return await Task.FromResult(results);
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
