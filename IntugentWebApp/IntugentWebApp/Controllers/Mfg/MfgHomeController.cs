using IntugentWebApp.Models;
using Microsoft.AspNetCore.Http;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;

namespace IntugentWebApp.Controllers.Mfg
{
    public class MfgHomeController
    {
        private static IConfiguration _configuration { get; set; }

        public static void Initialize(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public static async Task<List<MFGSearchResult>> GetSearchResults()
        {
            List<MFGSearchResult> results = new List<MFGSearchResult>();

            using (SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("Default")))
            {
                try
                {
                    connection.Open();

                    string query = "SELECT TOP(50) RN.[ID4ALL], convert(varchar(10), RN.[Test Date], 1) as sDate,RN.[Shift], Format(RN.[Test Date], 'MM/dd/yy - hh:mm tt') as sTestTime,Format(R2.[Finished Board Time Stamp FG], 'MM/dd/yy - hh:mm tt') as sFGTestTime, RN.[Product ID], RN.[Product ID] +' - ' +  R1.[Product Description] as Product  " +
                                    " ,R4.slocation, R5.sName  FROM[In Process Identify] as RN Left Join[Product Matrix] as R1 on RN.[Product ID] = R1.[Product Code] " +
                                    " Left Join[Finished Goods] as R2 on RN.[ID4all] = R2.[ID4ALL FG] Left Join[Dimensional Stability] as R3 on RN.[ID4all] = R3.[ID4all]" +
                                    " Left join tblLocations as R4 on RN.Location = R4.ID Left join tblLists as R5 on RN.Shift = R5.ID   ";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                MFGSearchResult result = new MFGSearchResult
                                {
                                    ID = reader.GetInt32(reader.GetOrdinal("ID4ALL")),
                                    Date = reader.IsDBNull(reader.GetOrdinal("sDate")) ? (DateTime?)null : DateTime.Parse(reader.GetString(reader.GetOrdinal("sDate"))),
                                    Shift = reader.IsDBNull(reader.GetOrdinal("sName")) ? null : reader.GetString(reader.GetOrdinal("sName")),
                                    GreenBoardTimeStamp = reader.IsDBNull(reader.GetOrdinal("sTestTime")) ? (DateTime?)null : DateTime.ParseExact(reader.GetString(reader.GetOrdinal("sTestTime")), "MM/dd/yy - hh:mm tt", CultureInfo.InvariantCulture),
                                    FGBoardTimeStamp = reader.IsDBNull(reader.GetOrdinal("sFGTestTime")) ? (DateTime?)null : DateTime.ParseExact(reader.GetString(reader.GetOrdinal("sFGTestTime")), "MM/dd/yy - hh:mm tt", CultureInfo.InvariantCulture),
                                    Product = reader.IsDBNull(reader.GetOrdinal("Product")) ? null : reader.GetString(reader.GetOrdinal("Product")),
                                    Location = reader.IsDBNull(reader.GetOrdinal("sLocation")) ? null : reader.GetString(reader.GetOrdinal("sLocation")),
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

        public static async Task<List<MFGSearchResult>> GetSearchResults(MfgSearchDatasetParams param)
        {
            List<MFGSearchResult> results = new List<MFGSearchResult>();
            (string sqlSearchParams, bool isParamPresent) = GetSearchCriteria(param);

            using (SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("Default")))
            {
                try
                {
                    connection.Open();

                    string query = "SELECT TOP(50) RN.[ID4ALL], convert(varchar(10), RN.[Test Date], 1) as sDate,RN.[Shift], Format(RN.[Test Date], 'MM/dd/yy - hh:mm tt') as sTestTime,Format(R2.[Finished Board Time Stamp FG], 'MM/dd/yy - hh:mm tt') as sFGTestTime, RN.[Product ID], RN.[Product ID] +' - ' +  R1.[Product Description] as Product  " +
                                    " ,R4.slocation, R5.sName  FROM[In Process Identify] as RN Left Join[Product Matrix] as R1 on RN.[Product ID] = R1.[Product Code] " +
                                    " Left Join[Finished Goods] as R2 on RN.[ID4all] = R2.[ID4ALL FG] Left Join[Dimensional Stability] as R3 on RN.[ID4all] = R3.[ID4all]" +
                                    " Left join tblLocations as R4 on RN.Location = R4.ID Left join tblLists as R5 on RN.Shift = R5.ID   ";

                    if(isParamPresent)
                    {
                        query += "WHERE " + sqlSearchParams;
                    }

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                MFGSearchResult result = new MFGSearchResult
                                {
                                    ID = reader.GetInt32(reader.GetOrdinal("ID4ALL")),
                                    Date = reader.IsDBNull(reader.GetOrdinal("sDate")) ? (DateTime?)null : DateTime.Parse(reader.GetString(reader.GetOrdinal("sDate"))),
                                    Shift = reader.IsDBNull(reader.GetOrdinal("sName")) ? null : reader.GetString(reader.GetOrdinal("sName")),
                                    GreenBoardTimeStamp = reader.IsDBNull(reader.GetOrdinal("sTestTime")) ? (DateTime?)null : DateTime.ParseExact(reader.GetString(reader.GetOrdinal("sTestTime")), "MM/dd/yy - hh:mm tt", CultureInfo.InvariantCulture),
                                    FGBoardTimeStamp = reader.IsDBNull(reader.GetOrdinal("sFGTestTime")) ? (DateTime?)null : DateTime.ParseExact(reader.GetString(reader.GetOrdinal("sFGTestTime")), "MM/dd/yy - hh:mm tt", CultureInfo.InvariantCulture),
                                    Product = reader.IsDBNull(reader.GetOrdinal("Product")) ? null : reader.GetString(reader.GetOrdinal("Product")),
                                    Location = reader.IsDBNull(reader.GetOrdinal("sLocation")) ? null : reader.GetString(reader.GetOrdinal("sLocation")),
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

        private static (string, bool) GetSearchCriteria(MfgSearchDatasetParams param)
        {
            bool isParamPresent = false;
            string sql = string.Empty;

            if(param.Group != null)
            {
                //sql += " RN.[Location] = " + param.Group + " AND ";
                
                sql += " RN.[Location] = " + 1 + " AND "; // hardcoded to 1 !!!!!!!!!!!!!!
                isParamPresent = true;
            }

            if(param.ManufacturedBeforeAt != DateOnly.MinValue)
            {
                sql += "RN.[Test Date] < '" + param.ManufacturedBeforeAt.ToString() + "'" + " AND ";
                isParamPresent = true;
            }

            if (param.ManufacturedAfterAt != DateOnly.MinValue)
            {
                sql += "RN.[Test Date] >= '" + param.ManufacturedAfterAt.ToString() + "'" + " AND ";
                isParamPresent = true;
            }

            if (param.TestingStatus != null)
            {
                if (param.TestingStatus.Equals("Testing in Progress"))
                    sql += " (RN.[IP Testing Complete] = 'false' or RN.[IP Testing Complete] is null) and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) " + " AND ";   //Testing in Progress
                else if (param.TestingStatus.Equals("In Process Complete"))
                    sql += " RN.[IP Testing Complete] = 'true'  and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) " + " AND "; //In Process Complete
                else if (param.TestingStatus.Equals("Finished Goods Complete"))
                    sql += " R2.[FG Testing Complete] = 'true'  and (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) " + " AND "; //FG Testing Complete
                else if (param.TestingStatus.Equals("All Active Datasets"))
                    sql += " (RN.[Abandoned] is null or RN.[Abandoned] ='false' ) " + " AND "; //All Active Datasets
                else if (param.TestingStatus.Equals("Abandoned"))
                    sql += " RN.[Abandoned] = 'true'" + " AND "; //All Active Datasets - Should be Abandoned !!!!!!!!!!!!!!!!!!!!

                isParamPresent = true;
            }

            if (param.AgedRValueTesting != null)
            {
                if (param.AgedRValueTesting.Equals("Not Done"))
                    sql += " (R2.[FG Aged R Value Complete] is null or R2.[FG Aged R Value Complete] = 'false') " + " AND ";
                else if (param.AgedRValueTesting.Equals("Done"))
                    sql += " R2.[FG Aged R Value Complete] = 'true'" + " AND ";

                isParamPresent = true;
                // we're showing 3rd option "All Datasets... how's that being handled !!!!!!!
            }

            if (param.DimensionStabilityTesting != null)
            {
                if (param.DimensionStabilityTesting.Equals("Not Done")) 
                    sql += " (R3.[DS Testing Complete] is null or R3.[DS Testing Complete] = 'false') " + " AND ";
                else if (param.DimensionStabilityTesting.Equals("Done")) 
                    sql += "R3.[DS Testing Complete] = 'true'" + " AND ";

                // we're showing 3rd option "All Datasets... how's that being handled !!!!!!!
                isParamPresent = true;
            }

            if (param.RunType != null)
            {
                if (param.RunType.Equals("Regular Mfg.")) 
                    sql += " RN.[Run Type] = 4" + " AND ";
                else if (param.RunType.Equals("Trial Run")) 
                    sql += " RN.[Run Type] = 5" + " AND ";

                // we're showing 3rd option "All Datasets... how's that being handled !!!!!!!
                isParamPresent = true;
            }

            // remove extra "AND" from the end
            if (sql.EndsWith(" AND "))
            {
                sql = sql.Substring(0, sql.Length - " AND ".Length);
            }

            return (sql, isParamPresent);
        }
    }
}
