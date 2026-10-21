using System;
using System.Data;
using System.Data.SqlClient;
using FreightAutomation.Utility;

namespace FreightAutomation.DAL
{
    public class FreightDAL
    {
        private readonly DBConnection dbConnection;

        public FreightDAL()
        {
            dbConnection = new DBConnection();
        }

        // ==========================================
        // GET NEXT UPLOAD ID
        // ==========================================

        public int GetNextUploadId()
        {
            using (SqlConnection con = dbConnection.GetConnection())
            {
                string query = @"
                    SELECT ISNULL(MAX(UploadId), 0) + 1
                    FROM FreightCal";

                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    con.Open();

                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
        }

        // ==========================================
        // INSERT EXCEL DATA
        // ==========================================

        public void InsertFreightData(DataTable sqlTable)
        {
            using (SqlConnection connection =
                dbConnection.GetConnection())
            {
                connection.Open();

                using (SqlBulkCopy bulkCopy =
                    new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName =  "dbo.FreightCal";

                    bulkCopy.ColumnMappings.Add( "Date", "Date");

                    bulkCopy.ColumnMappings.Add(   "LRNo", "LRNo");

                    bulkCopy.ColumnMappings.Add(  "JK", "JK");

                    bulkCopy.ColumnMappings.Add("CustomerName", "CustomerName");

                    bulkCopy.ColumnMappings.Add(   "ToLocation", "ToLocation");

                    bulkCopy.ColumnMappings.Add(   "VehicleNo", "VehicleNo");

                    bulkCopy.ColumnMappings.Add(  "VehicleType", "VehicleType");

                    bulkCopy.ColumnMappings.Add("NoOfTyres", "NoOfTyres");

                    bulkCopy.ColumnMappings.Add(  "Weight", "Weight");

                    bulkCopy.ColumnMappings.Add(  "InvoiceNo", "InvoiceNo");

                    bulkCopy.ColumnMappings.Add( "BeatCycleTiming", "BeatCycleTiming");

                    bulkCopy.ColumnMappings.Add( "Tyre", "Tyre");

                    bulkCopy.ColumnMappings.Add( "Rate1", "Rate1");

                    bulkCopy.ColumnMappings.Add( "UploadId", "UploadId");

                    bulkCopy.WriteToServer(sqlTable);
                }
            }
        }

        // ==========================================
        // CALCULATE TYRE AND RATE3
        // ==========================================

        public void CalculateFreightRate(int uploadId)
        {
            using (SqlConnection connection =
                dbConnection.GetConnection())
            {
                using (SqlCommand command =
                    new SqlCommand(
                        "CalculateFreightRate",
                        connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.Add(   "@UploadId",   SqlDbType.Int).Value = uploadId;

                    connection.Open();

                    command.ExecuteNonQuery();
                }
            }
        }
    }
}