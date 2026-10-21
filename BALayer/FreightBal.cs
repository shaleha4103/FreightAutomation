using System;
using System.Data;
using System.Web;
using ExcelDataReader;
using FreightAutomation.DAL;

namespace FreightAutomation.BAL
{
    public class FreightBAL
    {
        private readonly FreightDAL freightDAL;

        public FreightBAL()
        {
            freightDAL = new FreightDAL();
        }



        public int UploadExcel(HttpPostedFileBase excelfile)
        {
            DataTable excelTable;


            // ==========================================
            // 1. READ EXCEL FILE
            // ==========================================

            using (var reader =
                ExcelReaderFactory.CreateReader(
                    excelfile.InputStream))
            {
                var dataSet =
                    reader.AsDataSet(
                        new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = x =>
                                new ExcelDataTableConfiguration
                                {
                                    UseHeaderRow = true
                                }
                        });

                excelTable = dataSet.Tables[0];
            }


            // ==========================================
            // CHECK EXCEL DATA
            // ==========================================

            if (excelTable == null ||
                excelTable.Rows.Count == 0)
            {
                throw new Exception(
                    "The Excel file does not contain any data.");
            }


            // ==========================================
            // 2. CLEAN EXCEL COLUMN NAMES
            // ==========================================

            foreach (DataColumn column
                in excelTable.Columns)
            {
                column.ColumnName =
                    column.ColumnName
                    .Trim()
                    .Replace("\u00A0", " ");
            }


            // ==========================================
            // 3. GET NEW UPLOAD ID
            // ==========================================

            int uploadId =
                freightDAL.GetNextUploadId();


            // ==========================================
            // 4. CREATE DATA TABLE
            // ==========================================

            DataTable sqlTable =  new DataTable();


            sqlTable.Columns.Add( "Date", typeof(DateTime));

            sqlTable.Columns.Add( "LRNo", typeof(int));

            sqlTable.Columns.Add( "JK", typeof(string));

            sqlTable.Columns.Add( "CustomerName",  typeof(string));

            sqlTable.Columns.Add( "ToLocation", typeof(string));

            sqlTable.Columns.Add("VehicleNo",  typeof(string));

            sqlTable.Columns.Add( "VehicleType", typeof(string));

            sqlTable.Columns.Add( "NoOfTyres", typeof(int));

            sqlTable.Columns.Add( "Weight",  typeof(decimal));

            sqlTable.Columns.Add( "InvoiceNo",  typeof(string));

            sqlTable.Columns.Add( "BeatCycleTiming",  typeof(string));

            sqlTable.Columns.Add( "Tyre", typeof(int));

            sqlTable.Columns.Add(  "Rate1", typeof(decimal));

            sqlTable.Columns.Add( "UploadId",  typeof(int));


            // ==========================================
            // 5. COPY EXCEL DATA
            // ==========================================

            foreach (DataRow excelRow
                in excelTable.Rows)
            {
                DataRow row =
                    sqlTable.NewRow();


                // ==========================================
                // DATE
                // ==========================================

                string dateValue =  excelRow["DATE"]?.ToString().Trim();

                DateTime date;

                if (DateTime.TryParse(
                    dateValue,
                    out date))
                {
                    row["Date"] = date;
                }
                else
                {
                    row["Date"] = DBNull.Value;
                }


                // ==========================================
                // LR NO
                // ==========================================

                string lrNoValue =
                    excelRow["LR NO"]?.ToString().Trim();

                int lrNo;

                if (int.TryParse(
                    lrNoValue,
                    out lrNo))
                {
                    row["LRNo"] = lrNo;
                }
                else
                {
                    row["LRNo"] = DBNull.Value;
                }


                // ==========================================
                // JK
                // ==========================================

                row["JK"] =
                    excelRow["JK"]?.ToString().Trim()
                    ?? "";


                // ==========================================
                // CUSTOMER NAME
                // ==========================================

                row["CustomerName"] =
                    excelRow["CUSTOMER NAME"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // TO LOCATION
                // ==========================================

                row["ToLocation"] =
                    excelRow["TO LOCATION"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // VEHICLE NO
                // ==========================================

                row["VehicleNo"] =
                    excelRow["VEHICLE NO"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // VEHICLE TYPE
                // ==========================================

                row["VehicleType"] =
                    excelRow["VEHICLE TYPE"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // NO OF TYRES
                // ==========================================

                string noOfTyresValue =
                    excelRow["NO OF TYRES"]
                    ?.ToString()
                    .Trim();

                int noOfTyres;

                if (int.TryParse(
                    noOfTyresValue,
                    out noOfTyres))
                {
                    row["NoOfTyres"] =
                        noOfTyres;
                }
                else
                {
                    row["NoOfTyres"] =
                        DBNull.Value;
                }


                // ==========================================
                // WEIGHT
                // ==========================================

                string weightValue =
                    excelRow["WEIGHT"]
                    ?.ToString()
                    .Trim();

                decimal weight;

                if (decimal.TryParse(
                    weightValue,
                    out weight))
                {
                    row["Weight"] =
                        weight;
                }
                else
                {
                    row["Weight"] =
                        DBNull.Value;
                }


                // ==========================================
                // INVOICE NO
                // ==========================================

                row["InvoiceNo"] =
                    excelRow["INVOICE NO"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // BEAT CYCLE TIMING
                // ==========================================

                row["BeatCycleTiming"] =
                    excelRow["BEAT CYCLE TIMING"]
                    ?.ToString()
                    .Trim()
                    ?? "";


                // ==========================================
                // TYRE
                // ==========================================

                string tyreValue =
                    excelRow["Tyre"]
                    ?.ToString()
                    .Trim();

                int tyre;

                if (int.TryParse(
                    tyreValue,
                    out tyre))
                {
                    row["Tyre"] =
                        tyre;
                }
                else
                {
                    row["Tyre"] =
                        DBNull.Value;
                }


                // ==========================================
                // RATE 1
                // ==========================================

                string rate1Value =
                    excelRow["Rate 1"]
                    ?.ToString()
                    .Trim();

                decimal rate1;

                if (decimal.TryParse(
                    rate1Value,
                    out rate1))
                {
                    row["Rate1"] =
                        rate1;
                }
                else
                {
                    row["Rate1"] =
                        DBNull.Value;
                }


                // ==========================================
                // UPLOAD ID
                // ==========================================

                row["UploadId"] =
                    uploadId;



                sqlTable.Rows.Add(row);
            }


            // ==========================================
            // 6. INSERT DATA INTO DATABASE
            // ==========================================

            freightDAL.InsertFreightData(
                sqlTable);

            // ==========================================
            // CALCULATE TYRE + RATE1
            // RATE3 WILL CALCULATE AUTOMATICALLY
            // ==========================================

            freightDAL.CalculateFreightRate(uploadId);


            // ==========================================
            // 7. RETURN UPLOAD ID
            // ==========================================

            return uploadId;
        }
    }
}