using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;

namespace WinFormCCCDataSet
{
    public class General_Functions
    {        
        private static SqlConnection GetGeneralSqlConnection()
        {
            return new SqlConnection(GetGeneralSqlConnectionString());
        }

        // Get connection string to General DB Datasource.
        private static string GetGeneralSqlConnectionString()
        {
            string strSQLConnectString = ConfigurationManager.ConnectionStrings["AAB_GEN_DB"].ToString();
            return strSQLConnectString.ToString();
        }


        public DataTable GetExport1DataTable()
        {
            string strSQL0 = string.Empty;
            strSQL0 = "SELECT DISTINCT [account_type_name],[account_name] FROM [aab_export01] order by [account_type_name], [account_name]";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            //cmd.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dtaccount_type_nameaccount_name = new DataTable();
            da0.Fill(dtaccount_type_nameaccount_name);

            string strBuildQueryString = string.Empty;
            if (dtaccount_type_nameaccount_name.Rows.Count > 0)
            {
                foreach (DataRow dr1 in dtaccount_type_nameaccount_name.Rows)
                {
                    // Example: '' as 'Labor - Aluminum Body', 
                    strBuildQueryString += "'' as '" + dr1["account_type_name"].ToString() + " - " + dr1["account_name"].ToString() + "', ";
                }
            }

            string strSQL = "SELECT DISTINCT TOP 1000 A04.[posted_date] as 'Posted Date', A01.[repair_order_number] as 'RO Number', " +
                "'' as 'Location', " +
                "'' as 'Estimator', " +
                "'' as 'Insurance Company', " +
                "'' as 'Body Technician', " +
                "'' as 'Frame Technician', " +
                "'' as 'Mechanical Technician', " +
                "'' as 'Paint Technician', " +
                "'' as 'Grand Total', " +

                strBuildQueryString +

                "'' as 'EOL/BREAK', " +

                "'' as 'Owner', " +
                "'' as 'Vehicle Year', " +
                "'' as 'Vehicle Make', " +
                "'' as 'Vehicle Model', " +

                "'' as 'Total Loss', " +
                "'' as 'Primary Referral', " +
                "'' as 'Days in Shop', " +
                "'' as 'Vehicle In', " +
                "'' as 'Repairs Started', " +
                "'' as 'Repairs Completed', " +
                "'' as 'In - Start', " +
                "'' as 'In - Comp', " +
                "'' as 'In - Out', " +
                "'' as 'Start - Comp', " +
                "'' as 'Start - Out', " +
                "'' as 'Comp - Out', " +
                "'' as 'Labor Hours', " +
                "'' as 'Labor per Day', " +
                "'' as 'RO Hours per Day' " +

                " FROM [aab_export01] A01 (NOLOCK) INNER JOIN [aab_export04] A04 ON A01.[repair_order_number] = A04.[repair_order_number] ORDER BY A01.[repair_order_number] DESC";
            SqlCommand cmd = new SqlCommand(strSQL, GetGeneralSqlConnection());
            //cmd.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            int intTotalROstoProcess = dt.Rows.Count;

            foreach (DataRow dr in dt.Rows)
            {
                dr["Location"] = GetLocationForRONumber(dr["RO Number"].ToString());
                dr["Estimator"] = GetEstimatorForRONumber(dr["RO Number"].ToString());
                dr["Insurance Company"] = GetInsuranceCompanyForRONumber(dr["RO Number"].ToString());

                dr["Body Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(),"Body");
                dr["Frame Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Frame");
                dr["Mechanical Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Mechanical");
                dr["Paint Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Paint");

                if (dtaccount_type_nameaccount_name.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dtaccount_type_nameaccount_name.Rows)
                    {
                        // Example: //dr["Labor - Aluminum Body"] = GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), "Labor", "Aluminum Body").ToString("0.##");
                        dr[dr1["account_type_name"].ToString() + " - " + dr1["account_name"].ToString()] = GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), dr1["account_type_name"].ToString(), dr1["account_name"].ToString()).ToString("0.##");
                    }
                }

                dr["EOL/BREAK"] = "";
                dr["Owner"] = GetOwner(dr["RO Number"].ToString());
                dr["Vehicle Year"] = GetVehicleYear(dr["RO Number"].ToString());
                dr["Vehicle Make"] = GetVehicleMake(dr["RO Number"].ToString());
                dr["Vehicle Model"] = GetVehicleModel(dr["RO Number"].ToString());

                dr["Total Loss"] = GetExport05Columns(dr["RO Number"].ToString(), "is_total_loss");
                dr["Primary Referral"] = GetExport05Columns(dr["RO Number"].ToString(), "primary_referral_name");
                dr["Days in Shop"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_out_days")).ToString("0.##");
                dr["Vehicle In"] = GetExport05Columns(dr["RO Number"].ToString(), "vehicle_in_datetime");
                dr["Repairs Started"] = GetExport05Columns(dr["RO Number"].ToString(), "repair_started_datetime");
                dr["Repairs Completed"] = GetExport05Columns(dr["RO Number"].ToString(), "repair_completed_datetime");
                dr["In - Start"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_start_days")).ToString("0.##");
                dr["In - Comp"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_complete_days")).ToString("0.##");
                dr["In - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_out_days")).ToString("0.##");
                dr["Start - Comp"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "start_to_complete_days")).ToString("0.##");
                dr["Start - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "start_to_out_days")).ToString("0.##");
                dr["Comp - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "comp_to_out_days")).ToString("0.##");
                dr["Labor Hours"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "labor_hours_assigned")).ToString("0.##");
                dr["Labor per Day"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "labor_per_day")).ToString("0.##");
                dr["RO Hours per Day"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "repair_order_hours_per_day")).ToString("0.##");

            }
            return dt;
        }

        private string GetLocationForRONumber(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [repair_facility_name] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["repair_facility_name"].ToString();
                strReturnValue = strReturnValue.Trim();
                if (strReturnValue.Contains("-"))
                {
                    int intIndexOfHyphen = strReturnValue.IndexOf("-");
                    strReturnValue = strReturnValue.Substring(0, intIndexOfHyphen).Trim();
                }
            }
            return strReturnValue;
        }

        private string GetEstimatorForRONumber(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [service_writer_display_name] FROM [aab_export04] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["service_writer_display_name"].ToString().Trim();
            }
            return strReturnValue;
        }

        private string GetInsuranceCompanyForRONumber(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [master_carrier_name] FROM [aab_export05] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["master_carrier_name"].ToString().Trim();
            }
            return strReturnValue;
        }

        private string TechnicianForRONumber(string strRONumber, string strTechType)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            if (strTechType == "Body")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Body'";
            }
            if (strTechType == "Frame")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Frame'";
            }
            if (strTechType == "Mechanical")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Mechanical'";
            }
            if (strTechType == "Paint")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] LIKE '%Paint%'";
            }
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            //cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            //cmd0.Parameters.Add("@LaborAccountName", SqlDbType.NVarChar).Value = "'%" + strTechType + "%'";
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string strTemp = dt.Rows[0]["technician_display_name"].ToString().Trim();
                    if (strTemp.Length > 0)
                    {
                        strReturnValue = strTemp;
                    }
                }
            }
            return strReturnValue;
        }

        private decimal GetOrangeSalesAmountFromExport01(string strRONumber, string strAccountTypeName, string strAccountName)
        {
            decimal decReturnValue = 0;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT [sales_amount] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber AND [account_type_name] = @AccountTypeName AND [account_name] = @AccountName";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            cmd0.Parameters.Add("@AccountTypeName", SqlDbType.NVarChar).Value = strAccountTypeName;
            cmd0.Parameters.Add("@AccountName", SqlDbType.NVarChar).Value = strAccountName;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);
            
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    decReturnValue += Convert.ToDecimal(dr["sales_amount"]);
                }
            }
            return decReturnValue;
        }

        private string GetOwner(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [owner_name] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["owner_name"].ToString().Trim();
            }
            return strReturnValue;
        }

        private string GetVehicleYear(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [vehicle_year_make_model] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["vehicle_year_make_model"].ToString().Trim();
                if (strReturnValue.Length >= 4)
                {
                    strReturnValue = strReturnValue.Substring(0, 4); // Get Year, as it is the first 4 characters of this string.
                }
            }

            int value;
            if (int.TryParse(strReturnValue, out value))
            {
                // It's an int
            }
            else
            {
                // No it's not.
                strReturnValue = string.Empty;
            }

            return strReturnValue;
        }

        private string GetVehicleMake(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [vehicle_make_name] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["vehicle_make_name"].ToString().Trim();
            }
            return strReturnValue;
        }

        private string GetVehicleModel(string strRONumber)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [vehicle_year_make_model] FROM [aab_export01] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0]["vehicle_year_make_model"].ToString().Trim();
                string[] words = strReturnValue.Split(' ');

                int intLoopCounter = 0;
                strReturnValue = string.Empty;
                foreach (string word in words)
                {
                    intLoopCounter += 1;
                    if (intLoopCounter < 3)
                        continue;
                    strReturnValue += word + " ";
                }
            }
            return strReturnValue.Trim();
        }


        private string GetExport05Columns(string strRONumber, string strColumnToGet)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [" + strColumnToGet + "] FROM [aab_export05] (NOLOCK) WHERE [repair_order_number] = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                strReturnValue = dt.Rows[0][0].ToString();
            }
            return strReturnValue;
        }





    }
}