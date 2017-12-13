using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.SqlClient;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;

// using

namespace WinFormCCCDataSet
{
    public partial class frmMainForm : Form
    {
        //General_Functions tgf = null;
        public frmMainForm()
        {
            InitializeComponent();
        }

        private void frmMainForm_Load(object sender, EventArgs e)
        {
            FormLoad();
        }

        private void FormLoad()
        {
            // Initial states of various controls.
            btnExportToExcel.Enabled = false; // Initial state.
            lblCurrentROProcessing.Text = string.Empty;
            lblROCounter.Text = string.Empty;
            lblExportRow.Text = string.Empty;
            progressBar1.Visible = false;
            lblRebuildIndexes.Text = string.Empty;

            //// Rebuild All DB Table Indexes in Target DB.
            //RebuildAllDBTableIndexes();

            tmrStartUpTasks.Interval = 3000;
            tmrStartUpTasks.Enabled = true;

            try
            {
                dateTimePickerStart.Value = Convert.ToDateTime(GetPostedDateFromExport04("Earliest"));
                dateTimePickerEnd.Value = Convert.ToDateTime(GetPostedDateFromExport04("Latest"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString() + ". Please quit this program and make sure the application config file is pointed to the right data source, and they are online and accessible.","Error Encountered!");
                btnDoWork.Enabled = false;
                btnExportToExcel.Enabled = false;
                btnGetLastMonthOnly.Enabled = false;
            }
        }

        private void RebuildAllDBTableIndexes()
        {
            using (SqlConnection conn = new SqlConnection(GetGeneralSqlConnectionString()))
            {
                conn.Open();
                string strSQL = string.Empty;
                SqlCommand cmd = null;

                strSQL = "usp_rebuild_all_DB_indexes";
                cmd = new SqlCommand(strSQL, conn);
                cmd.CommandType = CommandType.StoredProcedure;

                try
                {
                    SqlDataReader rdr = cmd.ExecuteReader();
                    lblRebuildIndexes.Text = "Indexes Rebuilt: " + DateTime.Now.ToString();
                }
                catch (Exception ex)
                {
                    string strEx = ex.Message.ToString();
                    lblRebuildIndexes.Text = "Indexes Rebuilt: " + strEx;
                }
            }
        }

        private void btnDoWork_Click(object sender, EventArgs e)
        {
            //tgf = new General_Functions();
            //DataTable dt = tgf.GetExport1DataTable();
            //dgvMasterDataSet.DataSource = dt;

            // Rebuild All DB Table Indexes in Target DB.
            RebuildAllDBTableIndexes();


            // First make sure dates are in the right order.  Start Date has to be earlier than End Date.
            if (dateTimePickerStart.Value <= dateTimePickerEnd.Value)
            {
                // OK, we are in good shape.
            }
            else
            {
                DateTime dtStart = dateTimePickerStart.Value;

                dateTimePickerStart.Value = dateTimePickerEnd.Value;
                dateTimePickerEnd.Value = dtStart;
            }

            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            DataTable dt = GetExport1DataTable();            
            dgvMasterDataSet.DataSource = dt;
            ChangeDataGridViewHeaderTextForBreakColumnsAndRenameOtherColumns();

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;

            if (dt.Rows.Count > 0)
            {
                btnExportToExcel.Enabled = true;                
            }
            else
            {
                btnExportToExcel.Enabled = false;
            }
            progressBar1.Visible = false;

            // Rebuild All DB Table Indexes in Target DB.
            RebuildAllDBTableIndexes();
        }

        private void ChangeDataGridViewHeaderTextForBreakColumnsAndRenameOtherColumns()
        {
            for (int j = 0; j < dgvMasterDataSet.Columns.Count; j++)
            {
                if (dgvMasterDataSet.Columns[j].HeaderText.StartsWith("Break "))
                {
                    dgvMasterDataSet.Columns[j].HeaderText = string.Empty;
                }
            }
            for (int j = 0; j < dgvMasterDataSet.Columns.Count; j++)
            {
                if (dgvMasterDataSet.Columns[j].HeaderText.StartsWith("Labor - "))
                {
                    dgvMasterDataSet.Columns[j].HeaderText = dgvMasterDataSet.Columns[j].HeaderText.Substring(8, dgvMasterDataSet.Columns[j].HeaderText.Length - 8);
                }
            }

            for (int j = 0; j < dgvMasterDataSet.Columns.Count; j++)
            {
                if (dgvMasterDataSet.Columns[j].HeaderText.StartsWith("Flagged Hours - "))
                {
                    dgvMasterDataSet.Columns[j].HeaderText = dgvMasterDataSet.Columns[j].HeaderText.Replace("Flagged Hours - ", "Labor - ");
                }
            }
        }

        private static SqlConnection GetGeneralSqlConnection()
        {
            return new SqlConnection(GetGeneralSqlConnectionString());
        }

        // Get connection string to General DB Datasource.
        private static string GetGeneralSqlConnectionString()
        {
            string strSQLConnectString = ConfigurationManager.ConnectionStrings["AAB_GEN_DB"].ToString();
            return strSQLConnectString;
        }


        private DateTime? GetPostedDateFromExport04(string strType)
        {
            DateTime? dtReturnValue = null; 

            string strSQL0 = "SELECT DISTINCT [posted_date] FROM [aab_export04] ORDER BY [posted_date] ASC";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                if (strType == "Earliest")
                {
                    dtReturnValue = Convert.ToDateTime(dt.Rows[0][0]);
                }
                if (strType == "Latest")
                {
                    dtReturnValue = Convert.ToDateTime(dt.Rows[dt.Rows.Count - 1][0]).AddDays(0);
                }
            }
            return dtReturnValue;
        }

        public DataTable GetExport1DataTable()
        {
            string strBuildQueryString = string.Empty;
            DataTable dtaccount_type_nameaccount_name = null;  // For account_type_name and account_name in "if" clause below.
            DataTable dt_sales_item = null;  // For sales_items in "else" clause below.

            if (chkGroupBySalesItem.Checked == false)
            {
                string strSQL0 = string.Empty;
                //strSQL0 = "SELECT DISTINCT [account_type_name],[account_name] FROM [aab_export01] order by [account_type_name], [account_name]";
                strSQL0 = "SELECT DISTINCT [account_type_name],[account_name] FROM [aab_consolidated_all_accounts] order by [account_type_name], [account_name]";  // Get from [aab_consolidated_all_accounts] table now.
                SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
                SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
                dtaccount_type_nameaccount_name = new DataTable();
                da0.Fill(dtaccount_type_nameaccount_name);

                strBuildQueryString = string.Empty;
                if (dtaccount_type_nameaccount_name.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dtaccount_type_nameaccount_name.Rows)
                    {
                        // Example: '' as 'Labor - Aluminum Body', 
                        strBuildQueryString += "'' as '" + dr1["account_type_name"].ToString() + " - " + dr1["account_name"].ToString() + "', ";
                    }
                }

            }
            else // (chkGroupBySalesItem.Checked == true)
            {
                string strSQL0 = string.Empty;
                strSQL0 = "SELECT DISTINCT [sales_item] FROM [aab_consolidated_all_accounts] WHERE [sales_item] IS NOT NULL ORDER BY [sales_item]";  // Get from [aab_consolidated_all_accounts] table now.
                SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
                SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
                dt_sales_item = new DataTable();
                da0.Fill(dt_sales_item);

                strBuildQueryString = string.Empty;
                if (dt_sales_item.Rows.Count > 0)
                {
                    foreach (DataRow dr1 in dt_sales_item.Rows)
                    {
                        // Example: '' as 'Adjustments/Discounts', 
                        strBuildQueryString += "'' as '" + dr1["sales_item"].ToString() + "', ";
                    }
                }
            }


            // Get Labor Account Names - to be used for Flagged Hours
            string strSQL1 = string.Empty;
            //strSQL1 = "SELECT DISTINCT [labor_account_name] FROM [aab_export03] ORDER BY [labor_account_name] ASC";  // Get unique Labor Account Names from Export 03 - to be used for Flagged Hours
            //strSQL1 = "SELECT DISTINCT [account_name] as 'labor_account_name' FROM [aab_consolidated_all_accounts] ORDER BY [account_name] ASC";  // Get from [aab_consolidated_all_accounts] table now.
            strSQL1 = "SELECT DISTINCT [account_name] as 'labor_account_name' FROM [aab_consolidated_all_accounts] WHERE [account_type_name] = 'Labor' ORDER BY [account_name] ASC";
            SqlCommand cmd1 = new SqlCommand(strSQL1, GetGeneralSqlConnection());
            SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
            DataTable dtExport03FlaggedHours = new DataTable();
            da1.Fill(dtExport03FlaggedHours);
            
            string strBuildQueryString2 = string.Empty;
            if (dtExport03FlaggedHours.Rows.Count > 0)
            {
                foreach (DataRow dr2 in dtExport03FlaggedHours.Rows)
                {
                    // Example: '' as 'Flagged Hours - Aluminum Body', 
                    strBuildQueryString2 += "'' as 'Flagged Hours - " + dr2["labor_account_name"].ToString() + "', ";
                }
            }

            // Get Labor Account Names - to be used for Sales Rate Amount
            string strSQL2 = string.Empty;
            //strSQL2 = "SELECT DISTINCT [labor_account_name] FROM [aab_export03] ORDER BY [labor_account_name] ASC";
            //strSQL2 = "SELECT DISTINCT [account_name] as 'labor_account_name' FROM [aab_consolidated_all_accounts] ORDER BY [account_name] ASC";  // Get from [aab_consolidated_all_accounts] table now.
            strSQL2 = "SELECT DISTINCT [account_name] as 'labor_account_name' FROM [aab_consolidated_all_accounts] WHERE [account_type_name] = 'Labor' ORDER BY [account_name] ASC";  // Get from [aab_consolidated_all_accounts] table now.
            SqlCommand cmd2 = new SqlCommand(strSQL2, GetGeneralSqlConnection());
            SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
            DataTable dtExport03SalesRateAmount = new DataTable();
            da2.Fill(dtExport03SalesRateAmount);

            string strBuildQueryString3 = string.Empty;
            if (dtExport03FlaggedHours.Rows.Count > 0)
            {
                foreach (DataRow dr2 in dtExport03SalesRateAmount.Rows)
                {
                    // Example: '' as 'Aluminum Body', 
                    strBuildQueryString3 += "'' as 'Rate - " + dr2["labor_account_name"].ToString() + "', ";
                }
            }

            //string strSQL = "SELECT DISTINCT TOP 25 A04.[posted_date] as 'Posted Date', A01.[repair_order_number] as 'RO Number', " +
            string strSQL = "SELECT DISTINCT A04.[posted_date] as 'Closed Date', A01.[repair_order_number] as 'RO Number', " +
                "'' as 'Location', " +
                "'' as 'Estimator', " +
                "'' as 'Insurance Company', " +
                "'' as 'Body Technician', " +
                "'' as 'Frame Technician', " +
                "'' as 'Mechanical Technician', " +
                "'' as 'Paint Technician', " +
                "'' as 'Grand Total', " +

                strBuildQueryString +

                "'' as 'Break 1', " +

                "'' as 'Owner', " +
                "'' as 'Vehicle Year', " +
                "'' as 'Vehicle Make', " +
                "'' as 'Vehicle Model', " +

                "'' as 'Break 2', " +

                strBuildQueryString2 +

                "'' as 'Break 3', " +

                strBuildQueryString3 +

                "'' as 'Break 4', " +

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

                " FROM [aab_export01] A01 (NOLOCK) INNER JOIN [aab_export04] A04 ON A01.[repair_order_number] = A04.[repair_order_number] " +
                " WHERE (A04.[posted_date] BETWEEN @StartDate AND @EndDate)" +
                " ORDER BY A04.[posted_date] ASC, A01.[repair_order_number] ASC";
            SqlCommand cmd = new SqlCommand(strSQL, GetGeneralSqlConnection());
            cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = dateTimePickerStart.Value;
            cmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = dateTimePickerEnd.Value.AddDays(1).AddSeconds(-1);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            int intTotalROstoProcess = dt.Rows.Count;

            int intROCounter = 0;
            foreach (DataRow dr in dt.Rows)
            {
                // For testing only.
                //string strRONumberforTest = dr["RO Number"].ToString();
                //if (strRONumberforTest == "102969")
                //{
                //    int x = 1;
                //}

                dr["Location"] = GetLocationForRONumber(dr["RO Number"].ToString());
                dr["Estimator"] = GetEstimatorForRONumber(dr["RO Number"].ToString());
                dr["Insurance Company"] = GetInsuranceCompanyForRONumber(dr["RO Number"].ToString());

                dr["Body Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Body");
                dr["Frame Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Frame");
                dr["Mechanical Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Mechanical");
                dr["Paint Technician"] = TechnicianForRONumber(dr["RO Number"].ToString(), "Paint");

                dr["Grand Total"] =  GetEstimateGrossAmountForGrandTotalFromExport04(dr["RO Number"].ToString()).ToString("C", new CultureInfo("en-US"));

                if (chkGroupBySalesItem.Checked == false)
                {
                    if (dtaccount_type_nameaccount_name.Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dtaccount_type_nameaccount_name.Rows)
                        {
                            // Example: //dr["Labor - Aluminum Body"] = GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), "Labor", "Aluminum Body").ToString("0.##");
                            dr[dr1["account_type_name"].ToString() + " - " + dr1["account_name"].ToString()] = GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), dr1["account_type_name"].ToString(), dr1["account_name"].ToString()).ToString("C", new CultureInfo("en-US"));
                        }
                    }
                }
                else
                {
                    if (dt_sales_item.Rows.Count > 0)
                    {                        
                        foreach (DataRow dr1 in dt_sales_item.Rows)
                        {
                            decimal decSalesItemTotal = 0;

                            string strSQL09 = string.Empty;
                            strSQL09 = "SELECT DISTINCT [account_type_name],[account_name] FROM [aab_consolidated_all_accounts] WHERE [sales_item] = @SalesItem";  
                            SqlCommand cmd09 = new SqlCommand(strSQL09, GetGeneralSqlConnection());
                            cmd09.Parameters.Add("@SalesItem", SqlDbType.NVarChar).Value = dr1["sales_item"].ToString();
                            SqlDataAdapter da09 = new SqlDataAdapter(cmd09);
                            dtaccount_type_nameaccount_name = new DataTable();
                            da09.Fill(dtaccount_type_nameaccount_name);

                            if (dtaccount_type_nameaccount_name.Rows.Count > 0)
                            {
                                foreach (DataRow dr1inner in dtaccount_type_nameaccount_name.Rows)
                                {
                                    // Example: decSalesItemTotal += GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), "Labor", "Aluminum Body");
                                    decSalesItemTotal += GetOrangeSalesAmountFromExport01(dr["RO Number"].ToString(), dr1inner["account_type_name"].ToString(), dr1inner["account_name"].ToString());
                                }
                            }
                            //dr["Body Labor"] = decSalesItemTotal.ToString("C", new CultureInfo("en-US"));
                            dr[dr1["sales_item"].ToString()] = decSalesItemTotal.ToString("C", new CultureInfo("en-US"));
                        }                        
                    }
                }

                dr["Break 1"] = "";

                dr["Owner"] = GetOwner(dr["RO Number"].ToString()).ToUpper();
                dr["Vehicle Year"] = GetVehicleYear(dr["RO Number"].ToString());
                dr["Vehicle Make"] = GetVehicleMake(dr["RO Number"].ToString()).ToUpper();
                dr["Vehicle Model"] = GetVehicleModel(dr["RO Number"].ToString());

                dr["Break 2"] = "";

                if (dtExport03FlaggedHours.Rows.Count > 0)
                {
                    foreach (DataRow dr2 in dtExport03FlaggedHours.Rows)
                    {
                        dr["Flagged Hours - " + dr2["labor_account_name"].ToString()] = GetGreenValueToUseFromExport03(dr["RO Number"].ToString(), dr2["labor_account_name"].ToString()).ToString("0.0");
                    }
                }

                dr["Break 3"] = "";

                if (dtExport03SalesRateAmount.Rows.Count > 0)
                {
                    foreach (DataRow dr3 in dtExport03SalesRateAmount.Rows)
                    {
                        dr["Rate - " + dr3["labor_account_name"].ToString()] = GetSalesRateAmountFromExport03(dr["RO Number"].ToString(), dr3["labor_account_name"].ToString()).ToString("C", new CultureInfo("en-US"));
                    }
                }

                dr["Break 4"] = "";

                dr["Total Loss"] = GetExport05Columns(dr["RO Number"].ToString(), "is_total_loss");
                dr["Primary Referral"] = GetExport05Columns(dr["RO Number"].ToString(), "primary_referral_name");

                try
                {
                    dr["Days in Shop"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_out_days")).ToString("0.##");
                }
                catch
                {
                    dr["Days in Shop"] = "";
                }
                
                dr["Vehicle In"] = GetExport05Columns(dr["RO Number"].ToString(), "vehicle_in_datetime");
                dr["Repairs Started"] = GetExport05Columns(dr["RO Number"].ToString(), "repair_started_datetime");
                dr["Repairs Completed"] = GetExport05Columns(dr["RO Number"].ToString(), "repair_completed_datetime");

                try
                {
                    dr["In - Start"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_start_days")).ToString("0.##");
                }
                catch
                {
                    dr["In - Start"] = "";
                }
                
                try
                {
                    dr["In - Comp"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_complete_days")).ToString("0.##");
                }
                catch
                {
                    dr["In - Comp"] = "";
                }
                
                try
                {
                    dr["In - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "in_to_out_days")).ToString("0.##");
                }
                catch
                {
                    dr["In - Out"] = "";
                }
                
                try
                {
                    dr["Start - Comp"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "start_to_complete_days")).ToString("0.##");
                }
                catch
                {
                    dr["Start - Comp"] = "";
                }
                
                try
                {
                    dr["Start - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "start_to_out_days")).ToString("0.##");
                }
                catch
                {
                    dr["Start - Out"] = "";
                }
                
                try
                {
                    dr["Comp - Out"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "comp_to_out_days")).ToString("0.##");
                }
                catch
                {
                    dr["Comp - Out"] = "";
                }
                
                try
                {
                    dr["Labor Hours"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "labor_hours_assigned")).ToString("0.##");
                }
                catch
                {
                    dr["Labor Hours"] = "";
                }
                
                try
                {
                    dr["Labor per Day"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "labor_per_day")).ToString("0.##");
                }
                catch
                {
                    dr["Labor per Day"] = "";
                }
                
                try
                {
                    dr["RO Hours per Day"] = Convert.ToDecimal(GetExport05Columns(dr["RO Number"].ToString(), "repair_order_hours_per_day")).ToString("0.##");
                }
                catch
                {
                    dr["RO Hours per Day"] = "";
                }
                
               
                intROCounter += 1;
                //lblROCounter.Text = intROCounter.ToString();

                //Debug.WriteLine("Progress: " + intROCounter.ToString());
                progressBar1.Visible = true;

                decimal decProgressPercent = Convert.ToDecimal(intROCounter) / Convert.ToDecimal(intTotalROstoProcess);
                decProgressPercent = decProgressPercent * Convert.ToDecimal(100.0);
                progressBar1.Value = (int)decProgressPercent;
                lblCurrentROProcessing.Text = "Getting Row: " + intROCounter.ToString() + " of " + intTotalROstoProcess.ToString();
                lblROCounter.Text = progressBar1.Value.ToString() + "%";                
                Application.DoEvents();
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
                    //strReturnValue = strReturnValue.Substring(0, intIndexOfHyphen).Trim();
                    strReturnValue = strReturnValue.Substring(intIndexOfHyphen + 1, strReturnValue.Length - intIndexOfHyphen - 1).Trim();
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
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] (NOLOCK) where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Body'";
            }
            if (strTechType == "Frame")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] (NOLOCK) where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Frame'";
            }
            if (strTechType == "Mechanical")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] (NOLOCK) where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] = 'Mechanical'";
            }
            if (strTechType == "Paint")
            {
                strSQL0 = "SELECT [repair_order_number],[technician_display_name],[labor_account_name] FROM [aab_export03] (NOLOCK) where [repair_order_number] = '" + strRONumber + "' AND [labor_account_name] LIKE '%Paint%'";
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


        private decimal GetEstimateGrossAmountForGrandTotalFromExport04(string strRONumber)
        {
            decimal decReturnValue = 0;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT [estimate_gross_amount] FROM [aab_export04] (NOLOCK) WHERE repair_order_number = @RONumber";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    decReturnValue += Convert.ToDecimal(dr["estimate_gross_amount"]);
                }
            }
            return decReturnValue;
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

        private string GetLaborAccountsMappingForExport03(string strLaborAccountName)
        {
            string strReturnValue = string.Empty;
            string strSQL = "SELECT [xml_field_to_use] FROM [aab_labor_accounts_mapping_export03] (NOLOCK) WHERE [labor_account_name] = @LaborAccountName ";
            SqlCommand cmd = new SqlCommand(strSQL, GetGeneralSqlConnection());
            cmd.Parameters.Add("@LaborAccountName", SqlDbType.NVarChar).Value = strLaborAccountName;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dtReturnData = new DataTable();
            da.Fill(dtReturnData);

            if (dtReturnData.Rows.Count > 0)
            {
                strReturnValue = dtReturnData.Rows[0]["xml_field_to_use"].ToString();
            }
            else
            {
                strReturnValue = "sales_hours"; // default to sales_hours
            }

            return strReturnValue;
        }

        private decimal GetGreenValueToUseFromExport03(string strRONumber, string strLaborAccountName)
        {
            decimal decReturnValue = 0;

            string strSQL0 = string.Empty;
            //strSQL0 = "SELECT [assigned_hours_complete] FROM [aab_export03] (NOLOCK) WHERE [repair_order_number] = @RONumber AND [labor_account_name] = @LaborAccountName";
            //strSQL0 = "SELECT [sales_hours] FROM [aab_export03] (NOLOCK) WHERE [repair_order_number] = @RONumber AND [labor_account_name] = @LaborAccountName";

            strSQL0 = "SELECT " + GetLaborAccountsMappingForExport03(strLaborAccountName) + " FROM [aab_export03] (NOLOCK) WHERE [repair_order_number] = @RONumber AND [labor_account_name] = @LaborAccountName";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            cmd0.Parameters.Add("@LaborAccountName", SqlDbType.NVarChar).Value = strLaborAccountName;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    //decReturnValue += Convert.ToDecimal(dr["assigned_hours_complete"]);
                    //decReturnValue += Convert.ToDecimal(dr["sales_hours"]);
                    decReturnValue += Convert.ToDecimal(dr[0]);
                }
            }
            return decReturnValue;
        }


        private decimal GetSalesRateAmountFromExport03(string strRONumber, string strLaborAccountName)
        {
            decimal decReturnValue = 0;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT [sales_rate_amount] FROM [aab_export03] (NOLOCK) WHERE [repair_order_number] = @RONumber AND [labor_account_name] = @LaborAccountName";
            SqlCommand cmd0 = new SqlCommand(strSQL0, GetGeneralSqlConnection());
            cmd0.Parameters.Add("@RONumber", SqlDbType.NVarChar).Value = strRONumber;
            cmd0.Parameters.Add("@LaborAccountName", SqlDbType.NVarChar).Value = strLaborAccountName;
            SqlDataAdapter da0 = new SqlDataAdapter(cmd0);
            DataTable dt = new DataTable();
            da0.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    decReturnValue += Convert.ToDecimal(dr["sales_rate_amount"]);
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


        private string GetExport03Columns(string strRONumber, string strColumnToGet)
        {
            string strReturnValue = string.Empty;

            string strSQL0 = string.Empty;
            strSQL0 = "SELECT TOP 1 [" + strColumnToGet + "] FROM [aab_export03] (NOLOCK) WHERE [repair_order_number] = @RONumber";
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

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {
            // Now this is a green button.

            btnExportToExcel.Enabled = false;
            lblExportRow.Visible = true;

            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            Application.DoEvents();

            ExportToExcel();

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;

            btnExportToExcel.Enabled = true;
            lblExportRow.Visible = false;
            Application.DoEvents();
        }


        private void DataGridViewColumnInserts()
        {
            DataGridViewColumn dgvBlankColumnA = new DataGridViewColumn();
            dgvBlankColumnA.Name = "Blank A";
            dgvBlankColumnA.HeaderText = "";
            dgvBlankColumnA.CellTemplate = new DataGridViewTextBoxCell();
            dgvMasterDataSet.Columns.Insert(34, dgvBlankColumnA);

            DataGridViewColumn dgvBlankColumnB = new DataGridViewColumn();
            dgvBlankColumnB.Name = "Blank B";
            dgvBlankColumnB.HeaderText = "";
            dgvBlankColumnB.CellTemplate = new DataGridViewTextBoxCell();
            dgvMasterDataSet.Columns.Insert(51, dgvBlankColumnB);
        }

        private void ExportToExcel()
        {
            DataGridViewColumnInserts();

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Consolidated";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (int i = 0; i < dgvMasterDataSet.Rows.Count - 1; i++)
                {
                    lblExportRow.Text = "Exporting Row: " + cellRowIndex.ToString();
                    Application.DoEvents();

                    for (int j = 0; j < dgvMasterDataSet.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check. 
                        if (cellRowIndex == 1)
                        {
                            // Header information.
                            try
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvMasterDataSet.Columns[j].HeaderText;

                                if (cellColumnIndex >= 10 && cellColumnIndex <= 34)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);

                                if (cellColumnIndex >= 36 && cellColumnIndex <= 51)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Turquoise);

                                if (cellColumnIndex >= 53 && cellColumnIndex <= 76)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                                if (cellColumnIndex >= 78 && cellColumnIndex <= 81)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);

                                if (cellColumnIndex >= 83 && cellColumnIndex <= 106)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);

                                if (cellColumnIndex >= 108 && cellColumnIndex <= 131)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);

                                if (cellColumnIndex >= 133 && cellColumnIndex <= 147)
                                    worksheet.Cells[cellRowIndex, cellColumnIndex].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            }
                            catch
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = string.Empty;
                            }
                            
                            // First row of data.
                            try
                            {
                                worksheet.Cells[cellRowIndex + 1, cellColumnIndex] = dgvMasterDataSet.Rows[i].Cells[j].Value.ToString();

                                // For first column, format at date.
                                if (cellColumnIndex == 1)
                                {
                                    Microsoft.Office.Interop.Excel.Range range = worksheet.Cells[cellRowIndex + 1, cellColumnIndex] as Microsoft.Office.Interop.Excel.Range;
                                    range.NumberFormat = "MM/dd/yyyy";
                                }
                            }
                            catch
                            {
                                worksheet.Cells[cellRowIndex + 1, cellColumnIndex] = string.Empty;
                            }
                        }
                        else
                        {
                            // Every other row of data beyond row 1.
                            try
                            {
                                worksheet.Cells[cellRowIndex + 1, cellColumnIndex] = dgvMasterDataSet.Rows[i].Cells[j].Value.ToString();

                                // For first column, format at date.
                                if (cellColumnIndex == 1)
                                {
                                    Microsoft.Office.Interop.Excel.Range range = worksheet.Cells[cellRowIndex + 1, cellColumnIndex] as Microsoft.Office.Interop.Excel.Range;
                                    range.NumberFormat = "MM/dd/yyyy";
                                }
                            }
                            catch
                            {
                                worksheet.Cells[cellRowIndex + 1, cellColumnIndex] = string.Empty;
                            }
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Excel Spreadsheet export successful.","Export Complete");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

        }

        private void btnGetLastMonthOnly_Click(object sender, EventArgs e)
        {
            int intCurrentYear = DateTime.Now.Year;
            int intCurrentMonth = DateTime.Now.Month;

            if (intCurrentMonth == 1)
            {
                dateTimePickerStart.Value = new DateTime(intCurrentYear - 1, 12, 1);
                dateTimePickerEnd.Value = new DateTime(intCurrentYear, intCurrentMonth, 1).AddDays(-1);
            }
            else
            {
                dateTimePickerStart.Value = new DateTime(intCurrentYear, intCurrentMonth - 1, 1);
                dateTimePickerEnd.Value = new DateTime(intCurrentYear, intCurrentMonth, 1).AddDays(-1);
            }
        }

        private void tmrStartUpTasks_Tick(object sender, EventArgs e)
        {
            tmrStartUpTasks.Enabled = false;

            // Rebuild All DB Table Indexes in Target DB.
            RebuildAllDBTableIndexes();
        }

    }
}
