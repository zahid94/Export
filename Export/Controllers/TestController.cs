using Export.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Export.Controllers
{
    public class TestController : Controller
    {
        // GET: Test
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            if (file.ContentLength > 0)
            {
                string extension = System.IO.Path.GetExtension(file.FileName).ToLower();                
                string connString = "";




                string[] validFileTypes = { ".xls", ".xlsx" };
                
                string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), file.FileName);
                if (!Directory.Exists(path1))
                {
                    Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                }
                if (validFileTypes.Contains(extension))
                {
                    DataSet ds = new DataSet();
                    if (System.IO.File.Exists(path1))
                    { System.IO.File.Delete(path1); }
                    file.SaveAs(path1);

                    //Connection String to Excel Workbook
                    if (extension.Trim() == ".xls")
                    {
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        ds = Utility.ConvertXSLXtoDataSet(path1, connString);
                        ViewBag.Data = ds;
                    }
                    else if (extension.Trim() == ".xlsx")
                    {
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        ds = Utility.ConvertXSLXtoDataSet(path1, connString);
                        ViewBag.Data = ds;
                    }                   

                    for (int j = 0; j < ds.Tables.Count; j++)
                    {
                        if (ds.Tables[j].TableName == "'Balance Sheet$'")
                        {

                            string[] sub = new string[] { "Property and Assets", "Liabilities and Capital" };
                            string[] sub2 = new string[] {"Cash","Balance with Other Banks and Financial Institutions",
                                    "Money at call and on short notice","Investments","Loans and Advances/Investments",
                                    "Fixed Assets including Premises, Furniture and Fixtures","Other Assets","Non-Banking Assets",
                                    "Liabilities","Borrowings from Other Banks, Financial Institutions and Agents","AB Bank Subordinated Bond",
                                    "Deposits and Other Accounts","Other Liabilities","Shareholders’ Equity",
                                    "Equity attributable to equity holders of the parent company","Non-controlling interest",
                                    "Net assets value per share","Shares to calculate NAVPS"};

                            var head1 = "";
                            var head2 = "";
                            string company = string.Empty;
                            for (int i = 0; i < ds.Tables[j].Rows.Count; i++)
                            {
                                
                                var check = ds.Tables[j].Rows[i][0].ToString();
                                string conn = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                                SqlConnection con = new SqlConnection(conn);
                                string query1 = "";
                                
                                if (i < 3)
                                {
                                    continue;
                                }
                                else if (sub.Contains(check))
                                {
                                    head1 = check;
                                    query1 = "Insert into TempFinanceData(HeadLayer1,HeadLayer2,HeadLayer3,[2011],[2012],[2013],[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + head1 + "','" + head2 + "','" + "" + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + ds.Tables[j].Rows[i][3].ToString() + "','" + ds.Tables[j].Rows[i][4].ToString() + "','" + ds.Tables[j].Rows[i][5].ToString() + "','" + ds.Tables[j].Rows[i][6].ToString() + "','" + ds.Tables[j].Rows[i][7].ToString() + "','" + ds.Tables[j].Rows[i][8].ToString() + "','" + company + "')";

                                }
                                else if (sub2.Contains(check))
                                {
                                    head2 = check;
                                    query1 = "Insert into TempFinanceData(HeadLayer1,HeadLayer2,HeadLayer3,[2011],[2012],[2013],[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + head1 + "','" + head2 + "','" + "" + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + ds.Tables[j].Rows[i][3].ToString() + "','" + ds.Tables[j].Rows[i][4].ToString() + "','" + ds.Tables[j].Rows[i][5].ToString() + "','" + ds.Tables[j].Rows[i][6].ToString() + "','" + ds.Tables[j].Rows[i][7].ToString() + "','" + ds.Tables[j].Rows[i][8].ToString() + "','" + company + "')";
                                }
                                else
                                {
                                    query1 = "Insert into TempFinanceData(HeadLayer1,HeadLayer2,HeadLayer3,[2011],[2012],[2013],[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + head1 + "','" + head2 + "','" + check + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + ds.Tables[j].Rows[i][3].ToString() + "','" + ds.Tables[j].Rows[i][4].ToString() + "','" + ds.Tables[j].Rows[i][5].ToString() + "','" + ds.Tables[j].Rows[i][6].ToString() + "','" + ds.Tables[j].Rows[i][7].ToString() + "','" + ds.Tables[j].Rows[i][8].ToString() + "','" + company + "')";
                                }
                                con.Open();
                                SqlCommand cmd = new SqlCommand(query1, con);
                                cmd.ExecuteNonQuery();
                                con.Close();
                            }
                        }

                        //else if (ds.Tables[j].TableName == "'Income Statement$'")
                        //{
                        //    string[] sub = new string[] {"Operating Income","Net interest income/net profit on investments",
                        //        "Operating profit","Operating Profit","Total Provisions","Loss on Disposal of AB Exchange (UK) Ltd.",
                        //        "Profit Before Taxation","Provision for Taxation","Net Profit","Earnings per share (par value Taka 10)","Shares to Calculate EPS"};

                        //    var head1 = "";
                        //    for (int i = 0; i < ds.Tables[j].Rows.Count; i++)
                        //    {
                        //        string company = "";
                        //        var check = ds.Tables[j].Rows[i][0].ToString();
                        //        string conn = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                        //        SqlConnection con = new SqlConnection(conn);
                        //        string query1 = "";

                        //        //check insert
                        //        if (i < 3)
                        //        {
                        //            continue;
                        //        }
                        //        else if (sub.Contains(check))
                        //        {
                        //            head1 = check;
                        //            query1 = "Insert into IncomeState(Head1,Head2,[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + head1 + "','" + "" + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + ds.Tables[j].Rows[i][3].ToString() + "','" + ds.Tables[j].Rows[i][4].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + company + "')";

                        //        }
                        //        else
                        //        {
                        //            query1 = "Insert into IncomeState(Head1,Head2,[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + head1 + "','" + check + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','"+ ds.Tables[j].Rows[i][3].ToString() + "','"+ ds.Tables[j].Rows[i][4].ToString() + "','"+ ds.Tables[j].Rows[i][2].ToString() + "','" + company + "')";
                        //        }

                        //        con.Open();
                        //        SqlCommand cmd = new SqlCommand(query1, con);
                        //        cmd.ExecuteNonQuery();
                        //        con.Close();
                        //    }
                        //}
                        else if (ds.Tables[j].TableName == "Cashflow$")
                        {
                            
                            string[] sub = new string[] {"Net Cash Flows - Operating Activities","Operating profit before changes in operating assets and liabilities","Increase / (decrease) in operating assets and liabilities",
                            "Net Cash Flows - Investment Activities","Net Cash Flows - Financing Activities","Net Change in Cash Flows","Cash and Cash Equivalents at Beginning Period","Cash and Cash Equivalents at End of Period",
                            "Net Operating Cash Flow Per Share","Shares to Calculate NOCFPS"};
                            string conn = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                            SqlConnection con = new SqlConnection(conn);
                            string company = "";
                            var head1 = "";
                            for (int i = 0; i < ds.Tables[j].Rows.Count; i++)
                            {
                                var check = ds.Tables[j].Rows[i][0].ToString();
                                string query1 = "";
                                if (i<3)
                                {
                                    continue;
                                }
                                else if (sub.Contains(check))
                                {
                                    head1 = check;
                                    query1= "Insert into Cashflow(Head1,Head2,[2017],[2018],CompanyName) Values('" + head1 + "','" + "" + "','"  + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + company + "')";

                                }
                                else
                                {
                                    query1 = "Insert into Cashflow(Head1,Head2,[2017],[2018],CompanyName) Values('" + head1 + "','" + check + "','" + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','" + company + "')";
                                }

                                con.Open();
                                SqlCommand cmd = new SqlCommand(query1, con);
                                cmd.ExecuteNonQuery();
                                con.Close();
                            }
                        }
                        else if (ds.Tables[j].TableName == "Ratios$")
                        {
                            for (int i = 0; i < ds.Tables[j].Rows.Count; i++)
                            {
                                var check = ds.Tables[j].Rows[i][0].ToString();
                                string conn = ConfigurationManager.ConnectionStrings["db"].ConnectionString;
                                SqlConnection con = new SqlConnection(conn);
                                string query1 = "";
                                string company = "";
                                if (i < 3)
                                {
                                    continue;
                                }                                
                                else
                                {
                                    query1 = "Insert into Ratios(Head,[2013],[2014],[2015],[2016],[2017],[2018],CompanyName) Values('" + check + "','"  + ds.Tables[j].Rows[i][1].ToString() + "','" + ds.Tables[j].Rows[i][2].ToString() + "','"+ ds.Tables[j].Rows[i][3].ToString() + "','"+ ds.Tables[j].Rows[i][4].ToString() + "','"+ ds.Tables[j].Rows[i][5].ToString() + "','" + ds.Tables[j].Rows[i][6].ToString() + "','" + company + "')";
                                }

                                con.Open();
                                SqlCommand cmd = new SqlCommand(query1, con);
                                cmd.ExecuteNonQuery();
                                con.Close();
                            }
                        }
                        
                    }

                }
                else
                {
                    ViewBag.Error = "Please Upload Files in .xls, .xlsx or .csv format";

                }

            }

            return View();
        }
    }
}