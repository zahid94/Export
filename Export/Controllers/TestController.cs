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
        import import = new import();
        // GET: Test
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(List<HttpPostedFileBase> files)
        {
            foreach (var file in files)
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
                            //ViewBag.Data = ds;
                        }

                        //insert data to sql
                        for (int j = 0; j < ds.Tables.Count; j++)
                        {
                            //if (ds.Tables[j].TableName == "'1$'")
                            //{

                            //    import.BalanceSheet(ds, j, file.FileName);
                            //}

                             if (ds.Tables[j].TableName == "'2$'")
                            {
                                import.IncomeStatement(ds, j, file.FileName);
                            }
                            //else if (ds.Tables[j].TableName == "'3$'")
                            //{
                            //    import.CashFlow(ds, j, file.FileName);

                            //}
                            //else if (ds.Tables[j].TableName == "Ratios$" || ds.Tables[j].TableName == "Ratio$")
                            //{
                            //    import.Ratio(ds, j, file.FileName);
                            //}
                        }

                    }
                    else
                    {
                        ViewBag.Error = "Please Upload Files in .xls, .xlsx format";

                    }

                }
            }

            return View();
        }

        [HttpGet]
        public ActionResult Test()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Test(string path)
        {
            foreach (string filePath in Directory.GetFiles(path, "*.xlsx"))
            {
                string fileName = Path.GetFileName(filePath);
                string extension = System.IO.Path.GetExtension(fileName).ToLower();
                DataSet ds = new DataSet();
                string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                ds = Utility.ConvertXSLXtoDataSet(filePath, connString);
                for (int j = 0; j < ds.Tables.Count; j++)
                {
                    if (ds.Tables[j].TableName == "'1$'")
                    {

                        import.BalanceSheet(ds, j, fileName);
                    }
                    //else if (ds.Tables[j].TableName == "'2$'")
                    //{
                    //    import.IncomeStatement(ds, j, fileName);
                    //}
                    //else if (ds.Tables[j].TableName == "'3$'")
                    //{
                    //    import.CashFlow(ds, j, fileName);

                    //}
                    //else if (ds.Tables[j].TableName == "Ratios$" || ds.Tables[j].TableName == "Ratio$")
                    //{
                    //    import.Ratio(ds, j, fileName);
                    //}
                }

            }
            return View();
        }
    }
}