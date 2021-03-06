﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace Export.Models
{
    public class import
    {
        RowData dt = new RowData();   
        public void BalanceSheet( DataSet ds,int j,string fileName)
        {
            string[] sub = new string[] { "propertyandassets", "liabilitiesandcapital","assets", "equityandliabilities" };
            string[] sub2 = new string[] {"cash","noncurrentassets","currentassets","balancewithotherbanksandfinancialinstitutions",
                                    "moneyatcallandonshortnotice","currentliabilities","noncurrentliabilities","investments","loansandadvances/investments","totalassets",
                                    "fixedassetsincludingpremises,furnitureandfixtures","otherassets","non-bankingassets",
                                    "liabilities","borrowingsfromotherbanks,financialinstitutionsandagents","abbanksubordinatedbond",
                                    "depositsandotheraccounts","otherliabilities","shareholdersequity","shareholders’equity",
                                    "equityattributable\ntoequityholdersoftheparentcompany","non-controllinginterest",
                                    "netassetsvaluepershare","sharestocalculatenavps"};
            string[] Qyear = new string[] {"42916","43008", "43190","43281","43373","43555","43646","43738","30-Jun-17","30-Sep-17","31-Mar-18","30-Jun-18","30-Sep-18","31-Mar-19","30-Jun-19","30-Sep-19"};

            string head1 = string.Empty;
            string head2 = string.Empty;
            string head3 = string.Empty;

            int fRow = 0;
            string company = ds.Tables[j].Columns[0].ToString().Replace("'", string.Empty).Trim();
            bool YearlyBlc = false;
            bool QBlc = false;            
            for (int row = 0; row < ds.Tables[j].Rows.Count; row++)
            {
                dt = new RowData();
                string check = ds.Tables[j].Rows[row][0].ToString().Replace("'", string.Empty).Trim();
                SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
                string query1 = string.Empty;
                for (int col = 0; col < ds.Tables[j].Columns.Count; col++)
                {
                    var data = ds.Tables[j].Rows[row][col].ToString().Replace(",",string.Empty).Trim();
                    if (data==string.Empty)
                    {
                        continue;
                    }
                    if (data == "2018" || data == "2017" || data == "2016" || data == "2015" || data == "2014" || data == "2013")
                    {
                        fRow = row;
                        YearlyBlc = true;
                        break;
                    }
                    if (Qyear.Contains(data))
                    {
                        fRow = row;
                        QBlc = true;
                        break;
                    }
                    if (sub.Contains(string.Concat(check.Where(c=>!char.IsWhiteSpace(c))).ToLower()))
                    {
                        head1 = check;
                        head2 = string.Empty;
                        head3 = string.Empty;                        
                    }
                    else if (sub2.Contains(string.Concat(check.Where(c => !char.IsWhiteSpace(c))).ToLower()))
                    {
                        head2 = check;
                        head3 = string.Empty;                        
                    }
                    else
                    {
                        head3 = check;                        
                    }
                    var excelYear = ds.Tables[j].Rows[fRow][col].ToString();                    
                   dt= InsertData(excelYear, dt, data);
                }
                if (fRow==0|| fRow==row)
                {
                    continue;
                }

                if (YearlyBlc==true)
                {
                    query1 = "Insert into BalanceSheets(Head1,head2,Head3,Y2011,Y2012,Y2013,Y2014,Y2015,Y2016,Y2017,Y2018,CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + head3 + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (QBlc==true)
                {
                    query1 = "Insert into QBalanceSheet(Head1,Head2,Head3,[30Jun17],[30Sep17],[31Mar18],[30Jun18],[30Sep18],[31Mar19],[30Jun19],[30Sep19],CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + head3 + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (head2 == "Shares to calculate NAVPS")
                {
                    break;
                }
            }
        }
        public void IncomeStatement(DataSet ds , int j,string fileName)
        {
            string[] sub = new string[] {"operatingincome","netinterestincome/netprofitoninvestments",
                                "operatingexpense","operatingprofit","totalprovisions","lossondisposalofabexchange(uk)ltd.",
                                "profitbeforetaxation","provisionfortaxation","netprofit","earningspershare(parvaluetaka10)","sharestocalculateeps"};

            string[] Qyear = new string[] { "42916", "43008", "43190", "43281", "43373", "43555", "43646", "43738", "30-Jun-17", "30-Sep-17", "31-Mar-18", "30-Jun-18", "30-Sep-18", "31-Mar-19", "30-Jun-19", "30-Sep-19" };
            string head1 = string.Empty;
            string head2 = string.Empty;
            int fRow = 0;
            bool YearlyBlc = false;
            bool Qblc = false;
            for (int row = 0; row < ds.Tables[j].Rows.Count; row++)
            {
                string company = ds.Tables[j].Columns[0].ToString().Replace("'", string.Empty).Trim();
                string check = ds.Tables[j].Rows[row][0].ToString().Replace("'", string.Empty).Trim();
                SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
                string query1 = string.Empty;
                dt = new RowData();

                //check insert
                for (int col = 0; col < ds.Tables[j].Columns.Count; col++)
                {
                    var data = ds.Tables[j].Rows[row][col].ToString();
                    var excelYear = ds.Tables[j].Rows[fRow][col].ToString();

                    if (data==string.Empty)
                    {
                        continue;
                    }
                    if (data == "2018" || data == "2017" || data == "2016" || data == "2015" || data == "2014" || data == "2013")
                    {
                        fRow = row;
                        YearlyBlc = true;
                        break;
                    }
                    if (Qyear.Contains(data))
                    {
                        fRow = row;
                        Qblc = true;
                        break;
                    }
                    else if (sub.Contains(string.Concat(check.Where(c => !char.IsWhiteSpace(c))).ToLower()))
                    {
                        head1 = check;
                        head2 = string.Empty;
                    }
                    else
                    {
                        head2 = check;
                    }

                    dt = InsertData(excelYear, dt, data);
                }
                 if (fRow == 0 || fRow == row)
                {
                    continue;
                }
                if (YearlyBlc==true)
                {
                    query1 = "Insert into IncomeStates(Head1,Head2,Y2014,Y2015,Y2016,Y2017,Y2018,CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (Qblc==true)
                {
                    query1 = "Insert into QIncomeState(Head1,Head2,[30Jun17],[30Sep17],[31Mar18],[30Jun18],[30Sep18],[31Mar19],[30Jun19],[30Sep19],CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (head1== "Shares to Calculate EPS")
                {
                    break;
                }
            }
        }
        public void CashFlow(DataSet ds,int j,string fileName)
        {
            string[] sub = new string[] {"netcashflows-operatingactivities","operatingprofitbeforechangesinoperatingassetsandliabilities",   "increase/(decrease)inoperatingassetsandliabilities","netcashflows-investmentactivities","netcashflows-financingactivities",   "netchangeincashflows","cashandcashequivalentsatbeginningperiod","cashandcashequivalentsatendofperiod",
                            "netoperatingcashflowpershare","effectsofexchangeratechangesoncashandcashequivalents",              "adjustmentofdisposalofbits","sharestocalculatenocfps"};
            string[] Qyear = new string[] { "42916", "43008", "43190", "43281", "43373", "43555", "43646", "43738", "30-Jun-17", "30-Sep-17", "31-Mar-18", "30-Jun-18", "30-Sep-18", "31-Mar-19", "30-Jun-19", "30-Sep-19" };
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            string company = string.Empty;
            string head1 = string.Empty;
            string head2 = string.Empty;
            int coln = 0;
            int fRow = 0;
            bool YearlyBlc = false;
            bool Qblc = false;
            for (int row = 0; row < ds.Tables[j].Rows.Count; row++)
            {

                dt = new RowData();
                string query1 = string.Empty;
                for (int col = 0; col < ds.Tables[j].Columns.Count; col++)
                {
                    var data = ds.Tables[j].Rows[row][col].ToString();
                    var excelYear = ds.Tables[j].Rows[fRow][col].ToString();
                    if (data == string.Empty)
                    {
                        continue;
                    }
                    if (data == "2018" || data == "2017" || data == "2016" || data == "2015" || data == "2014" || data == "2013")
                    {
                        fRow = row;
                        YearlyBlc = true;
                        break;
                    }
                    if (Qyear.Contains(data))
                    {
                        fRow = row;
                        Qblc = true;
                        break;
                    }
                    if (company == string.Empty)
                    {
                        company = ds.Tables[j].Columns[col].ToString().Replace("'", string.Empty).Trim();
                        coln = col;
                    }
                    string check = ds.Tables[j].Rows[row][coln].ToString().Replace("'", string.Empty).Trim();
                    if (sub.Contains(string.Concat(check.Where(c => !char.IsWhiteSpace(c))).ToLower()))
                    {
                        head1 = check;
                        head2 = string.Empty;
                    }
                    else
                    {
                        head2 = check;
                    }
                    dt = InsertData(excelYear, dt, data);
                }
                if (fRow == 0 || fRow == row)
                {
                    continue;
                }
                if (YearlyBlc==true)
                {
                    query1 = "Insert into YCashflow(Head1,Head2,[Y2011],[Y2012],[Y2013],[Y2014],[Y2015],[Y2016],[Y2017],[Y2018],CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (Qblc==true)
                {
                    query1 = "Insert into QCashFlow(Head1,Head2,[30Jun17],[30Sep17],[31Mar18],[30Jun18],[30Sep18],[31Mar19],[30Jun19],[30Sep19],CompanyName,FileName) Values('" + head1 + "','" + head2 + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (head1== "Shares to Calculate NOCFPS")
                {
                    break;
                }
            }
        }
        public void Ratio(DataSet ds,int j,string fileName)
        {
            string[] Qyear = new string[] { "42916", "43008", "43190", "43281", "43373", "43555", "43646", "43738", "30-Jun-17", "30-Sep-17", "31-Mar-18", "30-Jun-18", "30-Sep-18", "31-Mar-19", "30-Jun-19", "30-Sep-19" };
            
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["db"].ConnectionString);
            string query1 = string.Empty;
            string company = string.Empty;            
            int fRow = 0;
            bool YearlyBlc = false;
            bool QBlc = false;
            for (int row = 0; row < ds.Tables[j].Rows.Count; row++)
            {
                dt = new RowData();
                company = ds.Tables[j].Columns[0].ToString().Replace("'", string.Empty).Trim();
                string check = ds.Tables[j].Rows[row][0].ToString().Replace("'", string.Empty).Trim();
                
                for (int col = 0; col < ds.Tables[j].Columns.Count; col++)
                {
                    var data = ds.Tables[j].Rows[row][col].ToString().Replace("%", string.Empty).Trim();
                    var excelYear = ds.Tables[j].Rows[fRow][col].ToString();
                    if (data == "2018"|| data == "2017"|| data == "2016"|| data == "2015"|| data == "2014"|| data == "2013")
                    {
                        fRow = row;
                        YearlyBlc = true;
                        break;
                    }
                    if (Qyear.Contains(data))
                    {
                        fRow = row;
                        QBlc = true;
                        break;
                    }
                    try
                    {
                        data= decimal.Round(decimal.Parse(data) * 100, 2).ToString();
                        dt = InsertData(excelYear, dt, data);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                if (fRow == 0 || row == fRow)
                {
                    continue;
                }
                if (YearlyBlc==true)
                {
                    query1 = "Insert into Ratios(Head1,Y2013,Y2014,Y2015,Y2016,Y2017,Y2018,CompanyName,FileName) Values('" + check + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (QBlc==true)
                {
                    query1 = "Insert into QRatio(Head1,[30Jun17],[30Sep17],[31Mar18],[30Jun18],[30Sep18],[31Mar19],[30Jun19],[30Sep19],CompanyName,FileName) Values('" + check + "','" + dt.fData1 + "','" + dt.fData2 + "','" + dt.fData3 + "','" + dt.fData4 + "','" + dt.fData5 + "','" + dt.fData6 + "','" + dt.fData7 + "','" + dt.fData8 + "','" + company + "','" + fileName + "')";
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query1, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
                if (check== "Advance/Investment to Deposit Ratio" || check== "Investment to Deposit Ratio")
                {
                    break;
                }
            }
        }
        public RowData InsertData(string excelYear,RowData dt,string data)
        {
            if (excelYear == "2011" || excelYear == "42916" || excelYear == "30-Jun-17")
            {
                dt.fData1 = data;
            }
            if (excelYear == "2012" || excelYear == "43008" || excelYear == "30-Sep-17")
            {
                dt.fData2 = data;
            }
            if (excelYear == "2013" || excelYear == "43190" || excelYear == "31-Mar-18")
            {
                dt.fData3 = data;
            }
            if (excelYear == "2014" || excelYear == "43281" || excelYear == "30-Jun-18")
            {
                dt.fData4 = data;
            }
            if (excelYear == "2015" || excelYear == "43373" || excelYear == "30-Sep-18")
            {
                dt.fData5 = data;
            }
            if (excelYear == "2016" || excelYear == "43555" || excelYear == "31-Mar-19")
            {
                dt.fData6 = data;
            }
            if (excelYear == "2017" || excelYear == "43646" || excelYear == "30-Jun-19")
            {
                dt.fData7 = data;
            }
            if (excelYear == "2018" || excelYear == "43738" || excelYear == "30-Sep-19")
            {
                dt.fData8 = data;
            }
            return dt;
        }
    }
}