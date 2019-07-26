using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace KursnaListaSAPB1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Izmjena kursne liste u toku. Molim sačekajte....");

            SAPbobsCOM.Company company;
            SAPbobsCOM.SBObob sBObob;
            SAPbobsCOM.Recordset recordset;

            company = new SAPbobsCOM.Company();

            company.Server = "SapServer";
            company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            company.CompanyDB = "SapDb";
            company.LicenseServer = "SapServer:30000";
            company.UserName = "user";
            company.Password = "pass";

            int result = company.Connect();

            string err = company.GetLastErrorDescription();

            DateTime danas = DateTime.Now;

            //DateTime danas = new DateTime(2018, 8, 19);

            string uri = "https://www.cbbh.ba/CurrencyExchange/GetXml?date=" + danas.Month.ToString() + " " + danas.Day.ToString() + " " + danas.Year.ToString();

            XDocument xDocument = XDocument.Load(uri);

            var tagovi = xDocument.Descendants("CurrencyExchangeItem")
                .Where(x => x.Elements("AlphaCode") != null && x.Elements("Middle") != null)
                .Select(x => new
                {
                    Valuta = x.Element("AlphaCode").Value,
                    Srednji = x.Element("Middle").Value
                });


            sBObob = (SAPbobsCOM.SBObob)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            recordset = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //recordset = sBObob.GetLocalCurrency();
            //recordset = sBObob.GetSystemCurrency();

            foreach (var tag in tagovi)
            {
                if (tag.Valuta.ToString() == "USD")
                {
                    sBObob.SetCurrencyRate("$", DateTime.Now, Convert.ToDouble(tag.Srednji), true);
                }

                if (tag.Valuta.ToString() == "CHF")
                {
                    sBObob.SetCurrencyRate("CHF", DateTime.Now, Convert.ToDouble(tag.Srednji), true);
                }

                if (tag.Valuta.ToString() == "GBP")
                {
                    sBObob.SetCurrencyRate("GBP", DateTime.Now, Convert.ToDouble(tag.Srednji), true);
                }

                if (tag.Valuta.ToString() == "EUR")
                {
                    sBObob.SetCurrencyRate("EUR", DateTime.Now, Convert.ToDouble(tag.Srednji), true);
                }
            }

            company.Disconnect();

            Marshal.FinalReleaseComObject(company);

            Console.WriteLine("Izmjena kursne liste završena!");
            Console.WriteLine("Izlazim iz programa....");

        }
    }
}
