using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using RandomQuoteGenerator.Models;

namespace RandomQuoteGenerator.Controllers
{
    public class QuoteController : Controller
    {
        public ExcelPackage _package;
        public List<string> listQuotes;
        public Random randomNumberGenerator;
        public QuoteManager objQuoteManager;
        // GET: Quote
        public ActionResult Index()
        {
            fillList();
            ViewBag.list = listQuotes;
            objQuoteManager = new QuoteManager();
            var quoteDetails = GetRandomQuote().Values.First().Split('-');
            objQuoteManager.QuoteName = quoteDetails[0];
            objQuoteManager.QuoteID = GetRandomQuote().Keys.First();
            objQuoteManager.Author = quoteDetails[1];
            TempData["Title"] = "Just a beginning!";
          
           
            return View(objQuoteManager);
        }

        private void CallButtonClickAction()
        {
            ActionResult t = ButtonClickAction();
        }

        public Dictionary<int, string>  GetRandomQuote()
        {
            var quoteDict = new Dictionary<int, string>();
             randomNumberGenerator = new Random();
            if (listQuotes == null)
            {
                listQuotes = Session["listOfQuotes"] as List<string>;
            }
            var randomNumber = randomNumberGenerator.Next(1,listQuotes.Count);
            quoteDict.Add(randomNumber+1, listQuotes[randomNumber].ToString());
            return quoteDict; 
        }

        public void fillList()
        {
            listQuotes = new List<string>();
            var lFileName = "C:\\Users\\shivi.gupta\\Documents\\QuoteDB.xlsx";
            Stream stream = new FileStream(lFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            _package = new ExcelPackage(stream);
            for (var i = 2; i <= _package.Workbook.Worksheets[1].Dimension.End.Row; ++i)
            {
                listQuotes.Add(_package.Workbook.Worksheets[1].Cells[i, 2].Text);
            }
            Session["listOfQuotes"] = listQuotes;
        }
        [OutputCache(Duration = 5)]
        public ActionResult ButtonClickAction()
        {
           
            var objQuoteManager1=new QuoteManager();
            var quoteDetails = GetRandomQuote().Values.First().Split('-');
            objQuoteManager1.QuoteName = quoteDetails[0];
            objQuoteManager1.QuoteID  = GetRandomQuote().Keys.First();
            objQuoteManager1.Author = quoteDetails[1];
            return View("Index",objQuoteManager1);
        }

    }
}