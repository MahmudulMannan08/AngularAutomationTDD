using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Newtonsoft.Json;
using RestSharp;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TimeUnityPortal
{
    using System;
    using System.Text;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Firefox;
    using OpenQA.Selenium.IE;
    using OpenQA.Selenium.Support.UI;
    using Protractor;
    using TechTalk.SpecFlow;
    //using System.Net.Http;
    //using System.Collections.Generic;

    [Binding]
    public class TestBase
    {
        public static string HostUrl = ConfigurationManager.AppSettings["ADMIN"];
        public static string Username = ConfigurationManager.AppSettings["ADMINUSER"];
        public static string Password = ConfigurationManager.AppSettings["ADMINPASS"];
        //public static string TransactionID = ConfigurationManager.AppSettings["ADMINTRANSACTIONID"];
        public static string fctURN = ConfigurationManager.AppSettings["PurchaseDealURN"];
        //public static string Request = ConfigurationManager.AppSettings["SERVICEREQUEST"];
        //public static string IntegrationLogin = ConfigurationManager.AppSettings["SERVICEINTEGRATIONLOGIN"];
        public static string IntegrationLogin = ConfigurationManager.AppSettings["INTEGRATIONLOGIN"];
        public static string IntegrationSendOrder = ConfigurationManager.AppSettings["INTEGRATIONSENDORDER"];
        public static string SeleniumDriver = ConfigurationManager.AppSettings["SELENIUM_DRIVER"];
        public static string loginRequestBody = "{'userName':'stester','password':'Itmagnet-03'}";
        //public static string loginRequestBody = "{'userNamee':'stester','password':'Itmagnet-03'}"; //Email on Error Test

        public static string tableRequestHeaderEndLast, tableRequestHeaderEndFirst;
        //public static string sendorderRequestBody = "{'dealUrn':null,'contactFirstName':'Contact First Name Here','contactLastName':'Contact Last Name here','mmsDeal':'true','mmsDealUrn':'MMS-3456789','matterNumber':'Test','fileNumber':'File Test','transactionType':'PurchaseNew','purchasePrice':450000.00,'closingDate':'2018-09-20T18:37:34.530Z','property':{'propertyType':'Single Family Dwelling','numberOfUnits':1,'zoning':'','occupancy':'Owner','address':{'unitNumber':'12','streetNumber':'2235','streetAddress1':'City Centre Dr','streetAddress2':'Square One','city':'Mississauga','province':'ON','postalCode':'L6H6E9'},'pins':[{'sourceId':56782,'value':'Pin789'}],'parcels':[{'estateType':'FeeSimple','sequence':'1','legalDescription':'Purchase New with Mortgage','condominiumPlan':'Smoke IT','interest':2.75},{'estateType':'Leasehold','sequence':'2','legalDescription':'Purchase New with Mortgage','condominiumPlan':'Full Furnished Condo','interest':1.50}]},'owners':[{'sourceId':1,'corporationName':'CANADA INC.','salutation':'Mr','firstName':'Mahmudul','middleName':'MM','lastName':'Mannan','dateOfBirth':'1985-01-01T18:37:34.530Z','phone':'905-280-8752','email':'mmannan @fct.ca','address':{'unitNumber':'1','streetNumber':'4597','streetAddress1':'Joshua Creek','streetAddress2':'JC','city':'Oakville','province':'ON','postalCode':'L4G3L3'}},{'sourceId':2,'salutation':'Ms','firstName':'Rifat','middleName':'RR','lastName':'Reza','dateOfBirth':'1990-01-01T18:37:34.530Z','phone':'905-280-8752','email':'rreza @fct.ca','address':{'unitNumber':'4','streetNumber':'5544','streetAddress1':'Silverstone Creek','streetAddress2':'SC','city':'Oakville','province':'ON','postalCode':'L4G3K9'}}],'mortgages':[{'sourceId':1,'primary':true,'referenceNumber':'MT2284745','amount':250000.00,'lenderName':'Toronto Dominion Bank of Canada','priority':1},{'sourceId':2,'primary':false,'referenceNumber':'MT2365456','amount':250000.00,'lenderName':'Royal Bank of Canada','priority':2}],'mortgagesProperties':[{'mortgageSourceId':1,'propertySourceId':0},{'mortgageSourceId':2,'propertySourceId':0}],'vendorSolicitor':{'firmname':'Vendor Firm','firstName':'Vendor first name','lastName':'Vendor last name'},'realEstateAgent':{'firmName':'Real Estate Firm','firstName':'Agent First Name','lastName':'Agent Last Name','phone':'4257945678'}}";
        public static RestClient Client = new RestClient();
        private static readonly Random Random = new Random((int)DateTime.Now.Ticks);
        public static DefaultWait<bool> apiWait = new DefaultWait<bool>(new bool())
        {
            Timeout = TimeSpan.FromMinutes(2),
            PollingInterval = TimeSpan.FromSeconds(10)
        };
        public static WebDriverWait Wait;
        public static NgWebDriver driver
        {
            get
            {
                if (!FeatureContext.Current.ContainsKey("browser"))
                {
                    FeatureContext.Current["browser"] = StartBrowser(SeleniumDriver);
                    //FeatureContext.Current["browser"] = "InternetExplorer";
                }

                return (NgWebDriver)FeatureContext.Current["browser"];
            }
        }

        public static string RandomString(int size)
        {
            var builder = new StringBuilder();
            for (var i = 0; i < size; i++)
            {
                var ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * Random.NextDouble() + 65)));
                builder.Append(ch);
            }

            return builder.ToString();
        }

        public static string RandomNumber(int size)
        {
            Random random = new Random();
            string r = "";
            int i;
            for (i = 0; i < size; i++)
            {
                r += random.Next(0, 9).ToString();
            }
            return r;
        }

        /*public static int GetRandomNumber(int minimum, int maximum)
        {
            Random random = new Random();
            return random.Next() * (maximum - minimum) + minimum;
        }*/

        public static int GetRandomNumber(int minimum, int maximum)
        {
            Random random = new Random();
            return random.Next(minimum, maximum);
        }

        public static string GetRandomString(List<string> stringList)
        {
            int index = Random.Next(stringList.Count);
            return stringList[index];
        }

        public static string NumberToString(int number, bool isCaps)
        {
            Char c = (Char)((isCaps ? 65 : 97) + (number - 1));

            return c.ToString();
        }

        public static DateTime RandomDay()
        {
            DateTime start = new DateTime(1920, 1, 1);
            //int range = (DateTime.Today - start).Days;
            int range = (new DateTime(2000, 1, 1) - start).Days;
            return start.AddDays(new Random().Next(range));
        }

        /*public static Func<DateTime> RandomDayFunc()
        {
            DateTime start = new DateTime(1900, 1, 1);
            Random gen = new Random();
            int range = ((TimeSpan)(DateTime.Today - start)).Days;
            return () => start.AddDays(gen.Next(range));
        }*/

        //public static IRestResponse PostRequest(string clientURL, string requestURL, string requestBody)
        //{
        //    //Client = new RestClient("http://IISPRILLCUNDQA1.PREFIRSTCDN.COM");

        //    // client.Authenticator = new HttpBasicAuthenticator(username, password);

        //    //var request = new RestRequest("LawyerIntegrationGateway/v1/users/login", Method.POST);

        //    //request.AddHeader("Content-type", "application/json");

        //    /*var reqbody = "{\"deal\":{\"dealId\":1343126,\"dealUrn\":\"18276065695\",\"lawFirm\":{\"firmName\":\"LLCand MMS Associate Firm \",\"lawyerFirstName\":\"LLC\",\"lawyerLastName\":\"Lawyer\",\"lawyerCrmId\":\"D874029CB11A4C27AF0F61F8B3FB7CD4\",\"contactFirstName\":\"Jack\",\"contactLastName\":\"Sparrow\"},\"transactionType\":\"Purchase New\",\"purchasePrice\":522000,\"closingDate\":\"2018-08-30T00:00:00.000Z\",\"finalPolicyDestination\":\"Lawyer\",\"fileNumber\":\"File Test\",\"services\":[\"Owner Policy\",\"Loan Policy\"],\"properties\":[{\"sourceId\":0,\"primary\":true,\"propertyType\":\"Single Family Dwelling\",\"numberOfUnits\":0,\"zoning\":\"Residential\",\"occupancy\":\"Rental\",\"address\":{\"unitNumber\":\"12\",\"streetNumber\":\"2235\",\"streetAddress1\":\"Sheridan Garden Dr\",\"streetAddress2\":\"test\",\"city\":\"Oakville\",\"province\":\"ON\",\"postalCode\":\"L6H6E6\"},\"pins\":[{\"sourceId\":12345,\"value\":\"Pin777\"}],\"parcels\":[{\"estateType\":\"Fee Simple\",\"sequence\":\"1\",\"legalDescription\":\"This is a Test\"}],\"additionalPropertyTransactionType\":\"Existing Owner\",\"additionalPropertyPolicyRequired\":false,\"questionSetId\":0}],\"owners\":[{\"sourceId\":1,\"corporationName\":\"FCT 2018\",\"salutation\":\"\",\"firstName\":\"\",\"middleName\":\"\",\"lastName\":\"\",\"dateOfBirth\":\"0001-01-01T00:00:00.000Z\",\"phone\":\"\",\"email\":\"\",\"ownerType\":\"CorporationName\"}],\"ownersProperties\":[{\"ownerSourceId\":1,\"propertySourceId\":0,\"selected\":true}],\"mortgages\":[{\"primary\":true,\"mortgageType\":\"Mortgage\",\"amount\":\"200000\",\"lenderName\":\"TD\",\"loanPolicySelected\":false,\"sourceId\":1,\"referenceNumber\":\"MTG1234678\",\"originalRegistrationDate\":\"0001-01-01T00:00:00.000Z\"},{\"primary\":true,\"mortgageType\":\"Mortgage\",\"amount\":\"250000\",\"lenderName\":\"RBC\",\"loanPolicySelected\":false,\"sourceId\":2,\"referenceNumber\":\"MTG1234678\",\"originalRegistrationDate\":\"0001-01-01T00:00:00.000Z\"}],\"mortgagesProperties\":[{\"mortgageSourceId\":1,\"propertySourceId\":0,\"priority\":1,\"selected\":true},{\"mortgageSourceId\":2,\"propertySourceId\":0,\"priority\":1,\"selected\":true}],\"notes\":\"\",\"rotReceived\":true,\"mmsDealUrn\":\"01234567891\"}}";
        //    var reqbody2 = "{\"dealUrn\":18289073145,\"contactFirstName\":\"Contact First Name Here\",\"contactLastName\":\"Contact Last Name here\",\"mmsDeal\":\"true\",\"mmsDealUrn\":\"MMS-3456789\",\"matterNumber\":\"Test\",\"fileNumber\":\"File Test\",\"transactionType\":\"PurchaseNew\",\"purchasePrice\":450000.00,\"closingDate\":\"2018-09-20T18:37:34.530Z\",\"property\":{\"propertyType\":\"SingleFamilyDwelling\",\"numberOfUnits\":1,\"zoning\":\"\",\"occupancy\":\"Owner\",\"address\":{\"unitNumber\":\"12\",\"streetNumber\":\"2235\",\"streetAddress1\":\"City Centre Dr\",\"streetAddress2\":\"Square One\",\"city\":\"Mississauga\",\"province\":\"ON\",\"postalCode\":\"L6H6E9\"},\"pins\":[{\"sourceId\":56782,\"value\":\"Pin789\"}],\"parcels\":[{\"estateType\":\"FeeSimple\",\"sequence\":\"A\",\"legalDescription\":\"Purchase New with Mortgage\"},{\"estateType\":\"Leasehold\",\"sequence\":\"B\",\"legalDescription\":\"Purchase New with Mortgage\",\"condominiumPlan\":\"Full Furnished Condo\",\"interest\":1.50},{\"estateType\":\"Easement\",\"sequence\":\"C\",\"legalDescription\":\"Something\",\"condominiumPlan\":\"Full Furnished Condo\",\"interest\":1.50},{\"estateType\":\"Other\",\"sequence\":\"D\",\"legalDescription\":\"Something\"}]},\"owners\":[{\"sourceId\":1,\"corporationName\":\"CANADA INC.\",\"salutation\":\"Mr\",\"firstName\":\"Mahmudul\",\"middleName\":\"MM\",\"lastName\":\"Mannan\",\"dateOfBirth\":\"1985-01-01T18:37:34.530Z\",\"phone\":\"905-280-8752\",\"email\":\"mmannan@fct.ca\",\"address\":{\"unitNumber\":\"1\",\"streetNumber\":\"4597\",\"streetAddress1\":\"Joshua Creek\",\"streetAddress2\":\"JC\",\"city\":\"Oakville\",\"province\":\"ON\",\"postalCode\":\"L4G3L3\"}},{\"sourceId\":2,\"salutation\":\"Ms\",\"firstName\":\"Rifat\",\"middleName\":\"RR\",\"lastName\":\"Reza\",\"dateOfBirth\":\"1990-01-01T18:37:34.530Z\",\"phone\":\"905-280-8752\",\"email\":\"rreza@fct.ca\",\"address\":{\"unitNumber\":\"4\",\"streetNumber\":\"5544\",\"streetAddress1\":\"Silverstone Creek\",\"streetAddress2\":\"SC\",\"city\":\"Oakville\",\"province\":\"ON\",\"postalCode\":\"L4G3K9\"}}],\"mortgages\":[{\"sourceId\":1,\"primary\":true,\"referenceNumber\":\"MT2284745\",\"amount\":250000.00,\"lenderName\":\"Toronto Dominion Bank of Canada\",\"priority\":1},{\"sourceId\":2,\"primary\":false,\"referenceNumber\":\"MT2365456\",\"amount\":150000.00,\"lenderName\":\"Royal Bank of Canada\",\"priority\":2}],\"mortgagesProperties\":[{\"mortgageSourceId\":1,\"propertySourceId\":0},{\"mortgageSourceId\":2,\"propertySourceId\":0}]}";
        //    var json = JsonConvert.SerializeObject(reqbody);
        //    //request.AddParameter("application/json; charset=utf-8", json, ParameterType.RequestBody);
        //    request.AddJsonBody(json);*/

        //    /*request.AddJsonBody(
        //        new
        //        {
        //            userName = "stester",
        //            password = "Itmagnet - 03",

        //        }); // AddJsonBody serializes the object automatically*/

        //    Client = new RestClient(clientURL);
        //    var request = new RestRequest(requestURL, Method.POST);
        //    request.AddHeader("Content-Type", "application/json; charset=utf-8");
        //    //var jsonRequest = JsonConvert.SerializeObject(requestBody);
        //    request.AddJsonBody(requestBody);
        //    //request.AddParameter("userName", "stester");
        //    //request.AddParameter("password", "Itmagnet-03");

        //    /*var jsonRequest = JsonConvert.SerializeObject(requestBody);
        //    request.AddJsonBody(jsonRequest);*/

        //    IRestResponse response = Client.Execute(request);
        //    //var content = response.Content;
        //    //return content;
        //    return response;
        //}

        /*public static IRestResponse PostRequest(string requestUrl, string requestBody)
        {
            var Client = new RestClient(requestUrl);
            var Request = new RestRequest(Method.POST);

            //Request.AddHeader("cache-control", "no-cache");
            Request.AddHeader("content-type", "application/json");
            Request.AddParameter("application/json", requestBody, ParameterType.RequestBody);
            IRestResponse Response = Client.Execute(Request);

            return Response;
        }*/

        public static IRestResponse PostRequest(string requestUrl, string requestBody, string requestName, string token)
        {
            var Client = new RestClient(requestUrl);
            var Request = new RestRequest(Method.POST);

            if (requestName == "login")
            {
                Request.AddHeader("content-type", "application/json");
            }

            if (requestName == "send order")
            {
                Request.AddHeader("content-type", "application/json");
                Request.AddHeader("authorization", token);
                Request.AddHeader("xFCTAuthorization", "{'authenticatedFctUser': 'stester','userContext':{'partnerUserName': 'stester','firstName': 'Service','lastName': 'Tester','businessRole': 'LAWYER','fctUserName': ''}}");
                Request.AddHeader("language", "en");
            }
            Request.AddParameter("application/json", requestBody, ParameterType.RequestBody);
            IRestResponse Response = Client.Execute(Request);

            return Response;
        }

        public static string [] ReadDataFromExcel(string cellStart, string cellEnd, int sheetNumber)
        {
            Application oXL;
            _Workbook oWB;
            _Worksheet oSheet;
            Range oRng;

            try
            {
                oXL = new Application();
                oWB = oXL.Workbooks.Open(@"C:\Projects\Time Unity\SELENIUM WORKS\Automation Docs\Integration Service Automation.xlsx");
                oSheet = oWB.Worksheets.get_Item(sheetNumber);

                string[] RowItems;
                oRng = oSheet.get_Range(cellStart.ToString(), cellEnd.ToString());
                System.Array myvalues = (Array)oRng.Cells.Value;
                //string[] foo = someObjectArray.OfType<object>().Select(o => o.ToString()).ToArray();
                //string[] fo = myvalues.Select(o => o == null ? (string)null : o.ToString()).ToArray();
                /*
                 List<string> lst = new List<string>(); 
                foreach (object o in myvalues) 
                if (o==null)
                { 
                    lst.Add(null); 
                } 
                else 
                { 
                    lst.Add(o.ToString()); 
                } 
                string[] str2 = lst.ToArray();
                 */

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(oRng);
                Marshal.ReleaseComObject(oSheet);

                //close and release
                oWB.Close();
                Marshal.ReleaseComObject(oWB);

                //quit and release
                oXL.Quit();
                Marshal.ReleaseComObject(oXL);

                List<string> lst = new List<string>();

                foreach (object o in myvalues)
                {
                    lst.Add(o == null ? null : o.ToString());
                }

                RowItems = lst.ToArray();

                return RowItems;
            }
            catch (Exception e)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, e.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, e.Source);

                MessageBox.Show(errorMessage, "Error");

                return null;
            }
        }

        public static void CreateWriteExcel(string [] tableRequestHeaders, string [] requestValues, string[] tableResponseHeaders, string [] responseValues, string requestName)
        {
            Application oXL;
            _Workbook oWB;
            _Worksheet oSheet;
            Range oRng;

            try
            {
                //Start Excel and get Application object
                oXL = new Application();
                oXL.Visible = false;
                //oXL.Visible = true;

                //Get a new workbook
                //oWB = (_Workbook)(oXL.Workbooks.Add(""));
                /*oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (_Worksheet)oWB.ActiveSheet;
                var NumberOfRequestColumns = tableRequestHeaders.Length;
                for (var i = 0; i < NumberOfRequestColumns; i++)
                {
                    oSheet.Cells[1, i + 1] = tableRequestHeaders[i];
                }*/

                //Accommodating for columns more than 256 (Request Header)
                oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (_Worksheet)oWB.ActiveSheet;
                var RequestHeaderRowNumber = 1;
                var RequestHeaderColumnCounter = 1;
                var NumberOfRequestHeaderColumns = tableRequestHeaders.Length;

                for (var i = 0; i < NumberOfRequestHeaderColumns; i++)
                {
                    if (RequestHeaderColumnCounter > 256)
                    {
                        RequestHeaderRowNumber = RequestHeaderRowNumber + 2;
                        RequestHeaderColumnCounter = 1;
                    }
                    oSheet.Cells[RequestHeaderRowNumber, RequestHeaderColumnCounter] = tableRequestHeaders[i];
                    RequestHeaderColumnCounter++;
                }

                //Fill 2nd row with an array of request values
                //oSheet.get_Range("A2", "B2").Value2 = LoginRequestValues;
                //var tableRequestHeaderEnd = NumberToString(NumberOfRequestColumns, true);
                //oSheet.get_Range("A2", tableRequestHeaderEnd + "2").Value2 = requestValues;

                var ColumnGroup = NumberOfRequestHeaderColumns / 26;
                var RestColumnGroup = NumberOfRequestHeaderColumns - (ColumnGroup * 26);
                if ((decimal) ((decimal)NumberOfRequestHeaderColumns / (decimal)26) > 1)
                {
                    tableRequestHeaderEndLast = NumberToString(RestColumnGroup, true);
                    tableRequestHeaderEndFirst = NumberToString(ColumnGroup, true);

                    //oSheet.get_Range("A2", tableRequestHeaderEndFirst + tableRequestHeaderEndLast + "2").Value2 = requestValues;
                }
                else
                {
                    var tableRequestHeaderEnd = NumberToString(NumberOfRequestHeaderColumns, true);
                    //oSheet.get_Range("A2", tableRequestHeaderEnd + "2").Value2 = requestValues;
                }

                //Accommodating for columns more than 256 (Request Values)
                var RequestValueRowNumber = 2;
                var RequestValueColumnCounter = 1;
                var NumberOfRequestValueColumns = requestValues.Length;

                for (var j = 0; j < NumberOfRequestValueColumns; j++)
                {
                    if (RequestValueColumnCounter > 256)
                    {
                        RequestValueRowNumber = RequestValueRowNumber + 2;
                        RequestValueColumnCounter = 1;
                    }
                    oSheet.Cells[RequestValueRowNumber, RequestValueColumnCounter] = requestValues[j];
                    RequestValueColumnCounter++;
                }

                //Fill 4th row with an array of response headers
                /*var NumberOfResponseColumns = tableResponseHeaders.Length;
                for (var j = 0; j < NumberOfResponseColumns; j++)
                {
                    //oSheet.Cells[4, j + 1] = tableResponseHeaders[j];
                    oSheet.Cells[RequestHeaderRowNumber + 2, j + 1] = tableResponseHeaders[j];
                }*/

                oSheet.Cells[RequestValueRowNumber + 4, 1] = "<<RESPONSE>>";
                //Accommodating for columns more than 256 (Response Headers)
                var ResponseHeaderColumnCounter = 1;
                var ResponseHeaderRowNumber = RequestValueRowNumber + 6;
                var NumberOfResponseHeaderColumns = tableResponseHeaders.Length;
                for (var i = 0; i < NumberOfResponseHeaderColumns; i++)
                {
                    if (ResponseHeaderColumnCounter > 256)
                    {
                        ResponseHeaderRowNumber = ResponseHeaderRowNumber + 2;
                        ResponseHeaderColumnCounter = 1;
                    }
                    oSheet.Cells[ResponseHeaderRowNumber, ResponseHeaderColumnCounter] = tableResponseHeaders[i];
                    ResponseHeaderColumnCounter++;
                }

                //Fill 5th row with an array of response values
                //var tableResponseHeaderEnd = NumberToString(NumberOfResponseColumns, true);
                //oSheet.get_Range("A5", tableResponseHeaderEnd + "5").Value2 = responseValues;
                //oSheet.get_Range("A" + (RequestValueRowNumber + 2), tableResponseHeaderEnd + (RequestValueRowNumber + 2)).Value2 = responseValues;

                //Accommodating for columns more than 256 (Response Value)
                var ResponseValueColumnCounter = 1;
                var ResponseValueRowNumber = RequestValueRowNumber + 7;
                for (var j = 0; j < NumberOfResponseHeaderColumns; j++)
                {
                    if (ResponseValueColumnCounter > 256)
                    {
                        ResponseValueRowNumber = ResponseValueRowNumber + 2;
                        ResponseValueColumnCounter = 1;
                    }
                    oSheet.Cells[ResponseValueRowNumber, ResponseValueColumnCounter] = responseValues[j];
                    ResponseValueColumnCounter++;
                }

                //Format column headers as bold, vertical alignment = center
                //oSheet.get_Range("A1", tableRequestHeaderEnd + "1").Font.Bold = true;
                //oSheet.get_Range("A1", tableRequestHeaderEnd + "1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                /*if ((double) (NumberOfRequestHeaderColumns / 26) > 1)
                {
                    oSheet.get_Range("A1", tableRequestHeaderEndFirst + tableRequestHeaderEndLast + "1").Font.Bold = true;
                    oSheet.get_Range("A1", tableRequestHeaderEndFirst + tableRequestHeaderEndLast + "1").VerticalAlignment = XlVAlign.xlVAlignCenter;
                }
                else
                {
                    var tableRequestHeaderEnd = NumberToString(NumberOfRequestHeaderColumns, true);
                    oSheet.get_Range("A1", tableRequestHeaderEnd + "1").Font.Bold = true;
                    oSheet.get_Range("A1", tableRequestHeaderEnd + "1").VerticalAlignment = XlVAlign.xlVAlignCenter;
                }

                oSheet.get_Range("A4", tableResponseHeaderEnd + "4").Font.Bold = true;
                oSheet.get_Range("A4", tableResponseHeaderEnd + "4").VerticalAlignment = XlVAlign.xlVAlignCenter;*/

                //Accommodating for more than 256 columns (headers as bold, vertical alignment = center, Auto fit columns ) [Request]
                var NumberofRequestHeaderRows = NumberOfRequestHeaderColumns / 256;
                var NumberofRequestHeaderCellonLastRow = NumberOfRequestHeaderColumns - (NumberofRequestHeaderRows * 256);
                var RequestHeaderColumnGroup = NumberofRequestHeaderCellonLastRow / 26;
                var RequestHeaderRestColumnGroup = NumberofRequestHeaderCellonLastRow - (RequestHeaderColumnGroup * 26);
                var CounterRequest = 1;
                if ((double) (NumberOfRequestHeaderColumns / 256) > 0)
                {
                    for (var k = 0; k < (NumberOfRequestHeaderColumns / 256); k++)
                    {
                        oSheet.get_Range("A" + CounterRequest, "IV" + CounterRequest).Font.Bold = true;
                        oSheet.get_Range("A" + CounterRequest, "IV" + CounterRequest).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterRequest, "IV" + CounterRequest);
                        oRng.EntireColumn.AutoFit();

                        CounterRequest = CounterRequest + 2;
                    }

                    if ((decimal) ((decimal)NumberofRequestHeaderCellonLastRow / (decimal)26) > 1)
                    {
                        oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest).Font.Bold = true;
                        oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest);
                        oRng.EntireColumn.AutoFit();
                    }
                    else
                    {
                        oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest).Font.Bold = true;
                        oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest);
                        oRng.EntireColumn.AutoFit();
                    }
                }
                else
                {
                    if ((decimal) ((decimal)NumberofRequestHeaderCellonLastRow / (decimal)26) > 1)
                    {
                        oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest).Font.Bold = true;
                        oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterRequest, NumberToString(RequestHeaderColumnGroup, true) + NumberToString(RequestHeaderRestColumnGroup, true) + CounterRequest);
                        oRng.EntireColumn.AutoFit();
                    }
                    else
                    {
                        oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest).Font.Bold = true;
                        oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterRequest, NumberToString(NumberofRequestHeaderCellonLastRow, true) + CounterRequest);
                        oRng.EntireColumn.AutoFit();
                    }
                }

                //Accommodating for more than 256 columns (headers as bold, vertical alignment = center, Auto fit columns ) [Response]
                var NumberofResponseHeaderRows = NumberOfResponseHeaderColumns / 256;
                var NumberofResponseHeaderCellonLastRow = NumberOfResponseHeaderColumns - (NumberofResponseHeaderRows * 256);
                var ResponseHeaderColumnGroup = NumberofResponseHeaderCellonLastRow / 26;
                var ResponseHeaderRestColumnGroup = NumberofResponseHeaderCellonLastRow - (ResponseHeaderColumnGroup * 26);
                var CounterResponse = RequestValueRowNumber + 6;
                if ((double)(NumberOfResponseHeaderColumns / 256) > 0)
                {
                    for (var k = 0; k < (NumberOfResponseHeaderColumns / 256); k++)
                    {
                        oSheet.get_Range("A" + CounterResponse, "IV" + CounterResponse).Font.Bold = true;
                        oSheet.get_Range("A" + CounterResponse, "IV" + CounterResponse).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterResponse, "IV" + CounterResponse);
                        oRng.EntireColumn.AutoFit();

                        CounterResponse = CounterResponse + 2;
                    }

                    if ((decimal) ((decimal)NumberofResponseHeaderCellonLastRow / (decimal)26) > 1)
                    {
                        oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse).Font.Bold = true;
                        oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse);
                        oRng.EntireColumn.AutoFit();
                    }
                    else
                    {
                        oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse).Font.Bold = true;
                        oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse);
                        oRng.EntireColumn.AutoFit();
                    }
                }
                else
                {
                    if ((decimal) ((decimal)NumberofResponseHeaderCellonLastRow / (decimal)26) > 1)
                    {
                        oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse).Font.Bold = true;
                        oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterResponse, NumberToString(ResponseHeaderColumnGroup, true) + NumberToString(ResponseHeaderRestColumnGroup, true) + CounterResponse);
                        oRng.EntireColumn.AutoFit();
                    }
                    else
                    {
                        oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse).Font.Bold = true;
                        oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse).VerticalAlignment = XlVAlign.xlVAlignCenter;

                        oRng = oSheet.get_Range("A" + CounterResponse, NumberToString(NumberofResponseHeaderCellonLastRow, true) + CounterResponse);
                        oRng.EntireColumn.AutoFit();
                    }
                }

                //AutoFit columns
                /*if (NumberOfRequestHeaderColumns > NumberOfResponseColumns)
                {
                    if ((double) (NumberOfRequestHeaderColumns / 26) > 1)
                    {
                        oRng = oSheet.get_Range("A1", tableRequestHeaderEndFirst + tableRequestHeaderEndLast + "1");
                    }
                    else
                    {
                        var tableRequestHeaderEnd = NumberToString(NumberOfRequestHeaderColumns, true);
                        oRng = oSheet.get_Range("A1", tableRequestHeaderEnd + "1");
                    }
                    oRng.EntireColumn.AutoFit();
                }
                else
                {
                    oRng = oSheet.get_Range("A1", tableResponseHeaderEnd + "1");
                    oRng.EntireColumn.AutoFit();
                }*/

                //oXL.Visible = false;
                oXL.UserControl = false;

                System.IO.Directory.CreateDirectory(Directory.GetParent(Directory.GetCurrentDirectory()).FullName + "\\TestResults");
                System.IO.Directory.CreateDirectory(Directory.GetParent(Directory.GetCurrentDirectory()).FullName + "\\TestResults\\Excel_Output");
                /*oWB.SaveAs(Directory.GetParent(Directory.GetCurrentDirectory()).FullName + "\\TestResults\\Excel_Output\\Request" + " " + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_") + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);*/

                if (requestName == "Login")
                {
                    oWB.SaveAs(Directory.GetParent(Directory.GetCurrentDirectory()).FullName + "\\TestResults\\Excel_Output\\" + "LoginRequest" + " " + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_") + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                if (requestName == "SendOrder")
                {
                    oWB.SaveAs(Directory.GetParent(Directory.GetCurrentDirectory()).FullName + "\\TestResults\\Excel_Output\\" + "SendOrderRequest" + " " + DateTime.Now.ToString().Replace("/", "_").Replace(":", "_") + ".xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                oWB.Close();
            }
            catch (Exception e)
            {
                //Console.WriteLine(e);
                //throw;

                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, e.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, e.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        public static NgWebDriver StartBrowser(string browser)
        //public static IWebDriver StartBrowser(string browser)
        {
            IWebDriver driver;

            if (browser.Equals("Firefox"))
            {
                driver = new FirefoxDriver();
            }
            if (browser.Equals("InternetExplorer"))
            {
                driver = new InternetExplorerDriver(@"C:\Windows\SysWOW64\IEDriverServer.exe");
            }
            else
            {
                var options = new ChromeOptions();
                options.AddArguments("test-type");
                driver = new ChromeDriver(options);
            }

            var ngdriver = new NgWebDriver(driver);
            //ngdriver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(100));  //Obsolete
            //ngdriver.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(20));  //Obsolete
            ngdriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);
            ngdriver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(20);
            ngdriver.Manage().Window.Maximize();
            Wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            ngdriver.IgnoreSynchronization = true;
            return ngdriver;
            /*driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);
            driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromSeconds(20);
            driver.Manage().Window.Maximize();
            Wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            return driver;*/
        }

        [BeforeScenario]
        public void RecordLog()
        {
            //api.SetApiKey(accessToken);

            if (!ScenarioContext.Current.ContainsKey("report"))
            {
                var listenerId = RandomString(20);
                //var path = Directory.GetCurrentDirectory();
                //var path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                //var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                //var fileName = path + "\\" + FeatureContext.Current.FeatureInfo.Title + "-"
                //+ ScenarioContext.Current.ScenarioInfo.Title + "_" + listenerId + "_.txt";
                var fileName = "Test.txt";
                var textListener = new TextWriterTraceListener(fileName, listenerId);
                ScenarioContext.Current.Add("report", textListener);
                ScenarioContext.Current.Add("report_file_name", fileName);
                ScenarioContext.Current.Add("stepcounter", 1);
                ScenarioContext.Current.Add("timeStarted", DateTime.Now);
                Trace.Listeners.Add(textListener);
                Trace.AutoFlush = true;
                Trace.Indent();
                textListener.WriteLine("-----------------------------------------------------------");
                textListener.WriteLine("START OF SCENARIO: " + ScenarioContext.Current.ScenarioInfo.Title);
                textListener.WriteLine("-----------------------------------------------------------");
            }
        }

        [AfterScenario]
        public void CloseBrowser()
        {
            var textListener = ScenarioContext.Current.Get<TextWriterTraceListener>("report");

            if (!FeatureContext.Current.ContainsKey("browser"))
            {
                return;
            }

            var dateTime1 = ScenarioContext.Current.Get<DateTime>("timeStarted");
            var dateTime2 = DateTime.Now;
            var diff = dateTime2 - dateTime1;

            textListener.WriteLine("END OF SCENARIO: " + ScenarioContext.Current.ScenarioInfo.Title);
            textListener.WriteLine("-----------------------------------------------------------");
            textListener.WriteLine("STARTED AT: " + ScenarioContext.Current["timeStarted"] + ": COMPLETED AT: " + dateTime2 + ": EXECUTION TIME: " + diff);
            textListener.WriteLine("-----------------------------------------------------------");

            if (ScenarioContext.Current.TestError != null)
            {
                textListener.WriteLine(
                    "ERROR OCCURED: " + ScenarioContext.Current.TestError.Message + " OF TYPE: "
                    + ScenarioContext.Current.TestError.GetType().Name);
                textListener.Flush();
                textListener.Close();
                UIHelper.SendEmailOnErrorWithAttachment(UIHelper.TakeScreenshot());
            }
            else
            {
                textListener.Flush();
                textListener.Close();
            }

            if (ConfigurationManager.AppSettings["VERBOSE_MODE"].Equals("OFF"))
            {
                //UIHelper.SendEmailOnError(UIHelper.TakeScreenshot());
                driver.Quit();
                driver.WrappedDriver.Quit();
            }

            FeatureContext.Current.Remove("browser");
            FeatureContext.Current.Remove("report");
        }
    }
}
