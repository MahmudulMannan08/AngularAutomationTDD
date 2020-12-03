using System;
using System.Collections.Generic;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using RestSharp;
using TechTalk.SpecFlow;
using TimeUnityPortal.Pages;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace TimeUnityPortal.Steps
{
    [Binding]
    public class IntegrationRequestsSteps : TestBase
    {
        public IRestResponse IntegrationLoginResponse, IntegrationSendOrderResponse;
        public static string token, dealUrn, dealURL, PINRequest, ParcelRequest, IsParcel, OwnerType, OwnerRequest, MortgageRequest, Salutation, DateofBirth, OwnerPostalCode, PINHeaders, ParcelHeaders, OwnerHeaders, MortgagePropertyRequest, VendorFirmName, VendorFirstName, VendorLastName, AgentFirmName, AgentFirstName, AgentLastName, AgentPhone, BorrowerFirmName, BorrowerFirstName, BorrowerLastName, BorrowerPhone;
        public static int NumberofPin, NumberofParcel, NumberofOwner, NumberofMortgage, MortgagePriority;
        public static JObject jObjectLogin, jObjectSendOrder;
        public static string ContactFirstName = RandomString(6);
        public static string ContactLastName = RandomString(7);

        public static List<int> MortgagePriorityValues = new List<int>();
        public static List<string> OwnerTypeCollection = new List<string>();
        public static List<string> PINValues = new List<string>();
        public static List<string> ParcelRequestValues = new List<string>();
        public static List<string> OwnerValues = new List<string>();
        public static List<string> MortgageValues = new List<string>();

        public static string[] ExcelRowItems;

        public static List<string> MMSDealValues = new List<string>(new string[]{ "true", "false" });
        //public static int MMSDealValuesArrayPosition = GetRandomNumber(0, 1);
        //public static string MMSDeal = MMSDealValues[MMSDealValuesArrayPosition];
        public static string MMSDeal = GetRandomString(MMSDealValues);
        public static string MMSDealUrn = RandomString(3) + "-" + RandomNumber(7);
        public static string MMSDealUrnRequest;

        public static string MatterNumber = "Matter" + RandomNumber(5);
        public static string FileNumber = "File" + RandomNumber(5);

        //public static List<string> LoginRequestValues = new List<string>(new string[]{ "stester", "Itmagnet-03" });

        public static List<string> TransactionTypeValues = new List<string>(new string[]{ "Purchase New", "Purchase Resale", "Mortgage Only", "Existing Owner With Mortgage" });
        public static string TransactionType = GetRandomString(TransactionTypeValues);

        public static string PurchasePrice = GetRandomNumber(0, 9000000).ToString();

        public static JsonSerializerSettings settings = new JsonSerializerSettings{
            DateFormatString = "yyyy-MM-ddTHH:mm:ss.fffZ",
            DateTimeZoneHandling = DateTimeZoneHandling.Utc
        };
        public static string ClosingDateJson = JsonConvert.SerializeObject(DateTime.Now, settings).Replace("\\", "");
        public static string ClosingDate = ClosingDateJson.Replace("\"", "");

        public static List<string> PropertyTypeValues = new List<string>(new string[]{ "Single Family Dwelling", "Condominium Strata", "Mobile Home", "Multi Family Units", "Rooming House", "Vacant Land", "Live Work Units" });
        public static string PropertyType = GetRandomString(PropertyTypeValues);

        public static List<string> OccupancyValues = new List<string>(new string[]{ "Owner", "Rental" });
        public static string Occupancy = GetRandomString(OccupancyValues);

        public static string UnitNumber = RandomNumber(4);
        public static string StreetNumber = RandomNumber(4);
        public static string StreetAddress1 = RandomString(5) + " " + RandomString(4);
        public static string StreetAddress2 = RandomString(5) + " " + RandomString(4);

        public static List<string> CityList = new List<string>(new string[] { "Mississauga", "Brampton", "Oakville", "Hamilton", "London", "Ottawa", "Kitchener", "Markham" });
        public static string City = GetRandomString(CityList);

        public static List<string> ProvinceValues = new List<string>(new string[] {"ON", "AB", "BC", "MB", "NB", "NL", "NT", "NS", "NU", "PE", "QC", "SK", "YT"});
        public static string Province = GetRandomString(ProvinceValues);

        public static string PostalCode = RandomString(1) + RandomNumber(1) + RandomString(1) + RandomNumber(1) +
                                          RandomString(1) + RandomNumber(1);

        public static List<string> EstateTypeValues = new List<string>(new string[]{ "FeeSimple", "Leasehold", "Easement", "Other" });
        public static List<string> IsParcelValues = new List<string>(new string[]{"true", "false"});

        public static List<string> OwnerTypeValues = new List<string>(new string[] {"Company", "Person"});
        public static List<string> SalutationValues = new List<string>(new string[]{ "Mr", "Ms" });

        [Given(@"I make call to integration login POST request")]
        public void GivenIMakeCallToIntegrationLoginPOSTRequest()
        {
            //IntegrationLoginResponse = integrationRequests.IntegraionLoginResponse(Request, IntegrationLogin, loginRequestBody);
            IntegrationLoginResponse = PostRequest(IntegrationLogin, loginRequestBody, "login", "");

            //if (IntegrationLoginResponse.Content == null)
            if (IntegrationLoginResponse.StatusCode != HttpStatusCode.OK)
            {
                UIHelper.SendEmailonError("mahmudautomateerror@gmail.com", "mahmudautomateerror@gmail.com", "Test Header", "Test Body");
                Console.WriteLine("Integration Login POST request was unsuccessful");
            }
            else
            {
                Console.WriteLine("Integration Login POST request is made");
            }
        }

        [Then(@"I verify login response is valid")]
        public void ThenIVerifyLoginResponseIsValid()
        {
            //Assert.True(integrationRequests.IntegrationLoginStatus().Equals(true));
            Assert.AreEqual(HttpStatusCode.OK, IntegrationLoginResponse.StatusCode, "Integration Login POST request returned a valid response");
        }

        [Then(@"I store token from response")]
        public void ThenIStoreTokenFromResponse()
        {
            //var JsonResponseContent = JsonConvert.DeserializeObject(IntegrationLoginResponse.Content);
            jObjectLogin = JObject.Parse(IntegrationLoginResponse.Content);
            token = (string)jObjectLogin.SelectToken("token");
            Assert.True(token.Length>0, "Token was not found");
        }

        [Given(@"I make call to integration send order POST request")]
        public void GivenIMakeCallToIntegrationSendOrderPOSTRequest()
        {
            //Generating MMSDeal to be true or false randomly. If 'true' generate MMSDealUrn
            if (MMSDeal == "true")
            {
                MMSDealUrnRequest = "'mmsDealUrn':'"+ MMSDealUrn + "',";
            }

            if (MMSDeal == "false")
            {
                MMSDealUrnRequest = "";
            }

            //Generating dynamic PIN's. Number of PINs reside between 1 to 10. 10 is the maximum number of PINs
            NumberofPin = GetRandomNumber(1, 10);
            for (int i = 0; i< NumberofPin; )
            {
                //{'sourceId':56782,'value':'Pin789'}
                var SourceID = RandomNumber(4);
                var PinValue = RandomNumber(3);
                PINRequest = PINRequest + "{'sourceId':" + SourceID + ",'value':'Pin" + PinValue + "'}";

                PINValues.Add(SourceID);
                PINValues.Add("Pin" + PinValue);

                if (i < (NumberofPin-1))
                {
                    PINRequest = PINRequest + ",";
                }

                i++;
            }

            //Generating dynamic Parcels with Parcel B being POTL randomly. Number of Parcels reside between 1 to 15. 15 is the maximum number of Parcels
            NumberofParcel = GetRandomNumber(1, 15);
            IsParcel = GetRandomString(IsParcelValues);
            for (int j = 0; j<NumberofParcel; )
            {
                //{'estateType':'FeeSimple','sequence':'1','legalDescription':'Purchase New with Mortgage','condominiumPlan':'Smoke IT','interest':2.75}

                var EstateType = GetRandomString(EstateTypeValues);
                if ((IsParcel == "true") & (NumberToString(j + 1, true) == "B"))
                {
                    ParcelRequest = ParcelRequest + "{'estateType':'" + EstateType + "','sequence':'" + NumberToString(j + 1, true) + "','legalDescription':'" + "Transaction Type - " + TransactionType + "','condominiumPlan':'Condo Plan','interest':2.75}";

                    ParcelRequestValues.Add(EstateType);
                    ParcelRequestValues.Add(NumberToString(j + 1, true));
                    ParcelRequestValues.Add("Transaction Type - " + TransactionType);
                    ParcelRequestValues.Add("Condo Plan");
                    ParcelRequestValues.Add("2.75");
                }
                else
                {
                    ParcelRequest = ParcelRequest + "{'estateType':'" + EstateType + "','sequence':'" + NumberToString(j + 1, true) + "','legalDescription':'" + "Transaction Type - " + TransactionType + "'}";

                    ParcelRequestValues.Add(EstateType);
                    ParcelRequestValues.Add(NumberToString(j + 1, true));
                    ParcelRequestValues.Add("Transaction Type - " + TransactionType);
                }

                if (j < (NumberofParcel-1))
                {
                    ParcelRequest = ParcelRequest + ",";
                }

                j++;
            }

            //Generating dynamic owners with Owner type being Company / Person randomly. Number of Owners reside between 1 to 25. 25 is the maximum number of Owners
            NumberofOwner = GetRandomNumber(1, 25);
            for (int k = 0; k< NumberofOwner; )
            {
                OwnerType = GetRandomString(OwnerTypeValues);
                OwnerPostalCode = RandomString(1) + RandomNumber(1) + RandomString(1) + RandomNumber(1) + RandomString(1) + RandomNumber(1);
                var CorporationName = RandomString(6);
                var Phone = RandomNumber(10);
                var Email = RandomString(5).ToLower();
                var UnitNumber = RandomNumber(3);
                var StreetNumber = RandomNumber(4);
                var StreetAddress1 = RandomString(6) + " " + RandomString(4);
                var StreetAddress2 = RandomString(2);
                var FirstName = RandomString(5);
                var MiddleName = RandomString(2);
                var LastName = RandomString(5);
                //{'sourceId':1,'corporationName':'CANADA INC.','salutation':'Mr','firstName':'Mahmudul','middleName':'MM','lastName':'Mannan','dateOfBirth':'1985-01-01T18:37:34.530Z','phone':'905-280-8752','email':'mmannan @fct.ca','address':{'unitNumber':'1','streetNumber':'4597','streetAddress1':'Joshua Creek','streetAddress2':'JC','city':'Oakville','province':'ON','postalCode':'L4G3L3'}}
                if (OwnerType == "Company")
                {
                    OwnerRequest = OwnerRequest + "{'sourceId':" + (k+1) + ",'corporationName':'" + CorporationName + " Inc.','phone':'" + Phone + "','email':'" + Email + "@fct.ca','address':{'unitNumber':'" + UnitNumber + "','streetNumber':'" + StreetNumber + "','streetAddress1':'" + StreetAddress1 + "','streetAddress2':'" + StreetAddress2 + "','city':'" + City + "','province':'" + Province + "','postalCode':'" + OwnerPostalCode + "'}}";

                    //OwnerHeaders = OwnerHeaders + "\"" + "sourceId" + "\", " + "\"" + "corporationName" + "\", " + "\"" + "phone" + "\", " + "\"" + "email" + "\", " + "\"" + "unitNumber" + "\", " + "\"" + "streetNumber" + "\", " + "\"" + "streetAddress1" + "\", " + "\"" + "streetAddress2" + "\", " + "\"" + "city" + "\", " + "\"" + "province" + "\", " + "\"" + "postalCode" + "\", ";

                    OwnerTypeCollection.Add("Company");

                    OwnerValues.Add((k + 1).ToString());
                    OwnerValues.Add(CorporationName + " Inc.");
                    OwnerValues.Add(Phone);
                    OwnerValues.Add(Email + "@fct.ca");
                    OwnerValues.Add(UnitNumber);
                    OwnerValues.Add(StreetNumber);
                    OwnerValues.Add(StreetAddress1);
                    OwnerValues.Add(StreetAddress2);
                    OwnerValues.Add(City);
                    OwnerValues.Add(Province);
                    OwnerValues.Add(OwnerPostalCode);
                }
                else if (OwnerType == "Person")
                {
                    Salutation = GetRandomString(SalutationValues);
                    DateofBirth = JsonConvert.SerializeObject(RandomDay(), settings).Replace("\\", "").Replace("\"", "");  //Generating random date of birth in the range (1/1/1920 - 1/1/2000)
                    OwnerRequest = OwnerRequest + "{'sourceId':" + (k+1) + ",'salutation':'" + Salutation + "','firstName':'" + FirstName + "','middleName':'" + MiddleName + "','lastName':'" + LastName + "','dateOfBirth':'" + DateofBirth + "','phone':'" + Phone + "','email':'" + Email + "@fct.ca','address':{'unitNumber':'" + UnitNumber + "','streetNumber':'" + StreetNumber + "','streetAddress1':'" + StreetAddress1 + "','streetAddress2':'" + StreetAddress2 + "','city':'" + City + "','province':'" + Province + "','postalCode':'" + OwnerPostalCode + "'}}";

                    //OwnerHeaders = OwnerHeaders + "\"" + "sourceId" + "\", " + "\"" + "salutation" + "\", " + "\"" + "firstName" + "\", " + "\"" + "middleName" + "\", " + "\"" + "lastName" + "\", " + "\"" + "dateOfBirth" + "\", " + "\"" + "phone" + "\", " + "\"" + "email" + "\", " + "\"" + "unitNumber" + "\", " + "\"" + "streetNumber" + "\", " + "\"" + "streetAddress1" + "\", " + "\"" + "streetAddress2" + "\", " + "\"" + "city" + "\", " + "\"" + "province" + "\", " + "\"" + "postalCode" + "\", ";

                    OwnerTypeCollection.Add("Person");

                    OwnerValues.Add((k + 1).ToString());
                    OwnerValues.Add(Salutation);
                    OwnerValues.Add(FirstName);
                    OwnerValues.Add(MiddleName);
                    OwnerValues.Add(LastName);
                    OwnerValues.Add(DateofBirth);
                    OwnerValues.Add(Phone);
                    OwnerValues.Add(Email + "@fct.ca");
                    OwnerValues.Add(UnitNumber);
                    OwnerValues.Add(StreetNumber);
                    OwnerValues.Add(StreetAddress1);
                    OwnerValues.Add(StreetAddress2);
                    OwnerValues.Add(City);
                    OwnerValues.Add(Province);
                    OwnerValues.Add(OwnerPostalCode);
                }

                if (k < (NumberofOwner-1))
                {
                    OwnerRequest = OwnerRequest + ",";
                }

                k++;
            }

            //Generating dynamic Mortgages with random priority. Number of Mortgages reside between 1 to 9. 9 is the maximum number of Mortgages
            NumberofMortgage = GetRandomNumber(1, 9);
            for (int m = 1; m <= NumberofMortgage; m++)
            {
                MortgagePriorityValues.Add(m);
            }
            for (int l = 0; l<NumberofMortgage; )
            {
                //{'sourceId':1,'primary':true,'referenceNumber':'MT2284745','amount':250000.00,'lenderName':'Toronto Dominion Bank of Canada','priority':1}

                //Too Slow
                /*do
                {
                    MortgagePriority = GetRandomNumber(1, NumberofMortgage);
                } while (MortgagePriorityValues.Contains(MortgagePriority));
                MortgagePriorityValues.Add(MortgagePriority);*/

                MortgagePriority = new Random().Next(1, MortgagePriorityValues.Count) - 1;
                var ReferenceNumber = RandomNumber(7);
                var Amount = GetRandomNumber(1, 900000);
                var LenderName = RandomString(4) + " " + RandomString(6);
                var Priority = MortgagePriorityValues[MortgagePriority];

                MortgageRequest = MortgageRequest + "{'sourceId':" + (l+1) + ",'primary':true,'referenceNumber':'MT" + ReferenceNumber + "','amount':" + Amount + ".00,'lenderName':'" + LenderName + "','priority':" + Priority + "}";

                MortgageValues.Add((l + 1).ToString());
                MortgageValues.Add("true".ToLower());
                MortgageValues.Add("MT" + ReferenceNumber);
                MortgageValues.Add(Amount + ".00");
                MortgageValues.Add(LenderName);
                MortgageValues.Add(Priority.ToString());

                MortgagePriorityValues.RemoveAt(MortgagePriority);

                if (l < (NumberofMortgage - 1))
                {
                    MortgageRequest = MortgageRequest + ",";
                }

                l++;
            }

            //Generating number of MortgageProperties based on the random number of Mortgages above
            for (int n = 0; n< NumberofMortgage; )
            {
                //{'mortgageSourceId':1,'propertySourceId':0}
                MortgagePropertyRequest = MortgagePropertyRequest + "{'mortgageSourceId':" + (n+1) + ",'propertySourceId':0}";

                if (n < (NumberofMortgage - 1))
                {
                    MortgagePropertyRequest = MortgagePropertyRequest + ",";
                }

                n++;
            }

            VendorFirmName = RandomString(6);
            VendorFirstName = RandomString(5);
            VendorLastName = RandomString(4);
            AgentFirmName = RandomString(5);
            AgentFirstName = RandomString(4);
            AgentLastName = RandomString(4);
            AgentPhone = RandomNumber(10);
            BorrowerFirmName = RandomString(7);
            BorrowerFirstName = RandomString(4);
            BorrowerLastName = RandomString(4);
            BorrowerPhone = RandomNumber(10);

            string sendorderRequestBody = "{'dealUrn':null,'contactFirstName':'" + ContactFirstName + "','contactLastName':'" + ContactLastName + "','mmsDeal':'"+ MMSDeal + "'," + MMSDealUrnRequest + "'matterNumber':'" + MatterNumber + "','fileNumber':'" + FileNumber + "','transactionType':'" + TransactionType + "','purchasePrice':" + PurchasePrice + ".00,'closingDate':'" + ClosingDate + "','property':{'propertyType':'" + PropertyType + "','numberOfUnits':1,'zoning':'','occupancy':'" + Occupancy + "','address':{'unitNumber':'" + UnitNumber + "','streetNumber':'" + StreetNumber + "','streetAddress1':'" + StreetAddress1 + "','streetAddress2':'" + StreetAddress2 + "','city':'" + City + "','province':'" + Province + "','postalCode':'" + PostalCode + "'},'pins':[" + PINRequest + "],'parcels':[" + ParcelRequest + "]},'owners':[" + OwnerRequest + "],'mortgages':[" + MortgageRequest + "],'mortgagesProperties':[" + MortgagePropertyRequest + "],'vendorSolicitor':{'firmname':'Vendor " + VendorFirmName + "','firstName':'Vendor " + VendorFirstName + "','lastName':'Vendor " + VendorLastName + "'},'realEstateAgent':{'firmName':'Real Estate " + AgentFirmName + "','firstName':'Agent " + AgentFirstName + "','lastName':'Agent " + AgentLastName + "','phone':'" + AgentPhone + "'},'borrowerSolicitor':{'firmName':'" + BorrowerFirmName + "','firstName':'" + BorrowerFirstName + "','lastName':'" + BorrowerLastName + "','phone':'" + BorrowerPhone + "'}}";

            IntegrationSendOrderResponse = PostRequest(IntegrationSendOrder, sendorderRequestBody, "send order", token);

            Console.WriteLine(IntegrationSendOrderResponse.Content == null? "Integration Send Order POST request was unsuccessful" : "Integration Send Order POST request is made");
        }

        [Then(@"I verify send order response is valid")]
        public void ThenIVerifySendOrderResponseIsValid()
        {
            Assert.AreEqual(HttpStatusCode.OK, IntegrationSendOrderResponse.StatusCode, "Integration Send Order POST request returned a valid response");
        }

        [Then(@"I store dealURN from response")]
        public void ThenIStoreDealURNFromResponse()
        {
            jObjectSendOrder = JObject.Parse(IntegrationSendOrderResponse.Content);
            dealUrn = (string)jObjectSendOrder.SelectToken("dealUrn");
            Assert.True(dealUrn.Length > 0, "DealURN was not found");

            dealURL = (string)jObjectSendOrder.SelectToken("url");
            Assert.True(dealURL.Length > 0, "URL was not found");
        }

        [Then(@"Create and write Login request and response data to excel file")]
        public void ThenCreateAndWriteLoginRequestAndResponseDataToExcelFile()
        {
            /*Application oXL;
            _Workbook oWB;
            _Worksheet oSheet;
            Range oRng;

            try
            {
                //Start Excel and get Application object
                oXL = new Application();
                //oXL.Visible = true;
                oXL.Visible = false;

                //Get a new workbook
                //oWB = (_Workbook)(oXL.Workbooks.Add(""));
                oWB = (_Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (_Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell
                oSheet.Cells[1, 1] = "userName";
                oSheet.Cells[1, 2] = "password";

                //Format A1:B1 as bold, vertical alignment = center
                oSheet.get_Range("A1", "B1").Font.Bold = true;
                oSheet.get_Range("A1", "B1").VerticalAlignment = XlVAlign.xlVAlignCenter;

                //Fill A2:B2 with an array of values
                //string[,] LoginRequestValues = new string[1, 2];
                //LoginRequestValues[0, 0] = "stester";
                //LoginRequestValues[0, 1] = "Itmagnet-03";

                string[] LoginRequestValues = { "stester", "Itmagnet-03" };
                //LoginRequestValues[0] = "stester";
                //LoginRequestValues[1] = "Itmagnet-03";
                oSheet.get_Range("A2", "B2").Value2 = LoginRequestValues;

                //AutoFit columns A:B
                oRng = oSheet.get_Range("A1", "B1");
                oRng.EntireColumn.AutoFit();

                //oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs("C:\\Users\\mmannan\\source\\repos\\TimeUnityPortal\\TestResults\\Excel_Output\\LoginRequest" + new Random().Next(1, 1000) + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
            }*/

            string[] LoginRequestHeaders = { "Type", "User Name", "Password" };
            string[] LoginRequestValues = { "Request", "stester", "Itmagnet-03" };
            string[] LoginResponseHeaders = { "Type", "Token" };
            string[] LoginResponseValues = { "Response", (string)jObjectLogin.SelectToken("token") };

            CreateWriteExcel(LoginRequestHeaders, LoginRequestValues, LoginResponseHeaders, LoginResponseValues, "Login");
        }

        [Then(@"Create and write Send Order request and response data to excel file")]
        public void ThenCreateAndWriteSendOrderRequestAndResponseDataToExcelFile()
        {
            /*for (int j = 0; j < NumberofParcel; j++)
            {
                ParcelHeaders = ParcelHeaders + "\"" + "estateType" + "\", " + "\"" + "sequence" + "\", " + "\"" + "legalDescription" + "\", ";
            }*/

            //string[] SendOrderRequestHeaders = { "dealUrn", "contactFirstName", "contactLastName", "mmsDeal", "mmsDealUrn", "matterNumber", "fileNumber", "transactionType", "purchasePrice", "closingDate", "propertyType", "numberOfUnits", "zoning", "occupancy", "unitNumber", "streetNumber", "streetAddress1", "streetAddress2", "city", "province", "postalCode" };

            //List<string> SendOrderRequestHeaders = new List<string> { "Type", "dealUrn", "contactFirstName", "contactLastName", "mmsDeal", "mmsDealUrn", "matterNumber", "fileNumber", "transactionType", "purchasePrice", "closingDate", "propertyType", "numberOfUnits", "zoning", "occupancy", "unitNumber", "streetNumber", "streetAddress1", "streetAddress2", "city", "province", "postalCode" };
            List<string> SendOrderRequestHeaders = new List<string> { "Type", "dealUrn", "contactFirstName", "contactLastName" };
            List<string> SendOrderRequestValues = new List<string>{ "Request", "null", ContactFirstName, ContactLastName };

            if (MMSDeal == "true")
            {
                SendOrderRequestHeaders.Add("mmsDeal");
                SendOrderRequestHeaders.Add("mmsDealUrn");

                SendOrderRequestValues.Add(MMSDeal);
                SendOrderRequestValues.Add(MMSDealUrn);
            }

            else if (MMSDeal == "false")
            {
                SendOrderRequestHeaders.Add("mmsDeal");

                SendOrderRequestValues.Add(MMSDeal);
            }

            SendOrderRequestHeaders.Add("matterNumber");
            SendOrderRequestHeaders.Add("fileNumber");
            SendOrderRequestHeaders.Add("transactionType");
            SendOrderRequestHeaders.Add("purchasePrice");
            SendOrderRequestHeaders.Add("closingDate");
            SendOrderRequestHeaders.Add("propertyType");
            SendOrderRequestHeaders.Add("numberOfUnits");
            SendOrderRequestHeaders.Add("zoning");
            SendOrderRequestHeaders.Add("occupancy");
            SendOrderRequestHeaders.Add("unitNumber");
            SendOrderRequestHeaders.Add("streetNumber");
            SendOrderRequestHeaders.Add("streetAddress1");
            SendOrderRequestHeaders.Add("streetAddress2");
            SendOrderRequestHeaders.Add("city");
            SendOrderRequestHeaders.Add("province");
            SendOrderRequestHeaders.Add("postalCode");

            SendOrderRequestValues.Add(MatterNumber);
            SendOrderRequestValues.Add(FileNumber);
            SendOrderRequestValues.Add(TransactionType);
            SendOrderRequestValues.Add(PurchasePrice);
            SendOrderRequestValues.Add(ClosingDate);
            SendOrderRequestValues.Add(PropertyType);
            SendOrderRequestValues.Add("1");
            SendOrderRequestValues.Add("");
            SendOrderRequestValues.Add(Occupancy);
            SendOrderRequestValues.Add(UnitNumber);
            SendOrderRequestValues.Add(StreetNumber);
            SendOrderRequestValues.Add(StreetAddress1);
            SendOrderRequestValues.Add(StreetAddress2);
            SendOrderRequestValues.Add(City);
            SendOrderRequestValues.Add(Province);
            SendOrderRequestValues.Add(PostalCode);

            int ListElement = 0;
            for (int i = 0; i < NumberofPin; i++)
            {
                SendOrderRequestHeaders.Add("sourceId");
                SendOrderRequestHeaders.Add("value");

                SendOrderRequestValues.Add(PINValues[ListElement]);
                SendOrderRequestValues.Add(PINValues[ListElement + 1]);
                ListElement = ListElement + 2;
            }

            ListElement = 0;
            for (int j = 0; j < NumberofParcel; j++)
            {
                if ((IsParcel == "true") & (j == 1))
                {
                    SendOrderRequestHeaders.Add("estateType");
                    SendOrderRequestHeaders.Add("sequence");
                    SendOrderRequestHeaders.Add("legalDescription");
                    SendOrderRequestHeaders.Add("condominiumPlan");
                    SendOrderRequestHeaders.Add("interest");

                    SendOrderRequestValues.Add(ParcelRequestValues[3]);
                    SendOrderRequestValues.Add(ParcelRequestValues[4]);
                    SendOrderRequestValues.Add(ParcelRequestValues[5]);
                    SendOrderRequestValues.Add(ParcelRequestValues[6]);
                    SendOrderRequestValues.Add(ParcelRequestValues[7]);

                    ListElement = ListElement + 5;
                }
                else
                {
                    SendOrderRequestHeaders.Add("estateType");
                    SendOrderRequestHeaders.Add("sequence");
                    SendOrderRequestHeaders.Add("legalDescription");

                    SendOrderRequestValues.Add(ParcelRequestValues[ListElement]);
                    SendOrderRequestValues.Add(ParcelRequestValues[ListElement + 1]);
                    SendOrderRequestValues.Add(ParcelRequestValues[ListElement + 2]);

                    ListElement = ListElement + 3;
                }
            }

            string[] OwnerTypeFlag = OwnerTypeCollection.ToArray();
            ListElement = 0;
            for (int k = 0; k < NumberofOwner; k++)
            {
                if (OwnerTypeFlag[k] == "Company")
                {
                    SendOrderRequestHeaders.Add("sourceId");
                    SendOrderRequestHeaders.Add("corporationName");
                    SendOrderRequestHeaders.Add("phone");
                    SendOrderRequestHeaders.Add("email");
                    SendOrderRequestHeaders.Add("unitNumber");
                    SendOrderRequestHeaders.Add("streetNumber");
                    SendOrderRequestHeaders.Add("streetAddress1");
                    SendOrderRequestHeaders.Add("streetAddress2");
                    SendOrderRequestHeaders.Add("city");
                    SendOrderRequestHeaders.Add("province");
                    SendOrderRequestHeaders.Add("postalCode");

                    SendOrderRequestValues.Add(OwnerValues[ListElement]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 1]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 2]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 3]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 4]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 5]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 6]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 7]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 8]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 9]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 10]);

                    ListElement = ListElement + 11;
                }
                else if (OwnerTypeFlag[k] == "Person")
                {
                    SendOrderRequestHeaders.Add("sourceId");
                    SendOrderRequestHeaders.Add("salutation");
                    SendOrderRequestHeaders.Add("firstName");
                    SendOrderRequestHeaders.Add("middleName");
                    SendOrderRequestHeaders.Add("lastName");
                    SendOrderRequestHeaders.Add("dateOfBirth");
                    SendOrderRequestHeaders.Add("phone");
                    SendOrderRequestHeaders.Add("email");
                    SendOrderRequestHeaders.Add("unitNumber");
                    SendOrderRequestHeaders.Add("streetNumber");
                    SendOrderRequestHeaders.Add("streetAddress1");
                    SendOrderRequestHeaders.Add("streetAddress2");
                    SendOrderRequestHeaders.Add("city");
                    SendOrderRequestHeaders.Add("province");
                    SendOrderRequestHeaders.Add("postalCode");

                    SendOrderRequestValues.Add(OwnerValues[ListElement]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 1]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 2]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 3]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 4]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 5]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 6]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 7]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 8]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 9]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 10]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 11]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 12]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 13]);
                    SendOrderRequestValues.Add(OwnerValues[ListElement + 14]);

                    ListElement = ListElement + 15;
                }
            }

            ListElement = 0;
            for (int l = 0; l < NumberofMortgage; l++)
            {
                SendOrderRequestHeaders.Add("sourceId");
                SendOrderRequestHeaders.Add("primary");
                SendOrderRequestHeaders.Add("referenceNumber");
                SendOrderRequestHeaders.Add("amount");
                SendOrderRequestHeaders.Add("lenderName");
                SendOrderRequestHeaders.Add("priority");

                SendOrderRequestValues.Add(MortgageValues[ListElement]);
                SendOrderRequestValues.Add(MortgageValues[ListElement + 1]);
                SendOrderRequestValues.Add(MortgageValues[ListElement + 2]);
                SendOrderRequestValues.Add(MortgageValues[ListElement + 3]);
                SendOrderRequestValues.Add(MortgageValues[ListElement + 4]);
                SendOrderRequestValues.Add(MortgageValues[ListElement + 5]);

                ListElement = ListElement + 6;
            }

            for (int l = 0; l < NumberofMortgage; l++)
            {
                SendOrderRequestHeaders.Add("mortgageSourceId");
                SendOrderRequestHeaders.Add("propertySourceId");

                SendOrderRequestValues.Add((l + 1).ToString());
                SendOrderRequestValues.Add("0");
            }

            SendOrderRequestHeaders.Add("Vendorfirmname");
            SendOrderRequestHeaders.Add("VendorfirstName");
            SendOrderRequestHeaders.Add("VendorlastName");
            SendOrderRequestHeaders.Add("AgentfirmName");
            SendOrderRequestHeaders.Add("AgentfirstName");
            SendOrderRequestHeaders.Add("AgentlastName");
            SendOrderRequestHeaders.Add("Agentphone");
            SendOrderRequestHeaders.Add("BorrowerfirmName");
            SendOrderRequestHeaders.Add("BorrowerfirstName");
            SendOrderRequestHeaders.Add("BorrowerlastName");
            SendOrderRequestHeaders.Add("Borrowerphone");

            SendOrderRequestValues.Add(VendorFirmName);
            SendOrderRequestValues.Add(VendorFirstName);
            SendOrderRequestValues.Add(VendorLastName);
            SendOrderRequestValues.Add(AgentFirmName);
            SendOrderRequestValues.Add(AgentFirstName);
            SendOrderRequestValues.Add(AgentLastName);
            SendOrderRequestValues.Add(AgentPhone);
            SendOrderRequestValues.Add(BorrowerFirmName);
            SendOrderRequestValues.Add(BorrowerFirstName);
            SendOrderRequestValues.Add(BorrowerLastName);
            SendOrderRequestValues.Add(BorrowerPhone);

            string[] SendOrderResponseHeaders = { "Type", "dealUrn", "url" };
            string[] SendOrderResponseValues = { "Response", dealUrn, dealURL };

            //CreateWriteExcel(SendOrderRequestHeaders.ToArray(), SendOrderRequestValues, SendOrderResponseHeaders, SendOrderResponseValues, "SendOrder");
            CreateWriteExcel(SendOrderRequestHeaders.ToArray(), SendOrderRequestValues.ToArray(), SendOrderResponseHeaders, SendOrderResponseValues, "SendOrder");
        }

        [Given(@"I read Send Order request data from excel")]
        public void GivenIReadSendOrderRequestDataFromExcel()
        {
            ExcelRowItems = ReadDataFromExcel("C4", "BJ4", 1);
        }

        [Given(@"I read Send Order request data (.*), (.*) & (.*) from excel")]
        public void GivenIReadSendOrderRequestDataCBJFromExcel(string CellStart, string CellEnd, int SheetNumber)
        {
            ExcelRowItems = ReadDataFromExcel(CellStart, CellEnd, SheetNumber);
        }

        [Then(@"I make call to integration send order POST request with data from excel")]
        public void ThenIMakeCallToIntegrationSendOrderPOSTRequestWithDataFromExcel()
        {
            var dealUrnValue = ExcelRowItems[0];
            ContactFirstName = ExcelRowItems[1];
            ContactLastName = ExcelRowItems[2];
            MMSDeal = ExcelRowItems[3];

            if ((MMSDeal == "true") || (MMSDeal == "TRUE"))
            {
                MMSDealUrnRequest = "'mmsDealUrn':'" + MMSDealUrn + "',";
            }

            if ((MMSDeal == "false") || (MMSDeal == "FALSE"))
            {
                MMSDealUrnRequest = "";
            }

            MatterNumber = ExcelRowItems[4];
            FileNumber = ExcelRowItems[5];
            TransactionType = ExcelRowItems[6];
            PurchasePrice = ExcelRowItems[7];
            ClosingDate = ExcelRowItems[8];
            PropertyType = ExcelRowItems[9];
            var numberOfUnitsValue = ExcelRowItems[10];
            var zoningValue = ExcelRowItems[11];
            Occupancy = ExcelRowItems[12];
            UnitNumber = ExcelRowItems[13];
            StreetNumber = ExcelRowItems[14];
            StreetAddress1 = ExcelRowItems[15];
            StreetAddress2 = ExcelRowItems[16];
            City = ExcelRowItems[17];
            Province = ExcelRowItems[18];
            PostalCode = ExcelRowItems[19];

            var SourceID = ExcelRowItems[20];
            var PinValue = ExcelRowItems[21];
            PINRequest = "{'sourceId':" + SourceID + ",'value':'Pin" + PinValue + "'}";

            PINValues.Add(SourceID);
            PINValues.Add(PinValue);

            var EstateType = ExcelRowItems[22];
            var ParcelSequenceValue = ExcelRowItems[23];
            var legalDescriptionValue = ExcelRowItems[24];
            var condominiumPlanValue = ExcelRowItems[25];
            var interestValue = ExcelRowItems[26];
            ParcelRequest = "{'estateType':'" + EstateType + "','sequence':'" + ParcelSequenceValue + "','legalDescription':'" + legalDescriptionValue + "','condominiumPlan':'" + condominiumPlanValue + "','interest':" + interestValue + "}";

            ParcelRequestValues.Add(EstateType);
            ParcelRequestValues.Add(ParcelSequenceValue);
            ParcelRequestValues.Add(legalDescriptionValue);
            ParcelRequestValues.Add(condominiumPlanValue);
            ParcelRequestValues.Add(interestValue);

            var OwnersourceIdValue = ExcelRowItems[27];
            var CorporationName = ExcelRowItems[28];
            Salutation = ExcelRowItems[29];
            var FirstName = ExcelRowItems[30];
            var MiddleName = ExcelRowItems[31];
            var LastName = ExcelRowItems[32];
            var DateofBirth = ExcelRowItems[33];
            var Phone = ExcelRowItems[34];
            var Email = ExcelRowItems[35];
            UnitNumber = ExcelRowItems[36];
            StreetNumber = ExcelRowItems[37];
            StreetAddress1 = ExcelRowItems[38];
            StreetAddress2 = ExcelRowItems[39];
            City = ExcelRowItems[40];
            Province = ExcelRowItems[41];
            OwnerPostalCode = ExcelRowItems[42];
            OwnerRequest = "{'sourceId':" + OwnersourceIdValue + ",'corporationName':'" + CorporationName + "','salutation':'" + Salutation + "','firstName':'" + FirstName + "','middleName':'" + MiddleName + "','lastName':'" + LastName + "','dateOfBirth':'" + DateofBirth + "','phone':'" + Phone + "','email':'" + Email + "','address':{'unitNumber':'" + UnitNumber + "','streetNumber':'" + StreetNumber + "','streetAddress1':'" + StreetAddress1 + "','streetAddress2':'" + StreetAddress2 + "','city':'" + City + "','province':'" + Province + "','postalCode':'" + OwnerPostalCode + "'}}";

            OwnerTypeCollection.Add("Company");

            OwnerValues.Add(OwnersourceIdValue);
            OwnerValues.Add(CorporationName);
            OwnerValues.Add(Salutation);
            OwnerValues.Add(FirstName);
            OwnerValues.Add(MiddleName);
            OwnerValues.Add(LastName);
            OwnerValues.Add(DateofBirth);
            OwnerValues.Add(Phone);
            OwnerValues.Add(Email);
            OwnerValues.Add(UnitNumber);
            OwnerValues.Add(StreetNumber);
            OwnerValues.Add(StreetAddress1);
            OwnerValues.Add(StreetAddress2);
            OwnerValues.Add(City);
            OwnerValues.Add(Province);
            OwnerValues.Add(OwnerPostalCode);

            var MortgagesourceIdValue = ExcelRowItems[43];
            var MortgageprimaryValue = ExcelRowItems[44];
            var ReferenceNumber = ExcelRowItems[45];
            var Amount = ExcelRowItems[46];
            var LenderName = ExcelRowItems[47];
            var Priority = ExcelRowItems[48];
            MortgageRequest = "{'sourceId':" + MortgagesourceIdValue + ",'primary':" + MortgageprimaryValue.ToLower() + ",'referenceNumber':'" + ReferenceNumber + "','amount':" + Amount + ",'lenderName':'" + LenderName + "','priority':" + Priority + "}";

            MortgageValues.Add(MortgagesourceIdValue);
            MortgageValues.Add(MortgageprimaryValue);
            MortgageValues.Add(ReferenceNumber);
            MortgageValues.Add(Amount);
            MortgageValues.Add(LenderName);
            MortgageValues.Add(Priority.ToString());

            MortgagePropertyRequest = "{'mortgageSourceId':1,'propertySourceId':0}";

            VendorFirmName = ExcelRowItems[49];
            VendorFirstName = ExcelRowItems[50];
            VendorLastName = ExcelRowItems[51];
            AgentFirmName = ExcelRowItems[52];
            AgentFirstName = ExcelRowItems[53];
            AgentLastName = ExcelRowItems[54];
            AgentPhone = ExcelRowItems[55];
            BorrowerFirmName = ExcelRowItems[56];
            BorrowerFirstName = ExcelRowItems[57];
            BorrowerLastName = ExcelRowItems[58];
            BorrowerPhone = ExcelRowItems[59];

            string sendorderRequestBody = "{'dealUrn':" + dealUrnValue + ",'contactFirstName':'" + ContactFirstName + "','contactLastName':'" + ContactLastName + "','mmsDeal':'" + MMSDeal + "'," + MMSDealUrnRequest + "'matterNumber':'" + MatterNumber + "','fileNumber':'" + FileNumber + "','transactionType':'" + TransactionType + "','purchasePrice':" + PurchasePrice + ".00,'closingDate':'" + ClosingDate + "','property':{'propertyType':'" + PropertyType + "','numberOfUnits':" + numberOfUnitsValue + ",'zoning':'" + zoningValue + "','occupancy':'" + Occupancy 
                                          + "','address':{'unitNumber':'" + UnitNumber + "','streetNumber':'" + StreetNumber + "','streetAddress1':'" + StreetAddress1 + "','streetAddress2':'" + StreetAddress2 + "','city':'" + City + "','province':'" + Province + "','postalCode':'" + PostalCode + "'},'pins':[" + PINRequest + "],'parcels':[" + ParcelRequest + "]},'owners':[" + OwnerRequest + "],'mortgages':[" + MortgageRequest + "],'mortgagesProperties':[" + MortgagePropertyRequest + "],'vendorSolicitor':{'firmname':'" + VendorFirmName + "','firstName':'" +
                                          VendorFirstName + "','lastName':'" + VendorLastName + "'},'realEstateAgent':{'firmName':'" + AgentFirmName + "','firstName':'" + AgentFirstName + "','lastName':'" + AgentLastName + "','phone':'" + AgentPhone + "'},'borrowerSolicitor':{'firmName':'" + BorrowerFirmName + "','firstName':'" + BorrowerFirstName + "','lastName':'" + BorrowerLastName + "','phone':'" + BorrowerPhone + "'}}";

            IntegrationSendOrderResponse = PostRequest(IntegrationSendOrder, sendorderRequestBody, "send order", token);

            Console.WriteLine(IntegrationSendOrderResponse.Content == null ? "Integration Send Order POST request was unsuccessful" : "Integration Send Order POST request is made");
        }
    }
}