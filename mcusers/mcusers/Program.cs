using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
using System.IO;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Net.Mime;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Message = Google.Apis.Gmail.v1.Data.Message;

namespace mcusers
{
    class Program
    {
        private static GraphServiceClient _graphClient;
        public string OneDriveApiRoot { get; set; } = "https://api.onedrive.com/v1.0/";
        static void Main(string[] args)
        {
            //Looks for appsettings.json file
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }

            //Previous month and year
            string previousMonth = DateTime.Now.AddMonths(-1).ToString("MMMM");
            //string year = DateTime.Now.ToString("yyyy");
            //Logins to Microsoft Admin
            var client = GetAuthenticatedGraphClient(config);
            //Requests assigned licenses
            var graphRequest = client.Users.Request();
            var users = graphRequest.GetAsync().Result;
            var graphRequest2 = client.Users.Request();
            var licence = graphRequest2.Select("assignedLicenses").GetAsync().Result;
            //License names
            var graphRequest3 = client.SubscribedSkus.Request();
            var nazev = graphRequest3.GetAsync().Result;
            //UsageLocationCountry
            var graphRequest4 = client.Users.Request();
            var country = graphRequest4.Select("UsageLocation").GetAsync().Result;

            //string cesta = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string cesta = Environment.CurrentDirectory;

            //---------------Save into txt-------------------------
            /*
            using (var sw = new StreamWriter(cesta + "/licence.txt", false))
            {
                for (int i = 0; i < users.Count(); i++)
                {
                    sw.WriteLine(users[i].DisplayName);
                    foreach (var item in licence[i].AssignedLicenses)
                    {
                        string skuname = "neznamy";
                        foreach (var sku in nazev)
                        {
                            if (sku.SkuId == item.SkuId)
                            {
                                skuname = sku.SkuPartNumber;
                            }
                        }
                        sw.WriteLine(item.SkuId + ": " + skuname);
                    }
                    sw.WriteLine("------------------------------");
                }
            }
            */

            //Save into CSV - "," readable by Google Sheets
            using (var sw = new StreamWriter(cesta + "/licence_GoogleSheets_" + previousMonth + ".csv", false, Encoding.UTF8))

            {
                sw.WriteLine("Name,Licence,Price/Month(Euro),UsageLocation");
                for (int i = 0; i < users.Count(); i++)
                {
                    
                    foreach (var item in licence[i].AssignedLicenses)
                    {   
                        string skuname = "neznamy";
                        foreach (var sku in nazev)
                        {
                            if (sku.SkuId == item.SkuId)
                            {
                                skuname = sku.SkuPartNumber;

                            }
                        }
                        //Replaces skunames with real subscription names
                        switch (skuname)
                        {
                            case "O365_BUSINESS":
                                sw.WriteLine(users[i].DisplayName + "," + "Microsoft 365 Apps for business" + "," + "10.60" + "," + country[i].UsageLocation);
                                break;
                            case "FLOW_FREE":
                                sw.WriteLine(users[i].DisplayName + "," + "Microsoft Power Automate FREE" + "," + "" + "," + country[i].UsageLocation);
                                break;
                            case "MCOPSTNC":
                                sw.WriteLine(users[i].DisplayName + "," + "Communication Credits FREE" + "," + "" + "," + country[i].UsageLocation);
                                break;
                            case "MCOMEETADV":
                                sw.WriteLine(users[i].DisplayName + "," + "Audio Conferencing in Microsoft 365" + "," + "2.10" + "," + country[i].UsageLocation);
                                break;
                            case "TEAMS_EXPLORATORY":
                                sw.WriteLine(users[i].DisplayName + "," + "Microsoft Teams Exploratory FREE" + "," + "" + "," + country[i].UsageLocation);
                                break;
                            case "PROJECTCLIENT":
                                sw.WriteLine(users[i].DisplayName + "," + "Project for Office 365" + "," + "21.10" + "," + country[i].UsageLocation);
                                break;
                            case "VISIOCLIENT":
                                sw.WriteLine(users[i].DisplayName + "," + "Visio Plan 2" + "," + "12.6" + "," + country[i].UsageLocation);
                                break;
                            case "MCOSTANDARD":
                                sw.WriteLine(users[i].DisplayName + "," + "Skype for Business Online (Plan 2)" + "," + "4.60" + "," + country[i].UsageLocation);
                                break;
                            case "TEAMS_ESSENTIALS_AAD":
                                sw.WriteLine(users[i].DisplayName + "," + "Microsoft Teams Essentials" + "," + "3.40" + "," + country[i].UsageLocation);
                                break;
                            default:
                                sw.WriteLine(users[i].DisplayName + "," + skuname + "," + "Unknown" + "," + country[i].UsageLocation);
                                break;
                        }
                    }
                }
                sw.Close();
            }
            //Save into CSV - ";" readable by EXCEL
            using (var sw = new StreamWriter(cesta + "/licence_Excel_" + previousMonth + ".csv", false, Encoding.Latin1))
            {
                sw.WriteLine("sep=;");
                sw.WriteLine("Name;Licence;Price/Month(Euro);UsageLocation");
                for (int i = 0; i < users.Count(); i++)
                {
                    foreach (var item in licence[i].AssignedLicenses)
                    {
                        string skuname = "neznamy";
                        foreach (var sku in nazev)
                        {
                            if (sku.SkuId == item.SkuId)
                            {
                                skuname = sku.SkuPartNumber;
                            }
                        }
                        //Replaces skunames with real subscription names
                        switch (skuname)
                        {
                            case "O365_BUSINESS":
                                sw.WriteLine(users[i].DisplayName + ";" + "Microsoft 365 Apps for business" + ";" + "10,60" + ";" + country[i].UsageLocation);
                                break;
                            case "FLOW_FREE":
                                sw.WriteLine(users[i].DisplayName + ";" + "Microsoft Power Automate FREE" + ";" + "" + ";" + country[i].UsageLocation);
                                break;
                            case "MCOPSTNC":
                                sw.WriteLine(users[i].DisplayName + ";" + "Communication Credits FREE" + ";" + "" + ";" + country[i].UsageLocation);
                                break;
                            case "MCOMEETADV":
                                sw.WriteLine(users[i].DisplayName + ";" + "Audio Conferencing in Microsoft 365" + ";" + "2,10" + ";" + country[i].UsageLocation);
                                break;
                            case "TEAMS_EXPLORATORY":
                                sw.WriteLine(users[i].DisplayName + ";" + "Microsoft Teams Exploratory FREE" + ";" + "" + ";" + country[i].UsageLocation);
                                break;
                            case "PROJECTCLIENT":
                                sw.WriteLine(users[i].DisplayName + ";" + "Project for Office 365" + ";" + "21,10" + ";" + country[i].UsageLocation);
                                break;
                            case "VISIOCLIENT":
                                sw.WriteLine(users[i].DisplayName + ";" + "Visio Plan 2" + ";" + "12,6" + ";" + country[i].UsageLocation);
                                break;
                            case "MCOSTANDARD":
                                sw.WriteLine(users[i].DisplayName + ";" + "Skype for Business Online (Plan 2)" + ";" + "4,60" + ";" + country[i].UsageLocation);
                                break;
                            case "TEAMS_ESSENTIALS_AAD":
                                sw.WriteLine(users[i].DisplayName + ";" + "Microsoft Teams Essentials" + ";" + "3,40" + ";" + country[i].UsageLocation);
                                break;
                            default:
                                sw.WriteLine(users[i].DisplayName + ";" + skuname + ";" + "Unknown" + ";" + country[i].UsageLocation);
                                break;
                        }
                    }
                }
                sw.Close();
            }
            //Send E-mail
            try
            {
                string predmet = "Microsoft Users - Monthly Report - " + previousMonth;
                string zprava = "List of users with assigned licenses. Users with unassigned licenses aren't on the list.<br><br>\n\nThis is an automated message, do not reply. <br>\nIn case of help mail to";

                //-----------------Sending mail with Third-Party Unsecured access via SMTP----------------
                /*
                MailMessage mailInstance = new MailMessage("mail@YouSendfrom", "mail@Yousendto");
                mailInstance.CC.Add("");
                mailInstance.Subject = predmet;
                mailInstance.Body = zprava;
                mailInstance.Attachments.Add(new System.Net.Mail.Attachment("licence.csv"));
                // Optional
                SmtpClient mailSenderInstance = new SmtpClient("smtp-relay.gmail.com", 587);
                // 25 is the port of the SMTP host
                mailSenderInstance.Credentials = new System.Net.NetworkCredential("mail", "password");
                mailSenderInstance.EnableSsl = true;
                mailSenderInstance.Send(mailInstance);
                mailInstance.Dispose(); // Please remember to dispose this object
                */
                
                //---------------------Sending mail via Google Mail API----------------------
                var msg = new Google.Apis.Gmail.v1.Data.Message();

                MailMessage mail = new MailMessage();
                mail.Subject = predmet;
                mail.Body = zprava;
                mail.From = new MailAddress("");
                mail.IsBodyHtml = true;
                string attImg = "licence_GoogleSheets_" + previousMonth + ".csv";
                string attImg2 = "licence_Excel_" + previousMonth + ".csv";
                mail.Attachments.Add(new System.Net.Mail.Attachment(attImg));
                mail.Attachments.Add(new System.Net.Mail.Attachment(attImg2));
                mail.To.Add(new MailAddress(""));
                mail.CC.Add(new MailAddress(""));
                MimeKit.MimeMessage mimeMessage = MimeKit.MimeMessage.CreateFromMailMessage(mail);


                //Google API Credintials Login
                string[] Scopes = { GmailService.Scope.GmailSend };
                string ApplicationName = "Gmail API Microsoft Users";

                UserCredential credential;

                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                //Create Gmail API Service
                var service = new GmailService(new BaseClientService.Initializer() { HttpClientInitializer = credential, ApplicationName = ApplicationName });

                //Encode to string
                string Base64UrlEncode(string input)
                {
                    var data = Encoding.UTF8.GetBytes(input);
                    return Convert.ToBase64String(data).Replace("+", "-").Replace("/", "_").Replace("=", "");
                }
                //Send Email
                msg.Raw = Base64UrlEncode(mimeMessage.ToString());
                service.Users.Messages.Send(msg, "me").Execute();
                mail.Attachments.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            Console.WriteLine("\nGraph Request");
            Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
        }
        //Connecting Miscrosoft API
        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                string workingDirectory = Environment.CurrentDirectory;
                //string projectDirectory = System.IO.Directory.GetParent(workingDirectory).Parent.Parent.FullName;
                var config = new ConfigurationBuilder()
                    .SetBasePath(workingDirectory)
                    .AddJsonFile("appsettings.json", false, true)
                    .Build();

                if (
                    string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }
    }
}