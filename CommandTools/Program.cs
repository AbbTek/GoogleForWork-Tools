using CommandTools.Templates;
using Google.Apis.Admin.Directory.directory_v1;
using Google.Apis.Admin.Directory.directory_v1.Data;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.GData.Apps.GoogleMailSettings;
using Google.GData.Client;
using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CommandTools
{
    class Program
    {
        static string[] Scopes = {
            "https://apps-apis.google.com/a/feeds/emailsettings/2.0/",
            DirectoryService.Scope.AdminDirectoryUserReadonly
        };

        static string ApplicationName = "GoogleForWork-Tools";
        static void Main(string[] args)
        {
            bool show_help = false;
            string domain = "";
            string pathRazorTemplate = "";
            string pathSecretFile = "client_secret.json";
            string user = "";

            var p = new OptionSet() {
            { "d|domain=", "the domain where you'll change the signatures.",
              v => domain = v },
            { "t|template=",
                "the file path of the razor template that will used.",
              v => pathRazorTemplate = v },
            { "s|secret=", "the file with the secret from Google console.",
              v => pathSecretFile = v },
            { "u|user=", "an specific user to apply the signature.",
              v => user = v },
            { "a|all", "apply the signature to every user form the domain.",
              v => { if (v != null) user=""; } },
            { "h|help",  "show this message and exit",
              v => show_help = v != null },
        };

            List<string> extra;
            try
            {
                extra = p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("Error: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try --help for more information.");
                return;
            }

            if (show_help)
            {
                ShowHelp(p);
                return;
            }


            UserCredential credential = GetCredentials(pathSecretFile);

            var service = new DirectoryService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            IList<User> users = new List<User>();

            if (!string.IsNullOrWhiteSpace(user))
            {
                users.Add(service.Users.Get(user).Execute());
            }
            else
            {
                UsersResource.ListRequest request = service.Users.List();
                request.Domain = domain;
                request.Projection = UsersResource.ListRequest.ProjectionEnum.Full;
                users = request.Execute().UsersValue;
            }

            OAuth2Parameters parameter = new OAuth2Parameters()
            {
                AccessToken = credential.Token.AccessToken
            };

            var requestFactory = new GOAuth2RequestFactory("apps", ApplicationName, parameter);
            var serviceGmail = new GoogleMailSettingsService("revolute.academy", ApplicationName);
            serviceGmail.RequestFactory = requestFactory;

            Console.WriteLine("Users:");
            if (users != null)
            {
                foreach (var userItem in users)
                {
                    if (userItem.IsMailboxSetup.HasValue && userItem.IsMailboxSetup.Value)
                    {
                        var signature = Render.Execute(pathRazorTemplate, userItem);
                        var userName = GetUser(userItem.PrimaryEmail);
                        dynamic obj = userItem.Organizations;

                        serviceGmail.UpdateSignature(userName, signature);
                        Console.WriteLine("New signature for {0}", userName);
                    }
                }
            }
            else
            {
                Console.WriteLine("No users found.");
            }

            Console.WriteLine("Signatures updated");
        }

        private static UserCredential GetCredentials(string secretFilePath)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(secretFilePath, FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            return credential;
        }

        static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
        }

        private static string GetUser(string email)
        {
            return email.Substring(0, email.IndexOf("@"));
        }
    }
}
