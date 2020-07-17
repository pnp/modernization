using CommandLine;
using CommandLine.Text;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Framework.TimerJobs.Enums;
using SharePoint.Modernization.Scanner.Core.Telemetry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;

namespace SharePoint.Modernization.Scanner.Core
{
    /// <summary>
    /// Possible scanning modes
    /// </summary>
    public enum Mode
    {
        Full = 0,
        GroupifyOnly, // this mode is always included, is part of the default scan
        PageOnly,
        PublishingOnly,
        PublishingWithPagesOnly,
        ListOnly,
        WorkflowOnly,
        InfoPathOnly,
        BlogOnly,
        HomePageOnly,
    }

    /// <summary>
    /// Commandline options
    /// </summary>
    public class Options 
    {
        #region Private variables
        Parser parser;
        #endregion

        #region Construction
        public Options()
        {
            var versionInfo = VersionCheck.LatestVersion();
            this.CurrentVersion = versionInfo.Item1;
            this.NewVersion = versionInfo.Item2;            
        }
        #endregion


        // Important:
        // Still available: l

        #region Security related options
        [Option('i', "clientid", HelpText = "Client ID of the app-only principal used to scan your site collections", MutuallyExclusiveSet = "A")]
        public string ClientID { get; set; }

        [Option('s', "clientsecret", HelpText = "Client Secret of the app-only principal used to scan your site collections", MutuallyExclusiveSet = "B")]
        public string ClientSecret { get; set; }

        [Option('u', "user", HelpText = "User id used to scan/enumerate your site collections", MutuallyExclusiveSet = "A")]
        public string User { get; set; }

        [Option('p', "password", HelpText = "Password of the user used to scan/enumerate your site collections", MutuallyExclusiveSet = "B")]
        public string Password { get; set; }

        [Option('z', "azuretenant", HelpText = "Azure tenant (e.g. contoso.microsoftonline.com)")]
        public string AzureTenant { get; set; }

        [Option('y', "azureenvironment", HelpText = "Azure environment (only works for Azure AD Cert auth!). Possible values: Production, USGovernment, Germany, China", DefaultValue = AzureEnvironment.Production, Required = false)]
        public AzureEnvironment AzureEnvironment { get; set; }

        [Option('f', "certificatepfx", HelpText = "Path + name of the pfx file holding the certificate to authenticate")]
        public string CertificatePfx { get; set; }

        [Option('x', "certificatepfxpassword", HelpText = "Password of the pfx file holding the certificate to authenticate")]
        public string CertificatePfxPassword { get; set; }

        [Option('a', "tenantadminsite", HelpText = "Url to your tenant admin site (e.g. https://contoso-admin.contoso.com): only needed when your not using SPO MT")]
        public string TenantAdminSite { get; set; }
        #endregion

        #region Sites to scan
        [Option('t', "tenant", HelpText = "Tenant name, e.g. contoso when your sites are under https://contoso.sharepoint.com/sites. This is the recommended model for SharePoint Online MT as this way all site collections will be scanned")]
        public string Tenant { get; set; }

        [OptionList('r', "urls", HelpText = "List of (wildcard) urls (e.g. https://contoso.sharepoint.com/*,https://contoso-my.sharepoint.com,https://contoso-my.sharepoint.com/personal/*) that you want to get scanned. Ignored if -t or --tenant are provided.", Separator = ',')]
        public virtual IList<string> Urls { get; set; }

        [Option('o', "includeod4b", HelpText = "Include OD4B sites in the scan", DefaultValue = false)]
        public bool IncludeOD4B { get; set; }

        [Option('v', "csvfile", HelpText = "CSV file name (e.g. input.csv) which contains the list of site collection urls that you want to scan")]
        public virtual string CsvFile { get; set; }
        #endregion

        #region Scanner configuration
        [Option('h', "threads", HelpText = "Number of parallel threads, maximum = 100", DefaultValue = 10)]
        public int Threads { get; set; }
        #endregion

        #region File handling
        [Option('e', "separator", HelpText = "Separator used in output CSV files (e.g. \";\")", DefaultValue = ",")]
        public string Separator { get; set; }
        #endregion

        [Option('m', "mode", HelpText = "Execution mode. Use following modes: Full, GroupifyOnly, ListOnly, PageOnly, HomePageOnly, PublishingOnly, PublishingWithPagesOnly, WorkflowOnly, InfoPathOnly or BlogOnly. Omit or use full for a full scan", DefaultValue = Mode.Full, Required = false)]
        public Mode Mode { get; set; }

        [Option('b', "exportwebpartproperties", HelpText = "Export the web part property data", DefaultValue = false, Required = false)]
        public bool ExportWebPartProperties { get; set; }

        [Option('c', "skipusageinformation", HelpText = "Don't use search to get the site/page usage information and don't export that data", DefaultValue = false, Required = false)]
        public bool SkipUsageInformation { get; set; }

        [Option('j', "skipuserinformation", HelpText = "Don't include user information in the exported data", DefaultValue = false, Required = false)]
        public bool SkipUserInformation { get; set; }

        [Option('k', "skiplistsonlyblockedbyoobreaons", HelpText = "Exclude lists which are blocked due to out of the box reasons: base template, view type of field type", DefaultValue = false)]
        public bool ExcludeListsOnlyBlockedByOobReasons { get; set; }

        [Option('d', "skipreport", HelpText = "Don't generate an Excel report for the found data", DefaultValue = false, Required = false)]
        public bool SkipReport { get; set; }

        [OptionList('g', "exportpaths", HelpText = "List of paths (e.g. c:\\temp\\636529695601669598,c:\\temp\\636529695601656430) containing scan results you want to add to the report", Separator = ',')]
        public virtual IList<string> ExportPaths { get; set; }

        [Option('n', "disabletelemetry", HelpText = "We use telemetry to make this a better tool...but you're free to disable that", DefaultValue = false, Required = false)]
        public bool DisableTelemetry { get; set; }

        [Option('q', "dateformat", HelpText = "Date format to use for date export in the CSV files. Use M/d/yyyy or d/M/yyyy", DefaultValue = "M/d/yyyy", Required = false)]
        public string DateFormat { get; set; }

        [Option('w', "storedcertificate", HelpText = "Path to stored certificate in the form of StoreName|StoreLocation|Thumbprint. E.g. My|LocalMachine|3FG496B468BE3828E2359A8A6F092FB701C8CDB1", DefaultValue = "", Required = false)]
        public string StoredCertificate { get; set; }

        public X509Certificate2 AzureCert { get; set; }

        public string AccessToken { get; set; }

        /// <summary>
        /// Property holding the possible newer version
        /// </summary>
        public string NewVersion { get; set; }

        /// <summary>
        /// Property holding the current version
        /// </summary>
        public string CurrentVersion { get; set; }

        /// <summary>
        /// Are we using SharePoint App-Only?
        /// </summary>
        /// <returns>true if app-only, false otherwise</returns>
        public AuthenticationType AuthenticationTypeProvided()
        {
            if (!string.IsNullOrEmpty(ClientID) && !string.IsNullOrEmpty(ClientSecret))
            {
                return AuthenticationType.AppOnly;
            }
            else if (!string.IsNullOrEmpty(User) && !string.IsNullOrEmpty(Password))
            {
                return AuthenticationType.Office365;
            }
            else if (!string.IsNullOrEmpty(CertificatePfx) && !string.IsNullOrEmpty(CertificatePfxPassword) && !string.IsNullOrEmpty(ClientID) && !string.IsNullOrEmpty(AzureTenant))
            {
                return AuthenticationType.AzureADAppOnly;
            }
            else if (!string.IsNullOrEmpty(StoredCertificate) && !string.IsNullOrEmpty(ClientID) && !string.IsNullOrEmpty(AzureTenant))
            {
                return AuthenticationType.AzureADAppOnly;
            }
            else if (!string.IsNullOrEmpty(AccessToken))
            {
                return AuthenticationType.AccessToken;
            }
            else
            {
                throw new Exception("Clonflicting security parameters provided.");
            }
        }

        /// <summary>
        /// Validate the provided commandline options, will exit the program when not valid
        /// </summary>
        /// <param name="args">Command line arguments</param>
        public void ValidateOptions(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine(this.GetUsage());
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
            }

            var parser = new Parser(settings =>
            {
                settings.MutuallyExclusive = true;
                settings.HelpWriter = Parser.Default.Settings.HelpWriter;
                settings.CaseSensitive = false;
            });
            this.parser = parser;

            ValidateOptions(args, parser);
        }

        /// <summary>
        /// Validate the provided commandline options, will exit the program when not valid
        /// </summary>
        /// <param name="args">Command line arguments</param>
        /// <param name="parser">Parser object holding the commadline parsing settings</param>
        public virtual void ValidateOptions(string[] args, Parser parser)
        {
            this.parser = parser;

            if (!parser.ParseArguments(args, this))
            {
                Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
            }

            // perform additional validation
            if (!String.IsNullOrEmpty(this.ClientID) && String.IsNullOrEmpty(this.CertificatePfx) && String.IsNullOrEmpty(this.StoredCertificate))
            {
                if (String.IsNullOrEmpty(this.ClientSecret))
                {
                    Console.WriteLine("If you specify a client id you also need to specify a client secret");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }
            if (!String.IsNullOrEmpty(this.ClientSecret))
            {
                if (String.IsNullOrEmpty(this.ClientID))
                {
                    Console.WriteLine("If you specify a client secret you also need to specify a client id");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }

            if (!String.IsNullOrEmpty(this.CertificatePfx))
            {
                if (String.IsNullOrEmpty(this.CertificatePfxPassword))
                {
                    Console.WriteLine("If you specify a certificate you also need to specify a password for the certificate");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
                if (String.IsNullOrEmpty(this.AzureTenant))
                {
                    Console.WriteLine("If you specify a certificate you also need to specify the Azure Tenant");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
                if (String.IsNullOrEmpty(this.ClientID))
                {
                    Console.WriteLine("If you specify a certificate you also need to specify the clientid of the Azure application");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }

            if (!String.IsNullOrEmpty(this.CertificatePfxPassword))
            {
                if (String.IsNullOrEmpty(this.CertificatePfx))
                {
                    Console.WriteLine("If you specify a certifcate password you also need to specify a certificate");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
                if (String.IsNullOrEmpty(this.AzureTenant))
                {
                    Console.WriteLine("If you specify a certificate password you also need to specify the Azure Tenant");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
                if (String.IsNullOrEmpty(this.ClientID))
                {
                    Console.WriteLine("If you specify a certificate password you also need to specify the clientid of the Azure application");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }

            if (!String.IsNullOrEmpty(this.User))
            {
                if (String.IsNullOrEmpty(this.Password))
                {
                    Console.WriteLine("If you specify a user you also need to specify a password");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }
            if (!String.IsNullOrEmpty(this.Password))
            {
                if (String.IsNullOrEmpty(this.User))
                {
                    Console.WriteLine("If you specify a password you also need to specify a user");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }
            if (!String.IsNullOrEmpty(this.CsvFile))
            {
                if (!System.IO.File.Exists(this.CsvFile))
                {
                    Console.WriteLine("Failed to find csv file with urls. Please check file path provided.");
                    Environment.Exit(CommandLine.Parser.DefaultExitCodeFail);
                }
            }
        }

        /// <summary>
        /// Include detailed site page analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludePage(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.PageOnly || mode == Mode.HomePageOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include detailed site page analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IsHomePageOnly(Mode mode)
        {
            if (mode == Mode.HomePageOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include detailed publishing analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludePublishing(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.PublishingOnly)
            {
                return true;
            }

            if (mode == Mode.PublishingWithPagesOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include detailed publishing page analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludePublishingWithPages(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.PublishingWithPagesOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include detailed publishing page analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludeLists(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.ListOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include workflow analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludeWorkflow(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.WorkflowOnly)
            {
                return true;
            }

            return false;
        }


        /// <summary>
        /// Include InfoPath analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludeInfoPath(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.InfoPathOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Include Blog analysis
        /// </summary>
        /// <param name="mode">mode that was provided</param>
        /// <returns>True if included, false otherwise</returns>
        public static bool IncludeBlog(Mode mode)
        {
            if (mode == Mode.Full)
            {
                return true;
            }

            if (mode == Mode.BlogOnly)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Shows the usage information of the scanner
        /// </summary>
        /// <returns>String with the usage information</returns>
        [HelpOption]
        public string GetUsage()
        {
            var help = this.GetUsage("SharePoint PnP Modernization scanner");

            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("==========================================================");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("See the sp-dev-modernization repo for more information at:");
            help.AddPreOptionsLine("https://github.com/SharePoint/sp-dev-modernization/tree/master/Tools/SharePoint.Modernization");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Let the tool figure out your urls (works only for SPO MT):");
            help.AddPreOptionsLine("==========================================================");
            help.AddPreOptionsLine("Using Azure AD app-only:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -t <tenant> -i <your client id> -z <Azure AD domain> -f <PFX file> -x <PFX file password>");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -t contoso -i e5808e8b-6119-44a9-b9d8-9003db04a882 -z conto.onmicrosoft.com  -f apponlycert.pfx -x pwd");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Using app-only:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -t <tenant> -i <your client id> -s <your client secret>");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -t contoso -i 7a5c1615-997a-4059-a784-db2245ec7cc1 -s eOb6h+s805O/V3DOpd0dalec33Q6ShrHlSKkSra1FFw=");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Using credentials:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -t <tenant> -u <your user id> -p <your user password>");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -t contoso -u spadmin@contoso.onmicrosoft.com -p pwd");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Specifying url to your sites and tenant admin (needed for SPO with vanity urls):");
            help.AddPreOptionsLine("================================================================================");
            help.AddPreOptionsLine("Using Azure AD app-only:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -r <wildcard urls> -a <tenant admin site>  -i <your client id> -z <Azure AD domain> -f <PFX file> -x <PFX file password>");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -r \"https://teams.contoso.com/sites/*,https://my.contoso.com/personal/*\" -a https://contoso-admin.contoso.com -i e5808e8b-6119-44a9-b9d8-9003db04a882 -z conto.onmicrosoft.com  -f apponlycert.pfx -x pwd");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Using app-only:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -r <wildcard urls> -a <tenant admin site> -i <your client id> -s <your client secret>");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -r \"https://teams.contoso.com/sites/*,https://my.contoso.com/personal/*\" -a https://contoso-admin.contoso.com -i 7a5c1615-997a-4059-a784-db2245ec7cc1 -s eOb6h+s805O/V3DOpd0dalec33Q6ShrHlSKkSra1FFw=");
            help.AddPreOptionsLine("");
            help.AddPreOptionsLine("Using credentials:");
            help.AddPreOptionsLine("SharePoint.Modernization.Scanner.exe -r <wildcard urls> -a <tenant admin site> -u <your user id> -p <your user password>");
            help.AddPreOptionsLine("e.g. SharePoint.Modernization.Scanner.exe -r \"https://teams.contoso.com/sites/*,https://my.contoso.com/personal/*\" -a https://contoso-admin.contoso.com -u spadmin@contoso.com -p pwd");
            help.AddPreOptionsLine("");
            help.AddOptions(this);
            return help;
        }

        #region Usage
        /// <summary>
        /// Returns the scanner usage information
        /// </summary>
        /// <param name="scanner">Name of the scanner</param>
        /// <returns>HelpText instance holding the help information</returns>
        public HelpText GetUsage(string scanner)
        {
            string version = "";
            try
            {
                Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(UrlToFileName(assembly.EscapedCodeBase));
                version = fvi.FileVersion;
            }
            catch { }

            var help = new HelpText
            {
                Heading = new HeadingInfo(scanner, version),
                Copyright = new CopyrightInfo("SharePoint PnP", DateTime.Now.Year),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true,
                MaximumDisplayWidth = 120
            };
            return help;
        }

        /// <summary>
        /// Converts an URI based file name (file:///c:/temp/file.txt) to a regular path + filename
        /// </summary>
        /// <param name="url">File URI</param>
        /// <returns>File path + name</returns>
        public static string UrlToFileName(string url)
        {
            if (url.StartsWith("file://"))
            {
                return new Uri(url).LocalPath;
            }
            else
            {
                return url;
            }
        }
        #endregion

    }
}

