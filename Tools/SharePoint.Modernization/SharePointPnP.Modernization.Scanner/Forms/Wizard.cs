using AeroWizard;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using SharePoint.Modernization.Scanner.Core;
using SharePoint.Modernization.Scanner.Core.Telemetry;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADAL = Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace SharePoint.Modernization.Scanner.Forms
{
    public partial class Wizard : Form
    {
        private Options options;
        private ADAL.AuthenticationContext authContext = null;
        private const string AuthorityUri = "https://login.microsoftonline.com/common/oauth2/authorize";
        private bool accessTokenObtained = false;

        public Wizard(Options options)
        {
            this.options = options;
            InitializeComponent();
            this.Text = $"SharePoint Modernization Scanner (version {options.CurrentVersion})";

            if (!string.IsNullOrEmpty(options.NewVersion))
            {
                MessageBox.Show($"Scanner version {options.NewVersion} is available. You're currently running {options.CurrentVersion}. Download the latest version of the scanner from {VersionCheck.newVersionDownloadUrl}.", 
                                "Version check", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void wizardPageContainer1_Finished(object sender, EventArgs e)
        {
            // Configure options 

            // Authentication
            if (cmbAuthOption.SelectedIndex == 0)
            {
                // Azure AD App-Only
                options.ClientID = txtAuthAzureADId.Text;
                options.AzureTenant = txtAuthAzureADDomainName.Text;
                options.CertificatePfx = txtAuthAzureADCert.Text;
                options.CertificatePfxPassword = txtAuthAzureADCertPassword.Text;
                ProcessAuthenticationRegion(cmbAuthenticationRegionCert.SelectedIndex);
            }
            else if (cmbAuthOption.SelectedIndex == 1)
            {
                // Azure ACS App-Only
                options.ClientID = txtAzureACSClientId.Text;
                options.ClientSecret = txtAzureADClientSecret.Text;
            }
            else if (cmbAuthOption.SelectedIndex == 2)
            {
                // Username + pwd
                options.User = txtCredentialsUser.Text;
                options.Password = txtCredentialsPassword.Text;
            }

            // Mode
            if (cmbScanMode.SelectedIndex == 0)
            { 
                // Group
                options.Mode = Mode.GroupifyOnly;
            }
            else if (cmbScanMode.SelectedIndex == 1)
            {
                // Lists
                options.Mode = Mode.ListOnly;
            }
            else if (cmbScanMode.SelectedIndex == 2)
            {
                // Home Pages
                options.Mode = Mode.HomePageOnly;
            }
            else if (cmbScanMode.SelectedIndex == 3)
            {
                // Pages
                options.Mode = Mode.PageOnly;
            }
            else if (cmbScanMode.SelectedIndex == 4)
            {
                // Publishing
                options.Mode = Mode.PublishingOnly;
            }
            else if (cmbScanMode.SelectedIndex == 5)
            {
                // Publishing
                options.Mode = Mode.PublishingWithPagesOnly;
            }
            else if (cmbScanMode.SelectedIndex == 6)
            {
                // Workflow
                options.Mode = Mode.WorkflowOnly;
            }
            else if (cmbScanMode.SelectedIndex == 7)
            {
                // InfoPath
                options.Mode = Mode.InfoPathOnly;
            }
            else if (cmbScanMode.SelectedIndex == 8)
            {
                // Blog
                options.Mode = Mode.BlogOnly;
            }
            else if (cmbScanMode.SelectedIndex == 9)
            {
                // Full
                options.Mode = Mode.Full;
            }

            // sites scope
            if (cmbSiteSelectionOption.SelectedIndex == 0)
            {
                // tenant
                options.Tenant = txtSitesTenantName.Text;
            }
            else if (cmbSiteSelectionOption.SelectedIndex == 1)
            {
                // urls
                options.TenantAdminSite = txtSitesAdminCenterUrl.Text;
                IList<string> newlist = new List<string>();
                foreach (var url in lstSitesUrlsToScan.Items)
                {
                    newlist.Add(url.ToString());
                }
                options.Urls = newlist;
            }
            else if (cmbSiteSelectionOption.SelectedIndex == 2)
            {
                options.TenantAdminSite = txtSitesAdminCenterUrl2.Text;
                // CSV file
                options.CsvFile = txtSitesCSVFile.Text;
            }

            // General options
            options.Threads = Int32.Parse(nmThreads.Value.ToString());
            options.SkipUserInformation = tgSkipUser.Checked;
            options.SkipUsageInformation = tgOptionSkipUsage.Checked;
            options.SkipReport = tgSkipExcelReports.Checked;
            options.ExcludeListsOnlyBlockedByOobReasons = tgListBlockedDueToOOB.Checked;
            options.Separator = cmbSeparator.Text;
            options.DisableTelemetry = tgDisableTelemetry.Checked;
            options.DateFormat = cmbDateFormat.Text;

            Close();
        }

        private void wizardPageContainer1_SelectedPageChanged(object sender, EventArgs e)
        {
            string[] headers = new string[] { "" };
            if (wizardPageContainer1.SelectedPage.Text != null)
                headers = wizardPageContainer1.SelectedPage.Text.Split('|');
            headerLabel.Text = headers[0];
            if (headers.Length == 2)
                subHeaderLabel.Text = headers[1];
        }

        private void wizardPage2_Initialize(object sender, AeroWizard.WizardPageInitEventArgs e)
        {
            headerPanel.Visible = topDivider.Visible = true;
        }

        private void authPage_Initialize(object sender, AeroWizard.WizardPageInitEventArgs e)
        {
            if (this.cmbAuthenticationRegionCert.SelectedIndex < 0)
            {
                this.cmbAuthenticationRegionCert.SelectedIndex = 0;
            }
            if (this.cmbAuthOption.SelectedIndex < 0)
            {
                this.cmbAuthOption.SelectedIndex = 0;
            }
            if (this.cmbSiteSelectionOption.SelectedIndex < 0)
            {
                this.cmbSiteSelectionOption.SelectedIndex = 0;
            }
            if (this.cmbScanMode.SelectedIndex < 0)
            {
                this.cmbScanMode.SelectedIndex = 0;
            }
            if (this.cmbSeparator.SelectedIndex < 0)
            {
                this.cmbSeparator.SelectedIndex = 0;
            }
            if (this.cmbDateFormat.SelectedIndex < 0)
            {
                this.cmbDateFormat.SelectedIndex = 0;
            }
        }

        private void btnCertificate_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Certificate files (*.pfx)|*.pfx|All Files (*.*)|*.*"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtAuthAzureADCert.Text = dlg.FileName;
            }                
        }

        private void cmbAuthOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((sender as ComboBox).SelectedIndex == 0)
            {
                AuthOptionUI(true, false, false, false);
            }
            else if ((sender as ComboBox).SelectedIndex == 1)
            {
                AuthOptionUI(false, true, false, false);
            }
            else if ((sender as ComboBox).SelectedIndex == 2)
            {
                AuthOptionUI(false, false, true, false);
            }
            else if ((sender as ComboBox).SelectedIndex == 3)
            {
                AuthOptionUI(false, false, false, true);
            }
        }


        private void AuthOptionUI(bool azureAD, bool azureACS, bool credentials, bool twofactorAuth)
        {
            pnlAzureAD.Visible = azureAD;
            pnlAzureACS.Visible = azureACS;
            pnlCredentials.Visible = credentials;
            pnl2FA.Visible = twofactorAuth;

            if (azureAD)
            {
                txtAuthAzureADId.Focus();
            }
            else if (azureACS)
            {
                txtAzureACSClientId.Focus();
            }
            else if (credentials)
            {
                txtCredentialsUser.Focus();
            }
            else if (twofactorAuth)
            {
                txtSiteFor2FA.Focus();
            }
        }

        private void cmbSiteSelectionOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((sender as ComboBox).SelectedIndex == 0)
            {
                SiteOptionUI(true, false, false);
            }
            else if ((sender as ComboBox).SelectedIndex == 1)
            {
                SiteOptionUI(false, true, false);
            }
            else if ((sender as ComboBox).SelectedIndex == 2)
            {
                SiteOptionUI(false, false, true);
            }
        }

        private void SiteOptionUI(bool tenant, bool wildcard, bool filelist)
        {
            pnlSiteTenant.Visible = tenant;
            pnlSiteWildcard.Visible = wildcard;
            pnlSiteFiles.Visible = filelist;
            UrlsToScanUI();

            if (tenant)
            {
                txtSitesTenantName.Focus();
            }
            else if (wildcard)
            {
                txtSitesUrlToAdd.Focus();
            }
            else if (filelist)
            {
                txtSitesCSVFile.Focus();
            }
        }

        private void btnSelectCSVFile_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All Files (*.*)|*.*"
            };

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtSitesCSVFile.Text = dlg.FileName;
            }
        }

        private void btnSitesAddUrl_Click(object sender, EventArgs e)
        {
            var txt = txtSitesUrlToAdd.Text;

            if (!string.IsNullOrEmpty(txt))
            {
                // check wildcard is at the end 
                if (!ValidateSiteSelectionUrl(txt))
                {
                    MessageBox.Show($"The provided url {txt} is not a valid (wildcard) url");
                }
                else if (this.lstSitesUrlsToScan.Items.Contains(txt))
                {
                    MessageBox.Show($"The provided url {txt} was already added");
                }
                else
                {
                    this.lstSitesUrlsToScan.Items.Add(txt);
                }
            }

            UrlsToScanUI();
        }

        private bool ValidateSiteSelectionUrl(string url)
        {
            if (url.Contains("*") && url.IndexOf("*") != url.Length - 1)
            {
                return false;
            }

            var checkUrl = url.Replace("*", "");
            if (!Uri.TryCreate(checkUrl, UriKind.Absolute, out Uri validUri))
            {
                return false;
            }

            if (!validUri.Scheme.Equals("https", StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private void UrlsToScanUI()
        {
            btnSitesRemoveUrl.Enabled = lstSitesUrlsToScan.SelectedIndex > -1;
            btnSitesClearUrls.Enabled = lstSitesUrlsToScan.Items.Count > 0;
        }

        private void btnSitesClearUrls_Click(object sender, EventArgs e)
        {
            this.lstSitesUrlsToScan.Items.Clear();
            UrlsToScanUI();
        }

        private void lstSitesUrlsToScan_SelectedIndexChanged(object sender, EventArgs e)
        {
            UrlsToScanUI();
        }

        private void btnSitesRemoveUrl_Click(object sender, EventArgs e)
        {
            if (lstSitesUrlsToScan.SelectedIndex > -1)
            {
                lstSitesUrlsToScan.Items.RemoveAt(lstSitesUrlsToScan.SelectedIndex);
            }
            UrlsToScanUI();
        }

        private void scopePage_Initialize(object sender, AeroWizard.WizardPageInitEventArgs e)
        {
            this.pnlSiteFiles.Location = this.pnlSiteWildcard.Location;
            this.pnlSiteTenant.Location = this.pnlSiteWildcard.Location;
        }

        private void cmbScanMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbScanMode.SelectedIndex == 0)
            {
                // Group
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 1)
            {
                // Lists
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = true;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 2)
            {
                // Home Pages
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = true;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 3)
            {
                // Pages
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = true;
                tgModePages.Checked = true;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 4)
            {
                // Publishing
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = true;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 5)
            {
                // Publishing
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = true;
                tgModePublishingDetailed.Checked = true;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 6)
            {
                // Workflow scanning
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = true;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 7)
            {
                // InfoPath scanning
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = true;
                tgModeBlogUsage.Checked = false;
            }
            else if (cmbScanMode.SelectedIndex == 8)
            {
                // Blog scanning
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = false;
                tgModeHomePageOnly.Checked = false;
                tgModePages.Checked = false;
                tgModePublishing.Checked = false;
                tgModePublishingDetailed.Checked = false;
                tgModeClassicWorkflowUsage.Checked = false;
                tgModeInfoPathUsage.Checked = false;
                tgModeBlogUsage.Checked = true;
            }
            else if (cmbScanMode.SelectedIndex == 9)
            {
                // Full
                tgModeGroupConnect.Checked = true;
                tgModeList.Checked = true;
                tgModeHomePageOnly.Checked = true;
                tgModePages.Checked = true;
                tgModePublishing.Checked = true;
                tgModePublishingDetailed.Checked = true;
                tgModeClassicWorkflowUsage.Checked = true;
                tgModeInfoPathUsage.Checked = true;
                tgModeBlogUsage.Checked = true;
            }

            tgListBlockedDueToOOB.Enabled = cmbScanMode.SelectedIndex == 1 || cmbScanMode.SelectedIndex == 9;
            tgExportDetailedWebPartData.Enabled = cmbScanMode.SelectedIndex == 2 || cmbScanMode.SelectedIndex == 3 || cmbScanMode.SelectedIndex == 9;

        }

        private void llblAzureADAuth_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            llblAzureADAuth.LinkVisited = true;
            Process.Start("https://aka.ms/sppnp-modernizationscanner-azureadsetup");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            llblAzureACSHelp.LinkVisited = true;
            Process.Start("https://aka.ms/sppnp-modernizationscanner-azureacssetup");
        }

        private void llblScannerInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            llblScannerInfo.LinkVisited = true;
            Process.Start("https://aka.ms/sppnp-modernizationscanner");
        }

        private void llblModernizationGuidance_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            llblModernizationGuidance.LinkVisited = true;
            Process.Start("https://aka.ms/sppnp-modernize");
        }

        private void PageCommit(object sender, AeroWizard.WizardPageConfirmEventArgs e)
        {

            if (e.Page == this.authPage)
            {
                if (cmbAuthOption.SelectedIndex == 0)
                {
                    // Azure AD App-Only
                    if (string.IsNullOrEmpty(txtAuthAzureADId.Text) ||
                        string.IsNullOrEmpty(txtAuthAzureADDomainName.Text) ||
                        string.IsNullOrEmpty(txtAuthAzureADCert.Text) ||
                        string.IsNullOrEmpty(txtAuthAzureADCertPassword.Text))
                    {
                        MessageBox.Show("Please specify an id, domain, pfx file and pfx file password");
                        e.Cancel = true;
                        return;
                    }

                    if (!System.IO.File.Exists(txtAuthAzureADCert.Text))
                    {
                        MessageBox.Show("Please specify an existing CSV file name");
                        e.Cancel = true;
                        return;
                    }
                }
                else if (cmbAuthOption.SelectedIndex == 1)
                {
                    // Azure ACS App-Only
                    if (string.IsNullOrEmpty(txtAzureACSClientId.Text) ||
                        string.IsNullOrEmpty(txtAzureADClientSecret.Text))
                    {
                        MessageBox.Show("Please specify an id and secret");
                        e.Cancel = true;
                        return;
                    }
                }
                else if (cmbAuthOption.SelectedIndex == 2)
                {
                    // Username + pwd
                    if (string.IsNullOrEmpty(txtCredentialsUser.Text) ||
                        string.IsNullOrEmpty(txtCredentialsPassword.Text))
                    {
                        MessageBox.Show("Please specify an user id and password");
                        e.Cancel = true;
                        return;
                    }
                }
                else if (cmbAuthOption.SelectedIndex == 3)
                {
                    // 2FA
                    if (!this.accessTokenObtained)
                    {
                        MessageBox.Show("Please login before continuing");
                        e.Cancel = true;
                        return;
                    }
                }
            }
            else if (e.Page == this.scopePage)
            {
                if (cmbSiteSelectionOption.SelectedIndex == 0)
                {
                    // tenant
                    if (string.IsNullOrEmpty(txtSitesTenantName.Text))
                    {
                        MessageBox.Show("Please specify an tenant name");
                        e.Cancel = true;
                        return;
                    }
                }
                else if (cmbSiteSelectionOption.SelectedIndex == 1)
                {
                    // urls
                    if (lstSitesUrlsToScan.Items.Count == 0)
                    {
                        MessageBox.Show("Please specify at least 1 (wildcard) url");
                        e.Cancel = true;
                        return;
                    }
                }
                else if (cmbSiteSelectionOption.SelectedIndex == 2)
                {
                    // CSV file
                    if (string.IsNullOrEmpty(txtSitesCSVFile.Text))
                    {
                        MessageBox.Show("Please specify an CSV file name");
                        e.Cancel = true;
                        return;
                    }

                    if (!System.IO.File.Exists(txtSitesCSVFile.Text))
                    {
                        MessageBox.Show("Please specify an existing CSV file name");
                        e.Cancel = true;
                        return;
                    }
                }
            }

        }

        private void llblCSV_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            llblCSV.LinkVisited = true;
            Process.Start("https://aka.ms/sppnp-modernizationscanner-csvinput");
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void txtSitesAdminCenterUrl_TextChanged(object sender, EventArgs e)
        {

        }

        private async void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                await AdalLogin(true);
                this.wizardPageContainer1.NextPage();
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(ex.ToDetailedString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);                
            }
        }

        private async Task AdalLogin(bool forcePrompt)
        {
            var spUri = new Uri($"{txtSiteFor2FA.Text}");

            string resourceUri = spUri.Scheme + "://" + spUri.Authority;
            const string clientId = "9bc3ab49-b65d-410a-85ad-de819febfddc";
            const string redirectUri = "https://oauth.spops.microsoft.com/";

            ADAL.AuthenticationResult authenticationResult;

            if (authContext == null || forcePrompt)
            {
                ADAL.TokenCache cache = new ADAL.TokenCache();
                authContext = new ADAL.AuthenticationContext(AuthorityUri, cache);
            }
            try
            {
                if (forcePrompt) throw new ADAL.AdalSilentTokenAcquisitionException();
                authenticationResult = await authContext.AcquireTokenSilentAsync(resourceUri, clientId);
            }
            catch (ADAL.AdalSilentTokenAcquisitionException)
            {
                authenticationResult = await authContext.AcquireTokenAsync(resourceUri, clientId, new Uri(redirectUri), new PlatformParameters(PromptBehavior.Always, null), ADAL.UserIdentifier.AnyUser, null, null);
            }

            options.AccessToken = authenticationResult.AccessToken;
            accessTokenObtained = true;
        }

        private void ProcessAuthenticationRegion(int index)
        {
            if (index == 0)
            {
                options.AzureEnvironment = OfficeDevPnP.Core.AzureEnvironment.Production;
            }
            else if (index == 1)
            {
                options.AzureEnvironment = OfficeDevPnP.Core.AzureEnvironment.USGovernment;
            }
            else if (index == 2)
            {
                options.AzureEnvironment = OfficeDevPnP.Core.AzureEnvironment.China;
            }
            else if (index == 3)
            {
                options.AzureEnvironment = OfficeDevPnP.Core.AzureEnvironment.Germany;
            }
        }

        private void cmbAuthenticationRegionCert_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((sender as ComboBox).SelectedIndex == 0)
            {
                textBox3.Text = ".sharepoint.com";
            }
            else if ((sender as ComboBox).SelectedIndex == 1)
            {
                textBox3.Text = ".sharepoint.us";
            }
            else if ((sender as ComboBox).SelectedIndex == 2)
            {
                textBox3.Text = ".sharepoint.cn";
            }
            else if ((sender as ComboBox).SelectedIndex == 3)
            {
                textBox3.Text = ".sharepoint.de";
            }

        }
    }
}
