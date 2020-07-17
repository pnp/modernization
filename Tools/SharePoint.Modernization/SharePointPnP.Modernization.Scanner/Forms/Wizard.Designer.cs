namespace SharePoint.Modernization.Scanner.Forms
{
    partial class Wizard
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Wizard));
            this.headerPanel = new System.Windows.Forms.Panel();
            this.headerImage = new System.Windows.Forms.PictureBox();
            this.subHeaderLabel = new System.Windows.Forms.Label();
            this.headerLabel = new System.Windows.Forms.Label();
            this.topDivider = new System.Windows.Forms.Label();
            this.bottomDivider = new System.Windows.Forms.Label();
            this.commandPanel = new System.Windows.Forms.Panel();
            this.backButton = new System.Windows.Forms.Button();
            this.nextButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.wizardPageContainer1 = new AeroWizard.WizardPageContainer();
            this.authPage = new AeroWizard.WizardPage();
            this.llblModernizationGuidance = new System.Windows.Forms.LinkLabel();
            this.llblScannerInfo = new System.Windows.Forms.LinkLabel();
            this.llblAzureACSHelp = new System.Windows.Forms.LinkLabel();
            this.llblAzureADAuth = new System.Windows.Forms.LinkLabel();
            this.pnlAzureAD = new System.Windows.Forms.Panel();
            this.label19 = new System.Windows.Forms.Label();
            this.cmbAuthenticationRegionCert = new System.Windows.Forms.ComboBox();
            this.txtAuthAzureADCertPassword = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtAuthAzureADId = new System.Windows.Forms.TextBox();
            this.btnCertificate = new System.Windows.Forms.Button();
            this.lblAzureADApplicationId = new System.Windows.Forms.Label();
            this.txtAuthAzureADCert = new System.Windows.Forms.TextBox();
            this.lblAzureADDomainName = new System.Windows.Forms.Label();
            this.txtAuthAzureADDomainName = new System.Windows.Forms.TextBox();
            this.lblAzureADCertificateFile = new System.Windows.Forms.Label();
            this.pnlAzureACS = new System.Windows.Forms.Panel();
            this.txtAzureADClientSecret = new System.Windows.Forms.TextBox();
            this.txtAzureACSClientId = new System.Windows.Forms.TextBox();
            this.lblAzureACSSecret = new System.Windows.Forms.Label();
            this.lblAzureACSClientId = new System.Windows.Forms.Label();
            this.pnlCredentials = new System.Windows.Forms.Panel();
            this.label20 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.lblAuthCreds2FAWarning = new System.Windows.Forms.Label();
            this.txtCredentialsPassword = new System.Windows.Forms.TextBox();
            this.txtCredentialsUser = new System.Windows.Forms.TextBox();
            this.lblCredsPassword = new System.Windows.Forms.Label();
            this.lblCredsUserName = new System.Windows.Forms.Label();
            this.pnl2FA = new System.Windows.Forms.Panel();
            this.btnLogin = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.txtSiteFor2FA = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.cmbAuthOption = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.scopePage = new AeroWizard.WizardPage();
            this.pnlSiteFiles = new System.Windows.Forms.Panel();
            this.txtSitesAdminCenterUrl2 = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.llblCSV = new System.Windows.Forms.LinkLabel();
            this.btnSelectCSVFile = new System.Windows.Forms.Button();
            this.txtSitesCSVFile = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbSiteSelectionOption = new System.Windows.Forms.ComboBox();
            this.lblSiteSelectionOption = new System.Windows.Forms.Label();
            this.pnlSiteTenant = new System.Windows.Forms.Panel();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.lblSiteTenantName = new System.Windows.Forms.Label();
            this.txtSitesTenantName = new System.Windows.Forms.TextBox();
            this.pnlSiteWildcard = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.txtSitesAdminCenterUrl = new System.Windows.Forms.TextBox();
            this.btnSitesClearUrls = new System.Windows.Forms.Button();
            this.btnSitesRemoveUrl = new System.Windows.Forms.Button();
            this.btnSitesAddUrl = new System.Windows.Forms.Button();
            this.lstSitesUrlsToScan = new System.Windows.Forms.ListBox();
            this.txtSitesUrlToAdd = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.modePage = new AeroWizard.WizardPage();
            this.tgModeHomePageOnly = new System.Windows.Forms.CheckBox();
            this.tgModeBlogUsage = new System.Windows.Forms.CheckBox();
            this.tgModeInfoPathUsage = new System.Windows.Forms.CheckBox();
            this.tgModeClassicWorkflowUsage = new System.Windows.Forms.CheckBox();
            this.tgModePublishingDetailed = new System.Windows.Forms.CheckBox();
            this.tgModePublishing = new System.Windows.Forms.CheckBox();
            this.tgModePages = new System.Windows.Forms.CheckBox();
            this.tgModeList = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tgModeGroupConnect = new System.Windows.Forms.CheckBox();
            this.cmbScanMode = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.optionsPage = new AeroWizard.WizardPage();
            this.cmbDateFormat = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.tgDisableTelemetry = new System.Windows.Forms.CheckBox();
            this.cmbSeparator = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.tgSkipExcelReports = new System.Windows.Forms.CheckBox();
            this.tgExportDetailedWebPartData = new System.Windows.Forms.CheckBox();
            this.tgListBlockedDueToOOB = new System.Windows.Forms.CheckBox();
            this.tgSkipUser = new System.Windows.Forms.CheckBox();
            this.tgOptionSkipUsage = new System.Windows.Forms.CheckBox();
            this.nmThreads = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.ttAuthentication = new System.Windows.Forms.ToolTip(this.components);
            this.headerPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.headerImage)).BeginInit();
            this.commandPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).BeginInit();
            this.wizardPageContainer1.SuspendLayout();
            this.authPage.SuspendLayout();
            this.pnlAzureAD.SuspendLayout();
            this.pnlAzureACS.SuspendLayout();
            this.pnlCredentials.SuspendLayout();
            this.pnl2FA.SuspendLayout();
            this.scopePage.SuspendLayout();
            this.pnlSiteFiles.SuspendLayout();
            this.pnlSiteTenant.SuspendLayout();
            this.pnlSiteWildcard.SuspendLayout();
            this.modePage.SuspendLayout();
            this.optionsPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmThreads)).BeginInit();
            this.SuspendLayout();
            // 
            // headerPanel
            // 
            this.headerPanel.BackColor = System.Drawing.SystemColors.Window;
            this.headerPanel.Controls.Add(this.headerImage);
            this.headerPanel.Controls.Add(this.subHeaderLabel);
            this.headerPanel.Controls.Add(this.headerLabel);
            this.headerPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.headerPanel.Location = new System.Drawing.Point(0, 0);
            this.headerPanel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.headerPanel.Name = "headerPanel";
            this.headerPanel.Size = new System.Drawing.Size(720, 88);
            this.headerPanel.TabIndex = 2;
            // 
            // headerImage
            // 
            this.headerImage.Image = ((System.Drawing.Image)(resources.GetObject("headerImage.Image")));
            this.headerImage.Location = new System.Drawing.Point(564, 6);
            this.headerImage.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.headerImage.Name = "headerImage";
            this.headerImage.Size = new System.Drawing.Size(153, 75);
            this.headerImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.headerImage.TabIndex = 1;
            this.headerImage.TabStop = false;
            // 
            // subHeaderLabel
            // 
            this.subHeaderLabel.AutoSize = true;
            this.subHeaderLabel.Location = new System.Drawing.Point(18, 48);
            this.subHeaderLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.subHeaderLabel.Name = "subHeaderLabel";
            this.subHeaderLabel.Size = new System.Drawing.Size(382, 20);
            this.subHeaderLabel.TabIndex = 0;
            this.subHeaderLabel.Text = "Use this wizard to configure the options for your scan";
            // 
            // headerLabel
            // 
            this.headerLabel.AutoSize = true;
            this.headerLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headerLabel.Location = new System.Drawing.Point(18, 17);
            this.headerLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.headerLabel.Name = "headerLabel";
            this.headerLabel.Size = new System.Drawing.Size(300, 20);
            this.headerLabel.TabIndex = 0;
            this.headerLabel.Text = "SharePoint Modernization Scanner";
            // 
            // topDivider
            // 
            this.topDivider.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.topDivider.Dock = System.Windows.Forms.DockStyle.Top;
            this.topDivider.Location = new System.Drawing.Point(0, 88);
            this.topDivider.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.topDivider.Name = "topDivider";
            this.topDivider.Size = new System.Drawing.Size(720, 3);
            this.topDivider.TabIndex = 3;
            // 
            // bottomDivider
            // 
            this.bottomDivider.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.bottomDivider.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.bottomDivider.Enabled = false;
            this.bottomDivider.Location = new System.Drawing.Point(0, 497);
            this.bottomDivider.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.bottomDivider.Name = "bottomDivider";
            this.bottomDivider.Size = new System.Drawing.Size(720, 3);
            this.bottomDivider.TabIndex = 4;
            // 
            // commandPanel
            // 
            this.commandPanel.Controls.Add(this.backButton);
            this.commandPanel.Controls.Add(this.nextButton);
            this.commandPanel.Controls.Add(this.cancelButton);
            this.commandPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.commandPanel.Location = new System.Drawing.Point(0, 500);
            this.commandPanel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.commandPanel.Name = "commandPanel";
            this.commandPanel.Size = new System.Drawing.Size(720, 62);
            this.commandPanel.TabIndex = 5;
            // 
            // backButton
            // 
            this.backButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.backButton.Location = new System.Drawing.Point(256, 14);
            this.backButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.backButton.Name = "backButton";
            this.backButton.Size = new System.Drawing.Size(146, 35);
            this.backButton.TabIndex = 2;
            this.backButton.Tag = AeroWizard.WizardCommandButtonState.Disabled;
            this.backButton.Text = "< Back";
            this.backButton.UseVisualStyleBackColor = true;
            // 
            // nextButton
            // 
            this.nextButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.nextButton.Location = new System.Drawing.Point(405, 14);
            this.nextButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(146, 35);
            this.nextButton.TabIndex = 3;
            this.nextButton.Tag = AeroWizard.WizardCommandButtonState.Enabled;
            this.nextButton.Text = "Next >";
            this.nextButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(560, 14);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(146, 35);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Tag = AeroWizard.WizardCommandButtonState.Disabled;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // wizardPageContainer1
            // 
            this.wizardPageContainer1.BackButton = this.backButton;
            this.wizardPageContainer1.BackButtonText = "< Back";
            this.wizardPageContainer1.CancelButton = this.cancelButton;
            this.wizardPageContainer1.CancelButtonText = "Cancel";
            this.wizardPageContainer1.Controls.Add(this.authPage);
            this.wizardPageContainer1.Controls.Add(this.scopePage);
            this.wizardPageContainer1.Controls.Add(this.modePage);
            this.wizardPageContainer1.Controls.Add(this.optionsPage);
            this.wizardPageContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wizardPageContainer1.FinishButtonText = "&Start scan";
            this.wizardPageContainer1.Location = new System.Drawing.Point(0, 91);
            this.wizardPageContainer1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.wizardPageContainer1.Name = "wizardPageContainer1";
            this.wizardPageContainer1.NextButton = this.nextButton;
            this.wizardPageContainer1.Pages.Add(this.authPage);
            this.wizardPageContainer1.Pages.Add(this.scopePage);
            this.wizardPageContainer1.Pages.Add(this.modePage);
            this.wizardPageContainer1.Pages.Add(this.optionsPage);
            this.wizardPageContainer1.ShowProgressInTaskbarIcon = true;
            this.wizardPageContainer1.Size = new System.Drawing.Size(720, 406);
            this.wizardPageContainer1.TabIndex = 0;
            this.wizardPageContainer1.Finished += new System.EventHandler(this.wizardPageContainer1_Finished);
            this.wizardPageContainer1.SelectedPageChanged += new System.EventHandler(this.wizardPageContainer1_SelectedPageChanged);
            // 
            // authPage
            // 
            this.authPage.Controls.Add(this.llblModernizationGuidance);
            this.authPage.Controls.Add(this.llblScannerInfo);
            this.authPage.Controls.Add(this.llblAzureACSHelp);
            this.authPage.Controls.Add(this.llblAzureADAuth);
            this.authPage.Controls.Add(this.pnlAzureAD);
            this.authPage.Controls.Add(this.pnlAzureACS);
            this.authPage.Controls.Add(this.pnlCredentials);
            this.authPage.Controls.Add(this.pnl2FA);
            this.authPage.Controls.Add(this.cmbAuthOption);
            this.authPage.Controls.Add(this.label1);
            this.authPage.Name = "authPage";
            this.authPage.NextPage = this.scopePage;
            this.authPage.Size = new System.Drawing.Size(720, 406);
            this.authPage.TabIndex = 4;
            this.authPage.Text = "Configure authentication for the scan";
            this.authPage.Commit += new System.EventHandler<AeroWizard.WizardPageConfirmEventArgs>(this.PageCommit);
            this.authPage.Initialize += new System.EventHandler<AeroWizard.WizardPageInitEventArgs>(this.authPage_Initialize);
            // 
            // llblModernizationGuidance
            // 
            this.llblModernizationGuidance.AutoSize = true;
            this.llblModernizationGuidance.Location = new System.Drawing.Point(382, 360);
            this.llblModernizationGuidance.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llblModernizationGuidance.Name = "llblModernizationGuidance";
            this.llblModernizationGuidance.Size = new System.Drawing.Size(261, 20);
            this.llblModernizationGuidance.TabIndex = 15;
            this.llblModernizationGuidance.TabStop = true;
            this.llblModernizationGuidance.Text = "SharePoint Modernization guidance";
            this.llblModernizationGuidance.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblModernizationGuidance_LinkClicked);
            // 
            // llblScannerInfo
            // 
            this.llblScannerInfo.AutoSize = true;
            this.llblScannerInfo.Location = new System.Drawing.Point(28, 360);
            this.llblScannerInfo.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llblScannerInfo.Name = "llblScannerInfo";
            this.llblScannerInfo.Size = new System.Drawing.Size(314, 20);
            this.llblScannerInfo.TabIndex = 14;
            this.llblScannerInfo.TabStop = true;
            this.llblScannerInfo.Text = "SharePoint Modernization Scanner readme";
            this.llblScannerInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblScannerInfo_LinkClicked);
            // 
            // llblAzureACSHelp
            // 
            this.llblAzureACSHelp.AutoSize = true;
            this.llblAzureACSHelp.LinkArea = new System.Windows.Forms.LinkArea(6, 4);
            this.llblAzureACSHelp.Location = new System.Drawing.Point(27, 321);
            this.llblAzureACSHelp.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llblAzureACSHelp.Name = "llblAzureACSHelp";
            this.llblAzureACSHelp.Size = new System.Drawing.Size(594, 24);
            this.llblAzureACSHelp.TabIndex = 13;
            this.llblAzureACSHelp.TabStop = true;
            this.llblAzureACSHelp.Text = "Click here for more information on setting up Azure ACS App Only authentication";
            this.llblAzureACSHelp.UseCompatibleTextRendering = true;
            this.llblAzureACSHelp.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // llblAzureADAuth
            // 
            this.llblAzureADAuth.AutoSize = true;
            this.llblAzureADAuth.LinkArea = new System.Windows.Forms.LinkArea(6, 4);
            this.llblAzureADAuth.Location = new System.Drawing.Point(27, 295);
            this.llblAzureADAuth.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llblAzureADAuth.Name = "llblAzureADAuth";
            this.llblAzureADAuth.Size = new System.Drawing.Size(583, 24);
            this.llblAzureADAuth.TabIndex = 12;
            this.llblAzureADAuth.TabStop = true;
            this.llblAzureADAuth.Text = "Click here for more information on setting up Azure AD App Only authentication";
            this.llblAzureADAuth.UseCompatibleTextRendering = true;
            this.llblAzureADAuth.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblAzureADAuth_LinkClicked);
            // 
            // pnlAzureAD
            // 
            this.pnlAzureAD.Controls.Add(this.label19);
            this.pnlAzureAD.Controls.Add(this.cmbAuthenticationRegionCert);
            this.pnlAzureAD.Controls.Add(this.txtAuthAzureADCertPassword);
            this.pnlAzureAD.Controls.Add(this.label8);
            this.pnlAzureAD.Controls.Add(this.txtAuthAzureADId);
            this.pnlAzureAD.Controls.Add(this.btnCertificate);
            this.pnlAzureAD.Controls.Add(this.lblAzureADApplicationId);
            this.pnlAzureAD.Controls.Add(this.txtAuthAzureADCert);
            this.pnlAzureAD.Controls.Add(this.lblAzureADDomainName);
            this.pnlAzureAD.Controls.Add(this.txtAuthAzureADDomainName);
            this.pnlAzureAD.Controls.Add(this.lblAzureADCertificateFile);
            this.pnlAzureAD.Location = new System.Drawing.Point(3, 64);
            this.pnlAzureAD.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlAzureAD.Name = "pnlAzureAD";
            this.pnlAzureAD.Size = new System.Drawing.Size(704, 225);
            this.pnlAzureAD.TabIndex = 9;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(18, 191);
            this.label19.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(164, 20);
            this.label19.TabIndex = 12;
            this.label19.Text = "Authentication region:";
            // 
            // cmbAuthenticationRegionCert
            // 
            this.cmbAuthenticationRegionCert.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAuthenticationRegionCert.FormattingEnabled = true;
            this.cmbAuthenticationRegionCert.Items.AddRange(new object[] {
            "Production",
            "USGovernment",
            "China",
            "Germany"});
            this.cmbAuthenticationRegionCert.Location = new System.Drawing.Point(252, 189);
            this.cmbAuthenticationRegionCert.Name = "cmbAuthenticationRegionCert";
            this.cmbAuthenticationRegionCert.Size = new System.Drawing.Size(436, 28);
            this.cmbAuthenticationRegionCert.TabIndex = 11;
            this.cmbAuthenticationRegionCert.SelectedIndexChanged += new System.EventHandler(this.cmbAuthenticationRegionCert_SelectedIndexChanged);
            // 
            // txtAuthAzureADCertPassword
            // 
            this.txtAuthAzureADCertPassword.Location = new System.Drawing.Point(252, 151);
            this.txtAuthAzureADCertPassword.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAuthAzureADCertPassword.Name = "txtAuthAzureADCertPassword";
            this.txtAuthAzureADCertPassword.Size = new System.Drawing.Size(436, 26);
            this.txtAuthAzureADCertPassword.TabIndex = 10;
            this.ttAuthentication.SetToolTip(this.txtAuthAzureADCertPassword, "Password used to protect the PFX file");
            this.txtAuthAzureADCertPassword.UseSystemPasswordChar = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 155);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(191, 20);
            this.label8.TabIndex = 9;
            this.label8.Text = "Password for the PFX file:";
            // 
            // txtAuthAzureADId
            // 
            this.txtAuthAzureADId.Location = new System.Drawing.Point(252, 22);
            this.txtAuthAzureADId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAuthAzureADId.Name = "txtAuthAzureADId";
            this.txtAuthAzureADId.Size = new System.Drawing.Size(436, 26);
            this.txtAuthAzureADId.TabIndex = 5;
            this.ttAuthentication.SetToolTip(this.txtAuthAzureADId, "Application ID (guid) for the Azure AD app used to get app-only access");
            // 
            // btnCertificate
            // 
            this.btnCertificate.Location = new System.Drawing.Point(652, 106);
            this.btnCertificate.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCertificate.Name = "btnCertificate";
            this.btnCertificate.Size = new System.Drawing.Size(38, 35);
            this.btnCertificate.TabIndex = 8;
            this.btnCertificate.Text = "...";
            this.btnCertificate.UseVisualStyleBackColor = true;
            this.btnCertificate.Click += new System.EventHandler(this.btnCertificate_Click);
            // 
            // lblAzureADApplicationId
            // 
            this.lblAzureADApplicationId.AutoSize = true;
            this.lblAzureADApplicationId.Location = new System.Drawing.Point(15, 22);
            this.lblAzureADApplicationId.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAzureADApplicationId.Name = "lblAzureADApplicationId";
            this.lblAzureADApplicationId.Size = new System.Drawing.Size(185, 20);
            this.lblAzureADApplicationId.TabIndex = 2;
            this.lblAzureADApplicationId.Text = "Azure AD Application ID:";
            // 
            // txtAuthAzureADCert
            // 
            this.txtAuthAzureADCert.Location = new System.Drawing.Point(252, 109);
            this.txtAuthAzureADCert.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAuthAzureADCert.Name = "txtAuthAzureADCert";
            this.txtAuthAzureADCert.Size = new System.Drawing.Size(390, 26);
            this.txtAuthAzureADCert.TabIndex = 7;
            this.ttAuthentication.SetToolTip(this.txtAuthAzureADCert, "Full path the PFX file holding the certificate used to access the Azure AD app");
            // 
            // lblAzureADDomainName
            // 
            this.lblAzureADDomainName.AutoSize = true;
            this.lblAzureADDomainName.Location = new System.Drawing.Point(15, 68);
            this.lblAzureADDomainName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAzureADDomainName.Name = "lblAzureADDomainName";
            this.lblAzureADDomainName.Size = new System.Drawing.Size(185, 20);
            this.lblAzureADDomainName.TabIndex = 3;
            this.lblAzureADDomainName.Text = "Azure AD Domain name:";
            // 
            // txtAuthAzureADDomainName
            // 
            this.txtAuthAzureADDomainName.Location = new System.Drawing.Point(252, 68);
            this.txtAuthAzureADDomainName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAuthAzureADDomainName.Name = "txtAuthAzureADDomainName";
            this.txtAuthAzureADDomainName.Size = new System.Drawing.Size(436, 26);
            this.txtAuthAzureADDomainName.TabIndex = 6;
            this.ttAuthentication.SetToolTip(this.txtAuthAzureADDomainName, "Azure AD domain (e.g. contoso.onmicrosoft.com) where you\'ve created the Azure AD " +
        "for SharePoint app only access");
            // 
            // lblAzureADCertificateFile
            // 
            this.lblAzureADCertificateFile.AutoSize = true;
            this.lblAzureADCertificateFile.Location = new System.Drawing.Point(15, 109);
            this.lblAzureADCertificateFile.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAzureADCertificateFile.Name = "lblAzureADCertificateFile";
            this.lblAzureADCertificateFile.Size = new System.Drawing.Size(154, 20);
            this.lblAzureADCertificateFile.TabIndex = 4;
            this.lblAzureADCertificateFile.Text = "Certificate (PFX) file:";
            // 
            // pnlAzureACS
            // 
            this.pnlAzureACS.Controls.Add(this.txtAzureADClientSecret);
            this.pnlAzureACS.Controls.Add(this.txtAzureACSClientId);
            this.pnlAzureACS.Controls.Add(this.lblAzureACSSecret);
            this.pnlAzureACS.Controls.Add(this.lblAzureACSClientId);
            this.pnlAzureACS.Location = new System.Drawing.Point(3, 64);
            this.pnlAzureACS.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlAzureACS.Name = "pnlAzureACS";
            this.pnlAzureACS.Size = new System.Drawing.Size(704, 115);
            this.pnlAzureACS.TabIndex = 10;
            // 
            // txtAzureADClientSecret
            // 
            this.txtAzureADClientSecret.Location = new System.Drawing.Point(252, 62);
            this.txtAzureADClientSecret.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAzureADClientSecret.Name = "txtAzureADClientSecret";
            this.txtAzureADClientSecret.Size = new System.Drawing.Size(436, 26);
            this.txtAzureADClientSecret.TabIndex = 3;
            this.ttAuthentication.SetToolTip(this.txtAzureADClientSecret, "Secret (long string) of the created app principal");
            // 
            // txtAzureACSClientId
            // 
            this.txtAzureACSClientId.Location = new System.Drawing.Point(252, 20);
            this.txtAzureACSClientId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtAzureACSClientId.Name = "txtAzureACSClientId";
            this.txtAzureACSClientId.Size = new System.Drawing.Size(436, 26);
            this.txtAzureACSClientId.TabIndex = 2;
            this.ttAuthentication.SetToolTip(this.txtAzureACSClientId, "Client ID (guid) of the created app principal");
            // 
            // lblAzureACSSecret
            // 
            this.lblAzureACSSecret.AutoSize = true;
            this.lblAzureACSSecret.Location = new System.Drawing.Point(20, 66);
            this.lblAzureACSSecret.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAzureACSSecret.Name = "lblAzureACSSecret";
            this.lblAzureACSSecret.Size = new System.Drawing.Size(187, 20);
            this.lblAzureACSSecret.TabIndex = 1;
            this.lblAzureACSSecret.Text = "Azure ACS Client Secret:";
            // 
            // lblAzureACSClientId
            // 
            this.lblAzureACSClientId.AutoSize = true;
            this.lblAzureACSClientId.Location = new System.Drawing.Point(20, 20);
            this.lblAzureACSClientId.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAzureACSClientId.Name = "lblAzureACSClientId";
            this.lblAzureACSClientId.Size = new System.Drawing.Size(157, 20);
            this.lblAzureACSClientId.TabIndex = 0;
            this.lblAzureACSClientId.Text = "Azure ACS Client ID:";
            // 
            // pnlCredentials
            // 
            this.pnlCredentials.Controls.Add(this.label12);
            this.pnlCredentials.Controls.Add(this.lblAuthCreds2FAWarning);
            this.pnlCredentials.Controls.Add(this.txtCredentialsPassword);
            this.pnlCredentials.Controls.Add(this.txtCredentialsUser);
            this.pnlCredentials.Controls.Add(this.lblCredsPassword);
            this.pnlCredentials.Controls.Add(this.lblCredsUserName);
            this.pnlCredentials.Location = new System.Drawing.Point(5, 64);
            this.pnlCredentials.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlCredentials.Name = "pnlCredentials";
            this.pnlCredentials.Size = new System.Drawing.Size(704, 219);
            this.pnlCredentials.TabIndex = 11;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(15, 144);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(598, 20);
            this.label12.TabIndex = 5;
            this.label12.Text = "Note: - The provided account needs to have owner permissions to the scanned sites" +
    "";
            // 
            // lblAuthCreds2FAWarning
            // 
            this.lblAuthCreds2FAWarning.AutoSize = true;
            this.lblAuthCreds2FAWarning.Location = new System.Drawing.Point(60, 172);
            this.lblAuthCreds2FAWarning.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAuthCreds2FAWarning.Name = "lblAuthCreds2FAWarning";
            this.lblAuthCreds2FAWarning.Size = new System.Drawing.Size(429, 20);
            this.lblAuthCreds2FAWarning.TabIndex = 4;
            this.lblAuthCreds2FAWarning.Text = "- Credential fails when multi-factor authentication is required";
            // 
            // txtCredentialsPassword
            // 
            this.txtCredentialsPassword.Location = new System.Drawing.Point(252, 58);
            this.txtCredentialsPassword.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtCredentialsPassword.Name = "txtCredentialsPassword";
            this.txtCredentialsPassword.Size = new System.Drawing.Size(436, 26);
            this.txtCredentialsPassword.TabIndex = 3;
            this.txtCredentialsPassword.UseSystemPasswordChar = true;
            // 
            // txtCredentialsUser
            // 
            this.txtCredentialsUser.Location = new System.Drawing.Point(252, 18);
            this.txtCredentialsUser.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtCredentialsUser.Name = "txtCredentialsUser";
            this.txtCredentialsUser.Size = new System.Drawing.Size(436, 26);
            this.txtCredentialsUser.TabIndex = 2;
            this.ttAuthentication.SetToolTip(this.txtCredentialsUser, "User id (e.g. joe@contoso.com or kate@contoso.onmicrosoft.com)");
            // 
            // lblCredsPassword
            // 
            this.lblCredsPassword.AutoSize = true;
            this.lblCredsPassword.Location = new System.Drawing.Point(14, 58);
            this.lblCredsPassword.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCredsPassword.Name = "lblCredsPassword";
            this.lblCredsPassword.Size = new System.Drawing.Size(82, 20);
            this.lblCredsPassword.TabIndex = 1;
            this.lblCredsPassword.Text = "Password:";
            // 
            // lblCredsUserName
            // 
            this.lblCredsUserName.AutoSize = true;
            this.lblCredsUserName.Location = new System.Drawing.Point(14, 18);
            this.lblCredsUserName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCredsUserName.Name = "lblCredsUserName";
            this.lblCredsUserName.Size = new System.Drawing.Size(219, 20);
            this.lblCredsUserName.TabIndex = 0;
            this.lblCredsUserName.Text = "User (e.g. joe@contoso.com):";
            // 
            // pnl2FA
            // 
            this.pnl2FA.Controls.Add(this.btnLogin);
            this.pnl2FA.Controls.Add(this.label14);
            this.pnl2FA.Controls.Add(this.label15);
            this.pnl2FA.Controls.Add(this.txtSiteFor2FA);
            this.pnl2FA.Controls.Add(this.label17);
            this.pnl2FA.Location = new System.Drawing.Point(9, 65);
            this.pnl2FA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnl2FA.Name = "pnl2FA";
            this.pnl2FA.Size = new System.Drawing.Size(704, 154);
            this.pnl2FA.TabIndex = 12;
            // 
            // btnLogin
            // 
            this.btnLogin.Location = new System.Drawing.Point(592, 15);
            this.btnLogin.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(96, 35);
            this.btnLogin.TabIndex = 6;
            this.btnLogin.Text = "Login";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(15, 63);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(492, 20);
            this.label14.TabIndex = 5;
            this.label14.Text = "Note: - The scan can run for the lifetime of the received access token";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(60, 89);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(609, 20);
            this.label15.TabIndex = 4;
            this.label15.Text = "- Use this for scanning just a few sites, for more sites use one of the App-Only " +
    "models";
            // 
            // txtSiteFor2FA
            // 
            this.txtSiteFor2FA.Location = new System.Drawing.Point(250, 18);
            this.txtSiteFor2FA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSiteFor2FA.Name = "txtSiteFor2FA";
            this.txtSiteFor2FA.Size = new System.Drawing.Size(331, 26);
            this.txtSiteFor2FA.TabIndex = 2;
            this.ttAuthentication.SetToolTip(this.txtSiteFor2FA, "The multi factor login process requires a site that you can access today");
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(14, 18);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(206, 20);
            this.label17.TabIndex = 0;
            this.label17.Text = "Url of a site you can access:";
            // 
            // cmbAuthOption
            // 
            this.cmbAuthOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAuthOption.FormattingEnabled = true;
            this.cmbAuthOption.Items.AddRange(new object[] {
            "Azure AD App Only (preferred)",
            "Azure ACS App Only",
            "Username and password",
            "Username and password (Multi Factor Auth)"});
            this.cmbAuthOption.Location = new System.Drawing.Point(264, 26);
            this.cmbAuthOption.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbAuthOption.Name = "cmbAuthOption";
            this.cmbAuthOption.Size = new System.Drawing.Size(436, 28);
            this.cmbAuthOption.TabIndex = 1;
            this.cmbAuthOption.SelectedIndexChanged += new System.EventHandler(this.cmbAuthOption_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(229, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selected authentication option:";
            // 
            // scopePage
            // 
            this.scopePage.Controls.Add(this.pnlSiteFiles);
            this.scopePage.Controls.Add(this.cmbSiteSelectionOption);
            this.scopePage.Controls.Add(this.lblSiteSelectionOption);
            this.scopePage.Controls.Add(this.pnlSiteTenant);
            this.scopePage.Controls.Add(this.pnlSiteWildcard);
            this.scopePage.Name = "scopePage";
            this.scopePage.NextPage = this.modePage;
            this.scopePage.Size = new System.Drawing.Size(720, 406);
            this.scopePage.TabIndex = 1;
            this.scopePage.Tag = "";
            this.scopePage.Text = "Configure the sites to scan";
            this.scopePage.Commit += new System.EventHandler<AeroWizard.WizardPageConfirmEventArgs>(this.PageCommit);
            this.scopePage.Initialize += new System.EventHandler<AeroWizard.WizardPageInitEventArgs>(this.scopePage_Initialize);
            // 
            // pnlSiteFiles
            // 
            this.pnlSiteFiles.Controls.Add(this.txtSitesAdminCenterUrl2);
            this.pnlSiteFiles.Controls.Add(this.label18);
            this.pnlSiteFiles.Controls.Add(this.label16);
            this.pnlSiteFiles.Controls.Add(this.llblCSV);
            this.pnlSiteFiles.Controls.Add(this.btnSelectCSVFile);
            this.pnlSiteFiles.Controls.Add(this.txtSitesCSVFile);
            this.pnlSiteFiles.Controls.Add(this.label2);
            this.pnlSiteFiles.Location = new System.Drawing.Point(693, 96);
            this.pnlSiteFiles.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlSiteFiles.Name = "pnlSiteFiles";
            this.pnlSiteFiles.Size = new System.Drawing.Size(694, 197);
            this.pnlSiteFiles.TabIndex = 4;
            // 
            // txtSitesAdminCenterUrl2
            // 
            this.txtSitesAdminCenterUrl2.Location = new System.Drawing.Point(133, 158);
            this.txtSitesAdminCenterUrl2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSitesAdminCenterUrl2.Name = "txtSitesAdminCenterUrl2";
            this.txtSitesAdminCenterUrl2.Size = new System.Drawing.Size(529, 26);
            this.txtSitesAdminCenterUrl2.TabIndex = 12;
            this.ttAuthentication.SetToolTip(this.txtSitesAdminCenterUrl2, "SPO Admin center url (e.g. https://spoadmin.contoso.com). Only needed when using " +
        "SPO with vanity urls (= urls that not end on .sharepoint.com)");
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(14, 161);
            this.label18.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(111, 20);
            this.label18.TabIndex = 11;
            this.label18.Text = "Admin  center:";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(14, 132);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(553, 20);
            this.label16.TabIndex = 10;
            this.label16.Text = "Only when using vanity url\'s specify the full SharePoint Online admin center url";
            // 
            // llblCSV
            // 
            this.llblCSV.AutoSize = true;
            this.llblCSV.Location = new System.Drawing.Point(236, 72);
            this.llblCSV.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.llblCSV.Name = "llblCSV";
            this.llblCSV.Size = new System.Drawing.Size(90, 20);
            this.llblCSV.TabIndex = 3;
            this.llblCSV.TabStop = true;
            this.llblCSV.Text = "Learn more";
            this.llblCSV.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblCSV_LinkClicked);
            // 
            // btnSelectCSVFile
            // 
            this.btnSelectCSVFile.Location = new System.Drawing.Point(640, 20);
            this.btnSelectCSVFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSelectCSVFile.Name = "btnSelectCSVFile";
            this.btnSelectCSVFile.Size = new System.Drawing.Size(36, 35);
            this.btnSelectCSVFile.TabIndex = 2;
            this.btnSelectCSVFile.Text = "...";
            this.btnSelectCSVFile.UseVisualStyleBackColor = true;
            this.btnSelectCSVFile.Click += new System.EventHandler(this.btnSelectCSVFile_Click);
            // 
            // txtSitesCSVFile
            // 
            this.txtSitesCSVFile.Location = new System.Drawing.Point(236, 23);
            this.txtSitesCSVFile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSitesCSVFile.Name = "txtSitesCSVFile";
            this.txtSitesCSVFile.Size = new System.Drawing.Size(400, 26);
            this.txtSitesCSVFile.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 28);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 20);
            this.label2.TabIndex = 0;
            this.label2.Text = "Select CSV sites file: ";
            // 
            // cmbSiteSelectionOption
            // 
            this.cmbSiteSelectionOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSiteSelectionOption.FormattingEnabled = true;
            this.cmbSiteSelectionOption.Items.AddRange(new object[] {
            "Complete tenant (not possible with vanity urls)",
            "Selected site collections",
            "CSV file listing sites to scan"});
            this.cmbSiteSelectionOption.Location = new System.Drawing.Point(256, 25);
            this.cmbSiteSelectionOption.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbSiteSelectionOption.Name = "cmbSiteSelectionOption";
            this.cmbSiteSelectionOption.Size = new System.Drawing.Size(446, 28);
            this.cmbSiteSelectionOption.TabIndex = 1;
            this.cmbSiteSelectionOption.SelectedIndexChanged += new System.EventHandler(this.cmbSiteSelectionOption_SelectedIndexChanged);
            // 
            // lblSiteSelectionOption
            // 
            this.lblSiteSelectionOption.AutoSize = true;
            this.lblSiteSelectionOption.Location = new System.Drawing.Point(22, 25);
            this.lblSiteSelectionOption.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSiteSelectionOption.Name = "lblSiteSelectionOption";
            this.lblSiteSelectionOption.Size = new System.Drawing.Size(156, 20);
            this.lblSiteSelectionOption.TabIndex = 0;
            this.lblSiteSelectionOption.Text = "Site selection option:";
            // 
            // pnlSiteTenant
            // 
            this.pnlSiteTenant.Controls.Add(this.textBox3);
            this.pnlSiteTenant.Controls.Add(this.textBox2);
            this.pnlSiteTenant.Controls.Add(this.lblSiteTenantName);
            this.pnlSiteTenant.Controls.Add(this.txtSitesTenantName);
            this.pnlSiteTenant.Location = new System.Drawing.Point(693, 306);
            this.pnlSiteTenant.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlSiteTenant.Name = "pnlSiteTenant";
            this.pnlSiteTenant.Size = new System.Drawing.Size(687, 75);
            this.pnlSiteTenant.TabIndex = 2;
            // 
            // textBox3
            // 
            this.textBox3.Enabled = false;
            this.textBox3.Location = new System.Drawing.Point(554, 22);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(121, 26);
            this.textBox3.TabIndex = 3;
            this.textBox3.Text = ".sharepoint.com";
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(236, 22);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(64, 26);
            this.textBox2.TabIndex = 2;
            this.textBox2.Text = "https://";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // lblSiteTenantName
            // 
            this.lblSiteTenantName.AutoSize = true;
            this.lblSiteTenantName.Location = new System.Drawing.Point(9, 26);
            this.lblSiteTenantName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSiteTenantName.Name = "lblSiteTenantName";
            this.lblSiteTenantName.Size = new System.Drawing.Size(141, 20);
            this.lblSiteTenantName.TabIndex = 1;
            this.lblSiteTenantName.Text = "Your tenant name:";
            // 
            // txtSitesTenantName
            // 
            this.txtSitesTenantName.Location = new System.Drawing.Point(306, 22);
            this.txtSitesTenantName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSitesTenantName.Name = "txtSitesTenantName";
            this.txtSitesTenantName.Size = new System.Drawing.Size(240, 26);
            this.txtSitesTenantName.TabIndex = 0;
            this.ttAuthentication.SetToolTip(this.txtSitesTenantName, "Tenant name (e.g. contoso). Don\'t add the admin url (contoso-admin) here");
            // 
            // pnlSiteWildcard
            // 
            this.pnlSiteWildcard.Controls.Add(this.label10);
            this.pnlSiteWildcard.Controls.Add(this.label9);
            this.pnlSiteWildcard.Controls.Add(this.txtSitesAdminCenterUrl);
            this.pnlSiteWildcard.Controls.Add(this.btnSitesClearUrls);
            this.pnlSiteWildcard.Controls.Add(this.btnSitesRemoveUrl);
            this.pnlSiteWildcard.Controls.Add(this.btnSitesAddUrl);
            this.pnlSiteWildcard.Controls.Add(this.lstSitesUrlsToScan);
            this.pnlSiteWildcard.Controls.Add(this.txtSitesUrlToAdd);
            this.pnlSiteWildcard.Controls.Add(this.label4);
            this.pnlSiteWildcard.Controls.Add(this.label3);
            this.pnlSiteWildcard.Location = new System.Drawing.Point(10, 66);
            this.pnlSiteWildcard.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.pnlSiteWildcard.Name = "pnlSiteWildcard";
            this.pnlSiteWildcard.Size = new System.Drawing.Size(678, 315);
            this.pnlSiteWildcard.TabIndex = 3;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(9, 240);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(553, 20);
            this.label10.TabIndex = 9;
            this.label10.Text = "Only when using vanity url\'s specify the full SharePoint Online admin center url";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 271);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(111, 20);
            this.label9.TabIndex = 8;
            this.label9.Text = "Admin  center:";
            this.label9.Click += new System.EventHandler(this.label9_Click);
            // 
            // txtSitesAdminCenterUrl
            // 
            this.txtSitesAdminCenterUrl.Location = new System.Drawing.Point(129, 266);
            this.txtSitesAdminCenterUrl.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSitesAdminCenterUrl.Name = "txtSitesAdminCenterUrl";
            this.txtSitesAdminCenterUrl.Size = new System.Drawing.Size(529, 26);
            this.txtSitesAdminCenterUrl.TabIndex = 7;
            this.ttAuthentication.SetToolTip(this.txtSitesAdminCenterUrl, "SPO Admin center url (e.g. https://spoadmin.contoso.com). Only needed when using " +
        "SPO with vanity urls (= urls that not end on standard SharePoint online urls lik" +
        "e .sharepoint.com, .sharepoint.us,...)");
            this.txtSitesAdminCenterUrl.TextChanged += new System.EventHandler(this.txtSitesAdminCenterUrl_TextChanged);
            // 
            // btnSitesClearUrls
            // 
            this.btnSitesClearUrls.Location = new System.Drawing.Point(548, 154);
            this.btnSitesClearUrls.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSitesClearUrls.Name = "btnSitesClearUrls";
            this.btnSitesClearUrls.Size = new System.Drawing.Size(112, 35);
            this.btnSitesClearUrls.TabIndex = 6;
            this.btnSitesClearUrls.Text = "Clear";
            this.btnSitesClearUrls.UseVisualStyleBackColor = true;
            this.btnSitesClearUrls.Click += new System.EventHandler(this.btnSitesClearUrls_Click);
            // 
            // btnSitesRemoveUrl
            // 
            this.btnSitesRemoveUrl.Location = new System.Drawing.Point(548, 109);
            this.btnSitesRemoveUrl.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSitesRemoveUrl.Name = "btnSitesRemoveUrl";
            this.btnSitesRemoveUrl.Size = new System.Drawing.Size(112, 35);
            this.btnSitesRemoveUrl.TabIndex = 5;
            this.btnSitesRemoveUrl.Text = "Remove";
            this.btnSitesRemoveUrl.UseVisualStyleBackColor = true;
            this.btnSitesRemoveUrl.Click += new System.EventHandler(this.btnSitesRemoveUrl_Click);
            // 
            // btnSitesAddUrl
            // 
            this.btnSitesAddUrl.Location = new System.Drawing.Point(548, 45);
            this.btnSitesAddUrl.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSitesAddUrl.Name = "btnSitesAddUrl";
            this.btnSitesAddUrl.Size = new System.Drawing.Size(112, 35);
            this.btnSitesAddUrl.TabIndex = 4;
            this.btnSitesAddUrl.Text = "Add";
            this.btnSitesAddUrl.UseVisualStyleBackColor = true;
            this.btnSitesAddUrl.Click += new System.EventHandler(this.btnSitesAddUrl_Click);
            // 
            // lstSitesUrlsToScan
            // 
            this.lstSitesUrlsToScan.FormattingEnabled = true;
            this.lstSitesUrlsToScan.ItemHeight = 20;
            this.lstSitesUrlsToScan.Location = new System.Drawing.Point(129, 83);
            this.lstSitesUrlsToScan.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.lstSitesUrlsToScan.Name = "lstSitesUrlsToScan";
            this.lstSitesUrlsToScan.Size = new System.Drawing.Size(396, 144);
            this.lstSitesUrlsToScan.TabIndex = 3;
            this.lstSitesUrlsToScan.SelectedIndexChanged += new System.EventHandler(this.lstSitesUrlsToScan_SelectedIndexChanged);
            // 
            // txtSitesUrlToAdd
            // 
            this.txtSitesUrlToAdd.Location = new System.Drawing.Point(129, 45);
            this.txtSitesUrlToAdd.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtSitesUrlToAdd.Name = "txtSitesUrlToAdd";
            this.txtSitesUrlToAdd.Size = new System.Drawing.Size(396, 26);
            this.txtSitesUrlToAdd.TabIndex = 2;
            this.ttAuthentication.SetToolTip(this.txtSitesUrlToAdd, "Add one or more (wildcard) urls (e.g. https://contoso.sharepoint.com/sites/a*)");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 45);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 20);
            this.label4.TabIndex = 1;
            this.label4.Text = "Url to add:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 17);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(522, 20);
            this.label3.TabIndex = 0;
            this.label3.Text = "Add one or more (wildcard) url\'s. Url\'s ending on a * are valid wild card urls";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // modePage
            // 
            this.modePage.Controls.Add(this.tgModeHomePageOnly);
            this.modePage.Controls.Add(this.tgModeBlogUsage);
            this.modePage.Controls.Add(this.tgModeInfoPathUsage);
            this.modePage.Controls.Add(this.tgModeClassicWorkflowUsage);
            this.modePage.Controls.Add(this.tgModePublishingDetailed);
            this.modePage.Controls.Add(this.tgModePublishing);
            this.modePage.Controls.Add(this.tgModePages);
            this.modePage.Controls.Add(this.tgModeList);
            this.modePage.Controls.Add(this.label6);
            this.modePage.Controls.Add(this.tgModeGroupConnect);
            this.modePage.Controls.Add(this.cmbScanMode);
            this.modePage.Controls.Add(this.label5);
            this.modePage.Name = "modePage";
            this.modePage.NextPage = this.optionsPage;
            this.modePage.Size = new System.Drawing.Size(720, 406);
            this.modePage.TabIndex = 2;
            this.modePage.Tag = "";
            this.modePage.Text = "Select the scan mode";
            // 
            // tgModeHomePageOnly
            // 
            this.tgModeHomePageOnly.AutoCheck = false;
            this.tgModeHomePageOnly.AutoSize = true;
            this.tgModeHomePageOnly.Enabled = false;
            this.tgModeHomePageOnly.Location = new System.Drawing.Point(50, 168);
            this.tgModeHomePageOnly.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModeHomePageOnly.Name = "tgModeHomePageOnly";
            this.tgModeHomePageOnly.Size = new System.Drawing.Size(452, 24);
            this.tgModeHomePageOnly.TabIndex = 12;
            this.tgModeHomePageOnly.TabStop = false;
            this.tgModeHomePageOnly.Text = "Wiki/Webpart Page transformation readiness (home pages)";
            this.tgModeHomePageOnly.UseVisualStyleBackColor = true;
            // 
            // tgModeBlogUsage
            // 
            this.tgModeBlogUsage.AutoSize = true;
            this.tgModeBlogUsage.Enabled = false;
            this.tgModeBlogUsage.Location = new System.Drawing.Point(50, 375);
            this.tgModeBlogUsage.Name = "tgModeBlogUsage";
            this.tgModeBlogUsage.Size = new System.Drawing.Size(169, 24);
            this.tgModeBlogUsage.TabIndex = 11;
            this.tgModeBlogUsage.Text = "Classic Blog usage";
            this.tgModeBlogUsage.UseVisualStyleBackColor = true;
            // 
            // tgModeInfoPathUsage
            // 
            this.tgModeInfoPathUsage.AutoSize = true;
            this.tgModeInfoPathUsage.Enabled = false;
            this.tgModeInfoPathUsage.Location = new System.Drawing.Point(50, 345);
            this.tgModeInfoPathUsage.Name = "tgModeInfoPathUsage";
            this.tgModeInfoPathUsage.Size = new System.Drawing.Size(144, 24);
            this.tgModeInfoPathUsage.TabIndex = 10;
            this.tgModeInfoPathUsage.Text = "InfoPath usage";
            this.tgModeInfoPathUsage.UseVisualStyleBackColor = true;
            // 
            // tgModeClassicWorkflowUsage
            // 
            this.tgModeClassicWorkflowUsage.AutoSize = true;
            this.tgModeClassicWorkflowUsage.Enabled = false;
            this.tgModeClassicWorkflowUsage.Location = new System.Drawing.Point(50, 311);
            this.tgModeClassicWorkflowUsage.Name = "tgModeClassicWorkflowUsage";
            this.tgModeClassicWorkflowUsage.Size = new System.Drawing.Size(198, 24);
            this.tgModeClassicWorkflowUsage.TabIndex = 9;
            this.tgModeClassicWorkflowUsage.Text = "Classic workflow usage";
            this.tgModeClassicWorkflowUsage.UseVisualStyleBackColor = true;
            // 
            // tgModePublishingDetailed
            // 
            this.tgModePublishingDetailed.AutoCheck = false;
            this.tgModePublishingDetailed.AutoSize = true;
            this.tgModePublishingDetailed.Enabled = false;
            this.tgModePublishingDetailed.Location = new System.Drawing.Point(78, 274);
            this.tgModePublishingDetailed.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModePublishingDetailed.Name = "tgModePublishingDetailed";
            this.tgModePublishingDetailed.Size = new System.Drawing.Size(401, 24);
            this.tgModePublishingDetailed.TabIndex = 8;
            this.tgModePublishingDetailed.TabStop = false;
            this.tgModePublishingDetailed.Text = "Publishing portal transformation readiness (detailed)";
            this.tgModePublishingDetailed.UseVisualStyleBackColor = true;
            // 
            // tgModePublishing
            // 
            this.tgModePublishing.AutoCheck = false;
            this.tgModePublishing.AutoSize = true;
            this.tgModePublishing.Enabled = false;
            this.tgModePublishing.Location = new System.Drawing.Point(50, 237);
            this.tgModePublishing.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModePublishing.Name = "tgModePublishing";
            this.tgModePublishing.Size = new System.Drawing.Size(331, 24);
            this.tgModePublishing.TabIndex = 6;
            this.tgModePublishing.TabStop = false;
            this.tgModePublishing.Text = "Publishing portal transformation readiness";
            this.tgModePublishing.UseVisualStyleBackColor = true;
            // 
            // tgModePages
            // 
            this.tgModePages.AutoCheck = false;
            this.tgModePages.AutoSize = true;
            this.tgModePages.Enabled = false;
            this.tgModePages.Location = new System.Drawing.Point(78, 202);
            this.tgModePages.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModePages.Name = "tgModePages";
            this.tgModePages.Size = new System.Drawing.Size(430, 24);
            this.tgModePages.TabIndex = 5;
            this.tgModePages.TabStop = false;
            this.tgModePages.Text = "Wiki/Webpart Page transformation readiness (alll pages)";
            this.tgModePages.UseVisualStyleBackColor = true;
            // 
            // tgModeList
            // 
            this.tgModeList.AutoCheck = false;
            this.tgModeList.AutoSize = true;
            this.tgModeList.Enabled = false;
            this.tgModeList.Location = new System.Drawing.Point(50, 134);
            this.tgModeList.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModeList.Name = "tgModeList";
            this.tgModeList.Size = new System.Drawing.Size(266, 24);
            this.tgModeList.TabIndex = 4;
            this.tgModeList.TabStop = false;
            this.tgModeList.Text = "Modern list experience readiness";
            this.tgModeList.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(22, 65);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(356, 20);
            this.label6.TabIndex = 3;
            this.label6.Text = "Components included in the selected scan mode:";
            // 
            // tgModeGroupConnect
            // 
            this.tgModeGroupConnect.AutoCheck = false;
            this.tgModeGroupConnect.AutoSize = true;
            this.tgModeGroupConnect.Enabled = false;
            this.tgModeGroupConnect.Location = new System.Drawing.Point(50, 97);
            this.tgModeGroupConnect.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgModeGroupConnect.Name = "tgModeGroupConnect";
            this.tgModeGroupConnect.Size = new System.Drawing.Size(312, 24);
            this.tgModeGroupConnect.TabIndex = 2;
            this.tgModeGroupConnect.TabStop = false;
            this.tgModeGroupConnect.Text = "Office 365 Group connection readiness";
            this.tgModeGroupConnect.UseVisualStyleBackColor = true;
            // 
            // cmbScanMode
            // 
            this.cmbScanMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbScanMode.FormattingEnabled = true;
            this.cmbScanMode.Items.AddRange(new object[] {
            "Office 365 Group connection readiness",
            "Modern list experience readiness",
            "Wiki/Webpart Page transformation readiness (home pages)",
            "Wiki/Webpart Page transformation readiness (all pages)",
            "Publishing portal transformation readiness",
            "Publishing portal transformation readiness (detailed)",
            "Classic workflow usage",
            "InfoPath usage",
            "Classic Blog usage",
            "All of the above (full scan)"});
            this.cmbScanMode.Location = new System.Drawing.Point(256, 14);
            this.cmbScanMode.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbScanMode.Name = "cmbScanMode";
            this.cmbScanMode.Size = new System.Drawing.Size(446, 28);
            this.cmbScanMode.TabIndex = 1;
            this.cmbScanMode.SelectedIndexChanged += new System.EventHandler(this.cmbScanMode_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(22, 18);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(215, 20);
            this.label5.TabIndex = 0;
            this.label5.Text = "Select the scan mode to use:";
            // 
            // optionsPage
            // 
            this.optionsPage.Controls.Add(this.cmbDateFormat);
            this.optionsPage.Controls.Add(this.label13);
            this.optionsPage.Controls.Add(this.tgDisableTelemetry);
            this.optionsPage.Controls.Add(this.cmbSeparator);
            this.optionsPage.Controls.Add(this.label11);
            this.optionsPage.Controls.Add(this.tgSkipExcelReports);
            this.optionsPage.Controls.Add(this.tgExportDetailedWebPartData);
            this.optionsPage.Controls.Add(this.tgListBlockedDueToOOB);
            this.optionsPage.Controls.Add(this.tgSkipUser);
            this.optionsPage.Controls.Add(this.tgOptionSkipUsage);
            this.optionsPage.Controls.Add(this.nmThreads);
            this.optionsPage.Controls.Add(this.label7);
            this.optionsPage.IsFinishPage = true;
            this.optionsPage.Name = "optionsPage";
            this.optionsPage.Size = new System.Drawing.Size(720, 406);
            this.optionsPage.TabIndex = 3;
            this.optionsPage.Text = "Configure additional scan options";
            // 
            // cmbDateFormat
            // 
            this.cmbDateFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDateFormat.FormattingEnabled = true;
            this.cmbDateFormat.Items.AddRange(new object[] {
            "M/d/yyyy",
            "d/M/yyyy"});
            this.cmbDateFormat.Location = new System.Drawing.Point(314, 346);
            this.cmbDateFormat.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbDateFormat.Name = "cmbDateFormat";
            this.cmbDateFormat.Size = new System.Drawing.Size(140, 28);
            this.cmbDateFormat.TabIndex = 10;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(63, 351);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(250, 20);
            this.label13.TabIndex = 9;
            this.label13.Text = "Date format in the generated files:";
            // 
            // tgDisableTelemetry
            // 
            this.tgDisableTelemetry.AutoSize = true;
            this.tgDisableTelemetry.Location = new System.Drawing.Point(63, 268);
            this.tgDisableTelemetry.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgDisableTelemetry.Name = "tgDisableTelemetry";
            this.tgDisableTelemetry.Size = new System.Drawing.Size(398, 24);
            this.tgDisableTelemetry.TabIndex = 7;
            this.tgDisableTelemetry.Text = "Disable in product feedback to Microsoft (telemetry)";
            this.tgDisableTelemetry.UseVisualStyleBackColor = true;
            // 
            // cmbSeparator
            // 
            this.cmbSeparator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSeparator.FormattingEnabled = true;
            this.cmbSeparator.Items.AddRange(new object[] {
            ",",
            ";"});
            this.cmbSeparator.Location = new System.Drawing.Point(314, 305);
            this.cmbSeparator.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbSeparator.Name = "cmbSeparator";
            this.cmbSeparator.Size = new System.Drawing.Size(46, 28);
            this.cmbSeparator.TabIndex = 8;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(63, 309);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(243, 20);
            this.label11.TabIndex = 7;
            this.label11.Text = "Separator for the generated files:";
            // 
            // tgSkipExcelReports
            // 
            this.tgSkipExcelReports.AutoSize = true;
            this.tgSkipExcelReports.Location = new System.Drawing.Point(63, 232);
            this.tgSkipExcelReports.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgSkipExcelReports.Name = "tgSkipExcelReports";
            this.tgSkipExcelReports.Size = new System.Drawing.Size(382, 24);
            this.tgSkipExcelReports.TabIndex = 6;
            this.tgSkipExcelReports.Text = "Don\'t generate an Excel report for the found data";
            this.tgSkipExcelReports.UseVisualStyleBackColor = true;
            // 
            // tgExportDetailedWebPartData
            // 
            this.tgExportDetailedWebPartData.AutoSize = true;
            this.tgExportDetailedWebPartData.Location = new System.Drawing.Point(63, 195);
            this.tgExportDetailedWebPartData.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgExportDetailedWebPartData.Name = "tgExportDetailedWebPartData";
            this.tgExportDetailedWebPartData.Size = new System.Drawing.Size(331, 24);
            this.tgExportDetailedWebPartData.TabIndex = 5;
            this.tgExportDetailedWebPartData.Text = "Export the detailed web part property data";
            this.tgExportDetailedWebPartData.UseVisualStyleBackColor = true;
            // 
            // tgListBlockedDueToOOB
            // 
            this.tgListBlockedDueToOOB.AutoSize = true;
            this.tgListBlockedDueToOOB.Location = new System.Drawing.Point(63, 158);
            this.tgListBlockedDueToOOB.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgListBlockedDueToOOB.Name = "tgListBlockedDueToOOB";
            this.tgListBlockedDueToOOB.Size = new System.Drawing.Size(495, 24);
            this.tgListBlockedDueToOOB.TabIndex = 4;
            this.tgListBlockedDueToOOB.Text = "Exclude lists which are only blocked due to out of the box reasons";
            this.tgListBlockedDueToOOB.UseVisualStyleBackColor = true;
            // 
            // tgSkipUser
            // 
            this.tgSkipUser.AutoSize = true;
            this.tgSkipUser.Location = new System.Drawing.Point(63, 122);
            this.tgSkipUser.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgSkipUser.Name = "tgSkipUser";
            this.tgSkipUser.Size = new System.Drawing.Size(390, 24);
            this.tgSkipUser.TabIndex = 3;
            this.tgSkipUser.Text = "Don\'t include user information in the exported data";
            this.tgSkipUser.UseVisualStyleBackColor = true;
            // 
            // tgOptionSkipUsage
            // 
            this.tgOptionSkipUsage.AutoSize = true;
            this.tgOptionSkipUsage.Location = new System.Drawing.Point(63, 85);
            this.tgOptionSkipUsage.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tgOptionSkipUsage.Name = "tgOptionSkipUsage";
            this.tgOptionSkipUsage.Size = new System.Drawing.Size(613, 24);
            this.tgOptionSkipUsage.TabIndex = 2;
            this.tgOptionSkipUsage.Text = "Don\'t use search to get the site/page usage information and don\'t export that dat" +
    "a";
            this.tgOptionSkipUsage.UseVisualStyleBackColor = true;
            // 
            // nmThreads
            // 
            this.nmThreads.Location = new System.Drawing.Point(333, 22);
            this.nmThreads.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.nmThreads.Minimum = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.nmThreads.Name = "nmThreads";
            this.nmThreads.Size = new System.Drawing.Size(105, 26);
            this.nmThreads.TabIndex = 1;
            this.nmThreads.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(58, 25);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(247, 20);
            this.label7.TabIndex = 0;
            this.label7.Text = "Number of parallel threads to use:";
            // 
            // Wizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(720, 562);
            this.Controls.Add(this.wizardPageContainer1);
            this.Controls.Add(this.bottomDivider);
            this.Controls.Add(this.commandPanel);
            this.Controls.Add(this.topDivider);
            this.Controls.Add(this.headerPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Wizard";
            this.Text = "SharePoint Modernization Scanner configuration";
            this.headerPanel.ResumeLayout(false);
            this.headerPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.headerImage)).EndInit();
            this.commandPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).EndInit();
            this.wizardPageContainer1.ResumeLayout(false);
            this.authPage.ResumeLayout(false);
            this.authPage.PerformLayout();
            this.pnlAzureAD.ResumeLayout(false);
            this.pnlAzureAD.PerformLayout();
            this.pnlAzureACS.ResumeLayout(false);
            this.pnlAzureACS.PerformLayout();
            this.pnlCredentials.ResumeLayout(false);
            this.pnlCredentials.PerformLayout();
            this.pnl2FA.ResumeLayout(false);
            this.pnl2FA.PerformLayout();
            this.scopePage.ResumeLayout(false);
            this.scopePage.PerformLayout();
            this.pnlSiteFiles.ResumeLayout(false);
            this.pnlSiteFiles.PerformLayout();
            this.pnlSiteTenant.ResumeLayout(false);
            this.pnlSiteTenant.PerformLayout();
            this.pnlSiteWildcard.ResumeLayout(false);
            this.pnlSiteWildcard.PerformLayout();
            this.modePage.ResumeLayout(false);
            this.modePage.PerformLayout();
            this.optionsPage.ResumeLayout(false);
            this.optionsPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nmThreads)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AeroWizard.WizardPageContainer wizardPageContainer1;
        private System.Windows.Forms.Panel headerPanel;
        private System.Windows.Forms.Label topDivider;
        private System.Windows.Forms.Label bottomDivider;
        private System.Windows.Forms.Button backButton;
        private System.Windows.Forms.Button cancelButton;
        private AeroWizard.WizardPage modePage;
        private AeroWizard.WizardPage scopePage;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Panel commandPanel;
        private System.Windows.Forms.Label subHeaderLabel;
        private System.Windows.Forms.Label headerLabel;
        private System.Windows.Forms.PictureBox headerImage;
        private AeroWizard.WizardPage optionsPage;
        private AeroWizard.WizardPage authPage;
        private System.Windows.Forms.ComboBox cmbAuthOption;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCertificate;
        private System.Windows.Forms.TextBox txtAuthAzureADCert;
        private System.Windows.Forms.TextBox txtAuthAzureADDomainName;
        private System.Windows.Forms.TextBox txtAuthAzureADId;
        private System.Windows.Forms.Label lblAzureADCertificateFile;
        private System.Windows.Forms.Label lblAzureADDomainName;
        private System.Windows.Forms.Label lblAzureADApplicationId;
        private System.Windows.Forms.Panel pnlAzureAD;
        private System.Windows.Forms.Panel pnlAzureACS;
        private System.Windows.Forms.TextBox txtAzureADClientSecret;
        private System.Windows.Forms.TextBox txtAzureACSClientId;
        private System.Windows.Forms.Label lblAzureACSSecret;
        private System.Windows.Forms.Label lblAzureACSClientId;
        private System.Windows.Forms.Panel pnlCredentials;
        private System.Windows.Forms.TextBox txtCredentialsPassword;
        private System.Windows.Forms.TextBox txtCredentialsUser;
        private System.Windows.Forms.Label lblCredsPassword;
        private System.Windows.Forms.Label lblCredsUserName;
        private System.Windows.Forms.ComboBox cmbSiteSelectionOption;
        private System.Windows.Forms.Label lblSiteSelectionOption;
        private System.Windows.Forms.Panel pnlSiteWildcard;
        private System.Windows.Forms.Panel pnlSiteTenant;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label lblSiteTenantName;
        private System.Windows.Forms.TextBox txtSitesTenantName;
        private System.Windows.Forms.Panel pnlSiteFiles;
        private System.Windows.Forms.TextBox txtSitesCSVFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnSelectCSVFile;
        private System.Windows.Forms.Button btnSitesClearUrls;
        private System.Windows.Forms.Button btnSitesRemoveUrl;
        private System.Windows.Forms.Button btnSitesAddUrl;
        private System.Windows.Forms.ListBox lstSitesUrlsToScan;
        private System.Windows.Forms.TextBox txtSitesUrlToAdd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox tgModePublishing;
        private System.Windows.Forms.CheckBox tgModePages;
        private System.Windows.Forms.CheckBox tgModeList;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox tgModeGroupConnect;
        private System.Windows.Forms.ComboBox cmbScanMode;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox tgModePublishingDetailed;
        private System.Windows.Forms.NumericUpDown nmThreads;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox tgSkipUser;
        private System.Windows.Forms.CheckBox tgOptionSkipUsage;
        private System.Windows.Forms.CheckBox tgListBlockedDueToOOB;
        private System.Windows.Forms.CheckBox tgExportDetailedWebPartData;
        private System.Windows.Forms.CheckBox tgSkipExcelReports;
        private System.Windows.Forms.LinkLabel llblAzureADAuth;
        private System.Windows.Forms.LinkLabel llblAzureACSHelp;
        private System.Windows.Forms.LinkLabel llblScannerInfo;
        private System.Windows.Forms.LinkLabel llblModernizationGuidance;
        private System.Windows.Forms.LinkLabel llblCSV;
        private System.Windows.Forms.TextBox txtAuthAzureADCertPassword;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtSitesAdminCenterUrl;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cmbSeparator;
        private System.Windows.Forms.CheckBox tgDisableTelemetry;
        private System.Windows.Forms.Label lblAuthCreds2FAWarning;
        private System.Windows.Forms.ToolTip ttAuthentication;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.CheckBox tgModeInfoPathUsage;
        private System.Windows.Forms.CheckBox tgModeClassicWorkflowUsage;
        private System.Windows.Forms.CheckBox tgModeBlogUsage;
        private System.Windows.Forms.ComboBox cmbDateFormat;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox tgModeHomePageOnly;
        private System.Windows.Forms.Panel pnl2FA;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtSiteFor2FA;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.TextBox txtSitesAdminCenterUrl2;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.ComboBox cmbAuthenticationRegionCert;
        private System.Windows.Forms.Label label20;
    }
}