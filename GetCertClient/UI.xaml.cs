using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Input;
using tvToolbox;

namespace GetCert2
{
    public partial class UI : SavedWindow
    {
        private tvProfile   moProfile;
        private DoGetCert   moDoGetCert;

        private const string mcsSanNameKey = "-Domain";
        private const string mcsTitleText = "GoGetCert Get Certificate";

        private double  miOriginalScreenHeight;
        private double  miOriginalScreenWidth;
        private double  miAdjustedWindowHeight;
        private double  miAdjustedWindowWidth;
        private int     miPreviousConfigWizardSelectedIndex  = -1;
        private int     miSanListTabs = 10;
        private int     miSanListItemsPerTab = 10;
        private bool    mbGetDefaultsDone;
        private bool    mbOrigUseStandAloneMode;
        private string  msGetSetConfigurationDefaultsError = null;      // This allows configuration errors to be displayed asynchronously.

        private Button              moStartStopButtonState = new Button();
        private DateTime            mdtNextStart = DateTime.MinValue;
        private List<string>        moCertNameList;
        private tvMessageBox        moNotifyWaitMessage;
        private WindowState         mePreviousWindowState;


        private UI() { }


        /// <summary>
        /// This constructor expects an
        /// application object to be provided.
        /// </summary>
        /// <param name="aoDoGetCert">
        /// The given application object fetches digital
        /// certificates from a certificate provider network.
        /// </param>
        public UI(DoGetCert aoDoGetCert)
        {
            InitializeComponent();

            // This loads window UI defaults from the given profile.
            base.Init();

            moProfile = aoDoGetCert.oProfile;
            moDoGetCert = aoDoGetCert;
            Env.AppendOutputTextLine += this.AppendOutputTextLine;
        }


        // This lets us handle windows messages.
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            HwndSource  loSource = PresentationSource.FromVisual(this) as HwndSource;
                        loSource.AddHook(WndProc);
        }

        [DllImport("user32")]
        public static extern int    RegisterWindowMessage(string message);
        public static readonly int  WM_SHOWME = RegisterWindowMessage("WM_SHOWME_GetCert");
        [DllImport("user32")]
        public static extern bool   PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        public static readonly int  HWND_BROADCAST = 0xffff;

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // This handles the "WM_SHOWME" message so another
            // instance can display this one before exiting.
            if( WM_SHOWME == msg )
            {
                this.ShowMe();

                if ( null != moNotifyWaitMessage )
                {
                    // Wait for the UI to redisplay.
                    while ( Visibility.Visible != this.Visibility )
                    {
                        System.Windows.Forms.Application.DoEvents();
                        System.Threading.Thread.Sleep(100);
                    }

                    moNotifyWaitMessage.Close();
                    moNotifyWaitMessage = null;
                }
            }

            return IntPtr.Zero;
        }


        /// <summary>
        /// This is used with "HideMe() / ShowMe()"
        /// to track the visible state of this window.
        /// </summary>
        public bool bNoPrompts
        {
            get
            {
                return moProfile.bValue("-NoPrompts", false);
            }
            set
            {
                moProfile["-NoPrompts"] = value;
            }
        }

        /// <summary>
        /// This is used with "HideMe() / ShowMe()"
        /// to track the visible state of this window.
        /// </summary>
        public bool bVisible
        {
            get
            {
                return mbVisible;
            }
            set
            {
                mbVisible = value;
            }
        }
        private bool mbVisible = false;

        /// <summary>
        /// This indicates when the
        /// main process is running.
        /// </summary>
        public bool bMainProcessRunning
        {
            get
            {
                return mbMainProcessRunning;
            }
            set
            {
                mbMainProcessRunning = value;

                if ( !mbMainProcessRunning )
                {
                    this.EnableButtons();
                }
                else
                {
                    this.DisableButtons();
                    this.HideMiddlePanels();
                    this.MiddlePanelOutputText.Visibility = Visibility.Visible;
                }

                System.Windows.Forms.Application.DoEvents();
            }
        }
        private bool mbMainProcessRunning;

        /// <summary>
        /// This contains all child windows
        /// opened by the main parent window
        /// (ie. by this window, eg. help).
        /// </summary>
        public List<ScrollingText> oOtherWindows
        {
            get
            {
                return moOtherWindows;
            }
            set
            {
                moOtherWindows = value;
            }
        }
        private List<ScrollingText> moOtherWindows = new List<ScrollingText>();

        /// <summary>
        /// This is the startup wait message
        /// object instanced within the parent
        /// application object.
        /// </summary>
        public tvMessageBox oStartupWaitMsg
        {
            get
            {
                return moStartupWaitMsg;
            }
            set
            {
                moStartupWaitMsg = value;
            }
        }
        private tvMessageBox moStartupWaitMsg = null;

        /// <summary>
        /// This text is used for any notify icon
        /// tooltip as well as the main window title.
        /// </summary>
        public string sTitleText
        {
            get
            {
                string  lsCertName = null;
                        if ( null == CertificateDomainName.SelectedValue )
                            lsCertName = this.CertificateDomainName.Text;
                        else
                            lsCertName = CertificateDomainName.SelectedValue.ToString();
                string  lsTitleText = null;
                        if ( String.IsNullOrEmpty(lsCertName) )
                        {
                            lsTitleText = mcsTitleText;
                            CertNameTitle.Content = "";
                            CertNameTitle.Visibility = System.Windows.Visibility.Collapsed;
                        }
                        else
                        {
                            lsTitleText = String.Format("GetCert ({0})", lsCertName);
                            CertNameTitle.Content = String.Format("({0})", lsCertName);
                            CertNameTitle.Visibility = System.Windows.Visibility.Visible;
                        }

                return String.Format("{0} {1}.{2}", lsTitleText
                        , System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major
                        , System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor
                        );
            }
        }

        public void AppendOutputTextLine(object source, string asTextLine)
        {
            this.ProcessOutput.Inlines.Add(asTextLine + Environment.NewLine);
            if ( this.ProcessOutput.Inlines.Count > moProfile.iValue("-ConsoleTextLineLength", 200) )
                this.ProcessOutput.Inlines.Remove(this.ProcessOutput.Inlines.FirstInline);
            this.scrProcessOutput.ScrollToEnd();
        }

        public void HandleShutdown()
        {
            // Save any setup changes.
            this.GetSetConfigurationDefaults();

            // Update the output text cache.
            this.GetSetOutputTextPanelCache();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Title = this.sTitleText;
            this.AdjustWindowSize();
            miOriginalScreenHeight = SystemParameters.PrimaryScreenHeight;
            miOriginalScreenWidth = SystemParameters.PrimaryScreenWidth;

            // This window is hidden by default. Only make it initially
            // visible if all setup steps have not yet been completed.
            // Otherwise it is displayed via its system tray icon.
            if ( !moProfile.bValue("-AllConfigWizardStepsCompleted", false) )
                this.ShowMe();

            // Turns off the "loading" message.
            if ( null != this.oStartupWaitMsg )
                this.oStartupWaitMsg.Close();

            // Determine the "Setup Done" panel text.
            this.SetupDoneText();

            if ( moProfile.bValue("-AllConfigWizardStepsCompleted", false) )
            {
                if ( moDoGetCert.bDoInAndOut() )
                {
                    this.Close();
                }
                else
                {
                    // Display all of the main application elements needed
                    // after the configuration wizard has been completed.
                    this.HideMiddlePanels();
                    this.MainButtonPanel.IsEnabled = true;
                    this.GetSetOutputTextPanelCache();

                    bool lbPreviousBackupError = this.ShowPreviousProcessStatus();

                    this.ShowMe();
                    this.DisplayOutputText();
                }
            }
            else
            {
                if ( moProfile.bValue("-LicenseAccepted", false) )
                {
                    this.DisplayWizard();
                }
                else
                {
                    const string lsLicenseCaption = "MIT License";
                    const string lsLicensePathFile = lsLicenseCaption + ".txt";

                    // Fetch license.
                    tvFetchResource.ToDisk(Application.ResourceAssembly.GetName().Name, lsLicensePathFile, null);

                    ScrollingText loLicense = null;

                    if ( !this.bNoPrompts )
                    {
                        tvMessageBox.ShowBriefly(this, string.Format("The \"{0}\" will now be displayed."
                                        + "\r\n\r\nPlease accept it if you would like to use this software."
                                , lsLicenseCaption), lsLicenseCaption, tvMessageBoxIcons.Information, 3000);

                        loLicense = new ScrollingText(File.ReadAllText(
                                              moProfile.sRelativeToProfilePathFile(lsLicensePathFile))
                                            , lsLicenseCaption, true);
                        loLicense.TextBackground = Brushes.LightYellow;
                        loLicense.OkButtonText = "Accept";
                        loLicense.bDefaultButtonDisabled = true;
                        loLicense.ShowDialog();
                    }

                    if ( null != loLicense && loLicense.bOkButtonClicked )
                    {
                        moProfile["-LicenseAccepted"] = true;
                        moProfile.Save();

                        this.DisplayWizard();
                    }
                }
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Save any setup changes.
            this.GetSetConfigurationDefaults();

            // Save the SAN list.
            this.SaveDomainList();

            // Update the output text cache.
            this.GetSetOutputTextPanelCache();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if ( 0 != moOtherWindows.Count )
                foreach (ScrollingText loWindow in moOtherWindows)
                    loWindow.Close();
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            // This kludge is needed since attempting to restore a window that is busy
            // processing (after a system minimize) will typically not be responsive.
            if ( this.bMainProcessRunning && WindowState.Minimized == this.WindowState )
                this.HideMe();
        }

        // Buttons that don't launch external processes are toggles.

        private void btnDoMainProcessNow_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            if ( this.bMainProcessRunning )
            {
                moDoGetCert.bMainLoopStopped = true;
                this.bMainProcessRunning = false;
            }
            else
            {
                this.HideMiddlePanels();
                this.GetSetConfigurationDefaults();

                if ( this.bValidateConfiguration() )
                    this.DoMainProcess();
                else
                    this.btnSetup_Click(null, null);
            }
        }

        private void btnSetup_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            if ( Visibility.Visible == this.MiddlePanelConfigWizard.Visibility )
            {
                this.GetSetConfigurationDefaults();
                this.HideMiddlePanels();
                this.DisplayOutputText();
            }
            else
            {
                this.HideMiddlePanels();
                this.DisplayWizard();

                this.ConfigWizardTabs.SelectedIndex = 0;
            }
        }

        private void btnSetupDone_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            if ( this.bValidateConfiguration() )
            {
                if ( moProfile.bValue("-AllConfigWizardStepsCompleted", false) )
                {
                    this.SetupDoneApplyProfileUpdates();

                    this.HideMiddlePanels();
                    this.MainButtonPanel.IsEnabled = true;

                    this.DoMainProcess();
                }
                else
                if ( !moProfile.bValue("-UseStandAloneMode", true) && mbOrigUseStandAloneMode )
                {
                    if ( tvMessageBoxResults.No == this.Show(@"
Changing from ""Use Stand Alone Mode"" to not ""Use Stand Alone Mode"" requires restarting this software.

Are you sure you want to restart it now?

Notes:

The host software must also be running when ""Use Stand Alone Mode"" is disabled.

Check the application log if this software fails to restart.

If you change your mind (ie. click ""No""), the ""Use Stand Alone Mode"" switch will again be checked.
"
                                , "Get Certificate"
                                , tvMessageBoxButtons.YesNo
                                , tvMessageBoxIcons.Question
                                )
                            )
                    {
                        this.UseStandAloneMode.IsChecked = mbOrigUseStandAloneMode;
                        this.GetSetConfigurationDefaults();
                    }
                    else
                    {
                        this.SetupDoneApplyProfileUpdates();
                        this.SetupDoneText();

                        this.Close();

                        Process.Start(moProfile.sExePathFile);
                    }
                }
                else
                if ( tvMessageBoxResults.Yes == this.Show(string.Format(@"
The get certificate process could take a minute or more to complete. Are you sure you want to run this now?

You can continue this later wherever you left off. "
+ @" You can also edit the profile file directly (""{0}"") for"
+ @" much more detailed configuration (see ""Help"").
"
                                , Path.GetFileName(moProfile.sLoadedPathFile)
                                )
                            , "Get Certificate"
                            , tvMessageBoxButtons.YesNo
                            , tvMessageBoxIcons.Question
                            )
                        )
                {
                    this.SetupDoneApplyProfileUpdates();
                    this.SetupDoneText();

                    this.HideMiddlePanels();
                    this.MainButtonPanel.IsEnabled = true;

                    this.DoMainProcess();
                }
            }
        }

        private void btnNextSetupStep_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.ConfigWizardTabs.SelectedIndex++;
        }

        private void btnEditSanList_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            moProfile.Reload();

            // Rebuild list (if we just did a review).
            if ( 0 != DomainListTabs.Items.Count && typeof(ComboBox) != ((Grid)((TabItem)DomainListTabs.Items[0]).Content).Children[0].GetType() )
                DomainListTabs.Items.Clear();

            if ( 0 == DomainListTabs.Items.Count )
            {
                this.LoadDomainList();
            }
            else
            {
                // The first domain name on the list must match the one on the main form.
                ComboBox    loComboBox = (ComboBox)((Grid)((TabItem)DomainListTabs.Items[0]).Content).Children[0];
                            loComboBox.Text = this.CertificateDomainName.Text.Trim();
            }

            btnDomainListSave.IsEnabled = true;
            MiddlePanelDomainList.IsOpen = true;
        }

        private void btnDomainListSave_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.SaveDomainList();

            string  lsMessage = this.bValidateConfigWizardSanList();
                    if ( !String.IsNullOrEmpty(lsMessage) )
                        this.ShowWarning(lsMessage, "Fix SAN List");
        }

        private void btnDomainListCancel_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            DomainListTabs.Items.Clear();
        }

        private void LoadDomainList()
        {
            if ( 0 == DomainListTabs.Items.Count )
            {
                moProfile.Reload();

                if ( null == this.CertificateDomainName.Text )
                    this.CertificateDomainName.Text = "";

                tvProfile loSanList = moProfile.oProfile("-SanList", String.Format("\r\n{0}=\"{1}\"\r\n", mcsSanNameKey, this.CertificateDomainName.Text.Trim()));

                for (int i=0; i < miSanListTabs; i++)
                {
                    TabItem loTabItem = new TabItem();
                            loTabItem.Header = (i + 1).ToString();
                            if ( 0 == i )
                                loTabItem.Header = "SAN Page " + loTabItem.Header;
                    Grid    loTabGrid = new Grid();
                            loTabGrid.ColumnDefinitions.Add(new ColumnDefinition());

                    for (int j=0; j < miSanListItemsPerTab; j++)
                    {
                        ComboBox    loComboBox = new ComboBox();
                                    loComboBox.Style = (Style)Resources["ConfigWizardComboBoxEditable"];
                                    loComboBox.LostKeyboardFocus += CertificateDomainName_LostKeyboardFocus;
                                    loComboBox.SelectionChanged += CertificateDomainName_SelectionChanged;
                                    loComboBox.ItemsSource = moCertNameList;

                        if ( 0 == j && 0 == i )
                        {
                            // The first domain name on the list must match the one on the main form.
                            loComboBox.Text = this.CertificateDomainName.Text.Trim();
                        }
                        else
                        {
                            int liSanListIndex = miSanListItemsPerTab * i + j;

                            if ( liSanListIndex < loSanList.Count )
                                loComboBox.Text = (string)loSanList[liSanListIndex];
                            else
                                loComboBox.Text = "";
                        }

                        loTabGrid.RowDefinitions.Add(new RowDefinition());
                        loTabGrid.Children.Add(loComboBox);
                        Grid.SetRow(loComboBox, j);
                    }

                    loTabItem.Content = loTabGrid;

                    DomainListTabs.Items.Add(loTabItem);
                }

                this.SaveDomainList();
            }
        }

        private void SaveDomainList()
        {
            // Only save the edit list (not the review).
            if ( 0 != DomainListTabs.Items.Count && typeof(ComboBox) == ((Grid)((TabItem)DomainListTabs.Items[0]).Content).Children[0].GetType() )
            {
                try
                {
                    tvProfile   loSanList = new tvProfile();
                                for (int i=0; i < DomainListTabs.Items.Count; i++)
                                {
                                    Grid loTabGrid = (Grid)((TabItem)DomainListTabs.Items[i]).Content;

                                    for (int j=0; j < loTabGrid.Children.Count; j++)
                                    {
                                        ComboBox    loComboBox = (ComboBox)loTabGrid.Children[j];
                                                    loComboBox.Text = loComboBox.Text.Trim();

                                        // The first domain name on the list must match the one on the main form.
                                        if ( 0 == j && 0 == i )
                                        {
                                            if ( "" != loComboBox.Text )
                                                this.CertificateDomainName.Text = loComboBox.Text;
                                            else
                                                loComboBox.Text = this.CertificateDomainName.Text.Trim();
                                        }

                                        if ( "" != loComboBox.Text )
                                            loSanList.Add(mcsSanNameKey, loComboBox.Text);
                                    }
                                }

                    moProfile["-SanList"] = loSanList.sCommandBlock();
                    moProfile.Save();
                }
                catch (InvalidCastException) {}
            }
        }

        private void btnReviewSanList_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            tvProfile loSanList = moProfile.oProfile("-SanList");

            DomainListTabs.Items.Clear();

            // Set to readonly labels.

            for (int i=0; i < miSanListTabs; i++)
            {
                TabItem loTabItem = new TabItem();
                        loTabItem.Header = (i + 1).ToString();
                        if ( 0 == i )
                            loTabItem.Header = "SAN Page " + loTabItem.Header;
                Grid    loTabGrid = new Grid();
                        loTabGrid.ColumnDefinitions.Add(new ColumnDefinition());

                for (int j=0; j < miSanListItemsPerTab; j++)
                {
                    Label   loLabel = new Label();
                            loLabel.Style = (Style)Resources["ConfigWizardComboBoxReadOnlyLabel"];

                    if ( 0 == j && 0 == i )
                    {
                        // The first domain name on the list must match the one on the main form.
                        loLabel.Content = this.CertificateDomainName.Text.Trim();
                    }
                    else
                    {
                        int liSanListIndex = miSanListItemsPerTab * i + j;

                        if ( liSanListIndex < loSanList.Count )
                            loLabel.Content = (string)loSanList[liSanListIndex];
                    }

                    loTabGrid.RowDefinitions.Add(new RowDefinition());
                    loTabGrid.Children.Add(loLabel);
                    Grid.SetRow(loLabel, j);
                }

                loTabItem.Content = loTabGrid;

                DomainListTabs.Items.Add(loTabItem);
            }


            btnDomainListSave.IsEnabled = false;
            MiddlePanelDomainList.IsOpen = true;
        }

        private void btnShowHelp_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.ShowHelp();
        }

        private void btnShowSite_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            Process.Start("https://GoGetCert.com");
        }

        private void btnShowLogs_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            string  lsPath = Path.GetDirectoryName(moProfile.sLoadedPathFile);
                    if ( String.IsNullOrEmpty(lsPath) )
                        lsPath = Directory.GetCurrentDirectory();

            Process.Start(moProfile.sValue("-WindowsExplorer", "explorer.exe"), lsPath);
        }

        private void btnClearDisplay_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.ProcessOutput.Inlines.Clear();
        }

        private void mnuExit_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            if ( MouseButton.Left == e.ChangedButton )
                this.DragMove();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            switch ( e.Key )
            {
                case Key.F1:
                    this.ShowHelp();
                    break;
            }
        }

        private void mnuMaximize_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.WindowState = WindowState.Maximized;
            this.Height = SystemParameters.MaximizedPrimaryScreenHeight;
            this.Width = SystemParameters.MaximizedPrimaryScreenWidth;

            mePreviousWindowState = WindowState.Maximized;
        }

        private void mnuRestore_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.WindowState = WindowState.Normal;
            this.Height = miAdjustedWindowHeight;
            this.Width = miAdjustedWindowWidth;

            mePreviousWindowState = WindowState.Normal;
        }

        private void mnuMinimize_Click(object sender, RoutedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.WindowState = WindowState.Minimized;
        }

        private void Window_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
        }


        // This collection of "Show" methods allows for pop-up messages
        // to be suppressed and written to the log file instead.


        private tvMessageBoxResults Show(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                )
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( !this.bNoPrompts )
                ltvMessageBoxResults = tvMessageBox.Show(
                          this
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        );

            return ltvMessageBoxResults;
        }

        private tvMessageBoxResults Show(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                , string asProfilePromptKey
                )
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( !this.bNoPrompts )
                ltvMessageBoxResults = tvMessageBox.Show(
                          this
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        , tvMessageBoxCheckBoxTypes.SkipThis
                        , moProfile
                        , asProfilePromptKey
                        );

            return ltvMessageBoxResults;
        }

        private tvMessageBoxResults Show(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                , tvMessageBoxCheckBoxTypes aetvMessageBoxCheckBoxType
                , string asProfilePromptKey
                )
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( !this.bNoPrompts )
                ltvMessageBoxResults = tvMessageBox.Show(
                          this
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        , tvMessageBoxCheckBoxTypes.SkipThis
                        , moProfile
                        , asProfilePromptKey
                        );

            return ltvMessageBoxResults;
        }

        private tvMessageBoxResults Show(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                , tvMessageBoxCheckBoxTypes aetvMessageBoxCheckBoxType
                , string asProfilePromptKey
                , tvMessageBoxResults aetvMessageBoxResult
                )
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( !this.bNoPrompts )
                ltvMessageBoxResults = tvMessageBox.Show(
                          this
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        , tvMessageBoxCheckBoxTypes.SkipThis
                        , moProfile
                        , asProfilePromptKey
                        , aetvMessageBoxResult
                        );

            return ltvMessageBoxResults;
        }

        private void ShowError(Exception aoException)
        {
            if ( this.bNoPrompts )
                Env.LogIt(aoException.Message);

            if ( !this.bNoPrompts )
                tvMessageBox.ShowError(this, aoException);
        }

        private void ShowError(string asMessageText)
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            if ( !this.bNoPrompts )
                tvMessageBox.ShowError(this, asMessageText);
        }

        public void ShowError(string asMessageText, string asMessageCaption)
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            if ( !this.bNoPrompts )
                tvMessageBox.ShowError(this, asMessageText, asMessageCaption);
        }

        public void ShowWarning(string asMessageText, string asMessageCaption)
        {
            if ( this.bNoPrompts )
                Env.LogIt(asMessageText);

            if ( !this.bNoPrompts )
                tvMessageBox.ShowWarning(this, asMessageText, asMessageCaption);
        }

        // This "HideMe() / ShowMe()" kludge is necessary
        // to avoid annoying flicker on some platforms.
        private void HideMe()
        {
            this.MainCanvas.Visibility = Visibility.Hidden;
            this.WindowState = mePreviousWindowState;
            this.Hide();
            this.bVisible = false;
        }

        private void ShowMe()
        {
            bool lTopmost = this.Topmost;

            this.AdjustWindowSize();
            this.MainCanvas.Visibility = Visibility.Visible;
            this.WindowState = mePreviousWindowState;
            this.Topmost = true;
            System.Windows.Forms.Application.DoEvents();
            this.Show();
            this.bVisible = true;
            System.Windows.Forms.Application.DoEvents();
            this.Topmost = lTopmost;
            System.Windows.Forms.Application.DoEvents();
        }

        private void AdjustWindowSize()
        {
            if ( WindowState.Maximized != this.WindowState
                    && (SystemParameters.PrimaryScreenHeight != miOriginalScreenHeight
                    || SystemParameters.MaximizedPrimaryScreenWidth != miOriginalScreenWidth)
                    )
            {
                // Adjust window size to optimize the display depending on screen size.
                // This is done here rather in "Window_Loaded()" in case the screen size
                // changes post startup (eg. via RDP).
                if ( SystemParameters.PrimaryScreenHeight <= 768 )
                {
                    this.mnuMaximize_Click(null, null);
                }
                else if ( SystemParameters.PrimaryScreenHeight <= 864 )
                {
                    this.Height = .90 * SystemParameters.MaximizedPrimaryScreenHeight;
                    this.Width = .90 * SystemParameters.MaximizedPrimaryScreenWidth;
                }
                else if ( SystemParameters.PrimaryScreenHeight <= 885 )
                {
                    this.Height = .80 * SystemParameters.MaximizedPrimaryScreenHeight;
                    this.Width = .80 * SystemParameters.MaximizedPrimaryScreenWidth;
                }
                else
                {
                    this.Height = 675;
                    this.Width = 900;
                }

                miAdjustedWindowHeight = this.Height;
                miAdjustedWindowWidth = this.Width;

                // "this.WindowStartupLocation = WindowStartupLocation.CenterScreen" fails. Who knew?
                this.Top = (SystemParameters.MaximizedPrimaryScreenHeight - this.Height) / 2;
                this.Left = (SystemParameters.MaximizedPrimaryScreenWidth - this.Width) / 2;
            }
        }

        private void GetSetConfigurationDefaults()
        {
            this.GetSetConfigurationDefaults(false);
        }
        private void GetSetConfigurationDefaults(bool abResetGetDefaultsDone)
        {
            moProfile.Reload();

            if ( abResetGetDefaultsDone )
                mbGetDefaultsDone = false;

            if ( !mbGetDefaultsDone )
            {
                try
                {
                    // Step 1
                    this.ContactEmailAddress.Text = moProfile.sValue("-ContactEmailAddress", "");

                    moCertNameList = new List<string>();
                    foreach(X509Certificate2 loCertificate in Env.oCurrentCertificateCollection)
                    {
                        string lsCertName = Env.sCertName(loCertificate);

                        if ( !moCertNameList.Contains(lsCertName) )
                            moCertNameList.Add(lsCertName);
                    }
                    moCertNameList.Sort();

                    this.CertificateDomainName.ItemsSource = moCertNameList;

                    if ( "" == moProfile.sValue("-CertificateDomainName", "") )
                        this.CertificateDomainName.Text = Env.sCurrentCertificateName;
                    else
                        this.CertificateDomainName.Text = moProfile.sValue("-CertificateDomainName", "");

                    this.UseStandAloneMode.IsChecked = moProfile.bValue("-UseStandAloneMode", true);
                    mbOrigUseStandAloneMode = moProfile.bValue("-UseStandAloneMode", true);
                    this.DoStagingTests.IsChecked = moProfile.bValue("-DoStagingTests", true);
                    this.RemoveReplacedCert.IsChecked = moProfile.bValue("-RemoveReplacedCert", false);
                }
                catch (Exception ex)
                {
                    msGetSetConfigurationDefaultsError = ex.Message;
                }

                mbGetDefaultsDone = true;
            }

            // Finish
            this.ReviewContactEmailAddress.Text = this.ContactEmailAddress.Text;
            this.ReviewCertificateDomainName.Text = this.CertificateDomainName.Text;
            this.ReviewUseStandAloneMode.IsChecked = this.UseStandAloneMode.IsChecked;
            this.ReviewDoStagingTests.IsChecked = this.DoStagingTests.IsChecked;
            this.ReviewRemoveReplacedCert.IsChecked = this.RemoveReplacedCert.IsChecked;

            if ( null == msGetSetConfigurationDefaultsError )
            {
                moProfile["-ContactEmailAddress"] = this.ReviewContactEmailAddress.Text.Trim();
                moProfile["-CertificateDomainName"] = this.ReviewCertificateDomainName.Text.Trim();
                moProfile["-UseStandAloneMode"] = this.ReviewUseStandAloneMode.IsChecked;
                moProfile["-DoStagingTests"] = this.ReviewDoStagingTests.IsChecked;
                moProfile["-RemoveReplacedCert"] = this.ReviewRemoveReplacedCert.IsChecked;
                moProfile.Save();
            }
        }

        private bool bValidateConfiguration()
        {
            return ( this.bValidateConfigWizardValues(true) );
        }

        private bool bValidateConfigWizardValues(bool abVerifyAllTabs)
        {
            string  lsCaption = "Please Fix Before You Finish";
            string  lsCommonMsg = " (Note: be sure to remove any leading or trailing spaces)";
            string  lsMessage = null;
            bool    lbHaveMovedForward = this.ConfigWizardTabs.SelectedIndex >= miPreviousConfigWizardSelectedIndex;

            Regex loEmailRegex = new Regex(moProfile.sValue("-RegexEmailAddress", @"^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)[a-zA-Z]{2,18}$"));
            Regex loDnsNameRegex = new Regex(moProfile.sValue("-RegexDnsNamePrimary", @"^([a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,18}$"));

            // Can't pass nulls to Regex.
            if ( null == this.ContactEmailAddress.Text )
                this.ContactEmailAddress.Text = "";
            if ( null == this.CertificateDomainName.Text )
                this.CertificateDomainName.Text = "";

            // Step 1
            if ( abVerifyAllTabs || lbHaveMovedForward && miPreviousConfigWizardSelectedIndex
                    == ItemsControl.ItemsControlFromItemContainer(
                    this.tabStep1).ItemContainerGenerator.IndexFromContainer(this.tabStep1)
                    )
            {
                if ( !loEmailRegex.IsMatch(this.ContactEmailAddress.Text) )
                    lsMessage += (String.IsNullOrEmpty(lsMessage) ? "" : Environment.NewLine + Environment.NewLine)
                            + string.Format("Enter a valid email address{0}. ", !this.ContactEmailAddress.Text.Contains(" ") ? "" :  lsCommonMsg)
                            ;
                if ( !loDnsNameRegex.IsMatch(this.CertificateDomainName.Text) )
                    lsMessage += (String.IsNullOrEmpty(lsMessage) ? "" : Environment.NewLine + Environment.NewLine)
                            + string.Format("Enter a valid domain name{0}. ", !this.CertificateDomainName.Text.Contains(" ") ? "" :  lsCommonMsg)
                            ;

                string  lsSanListMsgs = this.bValidateConfigWizardSanList();
                        if ( !String.IsNullOrEmpty(lsSanListMsgs) )
                            lsMessage += lsSanListMsgs;
            }

            if ( !String.IsNullOrEmpty(lsMessage) )
                this.ShowWarning(lsMessage, lsCaption);

            return String.IsNullOrEmpty(lsMessage);
        }

        private string bValidateConfigWizardSanList()
        {
            string  lsMessage = null;

            Regex loDnsNameRegex = new Regex(moProfile.sValue("-RegexDnsNamePrimary", @"^([a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,18}$"));
            Regex loDnsNameRegexSanList = new Regex(moProfile.sValue("-RegexDnsNameSanList", @"^([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)([a-zA-Z]{2,18}|)$"));

            // Rebuild (in case of review or SAN list was edited in the profile).
            DomainListTabs.Items.Clear();

            this.LoadDomainList();

            for (int i=0; i < DomainListTabs.Items.Count; i++)
            {
                Grid loTabGrid = (Grid)((TabItem)DomainListTabs.Items[i]).Content;

                for (int j=0; j < loTabGrid.Children.Count; j++)
                {
                    ComboBox loComboBox = (ComboBox)loTabGrid.Children[j];

                    if ( "" == loComboBox.Text || loDnsNameRegex.IsMatch(loComboBox.Text) || loDnsNameRegexSanList.IsMatch(loComboBox.Text) )
                    {
                        loComboBox.Foreground = Brushes.Black;
                    }
                    else
                    {
                        loComboBox.Foreground = Brushes.Red;
                        lsMessage += (String.IsNullOrEmpty(lsMessage) ? "" : Environment.NewLine + Environment.NewLine)
                                + string.Format("Domain name (\"{0}\") is invalid.", loComboBox.Text)
                                ;
                    }

                    string  lsDupMsg = this.bValidateConfigWizardSanList2(ref loComboBox, i, j);
                            if ( !String.IsNullOrEmpty(lsDupMsg) )
                            {
                                loComboBox.Foreground = Brushes.Red;
                                lsMessage += (String.IsNullOrEmpty(lsMessage) ? "" : Environment.NewLine + Environment.NewLine) + lsDupMsg;
                            }
                }
            }

            return lsMessage;
        }

        private string bValidateConfigWizardSanList2(ref ComboBox aoComboBox, int ai, int aj)
        {
            if ( "" == aoComboBox.Text)
                return null;

            string lsMessage = null;

            for (int i=0; i < DomainListTabs.Items.Count; i++)
            {
                Grid loTabGrid = (Grid)((TabItem)DomainListTabs.Items[i]).Content;

                for (int j=0; j < loTabGrid.Children.Count; j++)
                {
                    ComboBox loComboBox = (ComboBox)loTabGrid.Children[j];

                    if (       "" != loComboBox.Text
                            && miSanListTabs * ai + aj > miSanListTabs * i + j
                            && aoComboBox.Text.ToLower() == loComboBox.Text.ToLower()
                            )
                        lsMessage = string.Format("Domain name (\"{0}\") is a duplicate.", aoComboBox.Text);
                }
            }

            return lsMessage;
        }

        private void ConfigWizardTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MiddlePanelDomainList.IsOpen = false;

            this.GetSetConfigurationDefaults();
            this.bValidateConfigWizardValues(false);

            miPreviousConfigWizardSelectedIndex = this.ConfigWizardTabs.SelectedIndex;
        }

        private void DomainListTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void CertificateDomainName_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.Title = this.sTitleText;
            e.Handled = true;
        }

        private void CertificateDomainName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.Title = this.sTitleText;
            e.Handled = true;
        }

        private void UseStandAloneMode_CheckedChanged(object sender, EventArgs e)
        {
            // -RemoveReplacedCert only applies when -UseStandAloneMode is true.
            // Why? Because we can't risk a potentially forgotten and unchecked
            // buildup of old certificates on remote servers.
            if ( (bool)this.UseStandAloneMode.IsChecked )
            {
                this.RemoveReplacedCertRow.Height = GridLength.Auto;
                this.ReviewRemoveReplacedCertRow.Height = GridLength.Auto;
            }
            else
            {
                this.RemoveReplacedCertRow.Height = new GridLength(0, GridUnitType.Pixel); 
                this.ReviewRemoveReplacedCertRow.Height = new GridLength(0, GridUnitType.Pixel); 
            }
        }

        private void DisplayWizard()
        {
            this.GetSetConfigurationDefaults(true);
            this.MiddlePanelConfigWizard.Visibility = Visibility.Visible;

            if ( null != msGetSetConfigurationDefaultsError )
            {
                // This kludge (the "Replace()") is needed to correct for bad MS grammar.
                this.ShowError(msGetSetConfigurationDefaultsError.Replace("a unknown", "an unknown"), "Error Loading Configuration");
                msGetSetConfigurationDefaultsError = null;
            }
        }

        private void DisplayOutputText()
        {
            if ( "" != this.ProcessOutput.Text.Trim() )
                this.MiddlePanelOutputText.Visibility = Visibility.Visible;
        }

        private void HideMiddlePanels()
        {
            this.MiddlePanelOutputText.Visibility = Visibility.Collapsed;
            this.MiddlePanelConfigWizard.Visibility = Visibility.Collapsed;
        }

        private void DisableButtons()
        {
            moStartStopButtonState.Content = this.btnDoMainProcessNow.Content;
            moStartStopButtonState.ToolTip = this.btnDoMainProcessNow.ToolTip;
            moStartStopButtonState.Background = this.btnDoMainProcessNow.Background;
            moStartStopButtonState.FontSize = this.btnDoMainProcessNow.FontSize;

            this.btnDoMainProcessNow.Content = "STOP";
            this.btnDoMainProcessNow.ToolTip = "Stop the Process Now";
            this.btnDoMainProcessNow.Background = Brushes.Red;
            this.btnDoMainProcessNow.FontSize = 22;

            this.btnSetup.IsEnabled = false;
        }

        private void EnableButtons()
        {
            this.btnDoMainProcessNow.Content = moStartStopButtonState.Content;
            this.btnDoMainProcessNow.ToolTip = moStartStopButtonState.ToolTip;
            this.btnDoMainProcessNow.Background = moStartStopButtonState.Background;
            this.btnDoMainProcessNow.FontSize = moStartStopButtonState.FontSize;

            this.MainButtonPanel.IsEnabled = true;
            this.btnDoMainProcessNow.IsEnabled = true;
            this.btnSetup.IsEnabled = true;
        }

        public void GetSetOutputTextPanelCache()
        {
            if ( "" == this.ProcessOutput.Text )
            {
                this.ProcessOutput.Text = moProfile.sValue("-PreviousProcessOutputText", "") + Environment.NewLine;
                this.scrProcessOutput.ScrollToEnd();

                if ( "" == this.ProcessOutput.Text.Trim() )
                    this.ProcessOutput.Text = @"
Be sure to run the staging tests at least twice to verify old
certs and old bindings are being properly removed and replaced.

Click ""GO"" to get started.
                            ";
            }
            else
            {
                // This makes sure the last line of text appears before we save.
                System.Windows.Forms.Application.DoEvents();

                moProfile["-PreviousProcessOutputText"] = this.ProcessOutput.Text;
                moProfile.Save();
            }
        }

        private void SetupDoneText()
        {
            btnSetupDone.Content = "Setup Done - Run Process";
            txtSetupDone.Text = "Please review your choices from the previous setup steps, then click \"Setup Done - Run Process\".";
            txtSetupDoneDesc.Text = @"The get certificate process will run in ""staging mode"" until you change it. This will give you a chance to verify the process works before it runs automatically later.

If you would prefer to finish this setup at another time, you can exit now and continue later wherever you left off.";
        }

        private void SetupDoneApplyProfileUpdates()
        {
            moProfile["-AllConfigWizardStepsCompleted"] = true;
            moProfile.Save();                
        }

        private void ShowHelp()
        {
            // If a help window is already open, close it.
            foreach (ScrollingText loWindow in moOtherWindows)
            {
                if ( loWindow.Title.Contains("Help") )
                {
                    loWindow.Close();
                    moOtherWindows.Remove(loWindow);
                    break;
                }
            }

            ScrollingText   loHelp = new ScrollingText(moProfile["-Help"].ToString(), "Get Certificate Help", true);
                            loHelp.TextBackground = Brushes.Khaki;
                            loHelp.Show();

                            moOtherWindows.Add(loHelp);
        }

        private bool ShowPreviousProcessStatus()
        {
            bool lbPreviousProcessError = false;

            if ( moProfile.ContainsKey("-PreviousProcessOk") )
            {
                if ( !moProfile.bValue("-PreviousProcessOk", false) )
                {
                    lbPreviousProcessError = true;

                    this.Show("The previous get certificate process failed. Check the log for errors."
                            , "Get Certificate Failed"
                            , tvMessageBoxButtons.OK, tvMessageBoxIcons.Error
                            , tvMessageBoxCheckBoxTypes.SkipThis
                            , "-PreviousProcessFailed"
                            );
                }
                else
                {
                    int liPreviousProcessDays = (DateTime.Now - moProfile.dtValue("-PreviousProcessTime"
                                                    , DateTime.MinValue)).Days;

                    this.Show(string.Format(
                            "The previous get certificate process finished successfully ({0} {1} day{2} ago)."
                                    , liPreviousProcessDays < 1 ? "less than" : "about"
                                    , liPreviousProcessDays < 1 ? 1 : liPreviousProcessDays
                                    , liPreviousProcessDays <= 1 ? "" : "s"
                                    )
                            , "Get Certificate Finished"
                            , tvMessageBoxButtons.OK, tvMessageBoxIcons.Done
                            , tvMessageBoxCheckBoxTypes.SkipThis
                            , "-PreviousProcessFinished"
                            );
                }
            }

            return lbPreviousProcessError;
        }

        private void DoMainProcess()
        {
            this.bMainProcessRunning = true;

            tvProfile   loMinProfile = Env.oMinProfile(moProfile);
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);

            DoGetCert.ClearCache();
            moDoGetCert.bGetClientCertificates(lsHash, lbtArrayMinProfile);
            moDoGetCert.bReplaceSsoThumbprint(lsHash, lbtArrayMinProfile);
            moDoGetCert.bGetCertificate();
            this.ShowMe();

            this.bMainProcessRunning = false;
        }
    }
}
