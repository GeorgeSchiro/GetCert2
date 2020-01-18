using Microsoft.Web.Administration;
using System;
using System.Collections;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using tvToolbox;

namespace GetCert2
{
    /// <summary>
    /// GetCert2.exe
    ///
    /// Run this program. It will prompt you to create a default profile file:
    ///
    ///     GetCert2.exe.config
    ///
    /// The profile will contain help (see -Help) as well as default options.
    ///
    /// Note: This software creates its own support files (including DLLs or
    ///       other EXEs) in the folder that contains it. It will prompt you
    ///       to create its own desktop folder when you run it the first time.
    /// </summary>


    /// <summary>
    /// Get digital certificate from a certificate provider network.
    /// </summary>
    public class DoGetCert : System.Windows.Application
    {
        private bool        mbPowerScriptError;
        private ChannelFactory<GetCertService.IGetCertServiceChannel> moGetCertServiceFactory = null;
        private string      msPowerScriptOutput;

        public static string sFetchPrefix = "Resources.Fetch.";
        public static string sHostProcess = "GoPcBackup.exe";

        private DoGetCert()
        {
            tvMessageBox.ShowError(null, "Don't use this constructor!");
        }


        /// <summary>
        /// This constructor expects a profile object to be provided.
        /// </summary>
        /// <param name="aoProfile">
        /// The given profile object must either contain runtime options
        /// or it will be returned filled with default runtime options.
        /// </param>
        public DoGetCert(tvProfile aoProfile)
        {
            moProfile = aoProfile;

            this.SessionEnding += new System.Windows.SessionEndingCancelEventHandler(DoGetCert_SessionEnding);
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            DoGetCert       loMain  = null;
            tvProfile       loProfile = null;
            tvMessageBox    loStartupWaitMsg = null;
            string          lsExePathFile = Process.GetCurrentProcess().MainModule.FileName;

            try
            {
                // Make sure any previous parent instance is shutdown before proceeding
                // (but only after everything has been installed and configured first).
                if ( File.Exists(lsExePathFile + ".config") )
                    DoGetCert.bKillProcessParent(new tvProfile(args), Process.GetCurrentProcess());

                loProfile = DoGetCert.oHandleUpdates(ref args, ref loStartupWaitMsg);

                if ( !loProfile.bExit )
                {
                    loProfile.GetAdd("-Help",
                            @"
Introduction


This utility gets a digital certificate from the FREE ""Let's Encrypt"" certificate
provider network (see ""LetsEncrypt.org""). It installs the certificate in your
server's local computer certificate store and binds it to port 443 in IIS (it's 443
by default - any other port can be used instead).

If the current time is not within a given number of days prior to expiration
of the current digital certificate (eg. 30 days, see -ExpirationDaysBeforeRenewal
below), this software does nothing. Otherwise, the retrieval process begins.

It's as simple as that when the software runs in ""stand-alone"" mode (the default).

If ""stand-alone"" mode is disabled (see -UseStandAloneMode below), the certificate
retrieval process is used in concert with the secure certificate service (SCS),
see ""SafeTrust.org"".

If the software is not running in ""stand-alone"" mode, it also copies any new cert
to a file anywhere on the local area network to be picked up by the load balancer
administrator or process. It also replaces the SSO (single sign-on) certificate in
your central SSO configuration (ie. ADFS) and restarts the SSO service on all
servers in any defined SSO server farm. It also replaces all integrated application
SSO certificate references in any number of configuration files anywhere in the
local file system.

Also (when -UseStandAloneMode is False), this software checks for any previously
fetched certificate (by another client running the same software). If found,
it is downloaded directly from the SCS (rather than attempting retrieval from
the certificate provider network). If not found on the SCS, a certificate is
requested from the certificate provider network, uploaded to the SCS (for use
by other clients in the same server farm) and installed locally and bound to
port 443 in IIS. Finally, calls to the certificate provider network can be
overridden by providing this software access to a digital certificate file
anywhere on the local network.


Command-Line Usage


Open this utility's profile file to see additional options available. It is
usually located in the same folder as ""{EXE}"" and has the same name
with "".config"" added (see ""{INI}"").

Profile file options can be overridden with command-line arguments. The
keys for any ""-key=value"" pairs passed on the command-line must match
those that appear in the profile (with the exception of the ""-ini"" key).

For example, the following invokes the use of an alternative profile file:

    {EXE} -ini=NewProfile.txt

This tells the software to run in automatic mode:

    {EXE} -Auto


Author:  George Schiro (GeoCode@SafeTrust.org)

Date:    10/24/2019




Options and Features


The main options for this utility are listed below with their default values.
A brief description of each feature follows.

-AcmePsModuleUseGallery=False

    Set this switch True and the ""Powershell Gallery"" version of ""ACME-PS""
    will be used in lieu of the version embedded in the EXE (see -AcmePsPath
    below).

-AcmePsPath=""ACME-PS""

    ""ACME-PS"" is the primary tool used by this utility to communicate
    with the ""Let's Encrypt"" certificate network. By default, this key is
    set to the ""ACME-PS"" folder which, with no absolute path given, will
    be expected to be found within the folder that contains ""{INI}"".
    Set -AcmePsModuleUseGallery=True (see above) and the OS will look to
    find ""ACME-PS"" in its usual place as a module from the Powershell gallery.

-AdfsThumbprintFiles=""C:\inetpub\wwwroot\web.config""

    This is the path and filename of files that will have their ADFS certificate
    thumbprint replaced whenever the related ADFS certificate changes. Any files
    with the same name at all levels of the directory hierarchy may be updated,
    starting with the given base path. Wildcards may be used in the filename.

    This key may appear any number of times in the profile.

-Auto=False

    Set this switch True to run this utility one time (with no interactive UI)
    then shutdown automatically upon completion. This switch is useful if this
    software is run in a batch process or by a server job scheduler.

-CertificateDomainName= NO DEFAULT VALUE

    This is the subject name (ie. DNS name) of the certificate returned.

-CertificatePrivateKeyExportable=False

    Set this switch True (not recommended) to allow certificate private keys
    to be exportable from the local certificate store to a local disk file. Any
    SA with server access will then have access to the certificate's private key.

-CertificateRenewalDateOverride= NO DEFAULT VALUE

    Set this date value to override the date calculation that subtracts
    -ExpirationDaysBeforeRenewal days (see below) from the current certificate
    expiration date to know when to start fetching a new certificate.

-ContactEmailAddress= NO DEFAULT VALUE

    This is the contact email address the certificate network uses to send
    certificate expiration notices.

-CreateSanSites=True

    If a SAN specific website does not yet exist in IIS, it will be created
    automatically during the first run of the ""get certificate"" process for
    that SAN value. Set this switch False to have all SAN challenges routed
    through the IIS default website (such challenges will typically fail).

-DoStagingTests=True

    Initial testing is done with the certificate provider staging network. Set
    this switch False to use the live production certificate network.

-ExpirationDaysBeforeRenewal=30

    This is the number of days until certificate expiration before automated
    gets of the next new certificate are attempted.

-FetchSource=False

    Set this switch True to fetch the source code for this utility from the EXE.
    Look in the containing folder for a ZIP file with the full project sources.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-KillProcessForcedWaitMS=1000

    This is the maximum milliseconds to wait while force killing a process.

-KillProcessOrderlyWaitSecs=10

    This is the maximum number of seconds given to a process after a ""close""
    command is given before the process is forcibly terminated.

-LogEntryDateTimeFormatPrefix=""yyyy-MM-dd hh:mm:ss:fff tt  ""

    This format string is used to prepend a timestamp prefix to each log entry
    in the process log file (see -LogPathFile below).    

-LogFileDateFormat=""-yyyy-MM-dd""

    This format string is used to form the variable part of each log file output 
    filename (see -LogPathFile below). It is inserted between the filename and 
    the extension.

-LogPathFile=""Logs\{EXE}Log.txt""

    This is the output path\file that will contain the process log. The profile
    filename will be prepended to the default filename (see above) and the current
    date (see -LogFileDateFormat above) will be inserted between the filename and
    the extension.

-MaxCertRenewalLockDelaySecs=300

    Wait a random period each cycle (at most this many seconds) to allow different
    clients the opportunity to lock the certificate renewal (ie. only one client
    at a time per domain can communicate with the certificate provider network).

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. Messages
    are written to the log instead (see -LogPathFile above). You must use this
    switch whenever the software is run via a server computer batch job or job
    scheduler (ie. where no user interaction is permitted).

-PowerScriptPathFile=PowerScript.ps1

    This is the path\file location of the current (ie. temporary) Powershell script
    file.

-PowerScriptSleepMS=200

    This is the number of sleep milliseconds between loops while waiting for the
    Powershell script process to complete.

-PowerScriptTimeoutSecs=300

    This is the maximum number of seconds allocated to any Powershell script
    process to run prior to throwing a timeout exception.

-PowershellExeArgs=-NoProfile -ExecutionPolicy unrestricted -File ""{{0}}"" ""{{1}}""

    These are the arguments passed to the Windows Powershell EXE (see below).

-PowershellExePathFile=C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe

    This is the path\file location of the Windows Powershell EXE.

-RemoveReplacedCert=False

    Set this switch True and the old (ie. previously bound) certificate will be
    removed whenever a new retrieved certificate is bound to replace it.

    Note: this switch is ignored when -UseStandAloneMode is False.

-ReplaceTextSleepMS=200

    This is the number of sleep milliseconds between loops while waiting for the
    ADFS text replacement process to complete.

-ReplaceTextTimeoutSecs=120

    This is the maximum number of seconds allocated to the ADFS text replacement
    process to run prior to throwing a timeout exception.

-ResetStagingLogs=True

    Having to wade through several log sessions during testing can be cumbersome.
    So the default behavior is to clear the log file after it is uploaded to the SCS
    server (following each test). Setting this switch False will retain all previous
    log sessions on the client during testing.

-SanList= SEE PROFILE FOR DEFAULT VALUE

    This is the SAN list (subject alternative names) to appear on the certificate
    when it is generated. It will consist of -CertificateDomainName by default (see
    above). This list can be edited here directly or through the ""SAN List"" button
    in the {EXE} UI. Click the ""SAN List"" button to see the proper format here.

    Note: the {EXE} UI limits you to 100 SAN values (the certificate provider
          does the same). If you add more names than this limit, an error results and
          no certificate will be generated.

-SaveProfile=True

    Set this switch False to prevent saving to the profile file by this utility
    software itself. This is not recommended since process status information is 
    written to the profile after each run.

-SaveSansCmdLine=True

    Set this switch False to leave the profile file untouched after a command line
    has been passed to the EXE and merged with the profile. When true, everything
    but command line keys will be saved. When false, not even status information
    will be written to the profile file (ie. ""{INI}"").

-ScriptADFS= SEE PROFILE FOR DEFAULT VALUE

    This is the Powershell script that updates ADFS servers with new certificates.

-ScriptStage1= SEE PROFILE FOR DEFAULT VALUE

    There are multiple stages involved with the process of getting a certificate
    from the certificate provider network. Each stage has an associated Powershell 
    script. The stages are represented in this profile by -ScriptStage1 thru
    -ScriptStage7.

-ServiceNameLive=""LetsEncrypt""

    This is the name mapped to the live production certificate network service URL.

-ServiceNameStaging=""LetsEncrypt-Staging""

    This is the name mapped to the non-production (ie. ""staging"") certificate
    network service URL.

-ServiceReportEverything=True

    By default, all activity logged on the client during non-interactive mode
    is uploaded to the SCS server. This can be very helpful during automation
    testing. Once testing is complete, set this switch False to report errors
    only, thereby saving a considerable amount of network bandwidth.

    Note: ""non-interactive mode"" means the -Auto switch is set (see above).
          This switch is ignored when -UseStandAloneMode=True.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the profile
    file at startup in command-line format. This may be helpful as a diagnostic.

-SkipAdfsServer=False

    When a client is on an ADFS domain (ie. it's a member of an ADFS server farm),
    it will automatically attempt to update ADFS services with each new certificate
    retrieved. Set this switch True to disable ADFS updates (for this client only).

-SkipAdfsThumbprintUpdates=False

    When a client's domain references an ADFS domain, the client will automatically
    attempt to update configuration files with each new ADFS certificate thumbprint.
    Set this switch True to disable ADFS thumbprint configuration updates (for this
    client only).

-SubmissionRetries=999

    Pending submissions to the certificate provider network will be retried until
    they succeed or fail, by default, at most this many times.

-SubmissionWaitSecs=5

    These are the seconds of wait time after the DNS website challenge has been
    submitted to the certificate network as well as after the certificate request
    has been submitted. This is the amount of time during which the request should
    transition from a ""pending"" state to anything other than ""pending"".

-UseStandAloneMode=True

    Set this switch False and the software will use the SafeTrust Secure Certificate
    Service (see ""SafeTrust.org"") to manage certificates between several servers
    in a server farm, on SSO servers, SSO integrated application servers and load
    balancers.


Notes:

    There may be various other settings that can be adjusted also (user
    interface settings, etc). See the profile file (""{INI}"")
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added ""on the fly""
    (in order of execution) to ""{INI}"" as the software runs.

"
                            .Replace("{EXE}", Path.GetFileName(ResourceAssembly.Location))
                            .Replace("{INI}", Path.GetFileName(loProfile.sActualPathFile))
                            .Replace("{{", "{")
                            .Replace("}}", "}")
                            );

                    string lsFetchName = null;

                    // Fetch simple setup.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="Setup Application Folder.exe"), lsFetchName);

                    // Fetch source code.
                    if ( loProfile.bValue("-FetchSource", false) )
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name, "GetCert2.zip", null);

                    if ( !loProfile.bExit )
                    {
                        if ( !loProfile.bValue("-Auto", false) )
                        {
                            // Run in interactive mode.

                            try
                            {
                                loMain = new DoGetCert(loProfile);

                                // Load the UI.
                                UI  loUI = new UI(loMain);
                                    loUI.oStartupWaitMsg = loStartupWaitMsg;
                                    loMain.oUI = loUI;

                                loMain.Run(loUI);
                            }
                            catch (ObjectDisposedException) {}
                        }
                        else
                        {
                            // Run in batch mode.

                            DoGetCert   loDoDa = new DoGetCert(loProfile);
                                        if ( null != loStartupWaitMsg )
                                            loStartupWaitMsg.Close();
                                        bool lbGetCertificate = loDoDa.bGetCertificate();
                                        bool lbReplaceAdfsThumbprint = loDoDa.bReplaceAdfsThumbprint();

                            ChannelFactory<GetCertService.IGetCertServiceChannel> loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");

                            string lsProfile = DoGetCert.sProfile(loProfile);
                            string lsHash = HashClass.sHashIt(new tvProfile(lsProfile));
                            string lsLogFileTextReported = null;

                            if ( !lbGetCertificate || !lbReplaceAdfsThumbprint )
                            {
                                DoGetCert.ReportErrors(loProfile, loGetCertServiceFactory, lsHash, lsProfile, out lsLogFileTextReported);
                            }
                            else
                            {
                                DoGetCert.ReportEverything(loProfile, loGetCertServiceFactory, lsHash, lsProfile, out lsLogFileTextReported);
                            }

                            loProfile["-PreviousProcessOutputText"] = lsLogFileTextReported
                                    + "\r\nNote: the timestamp prepended to each text line implies the software was last run using the \"-Auto\" switch.";
                            loProfile.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DoGetCert.LogIt(loProfile, DoGetCert.sExceptionMessage(ex));

                try
                {
                    ChannelFactory<GetCertService.IGetCertServiceChannel> loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");

                    string lsProfile = DoGetCert.sProfile(loProfile);
                    string lsHash = HashClass.sHashIt(new tvProfile(lsProfile));
                    string lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(loProfile, loGetCertServiceFactory, lsHash, lsProfile, out lsLogFileTextReported);
                }
                catch {}
            }
            finally
            {
                if ( null != loMain && null != loMain.oUI )
                    loMain.oUI.Close();
            }
        }


        /// <summary>
        /// This is the main application profile object.
        /// </summary>
        public tvProfile oProfile
        {
            get
            {
                return moProfile;
            }
        }
        private tvProfile moProfile;

        /// <summary>
        /// This is the main application user interface (UI) object.
        /// </summary>
        public UI oUI
        {
            get;set;
        }

        public bool bCertificateSetupDone
        {
            get
            {
                return moProfile.bValue("-CertificateSetupDone", false);
            }
            set
            {
                if ( null != goSetupCertificate )
                {
                    // Get setup cert's private key file.
                    string  lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, goSetupCertificate);
                            if ( null != lsMachineKeyPathFile )
                            {
                                // Remove cert's private key file (sadly, the OS typically let's these things accumulate forever).
                                File.Delete(lsMachineKeyPathFile);
                            }
                }

                moProfile["-CertificateSetupDone"] = value;
                moProfile.Save();
            }
        }

        /// <summary>
        /// This switch is used within the main loop of the UI (see "this.oUI")
        /// to stop any loops running within this object as well.
        /// </summary>
        public bool bMainLoopStopped
        {
            get
            {
                return mbMainLoopStopped;
            }
            set
            {
                mbMainLoopStopped = value;
            }
        }
        private bool mbMainLoopStopped;

        /// <summary>
        /// This is used to suppress prompts when
        /// the software is run via a job scheduler.
        /// </summary>
        public bool bNoPrompts
        {
            get
            {
                return moProfile.bValue("-NoPrompts", false);
            }
        }

        /// <summary>
        /// This is the domain profile object.
        /// </summary>
        public tvProfile oDomainProfile
        {
            get
            {
                if ( null == moDomainProfile )
                {
                    if ( moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        moDomainProfile = new tvProfile();
                    }
                    else
                    {
                        if ( null == moGetCertServiceFactory )
                            moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");

                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                        {
                            string lsProfile = DoGetCert.sProfile(moProfile);
                            string lsHash = HashClass.sHashIt(new tvProfile(lsProfile));

                            moDomainProfile = new tvProfile(loGetCertServiceClient.sDomainProfile(lsHash, lsProfile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                    }
                }

                return moDomainProfile;
            }
        }
        private tvProfile moDomainProfile;

        /// <summary>
        /// Returns the current IIS port 443 bound certificate name (if any).
        /// </summary>
        public static string sCurrentCertificateName
        {
            get
            {
                return DoGetCert.sCertName(DoGetCert.oCurrentCertificate());
            }
        }

        /// <summary>
        /// Returns the current server computer name.
        /// </summary>
        public static string sCurrentComputerName
        {
            get
            {
                return DoGetCert.oEnvProfile().sValue("-COMPUTERNAME", "Computer name not found.");
            }
        }

        public bool bReplaceAdfsThumbprint()
        {
            bool lbReplaceAdfsThumbprint = false;

            // Do nothing if this is an ADFS server (or we are not doing thrumprint updates on this server).
            if ( this.oDomainProfile.bValue("-IsAdfsDomain", false) || moProfile.bValue("-SkipAdfsThumbprintUpdates", false) )
                return true;

            // Do nothing if no ADFS domain is defined for the current domain.
            string  lsAdfsDnsName = this.oDomainProfile.sValue("-AdfsDnsName", "");
                    if ( "" == lsAdfsDnsName )
                        return true;
            // At this point it is an error if no ADFS thumbprint exists for the current domain.
            string  lsAdfsThumbprint = this.oDomainProfile.sValue("-AdfsThumbprint", "");
                    if ( "" == lsAdfsThumbprint )
                    {
                        this.LogIt("");
                        throw new Exception(String.Format("The ADFS certificate thumbprint for the \"{0}\" domain has not yet been set!", lsAdfsDnsName));
                    }
            string  lsAdfsOldThumbprint = moProfile.sValue("-AdfsThumbprint", "");
                    if ( "" == lsAdfsOldThumbprint )
                    {
                        // Define the previous ADFS thumbprint if this is the first time through.
                        moProfile["-AdfsThumbprint"] = lsAdfsThumbprint;
                        moProfile.Save();
                        return true;
                    }
                    // There is nothing to do if the ADFS thumbprint hasn't changed.
                    if ( lsAdfsOldThumbprint == lsAdfsThumbprint )
                        return true;

                        // Add the default filespec.
                        moProfile.sValue("-AdfsThumbprintFiles", @"C:\inetpub\wwwroot\web.config");
            tvProfile   loFileList = new tvProfile();
                        foreach(DictionaryEntry loEntry in moProfile.oOneKeyProfile("-AdfsThumbprintFiles"))
                        {
                            // Change "-AdfsThumbprintFiles" to "-Files".
                            loFileList.Add("-Files", loEntry.Value);
                        }
            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(this.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(this.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = "ReplaceText.exe";
                        loProcess.StartInfo.Arguments = "-OldText={OldAdfsThumbprint} -NewText={NewAdfsThumbprint}"
                                .Replace("{OldAdfsThumbprint}", lsAdfsOldThumbprint).Replace("{NewAdfsThumbprint}", lsAdfsThumbprint)
                                + loFileList.sCommandLine()
                                ;
                        loProcess.StartInfo.UseShellExecute = false;
                        loProcess.StartInfo.RedirectStandardError = true;
                        loProcess.StartInfo.RedirectStandardInput = true;
                        loProcess.StartInfo.RedirectStandardOutput = true;
                        loProcess.StartInfo.CreateNoWindow = true;
            string      lsFetchName = null;
                        // Fetch ReplaceText.exe.
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="ReplaceText.exe"), lsFetchName);
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="ReplaceText.exe.txt"), lsFetchName);

            System.Windows.Forms.Application.DoEvents();

            mbPowerScriptError = false;
            msPowerScriptOutput = null;

            this.LogIt("");
            this.LogIt("Replacing ADFS Thumbprint ...");
            this.LogIt(loProcess.StartInfo.Arguments);

            loProcess.Start();
            loProcess.BeginErrorReadLine();
            loProcess.BeginOutputReadLine();

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(moProfile.dValue("-ReplaceTextTimeoutSecs", 120));

            // Wait for the process to finish.
            while ( !this.bMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !this.bMainLoopStopped )
                    Thread.Sleep(moProfile.iValue("-ReplaceTextSleepMS", 200));
            }

            if ( !this.bMainLoopStopped )
            {
                if ( !loProcess.HasExited )
                    this.LogIt(moProfile.sValue("-ReplaceTextTimeoutMsg", "*** thumbprint replacement sub-process timed-out ***\r\n\r\n"));

                if ( mbPowerScriptError || loProcess.ExitCode != 0 || !loProcess.HasExited )
                {
                    this.LogIt(DoGetCert.sExceptionMessage("The thumbprint replacement sub-process experienced a critical failure."));
                }
                else
                {
                    lbReplaceAdfsThumbprint = true;

                    this.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            DoGetCert.bKillProcess(moProfile, loProcess);

            moProfile["-AdfsThumbprint"] = lsAdfsThumbprint;
            moProfile.Save();

            File.Delete("ReplaceText.exe");

            return lbReplaceAdfsThumbprint;
        }

        /// <summary>
        /// Returns the current certificate collection.
        /// </summary>
        public static X509Certificate2Collection oCurrentCertificateCollection()
        {
            X509Certificate2Collection  loCurrentCertificateCollection = null;
                                        // Open the local machine certificate store (ie "Local Computer / Personal / Certificates").
            X509Store                   loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                                        loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadOnly);
                                        loCurrentCertificateCollection = loStore.Certificates;
                                        loStore.Close();

            return loCurrentCertificateCollection;
        }

        /// <summary>
        /// Returns the certificate name from a certificate.
        /// </summary>
        /// <param name="aoCertificate"></param>
        public static string sCertName(X509Certificate2 aoCertificate)
        {
            string lsCertificateName = null;

            if ( null != aoCertificate )
            {
                string[] lsSubjectArray = aoCertificate.Subject.Split(',');

                lsCertificateName = lsSubjectArray[0].Replace("CN=", "");
            }

            return lsCertificateName;
        }

        /// <summary>
        /// Returns the current "LogPathFile" name.
        /// </summary>
        public static string sLogPathFile(tvProfile aoProfile)
        {
            return DoGetCert.sUniqueIntOutputPathFile(
                        DoGetCert.sLogPathFileBase(aoProfile)
                    , aoProfile.sValue("-LogFileDateFormat", "-yyyy-MM-dd")
                    , true
                    );
        }

        public static string sProfile(tvProfile aoProfile)
        {
            X509Certificate2    loCurrentCertificate = DoGetCert.oCurrentCertificate(aoProfile.sValue("-CertificateDomainName" ,""));
            tvProfile           loProfile = new tvProfile(aoProfile.ToString());
                                loProfile.Remove("-Help");
                                loProfile.Remove("-PreviousProcessOutputText");
                                loProfile.Add("-CurrentLocalTime", DateTime.Now);
                                if ( null != loCurrentCertificate )
                                {
                                    DateTime    ldtBeforeExpirationDate;
                                                if ( aoProfile.ContainsKey("-CertificateRenewalDateOverride") )
                                                    ldtBeforeExpirationDate = aoProfile.dtValue("-CertificateRenewalDateOverride", DateTime.Now);
                                                else
                                                    ldtBeforeExpirationDate = loCurrentCertificate.NotAfter.AddDays(-aoProfile.iValue("-ExpirationDaysBeforeRenewal", 30));

                                    loProfile.Add("-CertificateRenewalDate", ldtBeforeExpirationDate);
                                }
                                loProfile.Add("-COMPUTERNAME", DoGetCert.sCurrentComputerName);

            return loProfile.ToString();
        }

        /// <summary>
        /// Write the given asMessageText to a text file as well as
        /// to the output console of the UI window (if it exists).
        /// </summary>
        /// <param name="asMessageText">The text message string to log.</param>
        public void LogIt(string asMessageText)
        {
            StreamWriter loStreamWriter = null;

            try
            {
                loStreamWriter = new StreamWriter(moProfile.sRelativeToProfilePathFile(DoGetCert.sLogPathFile(moProfile)), true);
                loStreamWriter.WriteLine(DateTime.Now.ToString(moProfile.sValueNoTrim(
                        "-LogEntryDateTimeFormatPrefix", "yyyy-MM-dd hh:mm:ss:fff tt  "))
                        + asMessageText);
            }
            catch { /* Can't log a log failure. */ }
            finally
            {
                if ( null != loStreamWriter )
                    loStreamWriter.Close();
            }

            if ( null != this.oUI )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                this.oUI.AppendOutputTextLine(asMessageText);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

        public void LogStage(string asStageId)
        {
            this.LogIt("");
            this.LogIt("");
            this.LogIt(String.Format("Stage {0} ...", asStageId));
            this.LogIt("");
        }

        public void LogSuccess()
        {
            this.LogIt("Success.");
        }

        public static void ReportErrors(tvProfile aoProfile, ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory
                , string asHash, string asProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = aoProfile.sRelativeToProfilePathFile(DoGetCert.sLogPathFile(aoProfile));
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( aoProfile.bValue("-UseStandAloneMode", true) )
                return;

            string  lsPreviousErrorsLogPathFile = aoProfile.sRelativeToProfilePathFile(aoProfile.sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                    if ( File.Exists(lsPreviousErrorsLogPathFile) )
                        File.AppendAllText(lsPreviousErrorsLogPathFile, asLogFileTextReported);
                    else
                        File.Copy(lsLogPathFile, lsPreviousErrorsLogPathFile, false);

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
            {
                loGetCertServiceClient.ReportErrors(asHash, asProfile, File.ReadAllText(lsPreviousErrorsLogPathFile));
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                File.Delete(lsPreviousErrorsLogPathFile);
            }
        }

        public static void ReportEverything(tvProfile aoProfile, ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory
                , string asHash, string asProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = aoProfile.sRelativeToProfilePathFile(DoGetCert.sLogPathFile(aoProfile));
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( aoProfile.bValue("-UseStandAloneMode", true) || !aoProfile.bValue("-ServiceReportEverything", true) )
                return;

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
            {
                loGetCertServiceClient.ReportEverything(asHash, asProfile, asLogFileTextReported);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( aoProfile.bValue("-DoStagingTests", true) && aoProfile.bValue("-ResetStagingLogs", true) )
                File.Delete(lsLogPathFile);
        }

        public void ShowError(string asMessageText, string asMessageCaption)
        {
            this.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowError(this.oUI, asMessageText, asMessageCaption);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

     
        // This kludge is necessary since neither "Process.CloseMainWindow()" nor "Process.Kill()" work reliably.
        private static bool bKillProcess(tvProfile aoProfile, Process aoProcess)
        {
            bool    lbKillProcess = true;
                    if ( null == aoProcess )
                        return lbKillProcess;

            int     liProcessId = 0;

            // First try ending the process in the usual way (ie. an orderly shutdown).

            Process loKillProcess = new Process();

            try
            {
                liProcessId = aoProcess.Id;

                loKillProcess = new Process();
                loKillProcess.StartInfo.FileName = "taskkill";
                loKillProcess.StartInfo.Arguments = "/pid " + liProcessId;
                loKillProcess.StartInfo.UseShellExecute = false;
                loKillProcess.StartInfo.CreateNoWindow = true;
                loKillProcess.Start();

                aoProcess.WaitForExit(1000 * aoProfile.iValue("-KillProcessOrderlyWaitSecs", 10));
            }
            catch (Exception ex)
            {
                lbKillProcess = false;
                DoGetCert.LogIt(aoProfile, String.Format("Orderly bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, DoGetCert.sExceptionMessage(ex)));
            }

            if ( lbKillProcess && !aoProcess.HasExited )
            {
                // Then try terminating the process forcefully (ie. slam it down).

                try
                {
                    loKillProcess = new Process();
                    loKillProcess.StartInfo.FileName = "taskkill";
                    loKillProcess.StartInfo.Arguments = "/f /pid " + liProcessId;
                    loKillProcess.StartInfo.UseShellExecute = false;
                    loKillProcess.StartInfo.CreateNoWindow = true;
                    loKillProcess.Start();

                    aoProcess.WaitForExit(aoProfile.iValue("-KillProcessForcedWaitMS", 1000));

                    lbKillProcess = aoProcess.HasExited;
                }
                catch (Exception ex)
                {
                    lbKillProcess = false;
                    DoGetCert.LogIt(aoProfile, String.Format("Forced bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, DoGetCert.sExceptionMessage(ex)));
                }
            }

            return lbKillProcess;
        }

        private static bool bKillProcessParent(tvProfile aoProfile, Process aoProcess)
        {
            bool lbKillProcess = true;

            foreach (Process loProcess in Process.GetProcessesByName(aoProcess.ProcessName))
            {
                if ( loProcess.Id != aoProcess.Id )
                    lbKillProcess = DoGetCert.bKillProcess(aoProfile, loProcess);
            }

            return lbKillProcess;
        }

        /// <summary>
        /// Returns the current IIS port 443 bound certificate (if any).
        /// </summary>
        private static X509Certificate2 oCurrentCertificate()
        {
            return DoGetCert.oCurrentCertificate(null);
        }
        private static X509Certificate2 oCurrentCertificate(string asCertName)
        {
            ServerManager   loServerManager = null;
            Site            loSite = null;
            Binding         loBinding = null;

            return DoGetCert.oCurrentCertificate(asCertName, out loServerManager, out loSite, out loBinding);
        }
        private static X509Certificate2 oCurrentCertificate(string asCertName, out ServerManager aoServerManager)
        {
            Site            loSite = null;
            Binding         loBinding = null;

            return DoGetCert.oCurrentCertificate(asCertName, out aoServerManager, out loSite, out loBinding);
        }
        private static X509Certificate2 oCurrentCertificate(string asCertName, out ServerManager aoServerManager, out Site aoSiteFound)
        {
            Binding         loBinding = null;

            return DoGetCert.oCurrentCertificate(asCertName, out aoServerManager, out aoSiteFound, out loBinding);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , out Binding aoBindingFound
                )
        {
            string[]        lsSanArray = null;
            Site            loDefaultSite = null;

            return DoGetCert.oCurrentCertificate(asCertName, out aoServerManager, out aoSiteFound, out aoBindingFound, lsSanArray, out loDefaultSite);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , out Site aoSiteFound
                , string[] asSanArray
                )
        {
            ServerManager   loServerManager = null;
            Binding         loBinding = null;
            Site            loDefaultSite = null;

            return DoGetCert.oCurrentCertificate(asCertName, out loServerManager, out aoSiteFound, out loBinding, asSanArray, out loDefaultSite);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , string[] asSanArray
                , out Site aoDefaultSiteFound
                )
        {
            Binding         loBinding = null;

            return DoGetCert.oCurrentCertificate(asCertName, out aoServerManager, out aoSiteFound, out loBinding, asSanArray, out aoDefaultSiteFound);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , out Binding aoBindingFound
                , string[] asSanArray
                , out Site aoDefaultSiteFound
                )
        {
            // ServerManager is the IIS ServerManager. It gives us the website for binding the cert.
            aoServerManager = new ServerManager();
            aoSiteFound = null;
            aoBindingFound = null;
            aoDefaultSiteFound = null;

            string              lsCertName = null == asCertName ? null : asCertName.ToLower();
            X509Certificate2    loCurrentCertificate = null;

            foreach(Site loSite in aoServerManager.Sites)
            {
                // First, look for a binding by matching certificate.
                foreach (Binding loBinding in loSite.Bindings)
                {
                    // Use the first binding found with a certificate (and a matching name - if lsCertName is not null).
                    foreach(X509Certificate2 loCertificate in DoGetCert.oCurrentCertificateCollection())
                    {
                        if ( null != loBinding.CertificateHash && loCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash)
                            && (null == lsCertName || lsCertName == DoGetCert.sCertName(loCertificate).ToLower()) )
                        {
                            // Binding found that matches a certificate and certificate name (or no name).
                            loCurrentCertificate = loCertificate;
                            aoSiteFound = loSite;
                            aoBindingFound = loBinding;
                            break;
                        }
                    }

                    if ( null != aoBindingFound )
                        break;
                }

                // Next, if no binding certificate match exists, look by binding hostname (ie. SNI).
                if ( null == aoBindingFound )
                    foreach (Binding loBinding in loSite.Bindings)
                    {
                        if ( null == loBinding.CertificateHash && null != lsCertName
                                && lsCertName == loBinding.Host.ToLower() && "https" == loBinding.Protocol )
                        {
                            // Site found with a binding hostname that matches lsCertName.
                            aoSiteFound = loSite;
                            aoBindingFound = loBinding;
                            break;
                        }

                        if ( null != aoBindingFound )
                            break;
                    }

                // Finally, with no binding found, try to match against the site name.
                if ( null == aoBindingFound && asCertName == loSite.Name )
                {
                    // Site found with a name that matches lsCertName.
                    aoSiteFound = loSite;
                    aoBindingFound = DoGetCert.oSslBinding(loSite);
                    break;
                }

                if ( null != aoBindingFound )
                    break;
            }

            // Finally, finally (no really), with no site found, find a related primary site to use for defaults.
            string lsPrimaryCertName = null == asSanArray || 0 == asSanArray.Length ? null : asSanArray[0].ToLower();

            if ( null == aoSiteFound && null != lsPrimaryCertName && lsCertName != lsPrimaryCertName )
            {
                // Use the primary site in the SAN array as defaults for any new site (to be created).
                DoGetCert.oCurrentCertificate(lsPrimaryCertName, out aoDefaultSiteFound, asSanArray);
            }

            return loCurrentCertificate;
        }

        private static Binding oSslBinding(Site aoSite)
        {
            Binding loSslBinding = null;

            if ( null != aoSite )
                foreach (Binding loBinding in aoSite.Bindings)
                {
                    if ( "https" == loBinding.Protocol )
                        loSslBinding = loBinding;
                }

            return loSslBinding;
        }

        private static tvProfile oEnvProfile()
        {
            if ( null == goEnvProfile )
            {
                goEnvProfile = new tvProfile();

                foreach(DictionaryEntry loEntry in Environment.GetEnvironmentVariables())
                {
                    goEnvProfile.Add("-" + loEntry.Key.ToString(), loEntry.Value);
                }
            }

            return goEnvProfile;
        }
        private static tvProfile goEnvProfile = null;

        public static string sExceptionMessage(Exception ex)
        {
            return "GetCertServiceFault: " + ex.Message + (null == ex.InnerException ? "": "; " + ex.InnerException.Message) + "\r\n" + ex.StackTrace;
        }
        public static string sExceptionMessage(string asMessage)
        {
            return "GetCertServiceFault: " + asMessage;
        }

        private static string sProcessExePathFile(string asExePathFile)
        {
            Process loProcessFound = null;

            return DoGetCert.sProcessExePathFile(asExePathFile, out loProcessFound);
        }
        private static string sProcessExePathFile(string asExePathFile, out Process aoProcessFound)
        {
            string  lsFindProcessExePathFile = null;
            string  lsProcessName = Path.GetFileNameWithoutExtension(asExePathFile);

            aoProcessFound = null;

            foreach (Process loProcess in Process.GetProcessesByName(lsProcessName))
            {
                aoProcessFound = loProcess;
                lsFindProcessExePathFile = loProcess.MainModule.FileName;
            }

            return lsFindProcessExePathFile;
        }

        /// <summary>
        /// Returns the current "LogPathFile" base name.
        /// </summary>
        private static string sLogPathFileBase(tvProfile aoProfile)
        {
            string  lsPath = "Logs";
            string  lsLogFile = "Log.txt";
            string  lsLogPathFileBase = lsLogFile;

            try
            {
                lsLogPathFileBase = aoProfile.sValue("-LogPathFile"
                        , Path.Combine(lsPath, Path.GetFileNameWithoutExtension(aoProfile.sLoadedPathFile) + lsLogFile));

                lsPath = Path.GetDirectoryName(aoProfile.sRelativeToProfilePathFile(lsLogPathFileBase));
                if ( !Directory.Exists(lsPath) )
                    Directory.CreateDirectory(lsPath);
            }
            catch {}

            return lsLogPathFileBase;
        }

        /// <summary>
        /// Returns a unique output pathfile based on an integer.
        /// </summary>
        /// <param name="asBasePathFile">The base pathfile string.
        /// It is segmented to form the output pathfile</param>
        /// <param name="asDateFormat">The format of the date inserted into the output filename.</param>
        /// <param name="abAppendOutput">If true, this boolean indicates that an existing file will be appended to. </param>
        /// <returns></returns>
        private static string sUniqueIntOutputPathFile(
                  string asBasePathFile
                , string asDateFormat
                , bool abAppendOutput
                )
        {
            // Get the path from the given asBasePathFile.
            string  lsOutputPath = Path.GetDirectoryName(asBasePathFile);
            // Make a filename from the given asBasePathFile and the current date.
            string  lsBaseFilename = Path.GetFileNameWithoutExtension(asBasePathFile)
                    + DateTime.Today.ToString(asDateFormat);
            // Get the filename extention from the given asBasePathFile.
            string  lsBaseFileExt = Path.GetExtension(asBasePathFile);

            string  lsOutputFilename = lsBaseFilename + lsBaseFileExt;
            int     liUniqueFilenameSuffix = 1;
            string  lsOutputPathFile = null;
            bool    lbDone = false;

            do
            {
                // If we are appending, we're done. Otherwise,
                // check for existence of the requested dated pathfile.
                lsOutputPathFile = Path.Combine(lsOutputPath, lsOutputFilename);
                lbDone = abAppendOutput | !File.Exists(lsOutputPathFile);

                // If the given pathfile already exists, create a variation on the dated
                // filename by appending an integer (see liUniqueFilenameSuffix above).
                // Keep trying until a unique dated pathfile is identified.
                if ( !lbDone )
                    lsOutputFilename = lsBaseFilename + "." + (++liUniqueFilenameSuffix).ToString() + lsBaseFileExt;
            }
            while ( !lbDone );

            return lsOutputPathFile;
        }

        private static void DoCheckIn(tvProfile aoProfile, string asHash, string asProfile, tvMessageBox aoStartupWaitMsg)
        {
            if ( aoProfile.bValue("-UseStandAloneMode", true) )
                return;

            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                if ( null != aoStartupWaitMsg )
                    aoStartupWaitMsg.Hide();

                DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                if ( null != aoStartupWaitMsg && !aoProfile.bExit )
                    aoStartupWaitMsg.Show();

                string lsPreviousErrorsLogPathFile = aoProfile.sRelativeToProfilePathFile(aoProfile.sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                string lsErrorLog = null;
                        if ( File.Exists(lsPreviousErrorsLogPathFile) )
                            lsErrorLog = File.ReadAllText(lsPreviousErrorsLogPathFile);

                loGetCertServiceClient.ClientCheckIn(asHash, asProfile, lsErrorLog);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                DoGetCert.LogIt(aoProfile, "");
                DoGetCert.LogIt(aoProfile, String.Format("Client checked in with the certificate repository{0}.", null == lsErrorLog ? "" : " (and reported previous errors)"));

                File.Delete(lsPreviousErrorsLogPathFile);
            }
        }

        public static void LogIt(tvProfile aoProfile, string asText)
        {
            string          lsLogPathFileBase = null;
            StreamWriter    loStreamWriter = null;

            try
            {
                if ( null == aoProfile )
                    aoProfile = new tvProfile();

                lsLogPathFileBase = DoGetCert.sLogPathFileBase(aoProfile);

                // Move down to the base folder if we're running an update.
                string  lsPathSep = Path.DirectorySeparatorChar.ToString();
                string  lsPath = Path.GetDirectoryName(aoProfile.sRelativeToProfilePathFile(lsLogPathFileBase)).Replace(
                                    String.Format("{0}{1}{0}",  lsPathSep, aoProfile.sValue("-UpdateFolder", "Update")), lsPathSep);
                        if ( !Directory.Exists(lsPath) )
                            Directory.CreateDirectory(lsPath);

                loStreamWriter = new StreamWriter(aoProfile.sRelativeToProfilePathFile(
                        DoGetCert.sUniqueIntOutputPathFile(lsLogPathFileBase
                        , aoProfile.sValue("-LogFileDateFormat", "-yyyy-MM-dd"), true)), true);
                loStreamWriter.WriteLine(DateTime.Now.ToString(aoProfile.sValueNoTrim(
                        "-LogEntryDateTimeFormatPrefix", "yyyy-MM-dd hh:mm:ss:fff tt  "))
                        + asText);
            }
            catch
            {
                try
                {
                    // Give it one last try ...
                    File.WriteAllText(lsLogPathFileBase, asText);
                }
                catch { /* Can't log a log failure. */ }
            }
            finally
            {
                if ( null != loStreamWriter )
                    loStreamWriter.Close();
            }
        }

        /// <summary>
        /// Initial client certificate setup.
        /// </summary>
        private static void SetupCertificate(tvProfile aoProfile, ChannelFactory<GetCertService.IGetCertServiceChannel> aoChannelFactory)
        {
            if ( aoProfile.bValue("-UseStandAloneMode", true) || aoProfile.bValue("-CertificateSetupDone", false) )
                return;

            if ( null == aoChannelFactory.Credentials.ClientCertificate.Certificate )
            {
                if (null == goSetupCertificate)
                    DoGetCert.SetupCertificate(aoProfile, new GetCertService.GetCertServiceClient());

                aoChannelFactory.Credentials.ClientCertificate.Certificate = goSetupCertificate;
            }
        }
        private static void SetupCertificate(tvProfile aoProfile, GetCertService.GetCertServiceClient aoGetCertServiceClient)
        {
            if ( aoProfile.bValue("-UseStandAloneMode", true) || aoProfile.bValue("-CertificateSetupDone", false)
                        || aoProfile.bValue("-NoPrompts", false) || aoProfile.bExit )
                return;

            if ( null == aoGetCertServiceClient.ClientCredentials.ClientCertificate.Certificate )
            {
                if ( null == goSetupCertificate && "" != aoProfile.sValue("-ContactEmailAddress" ,"") )
                    goSetupCertificate = DoGetCert.oCurrentCertificate(aoProfile.sValue("-CertificateDomainName" ,""));

                if ( null == goSetupCertificate )
                {
                    System.Windows.Forms.OpenFileDialog loOpenDialog = new System.Windows.Forms.OpenFileDialog();
                                                        loOpenDialog.FileName = "GetCertClientSetup.pfx";
                    System.Windows.Forms.DialogResult   leDialogResult = System.Windows.Forms.DialogResult.None;
                                                        leDialogResult = loOpenDialog.ShowDialog();

                    if ( System.Windows.Forms.DialogResult.OK != leDialogResult )
                    {
                        aoProfile.bExit = true;
                    }
                    else
                    {
                        string lsPassword = Microsoft.VisualBasic.Interaction.InputBox("Password?", Path.GetFileName(loOpenDialog.FileName), "", -1, 50);

                        if ( "" == lsPassword )
                        {
                            aoProfile.bExit = true;
                        }
                        else
                        {
                            try
                            {
                                goSetupCertificate = new X509Certificate2(loOpenDialog.FileName, lsPassword);
                            }
                            catch (Exception ex)
                            {
                                aoProfile.bExit = true;
                                DoGetCert.LogIt(aoProfile, DoGetCert.sExceptionMessage(ex));
                            }
                        }
                    }
                }

                aoGetCertServiceClient.ClientCredentials.ClientCertificate.Certificate = goSetupCertificate;
            }
        }
        private static X509Certificate2 goSetupCertificate = null;


        // This collection of UI "Show" methods within this non-interactive class allows
        // for pop-up messages to be displayed by using dispatcher calls into the UI object,
        // if it exists. Whether the UI object exists or not, each pop-up message is also
        // written to the log file.


        private tvMessageBoxResults Show(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                , string asProfilePromptKey
                )
        {
            this.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                ltvMessageBoxResults = tvMessageBox.Show(
                          this.oUI
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        , tvMessageBoxCheckBoxTypes.SkipThis
                        , moProfile
                        , asProfilePromptKey
                        );
            }
            ));
            System.Windows.Forms.Application.DoEvents();

            return ltvMessageBoxResults;
        }

        private tvMessageBoxResults ShowModeless(
                  string asMessageText
                , string asMessageCaption
                , tvMessageBoxButtons aetvMessageBoxButtons
                , tvMessageBoxIcons aetvMessageBoxIcon
                , string asProfilePromptKey
                )
        {
            this.LogIt(asMessageText);

            tvMessageBoxResults ltvMessageBoxResults = tvMessageBoxResults.None;

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                ltvMessageBoxResults = tvMessageBox.ShowModeless(
                          this.oUI
                        , asMessageText
                        , asMessageCaption
                        , aetvMessageBoxButtons
                        , aetvMessageBoxIcon
                        , tvMessageBoxCheckBoxTypes.SkipThis
                        , moProfile
                        , asProfilePromptKey
                        );
            }
            ));
            System.Windows.Forms.Application.DoEvents();

            return ltvMessageBoxResults;
        }

        private void ShowError(string asMessageText)
        {
            this.LogIt(asMessageText);

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowError(this.oUI, asMessageText);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

        private void ShowModelessError(string asMessageText, string asMessageCaption, string asProfilePromptKey)
        {
            this.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowModelessError(
                          this.oUI
                        , asMessageText
                        , asMessageCaption
                        , moProfile
                        , asProfilePromptKey
                        );
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }


        private void DoGetCert_SessionEnding(object sender, System.Windows.SessionEndingCancelEventArgs e)
        {
            this.LogIt("");
            this.LogIt(String.Format("{0} exiting due to log-off or system shutdown.", Path.GetFileName(System.Windows.Application.ResourceAssembly.Location)));

            if ( null != this.oUI )
            this.oUI.HandleShutdown();
        }

        private bool bFixTextFile(string asPathFile, string asText)
        {
            bool lbFixTextFile = false;

            this.LogIt("");
            this.LogIt(String.Format("Adjusting \"{0}\" ...", asPathFile));

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(moProfile.sRelativeToProfilePathFile(asPathFile)));
                File.WriteAllText(moProfile.sRelativeToProfilePathFile(asPathFile), asText);
                lbFixTextFile = true;
                this.LogIt("");
            }
            catch (Exception ex)
            {
                this.LogIt(DoGetCert.sExceptionMessage(ex));
            }

            return lbFixTextFile;
        }

        private bool bRunPowerScript(string asScript)
        {
            string lsOutput = null;

            return this.bRunPowerScript(out lsOutput, asScript);
        }
        private bool bRunPowerScript(string asScript, bool abNoSession)
        {
            string lsOutput = null;

            return this.bRunPowerScript(out lsOutput, asScript, abNoSession);
        }
        private bool bRunPowerScript(out string asOutput, string asScript)
        {
            return this.bRunPowerScript(out asOutput, asScript, false);
        }
        private bool bRunPowerScript(out string asOutput, string asScript, bool abNoSession)
        {
            bool            lbRunPowerScript = false;
            string          lsSessionScriptPathFile = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-PowerScriptSessionPathFile", "InGetCertSession.ps1"));
                            // Fetch session script.
                            tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                    , String.Format("{0}{1}", DoGetCert.sFetchPrefix, Path.GetFileName(lsSessionScriptPathFile)), lsSessionScriptPathFile);
            string          lsScriptPathFile = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-PowerScriptPathFile", "PowerScript.ps1"));
            StreamWriter    loStreamWriter   = null;

            try
            {
                loStreamWriter = new StreamWriter(lsScriptPathFile);
                loStreamWriter.WriteLine(asScript);
            }
            catch (Exception ex)
            {
                this.LogIt(DoGetCert.sExceptionMessage(ex));
            }
            finally
            {
                if ( null != loStreamWriter )
                    loStreamWriter.Close();
            }

            string  lsProcessPathFile = moProfile.sValue("-PowerShellExePathFile", @"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe");
            string  lsProcessArgs = moProfile.sValue("-PowerShellExeArgs", @"-NoProfile -ExecutionPolicy unrestricted -File ""{0}"" ""{1}""");
                    if ( abNoSession )
                        lsProcessArgs = String.Format(lsProcessArgs, lsScriptPathFile, "");
                    else
                        lsProcessArgs = String.Format(lsProcessArgs, lsSessionScriptPathFile, lsScriptPathFile);
            string  lsLogScript = !asScript.Contains("-CertificateKey") ? asScript : asScript.Substring(0, asScript.IndexOf("-CertificateKey"));

            if ( !abNoSession )
                this.LogIt(lsLogScript);

            Process loProcess = new Process();
                    loProcess.ErrorDataReceived += new DataReceivedEventHandler(this.PowerScriptProcessErrorHandler);
                    loProcess.OutputDataReceived += new DataReceivedEventHandler(this.PowerScriptProcessOutputHandler);
                    loProcess.StartInfo.FileName = lsProcessPathFile;
                    loProcess.StartInfo.Arguments = lsProcessArgs;
                    loProcess.StartInfo.UseShellExecute = false;
                    loProcess.StartInfo.RedirectStandardError = true;
                    loProcess.StartInfo.RedirectStandardInput = true;
                    loProcess.StartInfo.RedirectStandardOutput = true;
                    loProcess.StartInfo.CreateNoWindow = true;

            System.Windows.Forms.Application.DoEvents();

            mbPowerScriptError = false;
            msPowerScriptOutput = null;

            loProcess.Start();
            loProcess.BeginErrorReadLine();
            loProcess.BeginOutputReadLine();

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(moProfile.dValue("-PowerScriptTimeoutSecs", 300));

            // Wait for the process to finish.
            while ( !this.bMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !this.bMainLoopStopped )
                    Thread.Sleep(moProfile.iValue("-PowerScriptSleepMS", 200));
            }

            if ( !this.bMainLoopStopped )
            {
                if ( !loProcess.HasExited )
                    this.LogIt(moProfile.sValue("-PowerScriptTimeoutMsg", "*** sub-process timed-out ***\r\n\r\n"));

                if ( mbPowerScriptError || loProcess.ExitCode != 0 || !loProcess.HasExited )
                {
                    this.LogIt(DoGetCert.sExceptionMessage("The sub-process experienced a critical failure."));
                }
                else
                {
                    lbRunPowerScript = true;
                    File.Delete(lsScriptPathFile);

                    if ( !abNoSession )
                        this.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            DoGetCert.bKillProcess(moProfile, loProcess);

            asOutput = msPowerScriptOutput;

            return lbRunPowerScript;
        }

        private void PowerScriptProcessOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length && !outLine.Data.Contains("WARNING: PSSession GetCert") )
            {
                msPowerScriptOutput += outLine.Data;
                this.LogIt(outLine.Data);
            }
        }

        private void PowerScriptProcessErrorHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length )
            {
                msPowerScriptOutput += outLine.Data;
                mbPowerScriptError = true;
                this.LogIt(outLine.Data);
            }
        }

        private bool bApplyNewCertToADFS(X509Certificate2 aoNewCertificate, X509Certificate2 aoOldCertificate, string asHash, string asProfile)
        {
            // Do nothing if running stand-alone or this isn't an ADFS server (or we are not doing ADFS certificate updates on this server).
            if ( moProfile.bValue("-UseStandAloneMode", true) || !this.oDomainProfile.bValue("-IsAdfsDomain", false)
                        || null == aoOldCertificate || moProfile.bValue("-SkipAdfsServer", false) )
                return true;
            
            this.LogIt("");
            this.LogIt("But first, replace ADFS certificates ...");

            bool lbApplyNewCertToADFS = this.bRunPowerScript(moProfile.sValue("-ScriptADFS", @"
$ServiceName = ""adfssrv""
# Wait at most 3 minutes for the ADFS service to start (after a likely reboot).
if ( (Get-Service -Name $ServiceName).Status -ne ""Running"" ) {Start-Sleep -s 180}

Add-PSSnapin Microsoft.Adfs.Powershell 2> $null

if ( (Get-AdfsCertificate -CertificateType ""Service-Communications"").Thumbprint -eq ""{NewCertificateThumbprint}"" )
{
    echo ""ADFS certificates already replaced (by another server in the farm).""
    echo ""The ADFS service will be restarted.""
}
else
{
    if ( !$Errors ) {try {Add-AdfsCertificate    -CertificateType ""Token-Decrypting""       -Thumbprint ""{NewCertificateThumbprint}"" -IsPrimary} catch {$Errors = $error}}
    if ( !$Errors ) {try {Remove-AdfsCertificate -CertificateType ""Token-Decrypting""       -Thumbprint ""{OldCertificateThumbprint}""} catch {}}

    if ( !$Errors ) {try {Add-AdfsCertificate    -CertificateType ""Token-Signing""          -Thumbprint ""{NewCertificateThumbprint}"" -IsPrimary} catch {$Errors = $error}}
    if ( !$Errors ) {try {Remove-AdfsCertificate -CertificateType ""Token-Signing""          -Thumbprint ""{OldCertificateThumbprint}""} catch {}}

    if ( !$Errors ) {try {Set-AdfsCertificate    -CertificateType ""Service-Communications"" -Thumbprint ""{NewCertificateThumbprint}""} catch {$Errors = $error}}
}

Restart-Service -Name $ServiceName
                    ")
                    .Replace("{NewCertificateThumbprint}", aoNewCertificate.Thumbprint)
                    .Replace("{OldCertificateThumbprint}", aoOldCertificate.Thumbprint)
                    );

            if ( lbApplyNewCertToADFS )
            {
                this.LogIt("");
                this.LogIt("Updating ADFS token signing certificate thumbprint in the database ...");

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.SetAdfsThumbprint(asHash, asProfile, aoNewCertificate.Thumbprint);

                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }

            return lbApplyNewCertToADFS;
        }

        private bool bCertificateNotExpiring(X509Certificate2 aoOldCertificate)
        {
            bool lbCertificateNotExpiring = false;

            if ( moProfile.bValue("-DoStagingTests", true) )
            {
                this.LogIt("");
                this.LogIt("( staging mode is in effect (-DoStagingTests=True) )");
            }

            if ( null != aoOldCertificate )
            {
                bool lbCheckExpiration = true;

                if ( lbCheckExpiration )
                    lbCheckExpiration = !moProfile.bValue("-DoStagingTests", true);

                if ( lbCheckExpiration )
                {
                    lbCheckExpiration = aoOldCertificate.Verify();
                    if ( !lbCheckExpiration )
                    {
                        this.LogIt(String.Format("The current certificate ({0}) appears to be invalid (if you know otherwise, make sure internet connectivity is available and the system clock is accurate).", aoOldCertificate.Subject));
                        this.LogIt("Expiration status will therefore be ignored and the process will now run.");
                    }
                }

                if ( lbCheckExpiration )
                {
                    DateTime    ldtBeforeExpirationDate;
                                if ( moProfile.ContainsKey("-CertificateRenewalDateOverride") )
                                    ldtBeforeExpirationDate = moProfile.dtValue("-CertificateRenewalDateOverride", DateTime.Now);
                                else
                                    ldtBeforeExpirationDate = aoOldCertificate.NotAfter.AddDays(-moProfile.iValue("-ExpirationDaysBeforeRenewal", 30));

                    if ( DateTime.Now < ldtBeforeExpirationDate )
                    {
                        this.LogIt(String.Format("Nothing to do until {0}. The current \"{1}\" certificate doesn't expire until {2}."
                            , ldtBeforeExpirationDate.ToString(), aoOldCertificate.Subject, aoOldCertificate.NotAfter.ToString()));

                        lbCertificateNotExpiring = true;
                    }
                }
            }

            return lbCertificateNotExpiring;
        }

        private static tvProfile oHandleUpdates(ref string[] args, ref tvMessageBox aoStartupWaitMsg)
        {
            tvProfile   loProfile = null;
            tvProfile   loCmdLine = null;
            string      lsRunKey = "-UpdateRunExePathFile";
            string      lsDelKey = "-UpdateDeletePathFile";
            string      lsUpdateRunExePathFile = null;
            string      lsUpdateDeletePathFile = null;
            string      lsProfile = null;
            string      lsHash = null;

            try
            {
                loCmdLine = new tvProfile();
                loCmdLine.LoadFromCommandLineArray(args, tvProfileLoadActions.Append);

                if ( 0 != args.Length )
                {
                    lsUpdateRunExePathFile = (string)loCmdLine[lsRunKey];
                    lsUpdateDeletePathFile = (string)loCmdLine[lsDelKey];
                    loCmdLine.Remove(lsRunKey);
                    loCmdLine.Remove(lsDelKey);
                    args = loCmdLine.sCommandLineArray();
                }

                loProfile = new tvProfile(args, true);
                if ( loProfile.bExit )
                    return loProfile;

                DoGetCert.VerifyDependenciesInit(loProfile);
                DoGetCert.VerifyDependencies(loProfile);
                if ( loProfile.bExit )
                    return loProfile;

                // Only do the following fetches before initial setup and if the containing application folder has been created.
                if ( "" == loProfile.sValue("-ContactEmailAddress", "")
                        && Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Contains(ResourceAssembly.GetName().Name) )
                {
                    //if ( null == DoGetCert.sProcessExePathFile(DoGetCert.sHostProcess) )
                    //{
                    //    // Fetch host (if it's not already running).
                    //    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    //            , String.Format("{0}{1}", DoGetCert.sFetchPrefix, DoGetCert.sHostProcess), DoGetCert.sHostProcess);
                    //}

                    // Discard the default profile. Replace it with the WCF version fetched below.
                    File.Delete(loProfile.sLoadedPathFile);

                    // Fetch WCF config.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", DoGetCert.sFetchPrefix, "GetCert2.exe.config"), loProfile.sLoadedPathFile);

                    DoGetCert.ResetConfigMechanism(loProfile);
                }

                //if ( !loProfile.bExit && null != DoGetCert.sProcessExePathFile(DoGetCert.sHostProcess) )
                //{
                //    // Remove host (if it's already running).
                //    File.Delete(Path.Combine(Path.GetDirectoryName(loProfile.sLoadedPathFile), DoGetCert.sHostProcess));
                //}

                if ( loProfile.bExit || loProfile.bValue("-UseStandAloneMode", true) )
                    return loProfile;

                lsProfile = DoGetCert.sProfile(loProfile);
                lsHash = HashClass.sHashIt(new tvProfile(lsProfile));

                if ( !loProfile.bValue("-NoPrompts", false) )
                {
                    aoStartupWaitMsg = new tvMessageBox();
                    aoStartupWaitMsg.ShowWait(
                            null, Path.GetFileNameWithoutExtension(loProfile.sExePathFile) + " loading, please wait ...", 250);
                }

                if ( null == lsUpdateRunExePathFile && null == lsUpdateDeletePathFile )
                {
                    // Check-in with the GetCert service.
                    DoGetCert.DoCheckIn(loProfile, lsHash, lsProfile, aoStartupWaitMsg);

                    // Does the "hosts" file need updating?
                    DoGetCert.HandleHostsEntryUpdate(loProfile, lsHash, lsProfile);

                    // A not-yet-updated EXE is running. Does it need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleClientExeUpdate(loProfile, lsHash, lsProfile, lsRunKey, loCmdLine);
                }
                else
                {
                    // An updated EXE is running. It must be copied, relaunched then deleted.

                    if ( null == lsUpdateDeletePathFile )
                    {
                        File.Copy(loProfile.sExePathFile, lsUpdateRunExePathFile, true);

                        Process loProcess = new Process();
                                loProcess.StartInfo.FileName = lsUpdateRunExePathFile;
                                loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2}"
                                        , lsDelKey, loProfile.sExePathFile, loCmdLine.sCommandLine());
                                loProcess.Start();

                                System.Windows.Forms.Application.DoEvents();

                        loProfile.bExit = true;
                    }
                    else
                    {
                        string lsCurrentVersion = FileVersionInfo.GetVersionInfo(lsUpdateDeletePathFile).FileVersion;

                        try
                        {
                            Directory.Delete(Path.GetDirectoryName(lsUpdateDeletePathFile), true);
                        }
                        catch
                        {
                            Thread.Sleep(2000);
                            // These deletes sometimes fail due to parent shutdown
                            // timing. Wait a couple of seconds and try again.
                            Directory.Delete(Path.GetDirectoryName(lsUpdateDeletePathFile), true);
                        }

                        DoGetCert.LogIt(loProfile, String.Format(
                                "Software update successfully completed. Version {0} is now running.", lsCurrentVersion));
                    }
                }

                if ( !loProfile.bExit )
                {
                    loProfile = DoGetCert.HandleClientCfgUpdate(loProfile, lsHash, lsProfile);
                    loProfile = DoGetCert.HandleClientIniUpdate(loProfile, lsHash, lsProfile, args);

                    DoGetCert.HandleHostUpdates(loProfile, lsHash, lsProfile);
                }
            }
            catch (Exception ex)
            {
                if ( null != loProfile )
                    loProfile.bExit = true;

                DoGetCert.LogIt(loProfile, DoGetCert.sExceptionMessage(ex));

                if ( !loProfile.bValue("-UseStandAloneMode", true) && !loProfile.bValue("-CertificateSetupDone", false) )
                {
                    DoGetCert.LogIt(loProfile, @"

Checklist for {EXE} setup (SCS version):

    On this local server:

        .) Once ""-CertificateDomainName"" and ""-ContactEmailAddress"" have been
           provided (after the initial use of the ""GetCertClientSetup"" certificate),
           the SCS will expect to use a domain specific certificate for further
           communications. You can blank out either ""-CertificateDomainName"" or
           ""-ContactEmailAddress"" in the profile file to force the use of the
           ""GetCertClientSetup"" certificate again (instead of an existing domain
           certificate).

        .) A ""Host process can't be located on disk. Can't continue."" message means
           the host software (ie. the task scheduler) must also be installed to continue.
           Look in the ""{EXE}"" folder for the host EXE and run it.

    On the SCS server:

        .) The ""-StagingTestsEnabled"" switch must be set to ""True"" for the staging
           tests (it gets reset automatically every night).

        .) The ""GetCertClientSetup"" certificate must be added
           to the trusted people store.

        .) The ""GetCertClientSetup"" certificate must be removed 
           from the ""untrusted"" certificate store (if it's there).
"
                            .Replace("{EXE}", Path.GetFileName(ResourceAssembly.Location))
                            )
                            ;
                }

                try
                {
                    string lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(loProfile, new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService")
                                            , lsHash, lsProfile, out lsLogFileTextReported);
                }
                catch {}
            }

            return loProfile;
        }

        private static string sHostExePathFile(tvProfile aoProfile, out Process aoProcessFound)
        {
            return DoGetCert.sHostExePathFile(aoProfile, out aoProcessFound, false);
        }
        private static string sHostExePathFile(tvProfile aoProfile, out Process aoProcessFound, bool abQuietMode)
        {
            string  lsHostExeFilename = DoGetCert.sHostProcess;
            string  lsHostExePathFile = DoGetCert.sProcessExePathFile(lsHostExeFilename, out aoProcessFound);

            if ( null == lsHostExePathFile )
            {
                if ( !abQuietMode )
                    DoGetCert.LogIt(aoProfile, "Host image can't be located on disk based on the currently running process. Trying typical locations ...");

                lsHostExePathFile = @"C:\ProgramData\GoPcBackup\GoPcBackup.exe";
                if ( !File.Exists(lsHostExePathFile) )
                    lsHostExePathFile = @"C:\Program Files\GoPcBackup\GoPcBackup.exe";

                if ( !File.Exists(lsHostExePathFile) )
                {
                    DoGetCert.LogIt(aoProfile, "Host process can't be located on disk. Can't continue.");
                    throw new Exception("Exiting ...");
                }

                if ( !abQuietMode )
                    DoGetCert.LogIt(aoProfile, "Host process image found on disk. It will be restarted.");
            }

            return lsHostExePathFile;
        }

        private static string sStopHostExePathFile(tvProfile aoProfile)
        {
            Process loProcessFound = null;
            string  lsHostExePathFile = DoGetCert.sHostExePathFile(aoProfile, out loProcessFound, true);

            // Stop the EXE (can't update it or the INI while it's running).

            if ( !DoGetCert.bKillProcess(aoProfile, loProcessFound) )
            {
                DoGetCert.LogIt(aoProfile, "Host process can't be stopped. Update can't be applied.");
                throw new Exception("Exiting ...");
            }
            else
            {
                DoGetCert.LogIt(aoProfile, "Host process stopped successfully.");
            }

            return lsHostExePathFile;
        }

        private static void HandleClientExeUpdate(tvProfile aoProfile, string asHash, string asProfile, string asRunKey, tvProfile aoCmdLine)
        {
            string  lsCurrentVersion = FileVersionInfo.GetVersionInfo(aoProfile.sExePathFile).FileVersion;
            byte[]  lbtArrayExeUpdate = null;

            // Look for GetCert2 EXE update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                lbtArrayExeUpdate = loGetCertServiceClient.btArrayGetCertExeUpdate(asHash, asProfile, lsCurrentVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lbtArrayExeUpdate )
            {
                DoGetCert.LogIt(aoProfile, String.Format("Client version {0} is running. No update.", lsCurrentVersion));
            }
            else
            {
                DoGetCert.LogIt(aoProfile, "");
                DoGetCert.LogIt(aoProfile, String.Format("Client version {0} is running. Update found ...", lsCurrentVersion));

                string  lsUpdatePath = Path.Combine(Path.GetDirectoryName(aoProfile.sExePathFile), aoProfile.sValue("-UpdateFolder", "Update"));
                        if ( !Directory.Exists(lsUpdatePath) )
                            Directory.CreateDirectory(lsUpdatePath);
                string  lsNewExePathFile = Path.Combine(lsUpdatePath, Path.GetFileName(aoProfile.sExePathFile));
                string  lsNewIniPathFile = Path.Combine(lsUpdatePath, Path.GetFileName(aoProfile.sLoadedPathFile));

                DoGetCert.LogIt(aoProfile, String.Format("Writing update to \"{0}\".", lsNewExePathFile));
                File.WriteAllBytes(lsNewExePathFile, lbtArrayExeUpdate);

                DoGetCert.LogIt(aoProfile, String.Format("Writing update profile to \"{0}\".", lsNewIniPathFile));
                File.Copy(aoProfile.sLoadedPathFile, lsNewIniPathFile, true);

                DoGetCert.LogIt(aoProfile, String.Format("The file version of \"{0}\" is {1}."
                        , lsNewExePathFile, FileVersionInfo.GetVersionInfo(lsNewExePathFile).FileVersion));

                DoGetCert.LogIt(aoProfile, String.Format("Starting update \"{0}\" ...", lsNewExePathFile));
                Process loProcess = new Process();
                        loProcess.StartInfo.FileName = lsNewExePathFile;
                        loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2}"
                                , asRunKey, aoProfile.sExePathFile, aoCmdLine.sCommandLine());
                        loProcess.Start();

                        System.Windows.Forms.Application.DoEvents();

                aoProfile.bExit = true;
            }
        }

        private static tvProfile HandleClientCfgUpdate(tvProfile aoProfile, string asHash, string asProfile)
        {
            string      lsCfgVersion = aoProfile.sValue("-CfgVersion", "1");
            string      lsCfgUpdate = null;

            // Look for WCF configuration update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                lsCfgUpdate = loGetCertServiceClient.sCfgUpdate(asHash, asProfile, lsCfgVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsCfgUpdate )
            {
                DoGetCert.LogIt(aoProfile, String.Format("WCF configuration version {0} is in use. No update.", lsCfgVersion));
            }
            else
            {
                DoGetCert.LogIt(aoProfile, String.Format("WCF configuration version {0} is in use. Update found ...", lsCfgVersion));

                // Overwrite WCF config with the current update.
                File.WriteAllText(aoProfile.sLoadedPathFile, lsCfgUpdate
                        .Replace("{CertificateDomainName}", aoProfile.sValue("-CertificateDomainName", "")));

                // Overwrite WCF config version number in the current profile.
                tvProfile   loWcfCfg = new tvProfile(aoProfile.sLoadedPathFile, true);
                            aoProfile["-CfgVersion"] = loWcfCfg.sValue("-CfgVersion", "1");
                            aoProfile.Save();

                DoGetCert.LogIt(aoProfile, String.Format(
                        "WCF configuration update successfully completed. Version {0} is now in use.", aoProfile.sValue("-CfgVersion", "1")));
            }

            return aoProfile;
        }

        private static tvProfile HandleClientIniUpdate(tvProfile aoProfile, string asHash, string asProfile, string[] args)
        {
            string lsIniVersion = aoProfile.sValue("-IniVersion", "1");
            string lsIniUpdate = null;

            // Look for profile update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                lsIniUpdate = loGetCertServiceClient.sIniUpdate(asHash, asProfile, lsIniVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsIniUpdate )
            {
                DoGetCert.LogIt(aoProfile, String.Format("Profile version {0} is in use. No update.", lsIniVersion));
            }
            else
            {
                DoGetCert.LogIt(aoProfile, String.Format("Profile version {0} is in use. Update found ...", lsIniVersion));

                tvProfile loIniUpdateProfile = new tvProfile(lsIniUpdate);

                // Reload the profile without merging the command-line.
                aoProfile = new tvProfile(aoProfile.sLoadedPathFile, true);
                // Merge in the updated keys.
                aoProfile.LoadFromCommandLine(loIniUpdateProfile.ToString(), tvProfileLoadActions.Merge);
                // Allow for saving everything (including merged keys).
                aoProfile.bSaveSansCmdLine = false;
                aoProfile.bSaveEnabled = true;
                // Save the updated profile.
                aoProfile.Save();

                // Finally, reload the updated profile (and merge in the command-line).
                aoProfile = new tvProfile(args, true);

                DoGetCert.LogIt(aoProfile, String.Format(
                        "Profile update successfully completed. Version {0} is now in use.", aoProfile.sValue("-IniVersion", "1")));
            }

            return aoProfile;
        }

        private static void HandleHostUpdates(tvProfile aoProfile, string asHash, string asProfile)
        {
            bool        lbRestartHost = false;
            Process     loHostProcess = null;
            string      lsHostExePathFile= DoGetCert.sHostExePathFile(aoProfile, out loHostProcess);
            string      lsHostExeVersion = FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion;
            byte[]      lbtArrayGpcExeUpdate = null;
                        // Look for an updated GoPcBackup.exe.
                        using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
                        {
                            DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                            lbtArrayGpcExeUpdate = loGetCertServiceClient.btArrayGoPcBackupExeUpdate(asHash, asProfile, lsHostExeVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
            string      lsHostIniVersion = aoProfile.sValue("-HostIniVersion", "1");
            string      lsHostIniUpdate  = null;
                        // Look for an updated GoPcBackup.exe.txt.
                        using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
                        {
                            DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                            lsHostIniUpdate = loGetCertServiceClient.sGpcIniUpdate(asHash, asProfile, lsHostIniVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

            if ( null == lbtArrayGpcExeUpdate && null == lsHostIniUpdate && null == loHostProcess )
                lbRestartHost = true;

            if ( null != lbtArrayGpcExeUpdate || null != lsHostIniUpdate )
            {
                lbRestartHost = true;

                DoGetCert.sStopHostExePathFile(aoProfile);

                // Write the updated EXE (if any).
                if ( null != lbtArrayGpcExeUpdate )
                {
                    DoGetCert.LogIt(aoProfile, String.Format("Host version {0} is in use. Update found ...", lsHostExeVersion));

                    // Write the EXE.
                    File.WriteAllBytes(lsHostExePathFile, lbtArrayGpcExeUpdate);

                    DoGetCert.LogIt(aoProfile, String.Format("Host update successfully completed. Version {0} is now in use."
                                                            , FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion));
                }

                // Write the updated INI (if any).
                if ( null != lsHostIniUpdate )
                {
                    DoGetCert.LogIt(aoProfile, String.Format("Host profile version {0} is in use. Update found ...", lsHostIniVersion));

                    tvProfile   loUpdateProfile = new tvProfile(lsHostIniUpdate);
                    tvProfile   loHostProfile = new tvProfile(lsHostExePathFile + ".txt", false);

                                // Remove all "GetCert" related keys in tasks, backups and cleanups. Then add in any updates.
                                tvProfile   loTasks = new tvProfile();
                                            // Place any tasks in the update on top.
                                            foreach(DictionaryEntry loEntry in loUpdateProfile.oProfile("-AddTasks"))
                                            {
                                                loTasks.Add("-Task", loEntry.Value);
                                            }
                                            foreach(DictionaryEntry loEntry in loHostProfile.oProfile("-AddTasks"))
                                            {
                                                tvProfile   loTask = new tvProfile(loEntry.Value.ToString());
                                                            if ( !loTask.sValue("-CommandEXE", "").ToLower().Contains("getcert.exe") )
                                                                loTasks.Add("-Task", loEntry.Value);
                                            }
                                tvProfile   loBackupSets = new tvProfile();
                                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-BackupSet"))
                                            {
                                                bool        lbAddSet = true;
                                                tvProfile   loSet = (new tvProfile(loEntry.Value.ToString())).oOneKeyProfile("-FolderToBackup");
                                                            foreach(DictionaryEntry loPathFile in loSet)
                                                            {
                                                                if ( loPathFile.Value.ToString().ToLower().Contains(@"\getcert") )
                                                                {
                                                                    lbAddSet = false;
                                                                    break;
                                                                }
                                                            }

                                                if ( lbAddSet )
                                                    loBackupSets.Add("-BackupSet", loEntry.Value);
                                            }
                                            foreach(DictionaryEntry loEntry in loUpdateProfile.oOneKeyProfile("-BackupSet"))
                                            {
                                                loBackupSets.Add("-BackupSet", loEntry.Value);
                                            }
                                tvProfile   loCleanupSets = new tvProfile();
                                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-CleanupSet"))
                                            {
                                                bool        lbAddSet = true;
                                                tvProfile   loSet = (new tvProfile(loEntry.Value.ToString())).oOneKeyProfile("-FilesToDelete");
                                                            foreach(DictionaryEntry loPathFile in loSet)
                                                            {
                                                                if ( loPathFile.Value.ToString().ToLower().Contains(@"\getcert\") )
                                                                {
                                                                    lbAddSet = false;
                                                                    break;
                                                                }
                                                            }

                                                if ( lbAddSet )
                                                    loCleanupSets.Add("-CleanupSet", loEntry.Value);
                                            }
                                            foreach(DictionaryEntry loEntry in loUpdateProfile.oOneKeyProfile("-CleanupSet"))
                                            {
                                                loCleanupSets.Add("-CleanupSet", loEntry.Value);
                                            }

                                // The following steps place these 3 keys on top of
                                // the profile for easier access (via text editor).
                                tvProfile   loNewProfile = new tvProfile();
                                tvProfile   loAddTasksProfile = new tvProfile();
                                            foreach(DictionaryEntry loEntry in loTasks)
                                                loAddTasksProfile.Add(loEntry);

                                loNewProfile["-AddTasks"] = loAddTasksProfile.sCommandBlock();
                                foreach(DictionaryEntry loEntry in loBackupSets)
                                    loNewProfile.Add(loEntry);
                                foreach(DictionaryEntry loEntry in loCleanupSets)
                                    loNewProfile.Add(loEntry);
                            
                                loHostProfile.Remove("-AddTasks");
                                loHostProfile.Remove("-BackupSet");
                                loHostProfile.Remove("-CleanupSet");

                                foreach(DictionaryEntry loEntry in loHostProfile)
                                    loNewProfile.Add(loEntry);

                                loHostProfile.Remove("*");

                                foreach(DictionaryEntry loEntry in loNewProfile)
                                    loHostProfile.Add(loEntry);

                                loUpdateProfile.Remove("-AddTasks");
                                loUpdateProfile.Remove("-BackupSet");
                                loUpdateProfile.Remove("-CleanupSet");

                                // Finally, merge in the update (sans keys already handled).
                                loHostProfile.LoadFromCommandLine(loUpdateProfile.ToString(), tvProfileLoadActions.Merge);

                    // Update GetCert profile with current GPC IniVersion.
                    aoProfile["-HostIniVersion"] = loHostProfile.sValue("-GpcIniVersion", "1");
                    aoProfile.Save();

                    // Remove the version key from the host profile.
                    loHostProfile.Remove("-GpcIniVersion");

                    // Write the INI and unlock it.
                    loHostProfile.Save();
                    loHostProfile.bEnableFileLock = false;

                    DoGetCert.LogIt(aoProfile, String.Format(
                            "Host profile update successfully completed. Version {0} is now in use.", aoProfile.sValue("-HostIniVersion", "1")));
                }
            }

            if ( lbRestartHost )
            {
                // Restart the host EXE (skip startup tasks to avoid conflicts with what's currently running).
                loHostProcess = new Process();
                loHostProcess.StartInfo.FileName = Path.Combine(Path.GetDirectoryName(lsHostExePathFile), "Startup.cmd");
                loHostProcess.StartInfo.Arguments = "-StartupTasksDisabled";
                loHostProcess.StartInfo.UseShellExecute = false;
                loHostProcess.StartInfo.CreateNoWindow = true;
                loHostProcess.Start();

                DoGetCert.LogIt(aoProfile, "Host process restarted successfully.");
            }
        }

        private static void HandleHostsEntryUpdate(tvProfile aoProfile, string asHash, string asProfile)
        {
            string lsHstVersion = aoProfile.sValue("-HostsEntryVersion", "1");
            string lsHstUpdate = null;

            // Look for hosts entry update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetupCertificate(aoProfile, loGetCertServiceClient);

                lsHstUpdate = loGetCertServiceClient.sHostsEntryUpdate(asHash, asProfile, lsHstVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsHstUpdate )
            {
                DoGetCert.LogIt(aoProfile, String.Format("Hosts entry version {0} is in use. No update.", lsHstVersion));
            }
            else
            {
                DoGetCert.LogIt(aoProfile, String.Format("Hosts entry version {0} is in use. Update found ...", lsHstVersion));

                tvProfile       loHstUpdateProfile = new tvProfile(lsHstUpdate);
                string          lsIpAddress = loHstUpdateProfile.sValue("-IpAddress", "0.0.0.0");
                string          lsHostsPathFile = @"C:\windows\system32\drivers\etc\hosts";
                StringBuilder   lsbHostsStream = new StringBuilder(File.ReadAllText(lsHostsPathFile));
                string          lsDnsName = "secure_certificate_service_goes_here";
                Regex           loRegExRemoveComments = new Regex(String.Format(@"([#][0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)(.*)({0}.*\r\n)", lsDnsName), RegexOptions.IgnoreCase);
                                foreach(Match loMatch in loRegExRemoveComments.Matches(lsbHostsStream.ToString()))
                                {
                                    lsbHostsStream.Replace(loMatch.Value, "");
                                }
                Regex           loRegExSwapAddress = new Regex(String.Format(@"([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)(.*)({0})", lsDnsName), RegexOptions.IgnoreCase);

                try
                {
                    Match   loMatchIpAddress = loRegExSwapAddress.Matches(lsbHostsStream.ToString())[0];
                            lsbHostsStream.Replace(loMatchIpAddress.Groups[1].Value, lsIpAddress);
                }
                catch (Exception ex)
                {
                    if ( lsbHostsStream.ToString().ToLower().Contains(lsDnsName.ToLower()) )
                    {
                        DoGetCert.LogIt(aoProfile,"Something's wrong with the local 'hosts' file format. Can't update IP address.");
                        DoGetCert.LogIt(aoProfile, DoGetCert.sExceptionMessage(ex));
                    }
                    else
                    {
                        lsbHostsStream.Append(String.Format("\r\n{0}\t{1}", lsIpAddress, lsDnsName));
                    }
                }

                // Update "hosts" file.
                File.WriteAllText(lsHostsPathFile, lsbHostsStream.ToString());

                // Update GetCert profile with current GPC IniVersion.
                aoProfile["-HostsEntryVersion"] = loHstUpdateProfile.sValue("-HostsEntryVersion", "1");
                aoProfile.Save();

                DoGetCert.LogIt(aoProfile, String.Format(
                        "Hosts entry update successfully completed. Version {0} is now in use.", aoProfile.sValue("-HostsEntryVersion", "1")));

                // Restart host (should restart this app as well - after a brief delay).
                Process loHostProcess = new Process();
                        loHostProcess.StartInfo.FileName = Path.Combine(Path.GetDirectoryName(DoGetCert.sStopHostExePathFile(aoProfile)), "Startup.cmd");
                        loHostProcess.StartInfo.Arguments = "-StartupTasksDelaySecs=10";
                        loHostProcess.StartInfo.UseShellExecute = false;
                        loHostProcess.StartInfo.CreateNoWindow = true;
                        loHostProcess.Start();

                DoGetCert.LogIt(aoProfile, "Host process restarted successfully.");

                aoProfile.bExit = true;
            }
        }

        // Reset WCF configuration, from "https://stackoverflow.com/questions/6150644/change-default-app-config-at-runtime".
        private static void ResetConfigMechanism(tvProfile aoProfile)
        {
            AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", aoProfile.sLoadedPathFile);

            typeof(ConfigurationManager)
                .GetField("s_initState", BindingFlags.NonPublic | 
                                         BindingFlags.Static)
                .SetValue(null, 0);

            typeof(ConfigurationManager)
                .GetField("s_configSystem", BindingFlags.NonPublic | 
                                            BindingFlags.Static)
                .SetValue(null, null);

            typeof(ConfigurationManager)
                .Assembly.GetTypes()
                .Where(x => x.FullName == 
                            "System.Configuration.ClientConfigPaths")
                .First()
                .GetField("s_current", BindingFlags.NonPublic | 
                                       BindingFlags.Static)
                .SetValue(null, null);
        }

        private static void VerifyDependenciesInit(tvProfile aoProfile)
        {
            string  lsFetchName = null;

            // Fetch IIS Manager DLL.
            tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="Microsoft.Web.Administration.dll"), lsFetchName);
        }

        private static void VerifyDependencies(tvProfile aoProfile)
        {
            if ( aoProfile.bExit )
                return;

            bool    lbMinVer = true;
                    if ( Environment.OSVersion.Version.Major < 6 )
                        lbMinVer = false;
                    if ( lbMinVer && Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor < 1 )
                        lbMinVer = false;
            int     liErrors = 0;
            string  lsErrors = "";

            if ( !lbMinVer )
            {
                liErrors++;
                lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                        + "Windows Server 2008 R2 or later must be installed.";
            }

            try
            {
                int liCount = new ServerManager().Sites.Count;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                liErrors++;
                lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                        + "IIS must be installed.";
            }

            using (Microsoft.Win32.RegistryKey loLocalMachineStoreRegistryKey
                    = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\PowerShell\3", false))
            {
                try
                {
                    if ( 0 == (int)loLocalMachineStoreRegistryKey.GetValue("Install", 0) )
                        throw new Exception();
                }
                catch
                {
                    liErrors++;
                    lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                            + "PowerShell 3 or later must be installed.";
                }
            }

            if ( "" != lsErrors )
            {
                aoProfile.bExit = true;

                lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                        + String.Format("Please install {0} and try again.", liErrors == 1 ? "it" : "them");

                if ( !aoProfile.bValue("-NoPrompts", false) )
                    tvMessageBox.Show(null, lsErrors);

                throw new Exception(lsErrors);
            }
        }


        /// <summary>
        /// Get digital certificate, install it and bind it to port 443 in IIS.
        /// </summary>
        public bool bGetCertificate()
        {
            bool        lbGetCertificate = true;
            string      lsCertName = null;
            string      lsCertPathFile = null;
            string      lsHash = null;
            string      lsProfile = null;
            X509Store   loStore = null;

            this.LogIt("");
            this.LogIt("Get certificate process started ...");
            this.bMainLoopStopped = false;

            try
            {
                moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                DoGetCert.SetupCertificate(moProfile, moGetCertServiceFactory);
                string[]            lsSanArray = moProfile.oProfile("-SanList").sOneKeyArray("*");
                                    // The SAN list is the authoritative source for domain names, unless there's only one domain.
                                    if (  1 == lsSanArray.Length )
                                    {
                                        lsSanArray[0] = moProfile.sValue("-CertificateDomainName" ,"");
                                    }
                                    else
                                    if ( lsSanArray.Length > 1 )
                                    {
                                        moProfile["-CertificateDomainName"] = lsSanArray[0];

                                        // Names with trailing periods must have the primary DNS name appended.
                                        for (int i=0; i < lsSanArray.Length; i++)
                                            if ( lsSanArray[i].EndsWith(".") )
                                                lsSanArray[i] = lsSanArray[i] + lsSanArray[0];
                                    }

                lsCertName = moProfile.sValue("-CertificateDomainName" ,"");
                lsCertPathFile = moProfile.sRelativeToProfilePathFile(Guid.NewGuid().ToString());
                lsProfile = DoGetCert.sProfile(moProfile);
                lsHash = HashClass.sHashIt(new tvProfile(lsProfile));

                string              lsGuid = moProfile.sValue("-InstanceGuid", Guid.NewGuid().ToString());
                string              lsDefaultPhysicalPath = moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot");
                bool                lbCertificatePassword = false;  // Set this "false" to use password only for load balancer file.
                string              lsCertificatePassword = HashClass.sHashPw(lsProfile);
                ServerManager       loServerManager = null;
                X509Certificate2    loOldCertificate = DoGetCert.oCurrentCertificate(lsCertName, out loServerManager);
                                    if ( this.bCertificateNotExpiring(loOldCertificate) )
                                    {
                                        // The certificate is not ready to expire soon. There is nothing to do.
                                        return true;
                                    }
                X509Certificate2    loNewCertificate = null;
                byte[]              lbtArrayNewCertificate = null;
                bool                lbRepositoryCertsOnly = this.oDomainProfile.bValue("-RepositoryCertsOnly", false);
                                    if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                    {
                                        lbtArrayNewCertificate = loGetCertServiceClient.btArrayNewCertificate(lsHash, lsProfile);
                                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                            loGetCertServiceClient.Abort();
                                        else
                                            loGetCertServiceClient.Close();

                                        if ( null == lbtArrayNewCertificate )
                                        {
                                            this.LogIt("No new certificate was found in the repository.");
                                        }
                                        else
                                        {
                                            this.LogIt("New certificate downloaded from the repository.");

                                            // Buffer certificate to disk (same as what's done below).
                                            File.WriteAllBytes(lsCertPathFile, lbtArrayNewCertificate);
                                        }
                                    }

                if ( null == lbtArrayNewCertificate && lbRepositoryCertsOnly )
                {
                    int liErrorCount = 1 + moProfile.iValue("-RepositoryCertsOnlyErrorCount", 0);
                        moProfile["-RepositoryCertsOnlyErrorCount"] = liErrorCount;
                        moProfile.Save();

                    if ( moProfile.bValue("-RepositoryCertsOnlyReportAllErrors", false)
                            || liErrorCount > moProfile.iValue("-RepositoryCertsOnlyMaxErrors", 2) )
                    {
                        lbGetCertificate = false;

                        this.LogIt("This is an error since -RepositoryCertsOnly=True.");
                    }
                    else
                    {
                        this.LogIt("-RepositoryCertsOnly=True. Will try again next cycle.");

                        return true;
                    }
                }
                if ( null == lbtArrayNewCertificate && !lbRepositoryCertsOnly )
                {
                    // Certificate renewal locks are only needed during automation mode.
                    if ( moProfile.bValue("-Auto", false) )
                    {
                                // Wait a random period each cycle to allow different clients the opportunity to lock the renewal.
                        int     liMaxCertRenewalLockDelaySecs = moProfile.iValue("-MaxCertRenewalLockDelaySecs", 300);
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    this.LogIt(String.Format("Waiting at most {0} seconds to set a certificate renewal lock for this domain ...", liMaxCertRenewalLockDelaySecs));
                                    System.Windows.Forms.Application.DoEvents();
                                    Thread.Sleep(1000 * new Random().Next(liMaxCertRenewalLockDelaySecs));
                                }
                        bool    lbLockCertificateRenewal = false;
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                {
                                    lbLockCertificateRenewal = loGetCertServiceClient.bLockCertificateRenewal(lsHash, lsProfile);
                                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                        loGetCertServiceClient.Abort();
                                    else
                                        loGetCertServiceClient.Close();
                                }
                                // If the renewal can't be locked, another client must be doing the
                                // certificate renewal already (or not all clients renewed previously or
                                // certificate overrides are reserved to another computer on the domain).
                                // Try again next time.
                                if ( lbLockCertificateRenewal )
                                {
                                    this.LogIt("Certificate renewal has been locked for this domain.");
                                }
                                else
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    this.LogIt("Certificate renewal can't be locked. Will try again next cycle.");

                                    return true;
                                }
                    }

                    this.LogIt("");
                    if ( "" != this.oDomainProfile.sValue("-CertOverridePathFile", "") )
                    {
                        this.LogIt(String.Format("Retrieving new certificate for \"{0}\" from the domain (via its own certificate override file) ...", lsCertName));

                        lsCertPathFile = this.oDomainProfile.sValue("-CertOverridePathFile", "");
                        lsCertificatePassword = this.oDomainProfile.sValue("-CertOverridePassword", "");
                    }
                    else
                    {
                        this.LogIt("");
                        this.LogIt(String.Format("Retrieving new certificate for \"{0}\" from the certificate provider network ...", lsCertName));

                        string  lsAcmePsFile = "ACME-PS.zip";
                        string  lsAcmePsPath = Path.GetFileNameWithoutExtension(lsAcmePsFile);
                        string  lsAcmePsPathFile = moProfile.sRelativeToProfilePathFile(lsAcmePsFile);

                        // Fetch AcmePs module.
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsAcmePsFile), lsAcmePsFile);
                        if ( !Directory.Exists(moProfile.sRelativeToProfilePathFile(lsAcmePsPath)) )
                            ZipFile.ExtractToDirectory(lsAcmePsPathFile, Path.GetDirectoryName(lsAcmePsPathFile));

                        // Don't use the powershell gallery installed AcmePs module (by default, use what's embedded instead).
                        if ( !moProfile.bValue("-AcmePsModuleUseGallery", false) )
                        {
                            string lsAcmePsPathFolder = moProfile.sValue("-AcmePsPath", lsAcmePsPath);

                            moProfile.sValue("-AcmePsPathHelp", String.Format(
                                    "Any alternative -AcmePsPath must include a subfolder named \"{0}\" that contains the files.", lsAcmePsPathFolder));

                            lsAcmePsPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsPath", lsAcmePsPath));
                        }

                        string lsAcmeWorkPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsWorkPath", "AcmeState"));

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            this.LogStage("1 - init ACME workspace");

                            // Delete ACME workspace.
                            if ( Directory.Exists(lsAcmeWorkPath) )
                                Directory.Delete(lsAcmeWorkPath, true);

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptSessionOpen", @"
New-PSSession -ComputerName localhost -Name GetCert
$session = Disconnect-PSSession -Name GetCert
                                    "), true);

                            this.LogIt("");

                            if ( lbGetCertificate )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage1", @"
Using Module ""{AcmePsPath}""

$global:state = New-ACMEState -Path ""{AcmeWorkPath}""
Get-ACMEServiceDirectory $global:state -ServiceName ""{AcmeServiceName}"" -PassThru
                                    ")
                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                    .Replace("{AcmeServiceName}", moProfile.bValue("-DoStagingTests", true) ? moProfile.sValue("-ServiceNameStaging", "LetsEncrypt-Staging")
                                                                                                            : moProfile.sValue("-ServiceNameLive", "LetsEncrypt"))
                                        );
                        }

                        string      lsSanPsStringArray = "(";
                                    foreach(string lsSanItem in lsSanArray)
                                    {
                                        lsSanPsStringArray += String.Format(",\"{0}\"",  lsSanItem);
                                    }
                                    lsSanPsStringArray = lsSanPsStringArray.Replace("(,", "(") + ")";

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            this.LogStage("2 - register domain contact, submit order & authorization request");

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2", @"
Using Module ""{AcmePsPath}""

New-ACMENonce      $global:state
New-ACMEAccountKey $global:state -PassThru
New-ACMEAccount    $global:state -EmailAddresses ""{ContactEmailAddress}"" -AcceptTOS

$SanList = {SanPsStringArray}
[AcmeIdentifier[]] $identifiers = $null
        foreach ($SAN in $SanList) { $identifiers += New-ACMEIdentifier $SAN }

$global:order = New-ACMEOrder $global:state -Identifiers $identifiers
$global:authZ = Get-ACMEAuthorization -State $global:state -Order $global:order

[int[]] $global:SanMap = $null
        foreach ($SAN in $SanList) { for ($i=0; $i -lt $global:authZ.Length; $i++) { if ( $global:authZ[$i].Identifier.value -eq $SAN ) { $global:SanMap += $i }}}

                                    ")
                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                    .Replace("{ContactEmailAddress}", moProfile.sValue("-ContactEmailAddress" ,""))
                                    .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                    );
                        }

                        ArrayList   loAcmeCleanedList = new ArrayList();

                        for (int liSanArrayIndex=0; liSanArrayIndex < lsSanArray.Length; liSanArrayIndex++)
                        {
                            string lsSanItem = lsSanArray[liSanArrayIndex];

                            if ( lbGetCertificate && !mbMainLoopStopped )
                            {
                                this.LogStage(String.Format("3 - Define DNS name to be challenged (\"{0}\"), setup domain challenge in IIS and submit it to certificate provider", lsSanItem));

                                Site    loSanSite = null;
                                Site    loDefaultSite = null;
                                        DoGetCert.oCurrentCertificate(lsSanItem, out loServerManager, out loSanSite, lsSanArray, out loDefaultSite);

                                // If -CreateSanSites is false, only create a default
                                // website (as needed) and use it for all SAN values.
                                if ( null == loSanSite && !moProfile.bValue("-CreateSanSites", true)
                                        && 0 != loServerManager.Sites.Count )
                                {
                                    this.LogIt(String.Format("No website found for \"{0}\".\r\n-CreateSanSites is \"False\". So no website created.\r\n", lsSanItem));

                                    loSanSite = loServerManager.Sites[0];
                                }

                                if ( null == loSanSite )
                                {
                                    string lsPhysicalPath = null;

                                    if ( 0 == loServerManager.Sites.Count )
                                    {
                                        this.LogIt("No default website could be found.");

                                        lsPhysicalPath = lsDefaultPhysicalPath;
                                    }
                                    else
                                    {
                                        if ( null != loDefaultSite )
                                            lsPhysicalPath = loDefaultSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;
                                        else
                                            lsPhysicalPath = loServerManager.Sites[0].Applications["/"].VirtualDirectories["/"].PhysicalPath;
                                    }

                                    // Leave hostname blank in the default website to allow for old browsers (ie. no SNI support).
                                    loSanSite = loServerManager.Sites.Add(
                                              lsSanItem
                                            , "http"
                                            , String.Format("*:80:{0}", 0 == loServerManager.Sites.Count ? "" : lsSanItem)
                                            , lsPhysicalPath
                                            );

                                    loServerManager.CommitChanges();

                                    this.LogIt(String.Format("No website found for \"{0}\". New website created.", lsSanItem));
                                    this.LogIt("");
                                }

                                string  lsAcmeBasePath = moProfile.sValue("-AcmeBasePath", @".well-known\acme-challenge");
                                string  lsAcmePath = Path.Combine(
                                          Environment.ExpandEnvironmentVariables(loSanSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                                        , lsAcmeBasePath);

                                // Cleanup ACME folder.
                                if ( Directory.Exists(lsAcmePath) && !loAcmeCleanedList.Contains(lsAcmePath) )
                                {
                                    Directory.Delete(lsAcmePath, true);
                                    Thread.Sleep(moProfile.iValue("-ACMEcleanupSleepMS", 200));

                                    // Don't cleanup the same folder more than once per session.
                                    loAcmeCleanedList.Add(lsAcmePath);
                                }

                                // Add IIS "web.config" prior to challenge submission.
                                if ( lbGetCertificate && !mbMainLoopStopped )
                                    lbGetCertificate = this.bFixTextFile(Path.Combine(lsAcmePath, moProfile.sValue("-AcmeWebConfigFilename", "web.config"))
                                            , moProfile.sValue("-WebConfigContent", @"
<?xml version=""1.0"" encoding=""UTF-8""?>
<configuration>
    <system.webServer>
        <staticContent>
            <mimeMap fileExtension=""."" mimeType=""text/plain"" />
        </staticContent>
    </system.webServer>
</configuration>
                                            "));

                                if ( lbGetCertificate && !mbMainLoopStopped )
                                    lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3", @"
Using Module ""{AcmePsPath}""

$challenge = Get-ACMEChallenge $global:state $global:authZ[$global:SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge $global:state
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                        .Replace("{AcmePath}", lsAcmePath)
                                        );
                            }
                        }

                        int liSubmissionRetries = moProfile.iValue("-SubmissionRetries", 999);

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(1000 * moProfile.iValue("-SubmissionWaitSecs", 5));

                                this.LogStage(String.Format("4 - update challenge{0} from certificate provider", 1 == lsSanArray.Length ? "" : "s"));

                                string  lsSubmissionPending = ": pending";
                                string  lsOutput = null;

                                lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage4", @"
Using Module ""{AcmePsPath}""

$global:order | Update-ACMEOrder $global:state -PassThru
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        );

                                if ( !lsOutput.Contains(lsSubmissionPending) || mbMainLoopStopped )
                                    break;
                            }
                        }

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            this.LogStage("5 - generate certificate request and submit");

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage5", @"
Using Module ""{AcmePsPath}""

$global:certKey = New-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Complete-ACMEOrder $global:state -Order $global:order -CertificateKey $global:certKey
                                    ")
                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                    );
                        }

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(1000 * moProfile.iValue("-SubmissionWaitSecs", 5));

                                this.LogStage("6 - update certificate request");

                                string  lsSubmissionPending = "CertificateRequest       : \r\nCrtPemFile               :";
                                string  lsOutput = null;

                                lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage6", @"
Using Module ""{AcmePsPath}""

$global:order | Update-ACMEOrder $global:state -PassThru;
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        );

                                if ( !lsOutput.Contains(lsSubmissionPending) || mbMainLoopStopped )
                                    break;
                            }
                        }

                        if ( lbGetCertificate && !mbMainLoopStopped )
                        {
                            this.LogStage("7 - get certificate");

                            File.Delete(lsCertPathFile);

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage7", @"
Using Module ""{AcmePsPath}""

Export-ACMECertificate $global:state -Order $global:order -CertificateKey $global:certKey -Path ""{CertificatePathFile}""
                                    ")
                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                    .Replace("{CertificatePassword}", lsCertificatePassword)
                                    .Replace("{CertificatePathFile}", lsCertPathFile)
                                    );
                        }
                    }
                }

                if ( lbGetCertificate && !mbMainLoopStopped )
                {
                    if ( !moProfile.bValue("-CertificatePrivateKeyExportable", false) || this.oDomainProfile.bValue("-CertPrvKeyExpOverride", false) )
                    {
                        if ( lbCertificatePassword )
                            loNewCertificate = new X509Certificate2(lsCertPathFile, lsCertificatePassword
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);
                        else
                            loNewCertificate = new X509Certificate2(lsCertPathFile, (string)null
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);
                    }
                    else 
                    {
                        this.LogIt("");
                        this.LogIt("The new certificate private key is exportable (\"-CertificatePrivateKeyExportable=True\"). This is not recommended.");

                        if ( lbCertificatePassword )
                            loNewCertificate = new X509Certificate2(lsCertPathFile, lsCertificatePassword
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
                        else
                            loNewCertificate = new X509Certificate2(lsCertPathFile, (string)null
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
                    }

                    // Upload the new certificate to the repository (assuming it wasn't just downloaded from there).
                    if ( null == lbtArrayNewCertificate )
                    {
                        // But first, copy it to the load balancer file (if one is defined).
                        string  lsLoadBalancerPathFile = this.oDomainProfile.sValue("-LoadBalancerPathFile", "");
                                if ( "" != lsLoadBalancerPathFile )
                                {
                                    this.LogIt("");
                                    this.LogIt(String.Format("Copying the new certificate to \"{0}\" for use on the \"{1}\" load balancer ..."
                                                    , lsLoadBalancerPathFile, lsCertName));

                                    X509Certificate2 loLbCertificate = new X509Certificate2(lsCertPathFile, lsCertificatePassword
                                            , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);

                                    File.WriteAllBytes(lsLoadBalancerPathFile
                                            , loLbCertificate.Export(X509ContentType.Pfx, this.oDomainProfile.sValue("-LoadBalancerPassword", "")));

                                    if ( !moProfile.bValue("-UseStandAloneMode", true) && lbGetCertificate && !mbMainLoopStopped )
                                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                    {
                                        loGetCertServiceClient.NotifyLoadBalancerCertificate(lsHash, lsProfile, DoGetCert.sCurrentComputerName);
                                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                            loGetCertServiceClient.Abort();
                                        else
                                            loGetCertServiceClient.Close();
                                    }

                                    this.LogSuccess();
                                }

                        if ( !moProfile.bValue("-UseStandAloneMode", true) && lbGetCertificate && !mbMainLoopStopped )
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                        {
                            lbGetCertificate = loGetCertServiceClient.bNewCertificateUploaded(
                                    lsHash, lsProfile, File.ReadAllBytes(lsCertPathFile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            if ( !lbGetCertificate )
                                this.LogIt(DoGetCert.sExceptionMessage("GetCertService.bNewCertificateUploaded: Failed uploading new certificate."));
                        }
                    }
                }

                if ( lbGetCertificate && !mbMainLoopStopped )
                {
                    this.LogIt("");
                    this.LogIt(String.Format("Certificate thumbprint: {0} (\"{1}\")", loNewCertificate.Thumbprint, DoGetCert.sCertName(loNewCertificate)));

                    this.LogIt("");
                    this.LogIt("Install and bind certificate in IIS ...");

                    if ( lbGetCertificate && !mbMainLoopStopped )
                    {
                        string lsMachineKeyPathFile = null;

                        // Select the local machine certificate store (ie "Local Computer / Personal / Certificates").
                        loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                        loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                        // Add the new cert to the cert store.
                        loStore.Add(loNewCertificate);

                        // Do ADFS stuff (while both old and new certs are in the store).
                        if ( null != loOldCertificate && loOldCertificate.Thumbprint != loNewCertificate.Thumbprint )
                            lbGetCertificate = this.bApplyNewCertToADFS(loNewCertificate, loOldCertificate, lsHash, lsProfile);

                        if ( !lbGetCertificate )
                        {
                            // Remove the new cert from the store.
                            loStore.Remove(loNewCertificate);

                            lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, loOldCertificate);
                            if ( null != lsMachineKeyPathFile )
                            {
                                // Remove cert's private key file (sadly, the OS typically let's these things accumulate forever).
                                File.Delete(lsMachineKeyPathFile);
                            }
                        }
                        else
                        {
                            if ( null != loOldCertificate && loOldCertificate.Thumbprint == loNewCertificate.Thumbprint )
                            {
                                // It is necessary to do this here only
                                // when old and new thumprints are the same.

                                // Get old cert's private key file.
                                lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, loOldCertificate);
                                if ( null != lsMachineKeyPathFile )
                                {
                                    // Remove old cert's private key file.
                                    File.Delete(lsMachineKeyPathFile);
                                }
                            }

                            if ( null != loOldCertificate && lbGetCertificate && !mbMainLoopStopped )
                            {
                                if ( !moProfile.bValue("-UseStandAloneMode", true)
                                        || (moProfile.bValue("-UseStandAloneMode", true)
                                         && moProfile.bValue("-RemoveReplacedCert", false))
                                        )
                                {
                                    if ( loNewCertificate.Thumbprint == loOldCertificate.Thumbprint )
                                    {
                                        this.LogIt(String.Format("New certificate (\"{0}\") same as old. Not removed from local store.", loOldCertificate.Thumbprint));
                                    }
                                    else
                                    {
                                        // Remove the old cert.
                                        loStore.Remove(loOldCertificate);

                                        // Get old cert's private key file.
                                        lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(moProfile, loOldCertificate);
                                        if ( null != lsMachineKeyPathFile )
                                        {
                                            // Remove old cert's private key file.
                                            File.Delete(lsMachineKeyPathFile);
                                        }

                                        this.LogIt(String.Format("Old certificate (\"{0}\") removed from the local store.", loOldCertificate.Thumbprint));
                                    }
                                }
                            }

                            // Setup new site-to-cert bindings and remove any old ones.
                            foreach(string lsSanItem in lsSanArray)
                            {
                                if ( !lbGetCertificate || mbMainLoopStopped )
                                    break;

                                Site    loOldSite = null;
                                Binding loOldBinding = null;
                                Site    loDefaultSite = null;

                                DoGetCert.oCurrentCertificate(lsSanItem, out loServerManager, out loOldSite, out loOldBinding, lsSanArray, out loDefaultSite);

                                if ( null == loOldSite && lsSanItem == lsCertName )
                                {
                                    // Use the default site (if it exists) for the primary domain.
                                    if ( 0 != loServerManager.Sites.Count )
                                        loOldSite = loServerManager.Sites[0];   // Default site.

                                    // Use the default SSL binding (if it exists).
                                    loOldBinding = DoGetCert.oSslBinding(loOldSite);
                                }

                                // Create a new site with the default port 80 binding (as needed).
                                if ( null == loOldSite )
                                {
                                    if ( !moProfile.bValue("-CreateSanSites", true) )
                                    {
                                        this.LogIt(String.Format("No website could be found for \"{0}\".\r\n-CreateSanSites is \"False\". So no website created.\r\n", lsSanItem));
                                    }
                                    else
                                    {
                                        string lsPhysicalPath = null;

                                        if ( 0 == loServerManager.Sites.Count )
                                        {
                                            lsPhysicalPath = lsDefaultPhysicalPath;
                                        }
                                        else
                                        {
                                            if ( null != loDefaultSite )
                                                lsPhysicalPath = loDefaultSite.Applications["/"].VirtualDirectories["/"].PhysicalPath;
                                            else
                                                lsPhysicalPath = loServerManager.Sites[0].Applications["/"].VirtualDirectories["/"].PhysicalPath;
                                        }

                                        loOldSite = loServerManager.Sites.Add(
                                                  lsSanItem
                                                , "http"
                                                , String.Format("*:80:{0}", lsSanItem)
                                                , lsPhysicalPath
                                                );

                                        this.LogIt(String.Format("No website found for \"{0}\". New website created.", lsSanItem));
                                    }
                                }
                                else
                                if  ( null != loOldBinding )
                                {
                                    // Remove old IIS binding.
                                    loOldSite.Bindings.Remove(loOldBinding);

                                    this.LogIt(String.Format("Old certificate port binding (\"{0}\") removed from IIS for \"{1}\".", loOldBinding.Protocol, lsSanItem));
                                }

                                string      lsBindingInformation = null != loOldBinding ? loOldBinding.BindingInformation : String.Format("*:443:{0}", lsSanItem);
                                string[]    lsBindingInfArray = lsBindingInformation.Split(':');
                                string      lsBindingPort = lsBindingInfArray[1];
                                string      lsBindingHost = lsBindingInfArray[2].ToLower();

                                try
                                {
                                    // If there's no binding host defined or this SAN is the primary domain (and it has no website of its
                                    // own), use the default binding (ditto if the default binding host on the default site is undefined).
                                    bool    lbUseDefaultBinding = ("" == lsBindingHost
                                                    || (   lsSanItem == lsCertName
                                                        && loOldSite.Name != lsSanItem
                                                        && 0 != loServerManager.Sites[0].Bindings.Count && "" == loServerManager.Sites[0].Bindings[0].Host
                                                        ));

                                    if ( lbUseDefaultBinding )
                                    {
                                        lsBindingHost = "";
                                        lsBindingInformation = String.Format("*:{0}:{1}", lsBindingPort, lsBindingHost);
                                        loOldSite.Bindings.Add(lsBindingInformation, loNewCertificate.GetCertHash(), loStore.Name);

                                        this.LogIt(String.Format("Default binding for new certificate (\"{0}\") applied.", loNewCertificate.Thumbprint));
                                    }
                                    else
                                    {
                                        Binding loNewBinding = loOldSite.Bindings.Add(lsBindingInformation, loNewCertificate.GetCertHash(), loStore.Name);

                                        try
                                        {
                                            // Attempt to set the SNI (Server Name Indication) flag.
                                            loNewBinding.SetAttributeValue("SslFlags", 1 /* SslFlags.Sni */);

                                            this.LogIt(String.Format("SNI binding applied for new certificate (\"{0}\").", loNewCertificate.Thumbprint));
                                        }
                                        catch (Exception)
                                        {
                                            // Reverting to the default binding usage (must be an older OS version).
                                            this.LogIt(String.Format("Default binding applied (on older OS) for new certificate (\"{0}\").", loNewCertificate.Thumbprint));
                                        }
                                    }

                                    loServerManager.CommitChanges();

                                    this.LogIt(String.Format("New certificate (\"{0}\") bound to port {1} in IIS for \"{2}\"."
                                                    , loNewCertificate.Thumbprint, lsBindingPort, lsSanItem));
                                }
                                catch (Exception ex)
                                {
                                    // Perhaps the certificate was just deleted from the store (leaving the binding behind).
                                    if ( ex.Message.Contains("Cannot add duplicate collection entry of type 'binding'") )
                                    {
                                        this.LogIt(String.Format("Duplicate IIS port {0} binding entry found and removed for \"{1}\".", lsBindingPort, lsSanItem));
                                    }
                                    else
                                    {
                                        this.LogIt(DoGetCert.sExceptionMessage(ex));
                                        lbGetCertificate = false;
                                    }
                                }
                            }

                            if ( lbGetCertificate )
                            {
                                this.LogSuccess();
                                this.LogIt("");
                                this.LogIt("A new certificate was successfully installed and bound in IIS.");
                            }

                            if ( lbGetCertificate && !mbMainLoopStopped && !this.bCertificateSetupDone )
                            {
                                // Replace the null certificate reference in the WCF configuration (used for initial setup only)
                                // with a reference to the newly installed certificate.
                                string  lsNewCfgContent = File.ReadAllText(moProfile.sLoadedPathFile);
                                        lsNewCfgContent = lsNewCfgContent.Replace(
                                                    "</serviceCertificate>\r\n          </clientCredentials>"
                                                ,String.Format("</serviceCertificate>\r\n            <clientCertificate findValue=\"{0}\" x509FindType=\"FindBySubjectName\" storeLocation=\"LocalMachine\" storeName=\"My\" />\r\n          </clientCredentials>"
                                                        , lsCertName)
                                                );

                                // Write the updated WCF configuration.
                                File.WriteAllText(moProfile.sLoadedPathFile, lsNewCfgContent);
                                DoGetCert.ResetConfigMechanism(moProfile);

                                this.bCertificateSetupDone = true;
                            }

                            if ( lbGetCertificate && !mbMainLoopStopped && !moProfile.bValue("-UseStandAloneMode", true) )
                            {
                                // At this point we need to load the new certificate into the service factory object (before removing the old one).
                                moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");

                                this.LogIt("Requesting removal of the old certificate from the repository.");

                                // Last step, remove old certificate from the repository (it can't be removed until the new certificate is 100% in place).
                                if ( lbGetCertificate && !mbMainLoopStopped )
                                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                {
                                    bool    bOldCertificateRemoved = loGetCertServiceClient.bOldCertificateRemoved(lsHash, lsProfile);
                                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                                loGetCertServiceClient.Abort();
                                            else
                                                loGetCertServiceClient.Close();

                                    if ( !bOldCertificateRemoved )
                                        this.LogIt(DoGetCert.sExceptionMessage("GetCertService.bOldCertificateRemoved: Failed removing old certificate."));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogIt(DoGetCert.sExceptionMessage(ex));
                lbGetCertificate = false;
            }
            finally
            {
                if ( null != loStore )
                    loStore.Close();

                if ( "" == this.oDomainProfile.sValue("-CertOverridePathFile", "") )
                {
                    // Remove the cert (ie. PFX file from cert provider).
                    File.Delete(lsCertPathFile);
                }
                else
                if ( lbGetCertificate && !mbMainLoopStopped )
                {
                    // Remove the cert (ie. PFX file from domain - only upon success).
                    File.Delete(lsCertPathFile);
                }

                this.LogIt("");

                this.bRunPowerScript(moProfile.sValue("-ScriptSessionClose", @"
$session = Connect-PSSession -ComputerName localhost -Name GetCert
$session | Remove-PSSession
Get-WSManInstance -ResourceURI Shell -Enumerate
                        "), true);

                if ( !moProfile.bValue("-UseStandAloneMode", true) && moProfile.bValue("-Auto", false) )
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.bUnlockCertificateRenewal(lsHash, lsProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }

            if ( mbMainLoopStopped )
            {
                this.LogIt("Stopped.");
                lbGetCertificate = false;
            }

            if ( lbGetCertificate )
            {
                moProfile.Remove("-RepositoryCertsOnlyErrorCount");
                moProfile.Remove("-CertificateRenewalDateOverride");
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = true;
                moProfile.Save();

                this.LogIt("");
                this.LogIt("The get certificate process completed successfully.");
            }
            else
            {
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = false;
                moProfile.Save();

                this.LogIt("");
                this.LogIt("At least one stage failed (or the process was stopped). Check log for errors.");
            }

            return lbGetCertificate;
        }
    }
}
