using Microsoft.Web.Administration;
using System;
using System.Collections;
using System.Collections.Generic;
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
        private ChannelFactory<GetCertService.IGetCertServiceChannel> moGetCertServiceFactory = null;

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
                    DoGetCert.bKillProcessParent(Process.GetCurrentProcess());

                loProfile = tvProfile.oGlobal(DoGetCert.oHandleUpdates(ref args, ref loStartupWaitMsg));

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
of the current digital certificate (eg. 30 days, see -RenewalDaysBeforeExpiration
below), this software does nothing. Otherwise, the retrieval process begins.

It's as simple as that when the software runs in ""stand-alone"" mode (the default).

If ""stand-alone"" mode is disabled (see -UseStandAloneMode below), the certificate
retrieval process is used in concert with the secure certificate service (SCS),
see ""SafeTrust.org"".

If the software is not running in ""stand-alone"" mode, it also copies any new cert
to a file anywhere on the local area network to be picked up by the load balancer
administrator or process. It also replaces the SSO (single sign-on) certificate in
your central SSO configuration (eg. ADFS) and restarts the SSO service on all
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
overridden by providing this software access to a secure digital certificate
file anywhere on the local network.


Command-Line Usage


Open this utility's profile file to see additional options available. It is
usually located in the same folder as ""{EXE}"" and has the same name
with "".config"" added (see ""{INI}"").

Profile file options can be overridden with command-line arguments. The
keys for any ""-key=value"" pairs passed on the command-line must match
those that appear in the profile (with the exception of the ""-ini"" key).

For example, the following invokes the use of an alternative profile file
(be sure to copy an existing profile file if you do this):

    {EXE} -ini=NewProfile.txt

This tells the software to run in automatic mode:

    {EXE} -Auto


Author:  George Schiro (GeoCode@SafeTrust.org)

Date:    10/24/2019




Options and Features


The main options for this utility are listed below with their default values.
A brief description of each feature follows.

-AcmePsModuleUseGallery=False

    Set this switch True and the ""PowerShell Gallery"" version of ""ACME-PS""
    will be used in lieu of the version embedded in the EXE (see -AcmePsPath
    below).

-AcmePsPath=""ACME-PS""

    ""ACME-PS"" is the tool used by this utility to communicate with the
    ""Let's Encrypt"" certificate network. By default, this key is set to the ""ACME-PS""
    folder which, with no absolute path given, will be expected to be found within
    the folder containing ""{INI}"". Set -AcmePsModuleUseGallery=True
    (see above) and the OS will look to find ""ACME-PS"" in its usual place as a
    module from the PowerShell gallery.

-Auto=False

    Set this switch True to run this utility one time (with no interactive UI)
    then shutdown automatically upon completion. This switch is useful if the
    software is run in a batch process or by a server job scheduler.

-CertificateDomainName= NO DEFAULT VALUE

    This is the subject name (ie. DNS name) of the certificate returned.

-CertificatePrivateKeyExportable=False

    Set this switch True (not recommended) to allow certificate private keys
    to be exportable from the local certificate store to a local disk file. Any
    SA with server access will then have access to the certificate's private key.

-CertificateRenewalDateOverride= NO DEFAULT VALUE

    Set this date value to override the date calculation that subtracts
    -RenewalDaysBeforeExpiration days (see below) from the current certificate
    expiration date to know when to start fetching a new certificate.

    Note: this parameter will be removed from the profile after a certificate
          has been successfully retrieved from the certificate provider network.

-ContactEmailAddress= NO DEFAULT VALUE

    This is the contact email address the certificate network uses to send
    certificate expiration notices.

-CreateSanSites=True

    If a SAN specific website does not yet exist in IIS, it will be created
    automatically during the first run of the ""get certificate"" process for
    that SAN value. Set this switch False to have all SAN challenges routed
    through the IIS default website (such challenges will typically fail). If
    they do fail, you will need to create your SAN specific sites manually.

    Note: each SAN value must be challenged and therefore SAN challenges
          must be routed through your web server just like your primary domain.
          That means every SAN value must have a corresponding entry in the
          global internet DNS database.

          When a new website is created in IIS for a new SAN value, by default
          it is setup to use the same physical path as the primary domain.

-DoStagingTests=True

    Initial testing is done with the certificate provider staging network.
    Set this switch False to use the live production certificate network.

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

-LoadBalancerReleaseCert

    This switch indicates the new certificate has been released by the load
    balancer administrator or process.

    This switch would typically never appear in a profile file. It's meant to be
    used on the command-line only (in a script or a shortcut).

    Note: this switch is ignored when -UseStandAloneMode is True.

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

-NoIISBindingScript=""""

    This is the PowerShell script that binds a new certificate when IIS is not in
    use or the standard IIS binding procedure does not work for whatever reason.

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. Messages
    are written to the log instead (see -LogPathFile above). You must use this
    switch whenever the software is run via a server computer batch job or job
    scheduler (ie. where no user interaction is permitted).

-PowerScriptPathFile=PowerScript.ps1

    This is the path\file location of the current (ie. temporary) PowerShell script
    file.

-PowerScriptSleepMS=200

    This is the number of sleep milliseconds between loops while waiting for the
    PowerShell script process to complete.

-PowerScriptTimeoutSecs=300

    This is the maximum number of seconds allocated to any PowerShell script
    process to run prior to throwing a timeout exception.

-PowerShellExeArgs=-NoProfile -ExecutionPolicy unrestricted -File ""{{0}}"" ""{{1}}""

    These are the arguments passed to the Windows PowerShell EXE (see below).

-PowerShellExePathFile=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

    This is the path\file location of the Windows PowerShell EXE.

-RegexDnsNamePrimary=""^([a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,18}$""

    This regular expression is used to validate -CertificateDomainName (see above).

-RegexDnsNameSanList=""^([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)([a-zA-Z]{2,18}|)$""

    This regular expression is used to validate -SanList names (see below).

    Note: -RegexDnsNameSanList and -RegexDnsNamePrimary are used to validate
          SAN list names. A match of either pattern will pass validation.

-RegexEmailAddress=""^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)[a-zA-Z]{2,18}$""

    This regular expression is used to validate -ContactEmailAddress (see above).

-RemoveReplacedCert=False

    Set this switch True and the old (ie. previously bound) certificate will be
    removed whenever a new retrieved certificate is bound to replace it.

    Note: this switch is ignored when -UseStandAloneMode is False.

-RenewalDaysBeforeExpiration=30

    This is the number of days until certificate expiration before automated
    gets of the next new certificate are attempted.

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

    Here's a command-line example:

    -SanList=""-Domain='MyDomain.com' -Domain='www.MyDomain.com'""

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

-ScriptSSO= SEE PROFILE FOR DEFAULT VALUE

    This is the PowerShell script that updates SSO servers with new certificates.

-ScriptStage1= SEE PROFILE FOR DEFAULT VALUE

    There are multiple stages involved with the process of getting a certificate
    from the certificate provider network. Each stage has an associated PowerShell
    script snippet. The stages are represented in this profile by -ScriptStage1
    thru -ScriptStage7.

-ServiceNameLive=""LetsEncrypt""

    This is the name mapped to the live production certificate network service URL.

-ServiceNameStaging=""LetsEncrypt-Staging""

    This is the name mapped to the non-production (ie. ""staging"") certificate
    network service URL.

-ServiceReportEverything=True

    By default, all activity logged on the client during non-interactive mode
    is uploaded to the SCS server. This can be very helpful during automation
    testing. Once testing is complete, set this switch False to report errors
    only.

    Note: ""non-interactive mode"" means the -Auto switch is set (see above).
          This switch is ignored when -UseStandAloneMode=True.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the profile
    file at startup in command-line format. This may be helpful as a diagnostic.

-SingleSessionEnabled=False

    Set this switch True to run all PowerShell scripts in a single global session.

-SkipSsoServer=False

    When a client is on an SSO domain (ie. it's a member of an SSO server farm),
    it will automatically attempt to update SSO services with each new certificate
    retrieved. Set this switch True to disable SSO updates (for this client only).

-SkipSsoThumbprintUpdates=False

    When a client's domain references an SSO domain, the client will automatically
    attempt to update configuration files with each new SSO certificate thumbprint.
    Set this switch True to disable SSO thumbprint configuration updates (for this
    client only). See {SsoKey} below.

{SsoKey}=""C:\inetpub\wwwroot\web.config""

    This is the path and filename of files that will have their SSO certificate
    thumbprint replaced whenever the related SSO certificate changes. Each file
    with the same name at all levels of the directory hierarchy will be updated,
    starting with the given base path, if the old SSO certificate thumbprint is
    found there. See -SkipSsoThumbprintUpdates above.

    Note: This key may appear any number of times in the profile and wildcards
          can be used in the filename.

-SubmissionRetries=42

    Pending submissions to the certificate provider network will be retried until
    they succeed or fail, at most this many times. By default, the  process will
    retry for 7 minutes (-SubmissionRetries times -SubmissionWaitSecs, see below)
    for challenge status updates as well as certificate request status updates.

-SubmissionWaitSecs=10

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
                            .Replace("{SsoKey}", DoGetCert.sSsoThumbprintFilestKey)
                            .Replace("{{", "{")
                            .Replace("}}", "}")
                            );

                    string lsFetchName = null;

                    // Fetch simple setup.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="Setup Application Folder.exe")
                            , loProfile.sRelativeToProfilePathFile(lsFetchName));

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
                                    loMain.MainWindow = loUI;
                                    loMain.ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose;

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

                            ChannelFactory<GetCertService.IGetCertServiceChannel>   loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                                                                                    DoGetCert.SetCertificate(loGetCertServiceFactory);

                            tvProfile   loMinProfile = DoGetCert.oMinProfile(loProfile);
                            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                            string      lsHash = HashClass.sHashIt(loMinProfile);
                            string      lsLogFileTextReported = null;
                            bool        lbReplaceSsoThumbprint = loDoDa.bReplaceSsoThumbprint(loGetCertServiceFactory, lsHash, lbtArrayMinProfile);
                            bool        lbGetCertificate = loDoDa.bGetCertificate();

                                    // Reinitialize the communications channel in case the certificate was just replaced.
                                    loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                                    DoGetCert.SetCertificate(loGetCertServiceFactory);

                            if ( !lbGetCertificate || !lbReplaceSsoThumbprint )
                            {
                                DoGetCert.ReportErrors(loGetCertServiceFactory, lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
                            }
                            else
                            {
                                DoGetCert.ReportEverything(loGetCertServiceFactory, lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
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
                DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));

                try
                {
                    ChannelFactory<GetCertService.IGetCertServiceChannel>   loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                                                                            DoGetCert.SetCertificate(loGetCertServiceFactory);

                    tvProfile   loMinProfile = DoGetCert.oMinProfile(loProfile);
                    byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                    string      lsHash = HashClass.sHashIt(loMinProfile);
                    string      lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(loGetCertServiceFactory, lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
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
            get
            {
                return goUI;
            }
            set
            {
                goUI = value;
            }
        }
        private static UI goUI;

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
                            if ( null != lsMachineKeyPathFile && DoGetCert.sCertName(goSetupCertificate) == DoGetCert.sNewClientSetupCertName )
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
        private bool bPowerScriptError
        {
            get
            {
                return gbPowerScriptError;
            }
            set
            {
                gbPowerScriptError = value;
            }
        }
        private static bool gbPowerScriptError;

        /// <summary>
        /// This switch is used within the main loop of the UI (see "this.oUI")
        /// to stop any loops running within this object as well.
        /// </summary>
        private bool bPowerScriptSkipLog
        {
            get
            {
                return gbPowerScriptSkipLog;
            }
            set
            {
                gbPowerScriptSkipLog = value;
            }
        }
        private static bool gbPowerScriptSkipLog;

        /// <summary>
        /// This switch is used within the main loop of the UI (see "this.oUI")
        /// to stop any loops running within this object as well.
        /// </summary>
        private string sPowerScriptOutput
        {
            get
            {
                return gsPowerScriptOutput;
            }
            set
            {
                gsPowerScriptOutput = value;
            }
        }
        private static string gsPowerScriptOutput;


        public static string sContentKey  = "-Content";
        public static string sFetchPrefix = "Resources.Fetch.";
        public static string sHostProcess = "GoPcBackup.exe";
        public static string sNewClientSetupCertName = "GetCertClientSetup";
        public static string sSsoThumbprintFilestKey = "-SsoThumbprintFiles";


        /// <summary>
        /// This switch is used within the main loop of the UI (see "this.oUI")
        /// to stop any loops running within this object as well.
        /// </summary>
        public bool bMainLoopStopped
        {
            get
            {
                return gbMainLoopStopped;
            }
            set
            {
                gbMainLoopStopped = value;
            }
        }
        private static bool gbMainLoopStopped;

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
                        {
                            moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                            DoGetCert.SetCertificate(moGetCertServiceFactory);
                        }

                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                        {
                            tvProfile   loMinProfile = DoGetCert.oMinProfile(moProfile);
                            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                            string      lsHash = HashClass.sHashIt(loMinProfile);

                            moDomainProfile = new tvProfile(loGetCertServiceClient.sDomainProfile(lsHash, lbtArrayMinProfile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            if ( 0 == moDomainProfile.Count )
                                throw new Exception("The domain profile is empty. Can't continue.");
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
                string              lsCurrentCertificateName = null;
                X509Certificate2    loCurrentCertificate = DoGetCert.oCurrentCertificate();
                                    if ( null != loCurrentCertificate )
                                        lsCurrentCertificateName = DoGetCert.sCertName(loCurrentCertificate);

                return lsCurrentCertificateName;
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

        public bool bReplaceSsoThumbprint(ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory, string asHash, byte[] abtArrayMinProfile)
        {
            bool lbReplaceSsoThumbprint = false;

            // Do nothing if this is an SSO server (or we are not doing thrumprint updates on this server).
            if ( this.oDomainProfile.bValue("-IsSsoDomain", false) || moProfile.bValue("-SkipSsoThumbprintUpdates", false) )
                return true;

            // Do nothing if no SSO domain is defined for the current domain.
            string  lsSsoDnsName = this.oDomainProfile.sValue("-SsoDnsName", "");
                    if ( "" == lsSsoDnsName )
                        return true;

            DoGetCert.LogIt("");
            DoGetCert.LogIt("Checking for SSO thumbprint change ...");
            
            // At this point it is an error if no SSO thumbprint exists for the current domain.
            string      lsSsoThumbprint = this.oDomainProfile.sValue("-SsoThumbprint", "");
                        if ( "" == lsSsoThumbprint )
                        {
                            DoGetCert.LogIt("");
                            throw new Exception(String.Format("The SSO certificate thumbprint for the \"{0}\" domain has not yet been set!", lsSsoDnsName));
                        }
            string      lsSsoPreviousThumbprint = moProfile.sValue("-SsoThumbprint", "");
                        if ( !moProfile.ContainsKey(DoGetCert.sSsoThumbprintFilestKey) )
                        {
                            string  lsDefaultPhysicalPathFiles = Path.Combine(moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot"), "web.config");
                            Site[]  loSetupCertBoundSiteArray = DoGetCert.oSetupCertBoundSiteArray(new ServerManager());
                                    moProfile.Remove(DoGetCert.sSsoThumbprintFilestKey + "Note");
                                    if ( null == loSetupCertBoundSiteArray || 0 == loSetupCertBoundSiteArray.Length )
                                    {
                                        DoGetCert.LogIt(String.Format("{0} will be set to the IIS default since no sites could be found bound to the current certificate.", DoGetCert.sSsoThumbprintFilestKey));

                                        if ( null != loSetupCertBoundSiteArray )
                                        {
                                            moProfile.Add(DoGetCert.sSsoThumbprintFilestKey, Environment.ExpandEnvironmentVariables(lsDefaultPhysicalPathFiles));
                                            moProfile.Add(DoGetCert.sSsoThumbprintFilestKey + "Note", "No IIS websites could be found bound to the current certificate.");
                                        }
                                        else
                                        {
                                            DoGetCert.LogIt("");
                                            throw new Exception("The current setup certificate is null. This is an error.");
                                        }
                                    }
                                    else
                                    {
                                        foreach(Site loSite in loSetupCertBoundSiteArray)
                                            moProfile.Add(DoGetCert.sSsoThumbprintFilestKey
                                                    , Path.Combine(Environment.ExpandEnvironmentVariables(loSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                                                    , Path.GetFileName(lsDefaultPhysicalPathFiles)));
                                    }
                                    moProfile.Save();
                        }
                        if ( "" == lsSsoPreviousThumbprint )
                        {
                            // Define the previous SSO thumbprint (only if this is the first time through).
                            moProfile["-SsoThumbprint"] = lsSsoThumbprint;
                            moProfile.Save();

                            DoGetCert.LogIt(String.Format("No new SSO thumbprint found (caching old: \"{0}\").", lsSsoThumbprint));

                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
                            {
                                loGetCertServiceClient.NotifySsoThumbprintReplacementSuccess(asHash, abtArrayMinProfile);
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }

                            return true;
                        }
                        // There is nothing to do if the SSO thumbprint hasn't changed.
                        if ( lsSsoPreviousThumbprint == lsSsoThumbprint )
                        {
                            DoGetCert.LogIt("No new SSO thumbprint found.");
                            return true;
                        }
            tvProfile   loFileList = new tvProfile();
                        foreach(DictionaryEntry loEntry in moProfile.oOneKeyProfile(DoGetCert.sSsoThumbprintFilestKey))
                        {
                            // Change DoGetCert.sSsoThumbprintFilestKey to "-Files".
                            loFileList.Add("-Files", loEntry.Value);
                        }
            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = "ReplaceText.exe";
                        loProcess.StartInfo.Arguments = "-OldText={OldSsoThumbprint} -NewText={NewSsoThumbprint}"
                                .Replace("{OldSsoThumbprint}", lsSsoPreviousThumbprint).Replace("{NewSsoThumbprint}", lsSsoThumbprint)
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
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="ReplaceText.exe"), moProfile.sRelativeToProfilePathFile(lsFetchName));
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="ReplaceText.exe.txt"), moProfile.sRelativeToProfilePathFile(lsFetchName));

            System.Windows.Forms.Application.DoEvents();

            this.bPowerScriptError = false;
            this.sPowerScriptOutput = null;

            DoGetCert.LogIt("");
            DoGetCert.LogIt("Replacing SSO thumbprint ...");
            DoGetCert.LogIt(loProcess.StartInfo.Arguments);

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
                    DoGetCert.LogIt(moProfile.sValue("-ReplaceTextTimeoutMsg", "*** thumbprint replacement sub-process timed-out ***\r\n\r\n"));

                if ( this.bPowerScriptError || loProcess.ExitCode != 0 || !loProcess.HasExited )
                {
                    DoGetCert.LogIt(DoGetCert.sExceptionMessage("The thumbprint replacement sub-process experienced a critical failure."));
                }
                else
                {
                    lbReplaceSsoThumbprint = true;

                    DoGetCert.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            DoGetCert.bKillProcess(loProcess);

            if ( lbReplaceSsoThumbprint )
            {
                File.Delete("ReplaceText.exe");
                File.Delete("ReplaceText.exe.txt");

                moProfile["-SsoThumbprint"] = lsSsoThumbprint;
                moProfile.Save();

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifySsoThumbprintReplacementSuccess(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }
            else
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifySsoThumbprintReplacementFailure(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

            return lbReplaceSsoThumbprint;
        }

        public bool bUploadCertToLoadBalancer(string asHash, byte[] abtArrayMinProfile, string asCertName, string asCertPathFile, string asCertificatePassword)
        {
            bool    lbUploadCertToLoadBalancer = false;
            string  lsLoadBalancerPfxPathFile = this.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "");

            DoGetCert.LogIt("");
            DoGetCert.LogIt("Checking for load balancer ...");

            // Do nothing if a load balancer PFX file location is not defined.
            if ( "" == lsLoadBalancerPfxPathFile )
            {
                DoGetCert.LogIt("No load balancer found.");
                return true;
            }

            DoGetCert.LogIt(String.Format("Copying new certificate for use on the \"{0}\" load balancer ...", asCertName));

            X509Certificate2 loLbCertificate = new X509Certificate2(asCertPathFile, asCertificatePassword
                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable);

            Directory.CreateDirectory(Path.GetDirectoryName(lsLoadBalancerPfxPathFile));
            File.WriteAllBytes(lsLoadBalancerPfxPathFile
                    , loLbCertificate.Export(X509ContentType.Pfx, HashClass.sDecrypted(goSetupCertificate, this.oDomainProfile.sValue("-LoadBalancerPfxPassword", ""))));

            DoGetCert.LogSuccess();

            string  lsLoadBalancerExePathFile = this.oDomainProfile.sValue("-LoadBalancerExePathFile", "");
                    if ( "" == lsLoadBalancerExePathFile && !this.bMainLoopStopped )
                    {
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                        {
                            loGetCertServiceClient.NotifyLoadBalancerCertificatePending(asHash, abtArrayMinProfile);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

                        return true;
                    }

            Process     loProcess = new Process();
                        loProcess.ErrorDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessErrorHandler);
                        loProcess.OutputDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessOutputHandler);
                        loProcess.StartInfo.FileName = lsLoadBalancerExePathFile;
                        loProcess.StartInfo.Arguments = lsLoadBalancerPfxPathFile;
                        loProcess.StartInfo.UseShellExecute = false;
                        loProcess.StartInfo.RedirectStandardError = true;
                        loProcess.StartInfo.RedirectStandardInput = true;
                        loProcess.StartInfo.RedirectStandardOutput = true;
                        loProcess.StartInfo.CreateNoWindow = true;

            System.Windows.Forms.Application.DoEvents();

            this.bPowerScriptError = false;
            this.sPowerScriptOutput = null;

            if ( !this.bMainLoopStopped )
            {
                DoGetCert.LogIt("");
                DoGetCert.LogIt("Uploading certificate to load balancer directly ...");

                loProcess.Start();
                loProcess.BeginErrorReadLine();
                loProcess.BeginOutputReadLine();
            }

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(moProfile.dValue("-ReplaceLBcertTimeoutSecs", 120));

            // Wait for the process to finish.
            while ( !this.bMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !this.bMainLoopStopped )
                    Thread.Sleep(moProfile.iValue("-ReplaceLBcertSleepMS", 200));
            }

            if ( !this.bMainLoopStopped )
            {
                if ( !loProcess.HasExited )
                    DoGetCert.LogIt(moProfile.sValue("-ReplaceLBcertTimeoutMsg", "*** load balancer certificate replacement sub-process timed-out ***\r\n\r\n"));

                if ( this.bPowerScriptError || loProcess.ExitCode != 0 || !loProcess.HasExited )
                {
                    DoGetCert.LogIt(DoGetCert.sExceptionMessage("The load balancer certificate replacement sub-process experienced a critical failure."));
                }
                else
                {
                    File.Delete(lsLoadBalancerPfxPathFile);

                    lbUploadCertToLoadBalancer = true;

                    DoGetCert.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            DoGetCert.bKillProcess(loProcess);

            if ( lbUploadCertToLoadBalancer && !this.bMainLoopStopped )
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateExeSuccess(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            else
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateExeFailure(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

            return lbUploadCertToLoadBalancer;
        }

        public void ClearCache()
        {
            // Clear the domain profile cache.
            moDomainProfile = null;
        }

        public void LoadBalancerReleaseCert()
        {
            if ( moProfile.bValue("-UseStandAloneMode", true) )
                return;

            DoGetCert.LogIt("");
            DoGetCert.LogIt("The load balancer admin asserts the new certificate is ready for general release during the next maintenance window ...");

            string  lsCaption = "Release Load Balancer Certificate";
            string  lsLoadBalancerPfxComputer = this.oDomainProfile.sValue("-LoadBalancerPfxComputer", "-LoadBalancerPfxComputer not defined");
            string  lsLoadBalancerPfxPathFile = this.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "-LoadBalancerPfxPathFile not defined");
            bool    lbWrongServer = "" != lsLoadBalancerPfxComputer && lsLoadBalancerPfxComputer != DoGetCert.sCurrentComputerName;
                    if ( lbWrongServer )
                        this.Show(String.Format("The load balancer certificate (\"{0}\") is found on another server (ie. \"{1}\").", lsLoadBalancerPfxPathFile, lsLoadBalancerPfxComputer)
                                , lsCaption
                                , tvMessageBoxButtons.OK, tvMessageBoxIcons.Alert, "LoadBalancerCertWrongServer");
            bool    lbReleased = File.Exists(lsLoadBalancerPfxPathFile);
                    if ( !lbReleased && !lbWrongServer )
                        this.Show(String.Format("The load balancer certificate (\"{0}\") does not yet exist on this server or it has already been released.", lsLoadBalancerPfxPathFile)
                                , lsCaption
                                , tvMessageBoxButtons.OK, tvMessageBoxIcons.Alert, "LoadBalancerCertDoesNotExist");
            string      lsLogFileTextReported = null;
            tvProfile   loMinProfile = DoGetCert.oMinProfile(moProfile);
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);
            ChannelFactory<GetCertService.IGetCertServiceChannel>   loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                                                                    DoGetCert.SetCertificate(loGetCertServiceFactory);

            try
            {
                if ( lbReleased )
                    File.Delete(lsLoadBalancerPfxPathFile);
            }
            catch (Exception ex)
            {
                lbReleased = false;

                this.ShowError(String.Format("The load balancer certificate (\"{0}\") on \"{1}\" can't be removed! ({2})"
                                                , lsLoadBalancerPfxPathFile, DoGetCert.sCurrentComputerName, ex.Message), lsCaption);

                DoGetCert.ReportErrors(loGetCertServiceFactory, lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
            }

            if ( lbReleased || lbWrongServer )
            {

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = loGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateReady(lsHash, lbtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();

                    this.Show("Load balancer \"Certificate Ready\" notifications have been sent.", lsCaption
                            , tvMessageBoxButtons.OK, tvMessageBoxIcons.Done, "LoadBalancerCertReleased");
                }
            }
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
        public static string sLogPathFile()
        {
            return DoGetCert.sUniqueIntOutputPathFile(
                        DoGetCert.sLogPathFileBase()
                    , tvProfile.oGlobal().sValue("-LogFileDateFormat", "-yyyy-MM-dd")
                    , true
                    );
        }

        public static tvProfile oMinProfile(tvProfile aoProfile)
        {
            string      lsIpAddress = null;
                        DoGetCert.bRunPowerScript(out lsIpAddress, aoProfile.sValue("-IpAddressScript", "(ipconfig | findstr \"IPv4\").Split(\":\")[1].Trim()"), false, true);
                        if ( !String.IsNullOrEmpty(lsIpAddress) && aoProfile.sValue("-IpAddress", "IP address not found.") != lsIpAddress )
                        {
                            aoProfile["-IpAddress"] = lsIpAddress;
                            aoProfile.Save();
                        }
            tvProfile   loProfile = new tvProfile(aoProfile.ToString());
                        loProfile.Remove("-Help");
                        loProfile.Remove("-PreviousProcessOutputText");
                        loProfile.Remove("-IpAddress");
                        loProfile.Add("-CurrentLocalTime", DateTime.Now);
                        loProfile.Add("-COMPUTERNAME", DoGetCert.sCurrentComputerName);
                        loProfile.Add("-IpAddress", aoProfile.sValue("-IpAddress", "IP address not found."));

            return loProfile;
        }
        public static string sMinProfile(tvProfile aoProfile)
        {
            string      lsIpAddress = null;
                        DoGetCert.bRunPowerScript(out lsIpAddress, aoProfile.sValue("-IpAddressScript", "(ipconfig | findstr \"IPv4\").Split(\":\")[1].Trim()"), false, true);
                        if ( !String.IsNullOrEmpty(lsIpAddress) && aoProfile.sValue("-IpAddress", "IP address not found.") != lsIpAddress )
                        {
                            aoProfile["-IpAddress"] = lsIpAddress;
                            aoProfile.Save();
                        }
            tvProfile   loProfile = new tvProfile(aoProfile.ToString());
                        loProfile.Remove("-Help");
                        loProfile.Remove("-PreviousProcessOutputText");
                        loProfile.Remove("-IpAddress");
                        loProfile.Add("-CurrentLocalTime", DateTime.Now);
                        loProfile.Add("-COMPUTERNAME", DoGetCert.sCurrentComputerName);
                        loProfile.Add("-IpAddress", aoProfile.sValue("-IpAddress", "IP address not found."));

            return loProfile.ToString();
        }

        /// <summary>
        /// Write the given asText to a text file as well as to
        /// the output console of the UI window (if it exists).
        /// </summary>
        /// <param name="asText">The text string to log.</param>
        public static void LogIt(string asText)
        {
            string          lsLogPathFileBase = null;
            StreamWriter    loStreamWriter = null;

            try
            {
                lsLogPathFileBase = DoGetCert.sLogPathFileBase();

                // Move down to the base folder if we're running an update.
                string  lsPathSep = Path.DirectorySeparatorChar.ToString();
                string  lsPath = Path.GetDirectoryName(tvProfile.oGlobal().sRelativeToProfilePathFile(lsLogPathFileBase)).Replace(
                                    String.Format("{0}{1}{0}",  lsPathSep, tvProfile.oGlobal().sValue("-UpdateFolder", "Update")), lsPathSep);
                        Directory.CreateDirectory(lsPath);

                loStreamWriter = new StreamWriter(tvProfile.oGlobal().sRelativeToProfilePathFile(
                        DoGetCert.sUniqueIntOutputPathFile(lsLogPathFileBase
                        , tvProfile.oGlobal().sValue("-LogFileDateFormat", "-yyyy-MM-dd"), true)), true);
                loStreamWriter.WriteLine(DateTime.Now.ToString(tvProfile.oGlobal().sValueNoTrim(
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

            if ( null != goUI )
            goUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                goUI.AppendOutputTextLine(asText);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

        public void LogStage(string asStageId)
        {
            DoGetCert.LogIt("");
            DoGetCert.LogIt("");
            DoGetCert.LogIt(String.Format("Stage {0} ...", asStageId));
            DoGetCert.LogIt("");
        }

        public static void LogSuccess()
        {
            DoGetCert.LogIt("Success.");
        }

        public static void ReportErrors(ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory
                , string asHash, byte[] abtArrayMinProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(DoGetCert.sLogPathFile());
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return;

            string  lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                    if ( File.Exists(lsPreviousErrorsLogPathFile) )
                        File.AppendAllText(lsPreviousErrorsLogPathFile, asLogFileTextReported);
                    else
                        File.Copy(lsLogPathFile, lsPreviousErrorsLogPathFile, false);

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            loCompressContentProfile.Add(DoGetCert.sContentKey, File.ReadAllText(lsPreviousErrorsLogPathFile));

                loGetCertServiceClient.ReportErrors(asHash, abtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                File.Delete(lsPreviousErrorsLogPathFile);
            }
        }

        public static void ReportEverything(ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory
                , string asHash, byte[] abtArrayMinProfile, out string asLogFileTextReported)
        {
            string  lsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(DoGetCert.sLogPathFile());
                    asLogFileTextReported = File.ReadAllText(lsLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) || !tvProfile.oGlobal().bValue("-ServiceReportEverything", true) )
                return;

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            loCompressContentProfile.Add(DoGetCert.sContentKey, asLogFileTextReported);

                loGetCertServiceClient.ReportEverything(asHash, abtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( tvProfile.oGlobal().bValue("-DoStagingTests", true) && tvProfile.oGlobal().bValue("-ResetStagingLogs", true) )
                File.Delete(lsLogPathFile);
        }

        public void ShowError(string asMessageText, string asMessageCaption)
        {
            DoGetCert.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowError(this.oUI, asMessageText, asMessageCaption);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

     
        // This kludge is necessary since neither "Process.CloseMainWindow()" nor "Process.Kill()" work reliably.
        private static bool bKillProcess(Process aoProcess)
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

                aoProcess.WaitForExit(1000 * tvProfile.oGlobal().iValue("-KillProcessOrderlyWaitSecs", 10));
            }
            catch (Exception ex)
            {
                lbKillProcess = false;
                DoGetCert.LogIt(String.Format("Orderly bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, DoGetCert.sExceptionMessage(ex)));
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

                    aoProcess.WaitForExit(tvProfile.oGlobal().iValue("-KillProcessForcedWaitMS", 1000));

                    lbKillProcess = aoProcess.HasExited;
                }
                catch (Exception ex)
                {
                    lbKillProcess = false;
                    DoGetCert.LogIt(String.Format("Forced bKillProcess() Failed on PID={0} (\"{1}\")", liProcessId, DoGetCert.sExceptionMessage(ex)));
                }
            }

            return lbKillProcess;
        }

        private static bool bKillProcessParent(Process aoProcess)
        {
            bool lbKillProcess = true;

            foreach (Process loProcess in Process.GetProcessesByName(aoProcess.ProcessName))
            {
                if ( loProcess.Id != aoProcess.Id )
                    lbKillProcess = DoGetCert.bKillProcess(loProcess);
            }

            return lbKillProcess;
        }

        /// <summary>
        /// Returns the current IIS port 443 bound certificate (if any), otherwise whatever's in the local store (by name).
        /// </summary>
        private static X509Certificate2 oCurrentCertificate()
        {
            return DoGetCert.oCurrentCertificate(null);
        }
        private static X509Certificate2 oCurrentCertificate(string asCertName)
        {
            string[]        lsSanArray = null;
            ServerManager   loServerManager = null;
            Site            loSite = null;
            Site            loPrimarySiteForDefaults = null;

            return DoGetCert.oCurrentCertificate(asCertName, lsSanArray, out loServerManager, out loSite, out loPrimarySiteForDefaults);
          }
        private static X509Certificate2 oCurrentCertificate(string asCertName, out ServerManager aoServerManager)
        {
            string[]        lsSanArray = null;
            Site            loSite = null;
            Site            loPrimarySiteForDefaults = null;

            return DoGetCert.oCurrentCertificate(asCertName, lsSanArray, out aoServerManager, out loSite, out loPrimarySiteForDefaults);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out Site aoSiteFound
                )
        {
            ServerManager   loServerManager = null;
            Site            loPrimarySiteForDefaults = null;

            return DoGetCert.oCurrentCertificate(asCertName, asSanArray, out loServerManager, out aoSiteFound, out loPrimarySiteForDefaults);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                )
        {
            Site            loPrimarySiteForDefaults = null;

            return DoGetCert.oCurrentCertificate(asCertName, asSanArray, out aoServerManager, out aoSiteFound, out loPrimarySiteForDefaults);
        }
        private static X509Certificate2 oCurrentCertificate(
                  string asCertName
                , string[] asSanArray
                , out ServerManager aoServerManager
                , out Site aoSiteFound
                , out Site aoPrimarySiteForDefaults
                )
        {
            X509Certificate2    loCurrentCertificate = null;

            // ServerManager is the IIS ServerManager. It gives us the website objectfor binding the cert.
            aoServerManager = new ServerManager();
            aoSiteFound = null;
            aoPrimarySiteForDefaults = null;

            try
            {
                string  lsCertName = null == asCertName ? null : asCertName.ToLower();
                Binding loBindingFound = null;

                foreach(Site loSite in aoServerManager.Sites)
                {
                    // First, look for a binding by matching certificate.
                    foreach (Binding loBinding in loSite.Bindings)
                    {
                        // Select the first binding with a certificate (and a matching name - if lsCertName is not null).
                        foreach(X509Certificate2 loCertificate in DoGetCert.oCurrentCertificateCollection())
                        {
                            if ( null != loBinding.CertificateHash && loCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash)
                                && (null == lsCertName || lsCertName == DoGetCert.sCertName(loCertificate).ToLower()) )
                            {
                                // Binding found that matches a certificate and certificate name (or no name).
                                loCurrentCertificate = loCertificate;
                                aoSiteFound = loSite;
                                loBindingFound = loBinding;
                                break;
                            }
                        }

                        if ( null != loBindingFound )
                            break;
                    }

                    // Next, look for a binding by matching certificate (name match not required).
                    if ( null == loBindingFound && null != goSetupCertificate )
                    foreach (Binding loBinding in loSite.Bindings)
                    {
                        if ( null != loBinding.CertificateHash && goSetupCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash) )
                        {
                            // Binding found that matches the setup certificate.
                            loCurrentCertificate = null;
                            aoSiteFound = loSite;
                            loBindingFound = loBinding;
                            break;
                        }
                    }

                    // Next, if no binding certificate match exists, look by binding hostname (ie. SNI).
                    if ( null == loBindingFound )
                        foreach (Binding loBinding in loSite.Bindings)
                        {
                            if ( null == loBinding.CertificateHash && null != lsCertName
                                    && lsCertName == loBinding.Host.ToLower() && "https" == loBinding.Protocol )
                            {
                                // Site found with a binding hostname that matches lsCertName.
                                aoSiteFound = loSite;
                                loBindingFound = loBinding;
                                break;
                            }

                            if ( null != loBindingFound )
                                break;
                        }

                    // Finally, with no binding found, try to match against the site name.
                    if ( null == loBindingFound && asCertName == loSite.Name )
                        aoSiteFound = loSite;

                    if ( null != aoSiteFound )
                        break;
                }

                // Finally, finally (no really), with no site found, find a related primary site to use for defaults.
                string lsPrimaryCertName = null == asSanArray || 0 == asSanArray.Length ? null : asSanArray[0].ToLower();

                if ( null == aoSiteFound && null != lsPrimaryCertName && lsCertName != lsPrimaryCertName )
                {
                    // Use the primary site in the SAN array as defaults for any new site (to be created).
                    DoGetCert.oCurrentCertificate(lsPrimaryCertName, asSanArray, out aoPrimarySiteForDefaults);
                }

                // As a last resort (so much for finally), use the newest certificate found in the local store (by name).
                if ( null == loCurrentCertificate && null != lsCertName )
                    foreach(X509Certificate2 loCertificate in DoGetCert.oCurrentCertificateCollection())
                    {
                        if ( lsCertName == DoGetCert.sCertName(loCertificate).ToLower()
                                && (null == loCurrentCertificate || loCertificate.NotAfter > loCurrentCertificate.NotAfter) )
                        {
                            loCurrentCertificate = loCertificate;
                        }
                    }
            }
            catch (Exception ex)
            {
                DoGetCert.LogIt(ex.Message);
                DoGetCert.LogIt("IIS may not be fully configured. This will prevent certificate-to-site binding in IIS.");
            }

            return loCurrentCertificate;
        }

        private static Binding oSslBinding(Site aoSite)
        {
            if ( null == aoSite || null == aoSite.Bindings )
                return null;

            Binding loSslBinding = null;

            foreach (Binding loBinding in aoSite.Bindings)
                if ( null != loBinding && "https" == loBinding.Protocol )
                {
                    loSslBinding = loBinding;
                    break;
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
            return String.Format("GetCertServiceFault: {0}{1}{2}{3}", ex.Message, (null == ex.InnerException ? "": "; " + ex.InnerException.Message), Environment.NewLine, ex.StackTrace);
        }
        public static string sExceptionMessage(string asMessage)
        {
            return "GetCertServiceFault: " + asMessage;
        }

        /// <summary>
        /// Returns the current "LogPathFile" base name.
        /// </summary>
        private static string sLogPathFileBase()
        {
            string lsPath = "Logs";
            string lsLogFile = "Log.txt";
            string lsLogPathFileBase = lsLogFile;

            try
            {
                lsLogPathFileBase = tvProfile.oGlobal().sValue("-LogPathFile"
                        , Path.Combine(lsPath, Path.GetFileNameWithoutExtension(tvProfile.oGlobal().sLoadedPathFile) + lsLogFile));

                lsPath = Path.GetDirectoryName(tvProfile.oGlobal().sRelativeToProfilePathFile(lsLogPathFileBase));
                if ( !Directory.Exists(lsPath) )
                    Directory.CreateDirectory(lsPath);
            }
            catch {}

            return lsLogPathFileBase;
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

        private static bool bDoCheckIn(string asHash, byte[] abtArrayMinProfile, tvMessageBox aoStartupWaitMsg)
        {
            bool lbDoCheckIn = false;

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return true;

            DoGetCert.LogIt("Attempting check-in with the certificate repository ...");

            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                if ( null != aoStartupWaitMsg )
                    aoStartupWaitMsg.Hide();

                DoGetCert.SetCertificate(loGetCertServiceClient);

                if ( null != aoStartupWaitMsg && !tvProfile.oGlobal().bExit )
                    aoStartupWaitMsg.Show();

                string      lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                string      lsErrorLog = null;
                tvProfile   loCompressContentProfile = new tvProfile();
                        if ( File.Exists(lsPreviousErrorsLogPathFile) )
                        {
                            lsErrorLog = File.ReadAllText(lsPreviousErrorsLogPathFile);
                            loCompressContentProfile.Add(DoGetCert.sContentKey, lsErrorLog);
                        }

                lbDoCheckIn = loGetCertServiceClient.bClientCheckIn(asHash, abtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                DoGetCert.LogIt("");
                if ( !lbDoCheckIn )
                {
                    throw new Exception("Client FAILED to check-in with the certificate repository. Can't continue.");
                }
                else
                {
                    DoGetCert.LogIt(String.Format("Client successfully checked-in with the certificate repository{0}.", null == lsErrorLog ? "" : " (and reported previous errors)"));

                    File.Delete(lsPreviousErrorsLogPathFile);
                }
            }

            return lbDoCheckIn;
        }

        /// <summary>
        /// Set client certificate.
        /// </summary>
        public static void SetCertificate(ChannelFactory<GetCertService.IGetCertServiceChannel> aoChannelFactory)
        {
            tvProfile loProfile = tvProfile.oGlobal();

            if ( loProfile.bValue("-UseStandAloneMode", true) || loProfile.bExit )
                return;

            if ( null == aoChannelFactory.Credentials.ClientCertificate.Certificate )
            {
                if ( null == goSetupCertificate )
                    DoGetCert.SetCertificate(new GetCertService.GetCertServiceClient());

                aoChannelFactory.Credentials.ClientCertificate.Certificate = goSetupCertificate;
            }
            else
            if ( null == goSetupCertificate )
                goSetupCertificate = aoChannelFactory.Credentials.ClientCertificate.Certificate;
        }
        private static void SetCertificate(GetCertService.GetCertServiceClient aoGetCertServiceClient)
        {
            tvProfile loProfile = tvProfile.oGlobal();

            if ( loProfile.bValue("-UseStandAloneMode", true) || loProfile.bExit )
                return;

            if ( null == aoGetCertServiceClient.ClientCredentials.ClientCertificate.Certificate )
            {
                if ( null == goSetupCertificate && "" != loProfile.sValue("-ContactEmailAddress" ,"") )
                    goSetupCertificate = DoGetCert.oCurrentCertificate(loProfile.sValue("-CertificateDomainName" ,""));

                if ( null == goSetupCertificate && !loProfile.bValue("-CertificateSetupDone", false) && !loProfile.bValue("-NoPrompts", false) )
                {
                    System.Windows.Forms.OpenFileDialog loOpenDialog = new System.Windows.Forms.OpenFileDialog();
                                                        loOpenDialog.FileName = DoGetCert.sNewClientSetupCertName + ".pfx";
                    System.Windows.Forms.DialogResult   leDialogResult = System.Windows.Forms.DialogResult.None;
                                                        leDialogResult = loOpenDialog.ShowDialog();

                    if ( System.Windows.Forms.DialogResult.OK != leDialogResult )
                    {
                        loProfile.bExit = true;
                    }
                    else
                    {
                        string lsPassword = Microsoft.VisualBasic.Interaction.InputBox("Password?", Path.GetFileName(loOpenDialog.FileName), "", -1, 50);

                        if ( "" == lsPassword )
                        {
                            loProfile.bExit = true;
                        }
                        else
                        {
                            try
                            {
                                goSetupCertificate = new X509Certificate2(loOpenDialog.FileName, lsPassword);
                            }
                            catch (Exception ex)
                            {
                                loProfile.bExit = true;
                                DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));
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
            DoGetCert.LogIt(asMessageText);

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
            DoGetCert.LogIt(asMessageText);

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
            DoGetCert.LogIt(asMessageText);

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
            DoGetCert.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

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
            DoGetCert.LogIt("");
            DoGetCert.LogIt(String.Format("{0} exiting due to log-off or system shutdown.", Path.GetFileName(System.Windows.Application.ResourceAssembly.Location)));

            if ( null != this.oUI )
            this.oUI.HandleShutdown();
        }

        private bool bRunPowerScript(string asScript)
        {
            string lsOutput = null;

            return this.bRunPowerScript(out lsOutput, asScript);
        }
        private bool bRunPowerScript(string asScript, bool abOpenOrCloseSingleSession)
        {
            string lsOutput = null;

            return DoGetCert.bRunPowerScript(out lsOutput, asScript, abOpenOrCloseSingleSession, false);
        }
        private bool bRunPowerScript(out string asOutput, string asScript)
        {
            return DoGetCert.bRunPowerScript(out asOutput, asScript, false, false);
        }
        private static bool bRunPowerScript(out string asOutput, string asScript, bool abOpenOrCloseSingleSession, bool abSkipLog)
        {
            bool            lbSingleSessionEnabled = tvProfile.oGlobal().bValue("-SingleSessionEnabled", false);
            string          lsSingleSessionScriptPathFile = null;
                            if ( lbSingleSessionEnabled )
                            {
                                lsSingleSessionScriptPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PowerScriptSessionPathFile", "InGetCertSession.ps1"));

                                // Fetch global session script.
                                tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                        , String.Format("{0}{1}", DoGetCert.sFetchPrefix, Path.GetFileName(lsSingleSessionScriptPathFile)), lsSingleSessionScriptPathFile);
                            }
                            else
                            if ( abOpenOrCloseSingleSession )
                            {
                                // -SingleSessionEnabled is false, so ignore the single session "open" and "close" scripts.
                                asOutput = null;
                                return true;
                            }
            bool            lbRunPowerScript = false;
            string          lsScriptPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PowerScriptPathFile", "PowerScript.ps1"));
            StreamWriter    loStreamWriter   = null;
            string          lsScript = null;
                            string          lsKludgeHide1 = Guid.NewGuid().ToString();  // Apparently Regex can't handle "++" in the text to process.
                            StringBuilder   lsbReplaceProfileTokens = new StringBuilder(asScript.Replace("++", lsKludgeHide1));

                            // Look for string tokens of the form: {ProfileKey}, where "ProfileKey" is
                            // only expected to be found in the profile if it is prefixed by a hyphen.
                            Regex   loRegex = new Regex("{(.*?)}");
                                    foreach (Match loMatch in loRegex.Matches(lsbReplaceProfileTokens.ToString()))
                                    {
                                        string  lsKey = loMatch.Groups[1].ToString();
                                                if ( tvProfile.oGlobal().ContainsKey(lsKey) )
                                                    lsbReplaceProfileTokens.Replace(loMatch.Groups[0].ToString(), tvProfile.oGlobal()[lsKey].ToString());
                                    }

                            lsScript = lsbReplaceProfileTokens.ToString().Replace(lsKludgeHide1, "++");

            try
            {
                loStreamWriter = new StreamWriter(lsScriptPathFile);
                loStreamWriter.WriteLine(lsScript);
            }
            catch (Exception ex)
            {
                DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));
            }
            finally
            {
                if ( null != loStreamWriter )
                    loStreamWriter.Close();
            }

            string  lsProcessPathFile = tvProfile.oGlobal().sValue("-PowerShellExePathFile", @"C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe");
            string  lsProcessArgs = tvProfile.oGlobal().sValue("-PowerShellExeArgs", @"-NoProfile -ExecutionPolicy unrestricted -File ""{0}"" ""{1}""");
                    if ( !lbSingleSessionEnabled || abOpenOrCloseSingleSession )
                        lsProcessArgs = String.Format(lsProcessArgs, lsScriptPathFile, "");
                    else
                        lsProcessArgs = String.Format(lsProcessArgs, lsSingleSessionScriptPathFile, lsScriptPathFile);
            string  lsLogScript = !lsScript.Contains("-CertificateKey") ? lsScript : lsScript.Substring(0, lsScript.IndexOf("-CertificateKey"));

            if ( !abSkipLog && (!lbSingleSessionEnabled || !abOpenOrCloseSingleSession) )
                DoGetCert.LogIt(lsLogScript);

            gbPowerScriptSkipLog = abSkipLog;

            Process loProcess = new Process();
                    loProcess.ErrorDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessErrorHandler);
                    loProcess.OutputDataReceived += new DataReceivedEventHandler(DoGetCert.PowerScriptProcessOutputHandler);
                    loProcess.StartInfo.FileName = lsProcessPathFile;
                    loProcess.StartInfo.Arguments = lsProcessArgs;
                    loProcess.StartInfo.UseShellExecute = false;
                    loProcess.StartInfo.RedirectStandardError = true;
                    loProcess.StartInfo.RedirectStandardInput = true;
                    loProcess.StartInfo.RedirectStandardOutput = true;
                    loProcess.StartInfo.CreateNoWindow = true;

            System.Windows.Forms.Application.DoEvents();

            gbPowerScriptError = false;
            gsPowerScriptOutput = null;

            loProcess.Start();
            loProcess.BeginErrorReadLine();
            loProcess.BeginOutputReadLine();

            DateTime ltdProcessTimeout = DateTime.Now.AddSeconds(tvProfile.oGlobal().dValue("-PowerScriptTimeoutSecs", 300));

            // Wait for the process to finish.
            while ( !gbMainLoopStopped && !loProcess.HasExited && DateTime.Now < ltdProcessTimeout )
            {
                System.Windows.Forms.Application.DoEvents();

                if ( !gbMainLoopStopped )
                    Thread.Sleep(tvProfile.oGlobal().iValue("-PowerScriptSleepMS", 200));
            }

            if ( !gbMainLoopStopped )
            {
                if ( !loProcess.HasExited )
                    DoGetCert.LogIt(tvProfile.oGlobal().sValue("-PowerScriptTimeoutMsg", "*** sub-process timed-out ***\r\n\r\n"));

                if ( gbPowerScriptError || loProcess.ExitCode != 0 || !loProcess.HasExited )
                {
                    DoGetCert.LogIt(DoGetCert.sExceptionMessage("The sub-process experienced a critical failure."));
                }
                else
                {
                    lbRunPowerScript = true;

                    if ( null != lsScriptPathFile )
                        File.Delete(lsScriptPathFile);

                    if ( !abSkipLog && (!lbSingleSessionEnabled || !abOpenOrCloseSingleSession) )
                        DoGetCert.LogSuccess();
                }
            }

            loProcess.CancelErrorRead();
            loProcess.CancelOutputRead();

            DoGetCert.bKillProcess(loProcess);

            asOutput = gsPowerScriptOutput;
            gbPowerScriptSkipLog = false;

            return lbRunPowerScript;
        }

        private static void PowerScriptProcessOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length )
            {
                gsPowerScriptOutput += outLine.Data;

                if ( !gbPowerScriptSkipLog )
                    DoGetCert.LogIt(outLine.Data);
            }
        }

        private static void PowerScriptProcessErrorHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            if ( null != outLine.Data && 0 != outLine.Data.Length )
            {
                gsPowerScriptOutput += outLine.Data;
                gbPowerScriptError = true;
                DoGetCert.LogIt(outLine.Data);
            }
        }

        private bool bApplyNewCertToSSO(X509Certificate2 aoNewCertificate, X509Certificate2 aoOldCertificate, string asHash, byte[] abtArrayMinProfile)
        {
            // Do nothing if running stand-alone or this isn't an SSO server (or we are not doing SSO certificate updates on this server).
            if ( moProfile.bValue("-UseStandAloneMode", true) || !this.oDomainProfile.bValue("-IsSsoDomain", false)
                        || null == aoOldCertificate || moProfile.bValue("-SkipSsoServer", false) )
                return true;

            bool lbApplyNewCertToSSO = false;

            DoGetCert.LogIt("");
            DoGetCert.LogIt("But first, replace SSO certificates ...");

            string  lsOutput = null;
            bool    lbIsOlder = false;
            bool    lbIsProxy = false;
                    if ( !moProfile.ContainsKey("-IsSsoOlder") || !moProfile.ContainsKey("-IsSsoProxy") )
                    {
                        if ( !DoGetCert.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptSsoVerChk1", "try {(Get-AdfsProperties).HostName} catch {}"), false, true) )
                        {
                            DoGetCert.LogIt("Failed checking SSO vs. proxy.");
                            return false;
                        }

                        lbIsProxy = String.IsNullOrEmpty(lsOutput);

                        if ( lbIsProxy )
                        {
                            if ( !DoGetCert.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptSsoProxyVerChk", "try {(Get-WebApplicationProxySslCertificate).HostName} catch {}"), false, true) )
                            {
                                DoGetCert.LogIt("Failed checking SSO proxy version.");
                                return false;
                            }

                            lbIsOlder = String.IsNullOrEmpty(lsOutput);
                        }
                        else
                        {
                            if ( !DoGetCert.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptSsoVerChk2", "try {(Get-AdfsProperties).AuditLevel} catch {}"), false, true) )
                            {
                                DoGetCert.LogIt("Failed checking SSO version.");
                                return false;
                            }

                            lbIsOlder = String.IsNullOrEmpty(lsOutput);
                        }

                        moProfile["-IsSsoOlder"] = lbIsOlder;
                        moProfile["-IsSsoProxy"] = lbIsProxy;
                        moProfile.Save();
                    }
                    else
                    {
                        lbIsOlder = moProfile.bValue("-IsSsoOlder", false);
                        lbIsProxy = moProfile.bValue("-IsSsoProxy", false);
                    }
            string  lsDefaultScript = null;
                    if ( lbIsProxy )
                    {
                        if ( !lbIsOlder )
                            // SSO proxy server, 3+ (nothing to do for older proxies).
                            lsDefaultScript = @"
Set-WebApplicationProxySslCertificate -Thumbprint ""{NewCertificateThumbprint}""

Restart-Service -Name adfssrv
                                    ";
                    }
                    else
                    {
                        lsDefaultScript = @"
$ServiceName = ""adfssrv""
# Wait at most 3 minutes for the ADFS service to start (after a likely reboot).
if ( (Get-Service -Name $ServiceName).Status -ne ""Running"" ) {Start-Sleep -s 180}

Add-PSSnapin Microsoft.Adfs.PowerShell 2> $null

if ( (Get-AdfsCertificate -CertificateType ""Service-Communications"").Thumbprint -eq ""{NewCertificateThumbprint}"" )
{
    echo ""SSO certificates already replaced (by another server in the farm).""
    echo ""The ADFS service will be restarted.""
}
else
{
    Set-AdfsCertificate    -CertificateType ""Service-Communications"" -Thumbprint ""{NewCertificateThumbprint}""

    Add-AdfsCertificate    -CertificateType ""Token-Decrypting""       -Thumbprint ""{NewCertificateThumbprint}"" -IsPrimary
    Remove-AdfsCertificate -CertificateType ""Token-Decrypting""       -Thumbprint ""{OldCertificateThumbprint}""

    Add-AdfsCertificate    -CertificateType ""Token-Signing""          -Thumbprint ""{NewCertificateThumbprint}"" -IsPrimary
    Remove-AdfsCertificate -CertificateType ""Token-Signing""          -Thumbprint ""{OldCertificateThumbprint}""
}

Restart-Service -Name $ServiceName
                                ";

                        // For ADFS 3+, insert the "HTTP.SYS" SSL assignment (above service restart).
                        if ( !lbIsOlder )
                            lsDefaultScript = @"
#Error suppression is needed here since ""Set-AdfsSslCertificate"" always times out, even after a successful run.
try {Set-AdfsSslCertificate -Thumbprint ""{NewCertificateThumbprint}""} catch {}
                                    "
                                    + Environment.NewLine
                                    + lsDefaultScript;
                    }

            lbApplyNewCertToSSO = this.bRunPowerScript(moProfile.sValue("-ScriptSSO", lsDefaultScript)
                    .Replace("{NewCertificateThumbprint}", aoNewCertificate.Thumbprint)
                    .Replace("{OldCertificateThumbprint}", aoOldCertificate.Thumbprint)
                    );

            if ( null == lsDefaultScript )
                throw new Exception("-ScriptSSO is empty. This is an error.");

            if ( lbApplyNewCertToSSO && !lbIsProxy )
            {
                DoGetCert.LogIt("");
                DoGetCert.LogIt("Updating SSO token signing certificate thumbprint in the database ...");

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    lbApplyNewCertToSSO = loGetCertServiceClient.bSetSsoThumbprint(asHash, abtArrayMinProfile, aoNewCertificate.Thumbprint);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

                if ( lbApplyNewCertToSSO )
                    DoGetCert.LogSuccess();
                else
                    DoGetCert.LogIt(DoGetCert.sExceptionMessage("The SSO thumbprint update sub-process experienced a critical failure."));
            }

            return lbApplyNewCertToSSO;
        }

        private bool bCertificateNotExpiring(ChannelFactory<GetCertService.IGetCertServiceChannel> aoGetCertServiceFactory
                                                , string asHash, byte[] abtArrayMinProfile, X509Certificate2 aoOldCertificate)
        {
            bool lbCertificateNotExpiring = false;

            if ( moProfile.bValue("-DoStagingTests", true) )
            {
                DoGetCert.LogIt("( staging mode is in effect (-DoStagingTests=True) )");
                DoGetCert.LogIt("");
            }

            if ( null != aoOldCertificate )
            {
                bool    lbCheckExpiration = true;
                string  lsCertOverridePfxComputer = this.oDomainProfile.sValue("-CertOverridePfxComputer", "");
                string  lsCertOverridePfxPathFile = this.oDomainProfile.sValue("-CertOverridePfxPathFile", "");

                if ( "" != lsCertOverridePfxComputer + lsCertOverridePfxPathFile && ("" == lsCertOverridePfxComputer || "" == lsCertOverridePfxPathFile) )
                {
                    throw new Exception("Both CertOverridePfxComputer AND CertOverridePfxPathFile must be defined for certificate overrides to work properly.");
                }
                else
                if ( !this.oDomainProfile.bValue("-CertOverridePfxReady", false) && DoGetCert.sCurrentComputerName == lsCertOverridePfxComputer && File.Exists(lsCertOverridePfxPathFile) )
                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = aoGetCertServiceFactory.CreateChannel())
                    {
                        loGetCertServiceClient.NotifyCertOverrideCertificateReady(asHash, abtArrayMinProfile);
                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                            loGetCertServiceClient.Abort();
                        else
                            loGetCertServiceClient.Close();
                    }

                // Allow checking staging certificates for expiration only in non-interactive mode.

                if ( lbCheckExpiration && moProfile.bValue("-DoStagingTests", true) && !moProfile.bValue("-Auto", false) )
                    lbCheckExpiration = false;

                if ( lbCheckExpiration && !moProfile.bValue("-DoStagingTests", true) )
                {
                    // We are running in production mode. Only check valid production certificates for expiration.
                    lbCheckExpiration = aoOldCertificate.Verify();

                    if ( !lbCheckExpiration )
                    {
                        DoGetCert.LogIt(String.Format("The current certificate ({0}) appears to be invalid.", aoOldCertificate.Subject));
                        DoGetCert.LogIt("  (If you know otherwise, make sure internet connectivity is available and the system clock is accurate.)");
                        DoGetCert.LogIt("Expiration status will therefore be ignored and the process will now run.");
                    }
                }

                if ( lbCheckExpiration )
                {
                    DateTime    ldtBeforeExpirationDate;
                                if ( moProfile.ContainsKey("-CertificateRenewalDateOverride") )
                                {
                                    ldtBeforeExpirationDate = moProfile.dtValue("-CertificateRenewalDateOverride", DateTime.Now);
                                }
                                else
                                if ( moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    ldtBeforeExpirationDate = aoOldCertificate.NotAfter.AddDays(-moProfile.iValue("-RenewalDaysBeforeExpiration", 30));
                                }
                                else
                                {
                                    int liRenewalDaysBeforeExpiration = this.oDomainProfile.iValue("-RenewalDaysBeforeExpiration", 30);
                                        if ( moProfile.iValue("-RenewalDaysBeforeExpiration", 30) != liRenewalDaysBeforeExpiration )
                                        {
                                            moProfile["-RenewalDaysBeforeExpiration"] = liRenewalDaysBeforeExpiration;
                                            moProfile.Save();
                                        }

                                    ldtBeforeExpirationDate = aoOldCertificate.NotAfter.AddDays(-liRenewalDaysBeforeExpiration);
                                }

                    if ( DateTime.Now < ldtBeforeExpirationDate )
                    {
                        DoGetCert.LogIt(String.Format("Nothing to do until {0}. The current \"{1}\" certificate doesn't expire until {2}."
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
            tvProfile   loMinProfile = null;
            byte[]      lbtArrayMinProfile = null;
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

                loProfile = tvProfile.oGlobal(new tvProfile(args, true));
                if ( loProfile.bExit )
                    return loProfile;

                DoGetCert.VerifyDependenciesInit();
                DoGetCert.VerifyDependencies();
                if ( loProfile.bExit )
                    return loProfile;

                // Only do the following fetches before initial setup and if the containing application folder has been created.
                if ( "" == loProfile.sValue("-ContactEmailAddress", "") && "" == loProfile.sValue("-CertificateDomainName", "")
                        && Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Contains(Path.GetFileNameWithoutExtension(ResourceAssembly.Location)) )
                {
                    //if ( null == DoGetCert.sProcessExePathFile(DoGetCert.sHostProcess) )
                    //{
                    //    // Fetch host (if it's not already running).
                    //    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    //            , String.Format("{0}{1}", DoGetCert.sFetchPrefix, DoGetCert.sHostProcess), loProfile.sRelativeToProfilePathFile(DoGetCert.sHostProcess));
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

                loMinProfile = DoGetCert.oMinProfile(loProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);

                if ( !loProfile.bValue("-NoPrompts", false) )
                {
                    aoStartupWaitMsg = new tvMessageBox();
                    aoStartupWaitMsg.ShowWait(
                            null, Path.GetFileNameWithoutExtension(loProfile.sExePathFile) + " loading, please wait ...", 250);
                }

                if ( null == lsUpdateRunExePathFile && null == lsUpdateDeletePathFile )
                {
                    // Check-in with the GetCert service.
                    loProfile.bExit = !DoGetCert.bDoCheckIn(lsHash, lbtArrayMinProfile, aoStartupWaitMsg);

                    // Short-circuit updates until basic setup is complete.
                    if ( "" == loProfile.sValue("-ContactEmailAddress", "") )
                        return loProfile;

                    // Does the "hosts" file need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleHostsEntryUpdate(loProfile, lsHash, lbtArrayMinProfile);

                    // A not-yet-updated EXE is running. Does it need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleClientExeUpdate(loProfile, lsHash, lbtArrayMinProfile, lsRunKey, loCmdLine);
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

                        DoGetCert.LogIt(String.Format(
                                "Software update successfully completed. Version {0} is now running.", lsCurrentVersion));
                    }
                }

                if ( !loProfile.bExit )
                {
                    loProfile = DoGetCert.oHandleClientCfgUpdate(loProfile, lsHash, lbtArrayMinProfile);
                    loProfile = DoGetCert.oHandleClientIniUpdate(loProfile, lsHash, lbtArrayMinProfile, args);

                    DoGetCert.HandleHostUpdates(loProfile, lsHash, lbtArrayMinProfile);
                }
            }
            catch (Exception ex)
            {
                if ( null != loProfile )
                    loProfile.bExit = true;

                DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));

                if ( !loProfile.bValue("-UseStandAloneMode", true) && !loProfile.bValue("-CertificateSetupDone", false) )
                {
                    DoGetCert.LogIt(@"

Checklist for {EXE} setup (SCS version):

    On this local server:

        .) Once ""-CertificateDomainName"" and ""-ContactEmailAddress"" have been
           provided (after the initial use of the ""{CertName}"" certificate),
           the SCS will expect to use a domain specific certificate for further
           communications. You can blank out either ""-CertificateDomainName"" or
           ""-ContactEmailAddress"" in the profile file to force the use of the
           ""{CertName}"" certificate again (instead of an existing domain
           certificate).

        .) A ""Host process can't be located on disk. Can't continue."" message means
           the host software (ie. the task scheduler) must also be installed to continue.
           Look in the ""{EXE}"" folder for the host EXE and run it.

    On the SCS server:

        .) The ""-StagingTestsEnabled"" switch must be set to ""True"" for the staging
           tests (it gets reset automatically every night).

        .) The ""{CertName}"" certificate must be added
           to the trusted people store.

        .) The ""{CertName}"" certificate must be removed 
           from the ""untrusted"" certificate store (if it's there).
"
                            .Replace("{EXE}", Path.GetFileName(ResourceAssembly.Location))
                            .Replace("{CertName}", DoGetCert.sNewClientSetupCertName)
                            )
                            ;
                }

                try
                {
                    ChannelFactory<GetCertService.IGetCertServiceChannel>   loGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                                                                            DoGetCert.SetCertificate(loGetCertServiceFactory);
                    string lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(loGetCertServiceFactory, lsHash, lbtArrayMinProfile, out lsLogFileTextReported);
                }
                catch {}
            }

            return loProfile;
        }

        private static string sHostExePathFile(out Process aoProcessFound)
        {
            return DoGetCert.sHostExePathFile(out aoProcessFound, false);
        }
        private static string sHostExePathFile(out Process aoProcessFound, bool abQuietMode)
        {
            string  lsHostExeFilename = DoGetCert.sHostProcess;
            string  lsHostExePathFile = DoGetCert.sProcessExePathFile(lsHostExeFilename, out aoProcessFound);

            if ( null == lsHostExePathFile )
            {
                if ( !abQuietMode )
                    DoGetCert.LogIt("Host image can't be located on disk based on the currently running process. Trying typical locations ...");

                lsHostExePathFile = @"C:\ProgramData\GoPcBackup\GoPcBackup.exe";
                if ( !File.Exists(lsHostExePathFile) )
                    lsHostExePathFile = @"C:\Program Files\GoPcBackup\GoPcBackup.exe";

                if ( !File.Exists(lsHostExePathFile) )
                {
                    DoGetCert.LogIt("Host process can't be located on disk. Can't continue.");
                    throw new Exception("Exiting ...");
                }

                if ( !abQuietMode )
                    DoGetCert.LogIt("Host process image found on disk. It will be restarted.");
            }

            return lsHostExePathFile;
        }

        // Return an array of every site using the current setup certificate.
        private static Site[] oSetupCertBoundSiteArray(ServerManager aoServerManager)
        {
            if ( null == goSetupCertificate )
                return null;

            List<Site>  loList = new List<Site>();
            byte[]      lbtSetupCertHashArray = goSetupCertificate.GetCertHash();

            // Walk thru the site list looking for all sites bound to the setup certificate.
            foreach(Site loSite in aoServerManager.Sites)
            {
                foreach (Binding loBinding in loSite.Bindings)
                {
                    if ( null != loBinding.CertificateHash && lbtSetupCertHashArray.SequenceEqual(loBinding.CertificateHash) )
                    {
                        loList.Add(loSite);
                        break;
                    }
                }
            }

            return loList.ToArray();
        }

        // Amend the SAN list with every site (by name) using the current setup certificate.
        private static string[] sSanArrayAmended(ServerManager aoServerManager, string[] asSanArray)
        {
            Site[]  loSetupCertBoundSiteArray = DoGetCert.oSetupCertBoundSiteArray(aoServerManager);
                    if ( null == loSetupCertBoundSiteArray )
                        return asSanArray;

            List<string> loList = new List<string>();

            foreach(Site loSite in loSetupCertBoundSiteArray)
                loList.Add(loSite.Name);

            // Append the given SAN array.
            foreach(string lsSanItem in asSanArray)
            {
                if ( !loList.Contains(lsSanItem) )
                    loList.Add(lsSanItem);
            }

            return loList.ToArray();
        }

        private static string sStopHostExePathFile()
        {
            Process loProcessFound = null;
            string  lsHostExePathFile = DoGetCert.sHostExePathFile(out loProcessFound, true);

            // Stop the EXE (can't update it or the INI while it's running).

            if ( !DoGetCert.bKillProcess(loProcessFound) )
            {
                DoGetCert.LogIt("Host process can't be stopped. Update can't be applied.");
                throw new Exception("Exiting ...");
            }
            else
            {
                DoGetCert.LogIt("Host process stopped successfully.");
            }

            return lsHostExePathFile;
        }

        private static void HandleClientExeUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile, string asRunKey, tvProfile aoCmdLine)
        {
            string  lsCurrentVersion = FileVersionInfo.GetVersionInfo(aoProfile.sExePathFile).FileVersion;
            byte[]  lbtArrayExeUpdate = null;

            // Look for EXE update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetCertificate(loGetCertServiceClient);

                lbtArrayExeUpdate = loGetCertServiceClient.btArrayGetCertExeUpdate(asHash, abtArrayMinProfile, lsCurrentVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lbtArrayExeUpdate )
            {
                DoGetCert.LogIt(String.Format("Client version {0} is running. No update.", lsCurrentVersion));
            }
            else
            {
                DoGetCert.LogIt("");
                DoGetCert.LogIt(String.Format("Client version {0} is running. Update found ...", lsCurrentVersion));

                string  lsUpdatePath = Path.Combine(Path.GetDirectoryName(aoProfile.sExePathFile), aoProfile.sValue("-UpdateFolder", "Update"));
                        Directory.CreateDirectory(lsUpdatePath);
                string  lsNewExePathFile = Path.Combine(lsUpdatePath, Path.GetFileName(aoProfile.sExePathFile));
                string  lsNewIniPathFile = Path.Combine(lsUpdatePath, Path.GetFileName(aoProfile.sLoadedPathFile));

                DoGetCert.LogIt(String.Format("Writing update to \"{0}\".", lsNewExePathFile));
                File.WriteAllBytes(lsNewExePathFile, lbtArrayExeUpdate);

                DoGetCert.LogIt(String.Format("Writing update profile to \"{0}\".", lsNewIniPathFile));
                File.Copy(aoProfile.sLoadedPathFile, lsNewIniPathFile, true);

                DoGetCert.LogIt(String.Format("The file version of \"{0}\" is {1}."
                        , lsNewExePathFile, FileVersionInfo.GetVersionInfo(lsNewExePathFile).FileVersion));

                DoGetCert.LogIt(String.Format("Starting update \"{0}\" ...", lsNewExePathFile));
                Process loProcess = new Process();
                        loProcess.StartInfo.FileName = lsNewExePathFile;
                        loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2}"
                                , asRunKey, aoProfile.sExePathFile, aoCmdLine.sCommandLine());
                        loProcess.Start();

                        System.Windows.Forms.Application.DoEvents();

                aoProfile.bExit = true;
            }
        }

        private static tvProfile oHandleClientCfgUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            string      lsCfgVersion = aoProfile.sValue("-CfgVersion", "1");
            string      lsCfgUpdate = null;

            // Look for WCF configuration update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetCertificate(loGetCertServiceClient);

                lsCfgUpdate = loGetCertServiceClient.sCfgUpdate(asHash, abtArrayMinProfile, lsCfgVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsCfgUpdate )
            {
                DoGetCert.LogIt(String.Format("WCF configuration version {0} is in use. No update.", lsCfgVersion));
            }
            else
            {
                DoGetCert.LogIt(String.Format("WCF configuration version {0} is in use. Update found ...", lsCfgVersion));

                // Overwrite WCF config with the current update.
                File.WriteAllText(aoProfile.sLoadedPathFile, lsCfgUpdate
                        .Replace("{CertificateDomainName}", aoProfile.sValue("-CertificateDomainName", "")));

                // Overwrite WCF config version number in the current profile.
                tvProfile   loWcfCfg = new tvProfile(aoProfile.sLoadedPathFile, true);
                            aoProfile["-CfgVersion"] = loWcfCfg.sValue("-CfgVersion", "1");
                            aoProfile.Save();

                DoGetCert.LogIt(String.Format(
                        "WCF configuration update successfully completed. Version {0} is now in use.", aoProfile.sValue("-CfgVersion", "1")));
            }

            return aoProfile;
        }

        private static tvProfile oHandleClientIniUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile, string[] args)
        {
            string lsIniVersion = aoProfile.sValue("-IniVersion", "1");
            string lsIniUpdate = null;

            // Look for profile update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetCertificate(loGetCertServiceClient);

                lsIniUpdate = loGetCertServiceClient.sIniUpdate(asHash, abtArrayMinProfile, lsIniVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsIniUpdate )
            {
                DoGetCert.LogIt(String.Format("Profile version {0} is in use. No update.", lsIniVersion));
            }
            else
            {
                DoGetCert.LogIt(String.Format("Profile version {0} is in use. Update found ...", lsIniVersion));

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

                DoGetCert.LogIt(String.Format(
                        "Profile update successfully completed. Version {0} is now in use.", aoProfile.sValue("-IniVersion", "1")));
            }

            return aoProfile;
        }

        private static void HandleHostUpdates(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            bool        lbRestartHost = false;
            Process     loHostProcess = null;
            string      lsHostExePathFile= DoGetCert.sHostExePathFile(out loHostProcess);
            string      lsHostExeVersion = FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion;
            byte[]      lbtArrayGpcExeUpdate = null;
                        // Look for an updated GoPcBackup.exe.
                        using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
                        {
                            DoGetCert.SetCertificate(loGetCertServiceClient);

                            lbtArrayGpcExeUpdate = loGetCertServiceClient.btArrayGoPcBackupExeUpdate(asHash, abtArrayMinProfile, lsHostExeVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( null != lbtArrayGpcExeUpdate )
                            DoGetCert.LogIt(String.Format("Host version {0} is in use. Update found ...", lsHostExeVersion));

            string      lsHostIniVersion = aoProfile.sValue("-HostIniVersion", "1");
            string      lsHostIniUpdate  = null;
                        // Look for an updated GoPcBackup.exe.txt.
                        using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
                        {
                            DoGetCert.SetCertificate(loGetCertServiceClient);

                            lsHostIniUpdate = loGetCertServiceClient.sGpcIniUpdate(asHash, abtArrayMinProfile, lsHostIniVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( null != lsHostIniUpdate )
                            DoGetCert.LogIt(String.Format("Host profile version {0} is in use. Update found ...", lsHostIniVersion));

            if ( null == lbtArrayGpcExeUpdate && null == lsHostIniUpdate && null == loHostProcess )
                lbRestartHost = true;

            if ( null != lbtArrayGpcExeUpdate || null != lsHostIniUpdate )
            {
                lbRestartHost = true;

                DoGetCert.sStopHostExePathFile();

                // Write the updated EXE (if any).
                if ( null != lbtArrayGpcExeUpdate )
                {
                    // Write the EXE.
                    File.WriteAllBytes(lsHostExePathFile, lbtArrayGpcExeUpdate);

                    DoGetCert.LogIt(String.Format("Host update successfully completed. Version {0} is now in use."
                                                            , FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion));
                }

                // Write the updated INI (if any).
                if ( null != lsHostIniUpdate )
                {
                    tvProfile   loUpdateProfile = new tvProfile(lsHostIniUpdate);
                    tvProfile   loHostProfile = new tvProfile(lsHostExePathFile + ".txt", false);

                                // Remove all "GetCert" related keys in tasks, backups and cleanups
                                // (ie. see "loHostProfile.Remove" below). Then add in any updates.
                                tvProfile   loTasks = new tvProfile();
                                            // Place tasks in the update on top of the list.
                                            foreach(DictionaryEntry loEntry in loUpdateProfile.oProfile("-AddTasks"))
                                            {
                                                loTasks.Add("-Task", loEntry.Value);
                                            }
                                            // Look for all tasks in the host profile that are NOT GetCert related.
                                            foreach(DictionaryEntry loEntry in loHostProfile.oProfile("-AddTasks"))
                                            {
                                                tvProfile   loTask = new tvProfile(loEntry.Value.ToString());
                                                            if ( !loTask.sValue("-CommandEXE", "").ToLower().Contains(Path.GetFileName(ResourceAssembly.Location).ToLower()) )
                                                                loTasks.Add("-Task", loEntry.Value);
                                            }
                                // Look for EXE folder references in the existing backup and cleanup sets. If found, skip them. In other words,
                                // within the new loUpdateProfile, leave out all sets from loHostProfile that make any reference to this EXE.
                                string      lsExeFolderNameToLower = String.Format(@"\{0}", Path.GetFileNameWithoutExtension(ResourceAssembly.Location).ToLower());
                                tvProfile   loBackupSets = new tvProfile();
                                            foreach(DictionaryEntry loEntry in loHostProfile.oOneKeyProfile("-BackupSet"))
                                            {
                                                bool        lbAddSet = true;
                                                tvProfile   loSet = (new tvProfile(loEntry.Value.ToString())).oOneKeyProfile("-FolderToBackup");
                                                            foreach(DictionaryEntry loPathFile in loSet)
                                                            {
                                                                if ( loPathFile.Value.ToString().ToLower().Contains(lsExeFolderNameToLower) )
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
                                                                if ( loPathFile.Value.ToString().ToLower().Contains(lsExeFolderNameToLower) )
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

                    DoGetCert.LogIt(String.Format(
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

                DoGetCert.LogIt("Host process restarted successfully.");
            }
        }

        private static void HandleHostsEntryUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            string lsHstVersion = aoProfile.sValue("-HostsEntryVersion", "1");
            string lsHstUpdate = null;

            // Look for hosts entry update.
            using (GetCertService.GetCertServiceClient loGetCertServiceClient = new GetCertService.GetCertServiceClient())
            {
                DoGetCert.SetCertificate(loGetCertServiceClient);

                lsHstUpdate = loGetCertServiceClient.sHostsEntryUpdate(asHash, abtArrayMinProfile, lsHstVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lsHstUpdate )
            {
                DoGetCert.LogIt(String.Format("Hosts entry version {0} is in use. No update.", lsHstVersion));
            }
            else
            {
                DoGetCert.LogIt(String.Format("Hosts entry version {0} is in use. Update found ...", lsHstVersion));

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
                        DoGetCert.LogIt("Something's wrong with the local 'hosts' file format. Can't update IP address.");
                        DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));
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

                DoGetCert.LogIt(String.Format(
                        "Hosts entry update successfully completed. Version {0} is now in use.", aoProfile.sValue("-HostsEntryVersion", "1")));

                // Restart host (should restart this app as well - after a brief delay).
                Process loHostProcess = new Process();
                        loHostProcess.StartInfo.FileName = Path.Combine(Path.GetDirectoryName(DoGetCert.sStopHostExePathFile()), "Startup.cmd");
                        loHostProcess.StartInfo.Arguments = "-StartupTasksDelaySecs=10";
                        loHostProcess.StartInfo.UseShellExecute = false;
                        loHostProcess.StartInfo.CreateNoWindow = true;
                        loHostProcess.Start();

                DoGetCert.LogIt("Host process restarted successfully.");

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

        private static void VerifyDependenciesInit()
        {
            string  lsFetchName = null;

            // Fetch IIS Manager DLL.
            tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsFetchName="Microsoft.Web.Administration.dll"), tvProfile.oGlobal().sRelativeToProfilePathFile(lsFetchName));
        }

        private static void VerifyDependencies()
        {
            if ( tvProfile.oGlobal().bExit )
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

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
            {
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
                        = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine", false))
                {
                    try
                    {
                        if ( String.Compare("5.1", (string)loLocalMachineStoreRegistryKey.GetValue("PowerShellVersion", "")) > 0 )
                            throw new Exception();
                    }
                    catch
                    {
                        liErrors++;
                        lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                                + "PowerShell 5 or later must be installed.";
                    }
                }
            }

            if ( "" != lsErrors )
            {
                tvProfile.oGlobal().bExit = true;

                lsErrors = ("" == lsErrors ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                        + String.Format("Please install {0} and try again.", liErrors == 1 ? "it" : "them");

                if ( !tvProfile.oGlobal().bValue("-NoPrompts", false) )
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
            tvProfile   loMinProfile = null;
            byte[]      lbtArrayMinProfile = null;
            string      lsHash = null;
            X509Store   loStore = null;

            DoGetCert.LogIt("");
            DoGetCert.LogIt("Get certificate process started ...");
            this.bMainLoopStopped = false;

            try
            {
                moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                DoGetCert.SetCertificate(moGetCertServiceFactory);

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
                loMinProfile = DoGetCert.oMinProfile(moProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);

                bool                lbSingleSessionEnabled = moProfile.bValue("-SingleSessionEnabled", false);
                string              lsInstanceGuid = moProfile.sValue("-InstanceGuid", Guid.NewGuid().ToString());
                string              lsDefaultPhysicalPath = moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot");
                bool                lbCertificatePassword = false;  // Set this "false" to use password only for load balancer file.
                string              lsCertificatePassword = HashClass.sHashPw(loMinProfile);
                bool                lbCreateSanSites = moProfile.bValue("-CreateSanSites", true);
                ServerManager       loServerManager = null;
                X509Certificate2    loOldCertificate = DoGetCert.oCurrentCertificate(lsCertName, out loServerManager);
                                    if ( this.bCertificateNotExpiring(moGetCertServiceFactory, lsHash, lbtArrayMinProfile, loOldCertificate) )
                                    {
                                        // The certificate is not ready to expire soon. There is nothing to do.
                                        return true;
                                    }
                X509Certificate2    loNewCertificate = null;
                byte[]              lbtArrayNewCertificate = null;
                string              lsCertOverridePfxComputer = this.oDomainProfile.sValue("-CertOverridePfxComputer", "");
                string              lsCertOverridePfxPathFile = this.oDomainProfile.sValue("-CertOverridePfxPathFile", "");
                bool                lbCertOverrideApplies = "" != lsCertOverridePfxPathFile && "" != lsCertOverridePfxComputer;
                bool                lbLoadBalancerCertPending = this.oDomainProfile.bValue("-LoadBalancerPfxPending", false);
                bool                lbRepositoryCertsOnly = this.oDomainProfile.bValue("-RepositoryCertsOnly", false);
                DateTime            ldtMaintenanceWindowBeginTime = this.oDomainProfile.dtValue("-MaintenanceWindowBeginTime", DateTime.MinValue);
                DateTime            ldtMaintenanceWindowEndTime   = this.oDomainProfile.dtValue("-MaintenanceWindowEndTime",   DateTime.MaxValue);
                                    if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                    {
                                        if ( lbLoadBalancerCertPending )
                                        {
                                            int     liErrorCount = 1 + moProfile.iValue("-LoadBalancerPendingErrorCount", 0);
                                                    moProfile["-LoadBalancerPendingErrorCount"] = liErrorCount;
                                                    moProfile.Save();

                                            DoGetCert.LogIt("The load balancer administrator or process has not yet released the new certificate.");

                                            if ( liErrorCount > moProfile.iValue("-LoadBalancerPendingMaxErrors", 1) )
                                            {
                                                lbGetCertificate = false;

                                                DoGetCert.LogIt("This is an error.");
                                            }
                                            else
                                            {
                                                DoGetCert.LogIt("Will try again next cycle.");
                                                DoGetCert.LogIt("");

                                                return true;
                                            }
                                        }
                                        else
                                            using (GetCertService.IGetCertServiceChannel moGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                            {
                                                lbtArrayNewCertificate = moGetCertServiceClient.btArrayNewCertificate(lsHash, lbtArrayMinProfile);
                                                if ( CommunicationState.Faulted == moGetCertServiceClient.State )
                                                    moGetCertServiceClient.Abort();
                                                else
                                                    moGetCertServiceClient.Close();

                                                if ( null == lbtArrayNewCertificate )
                                                {
                                                    DoGetCert.LogIt("No new certificate was found in the repository.");
                                                }
                                                else
                                                if ( moProfile.bValue("-Auto", false) && (DateTime.Now < ldtMaintenanceWindowBeginTime || DateTime.Now >= ldtMaintenanceWindowEndTime) )
                                                {
                                                    lbGetCertificate = false;

                                                    DoGetCert.LogIt(String.Format("Can't proceed. The current time is \"{0}\". The maintenance window for the \"{1}\" domain is from \"{2}\" to \"{3}\"."
                                                            , DateTime.Now, lsCertName, ldtMaintenanceWindowBeginTime, ldtMaintenanceWindowEndTime));
                                                }
                                                else
                                                {
                                                    DoGetCert.LogIt("New certificate downloaded from the repository.");

                                                    // Buffer certificate to disk (same as what's done below).
                                                    Directory.CreateDirectory(Path.GetDirectoryName(lsCertPathFile));
                                                    File.WriteAllBytes(lsCertPathFile, lbtArrayNewCertificate);
                                                }
                                            }
                                    }

                if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && lbRepositoryCertsOnly && !lbCertOverrideApplies )
                {
                    int liErrorCount = 1 + moProfile.iValue("-RepositoryCertsOnlyErrorCount", 0);
                        moProfile["-RepositoryCertsOnlyErrorCount"] = liErrorCount;
                        moProfile.Save();

                    if ( liErrorCount > moProfile.iValue("-RepositoryCertsOnlyMaxErrors", 2) )
                    {
                        lbGetCertificate = false;

                        DoGetCert.LogIt("This is an error since -RepositoryCertsOnly=True.");
                    }
                    else
                    {
                        DoGetCert.LogIt("-RepositoryCertsOnly=True. Will try again next cycle.");
                        DoGetCert.LogIt("");

                        return true;
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && (!lbRepositoryCertsOnly || lbCertOverrideApplies) )
                {
                    if ( lbCertOverrideApplies )
                    {
                        DoGetCert.LogIt("");

                        if ( DoGetCert.sCurrentComputerName == lsCertOverridePfxComputer )
                        {
                            DoGetCert.LogIt(String.Format("A new certificate for \"{0}\" will come from the certificate override.", lsCertName));

                            lsCertPathFile = lsCertOverridePfxPathFile;
                        }
                        else
                        {
                            int liErrorCount = 1 + moProfile.iValue("-CertOverrideErrorCount", 0);
                                moProfile["-CertOverrideErrorCount"] = liErrorCount;
                                moProfile.Save();

                            DoGetCert.LogIt(String.Format("The certificate override for \"{0}\" will be found on another server (ie. \"{1}\").", lsCertName, lsCertOverridePfxComputer));

                            if ( liErrorCount > moProfile.iValue("-CertOverrideMaxErrors", 3) )
                            {
                                lbGetCertificate = false;

                                DoGetCert.LogIt("No more retries. This is an error.");
                                DoGetCert.LogIt("");
                            }
                            else
                            {
                                DoGetCert.LogIt("Will try again next cycle.");
                                DoGetCert.LogIt("");

                                return true;
                            }
                        }
                    }
                    else
                    {
                        // There is no certificate override file. So proceed to get a new certificate from the certificate provider network.

                                // Wait a random period each cycle to allow different clients the opportunity to lock the renewal (skip if manually running tests).
                        int     liMaxCertRenewalLockDelaySecs = moProfile.iValue("-MaxCertRenewalLockDelaySecs", 300);
                                if ( !moProfile.bValue("-UseStandAloneMode", true) && !moProfile.bValue("-DoStagingTests", true) && moProfile.bValue("-Auto", false) )
                                {
                                    DoGetCert.LogIt(String.Format("Waiting at most {0} seconds to set a certificate renewal lock for this domain ...", liMaxCertRenewalLockDelaySecs));
                                    System.Windows.Forms.Application.DoEvents();
                                    Thread.Sleep(1000 * new Random().Next(liMaxCertRenewalLockDelaySecs));
                                }
                        bool    lbLockCertificateRenewal = false;
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                                {
                                    lbLockCertificateRenewal = loGetCertServiceClient.bLockCertificateRenewal(lsHash, lbtArrayMinProfile);
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
                                    DoGetCert.LogIt("Certificate renewal has been locked for this domain.");
                                }
                                else
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    DoGetCert.LogIt("Certificate renewal can't be locked. Will try again next cycle.");
                                    DoGetCert.LogIt("");

                                    return true;
                                }

                        DoGetCert.LogIt("");
                        DoGetCert.LogIt(String.Format("Retrieving new certificate for \"{0}\" from the certificate provider network ...", lsCertName));

                        string  lsAcmePsFile = "ACME-PS.zip";
                        string  lsAcmePsPath = Path.GetFileNameWithoutExtension(lsAcmePsFile);
                        string  lsAcmePsPathFile = moProfile.sRelativeToProfilePathFile(lsAcmePsFile);

                        // Fetch AcmePs module.
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", DoGetCert.sFetchPrefix, lsAcmePsFile), lsAcmePsPathFile);
                        if ( !Directory.Exists(moProfile.sRelativeToProfilePathFile(lsAcmePsPath)) )
                            ZipFile.ExtractToDirectory(lsAcmePsPathFile, Path.GetDirectoryName(lsAcmePsPathFile));

                        // Don't use the PowerShell gallery installed AcmePs module (by default, use what's embedded instead).
                        if ( !moProfile.bValue("-AcmePsModuleUseGallery", false) )
                        {
                            string lsAcmePsPathFolder = moProfile.sValue("-AcmePsPath", lsAcmePsPath);

                            moProfile.sValue("-AcmePsPathHelp", String.Format(
                                    "Any alternative -AcmePsPath must include a subfolder named \"{0}\" that contains the files.", lsAcmePsPathFolder));

                            lsAcmePsPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsPath", lsAcmePsPath));
                        }

                        string      lsAcmeWorkPath = moProfile.sRelativeToProfilePathFile(moProfile.sValue("-AcmePsWorkPath", "AcmeState"));
                        string      lsSanPsStringArray = "(";
                                    foreach(string lsSanItem in lsSanArray)
                                    {
                                        lsSanPsStringArray += String.Format(",\"{0}\"",  lsSanItem);
                                    }
                                    lsSanPsStringArray = lsSanPsStringArray.Replace("(,", "(") + ")";


                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            this.LogStage("1 - init ACME workspace");

                            // Delete ACME workspace.
                            if ( Directory.Exists(lsAcmeWorkPath) )
                                Directory.Delete(lsAcmeWorkPath, true);

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptSessionOpen", @"
New-PSSession -ComputerName localhost -Name GetCert
$session = Disconnect-PSSession -Name GetCert
                                    "), true);

                            DoGetCert.LogIt("");

                            if ( lbGetCertificate )
                            {
                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage1s", @"
Using Module ""{AcmePsPath}""

$global:state = New-ACMEState -Path ""{AcmeWorkPath}""
Get-ACMEServiceDirectory $global:state -ServiceName ""{AcmeServiceName}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            .Replace("{AcmeServiceName}", moProfile.bValue("-DoStagingTests", true) ? moProfile.sValue("-ServiceNameStaging", "LetsEncrypt-Staging")
                                                                                                                    : moProfile.sValue("-ServiceNameLive", "LetsEncrypt"))
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage1", @"
Import-Module ""{AcmePsPath}""

New-ACMEState      -Path ""{AcmeWorkPath}""
Get-ACMEServiceDirectory ""{AcmeWorkPath}"" -ServiceName ""{AcmeServiceName}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            .Replace("{AcmeServiceName}", moProfile.bValue("-DoStagingTests", true) ? moProfile.sValue("-ServiceNameStaging", "LetsEncrypt-Staging")
                                                                                                                    : moProfile.sValue("-ServiceNameLive", "LetsEncrypt"))
                                            );
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            this.LogStage("2 - register domain contact, submit order & authorization request");

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2s", @"
Using Module ""{AcmePsPath}""

$SanList = {SanPsStringArray}

New-ACMENonce      $global:state
New-ACMEAccountKey $global:state
New-ACMEAccount    $global:state -EmailAddresses ""{-ContactEmailAddress}"" -AcceptTOS -PassThru
$global:order = New-ACMEOrder $global:state -Identifiers $SanList

$global:authZ = Get-ACMEAuthorization $global:state -Order $global:order

[int[]] $global:SanMap = $null
        foreach ($SAN in $SanList) { for ($i=0; $i -lt $global:authZ.Length; $i++) { if ( $global:authZ[$i].Identifier.value -eq $SAN ) { $global:SanMap += $i }}}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2", @"
Import-Module ""{AcmePsPath}""

New-ACMENonce      ""{AcmeWorkPath}""
New-ACMEAccountKey ""{AcmeWorkPath}""
New-ACMEAccount    ""{AcmeWorkPath}"" -EmailAddresses ""{-ContactEmailAddress}"" -AcceptTOS -PassThru
New-ACMEOrder      ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                        }

                        ArrayList   loAcmeCleanedList = new ArrayList();

                        for (int liSanArrayIndex=0; liSanArrayIndex < lsSanArray.Length; liSanArrayIndex++)
                        {
                            string lsSanItem = lsSanArray[liSanArrayIndex];

                            if ( lbGetCertificate && !this.bMainLoopStopped )
                            {
                                this.LogStage(String.Format("3 - Define DNS name to be challenged (\"{0}\"), setup domain challenge in IIS and submit it to certificate provider", lsSanItem));

                                Site    loSanSite = null;
                                Site    loPrimarySiteForDefaults = null;
                                        DoGetCert.oCurrentCertificate(lsSanItem, lsSanArray, out loServerManager, out loSanSite, out loPrimarySiteForDefaults);

                                // If -CreateSanSites is false, only create a default
                                // website (as needed) and use it for all SAN values.
                                if ( null == loSanSite && !lbCreateSanSites && 0 != loServerManager.Sites.Count )
                                {
                                    DoGetCert.LogIt(String.Format("No website found for \"{0}\". -CreateSanSites is \"False\". So no website created.\r\n", lsSanItem));

                                    loSanSite = loServerManager.Sites[0];
                                }

                                if ( null == loSanSite )
                                {
                                    string lsPhysicalPath = null;

                                    if ( 0 == loServerManager.Sites.Count )
                                    {
                                        DoGetCert.LogIt("No default website could be found.");

                                        lsPhysicalPath = lsDefaultPhysicalPath;
                                    }
                                    else
                                    {
                                        if ( null != loPrimarySiteForDefaults )
                                            lsPhysicalPath = loPrimarySiteForDefaults.Applications["/"].VirtualDirectories["/"].PhysicalPath;
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

                                    DoGetCert.LogIt(String.Format("No website found for \"{0}\". New website created.", lsSanItem));
                                    DoGetCert.LogIt("");
                                }

                                string  lsAcmeBasePath = moProfile.sValue("-AcmeBasePath", @".well-known\acme-challenge");
                                string  lsAcmePath = Path.Combine(
                                          Environment.ExpandEnvironmentVariables(loSanSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                                        , lsAcmeBasePath);

                                // Cleanup ACME folder.
                                if ( !loAcmeCleanedList.Contains(lsAcmePath) )
                                {
                                    if ( Directory.Exists(lsAcmePath) )
                                    {
                                        Directory.Delete(lsAcmePath, true);
                                        Thread.Sleep(moProfile.iValue("-ACMEcleanupSleepMS", 200));
                                    }

                                    // Add an IIS "web.config" (needed for challenge submissions - ie. filenames with no extensions).
                                    Directory.CreateDirectory(lsAcmePath);
                                    File.WriteAllText(Path.Combine(lsAcmePath, moProfile.sValue("-AcmeWebConfigFilename", "web.config"))
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

                                    // Don't cleanup the same folder more than once per session.
                                    loAcmeCleanedList.Add(lsAcmePath);
                                }

                                if ( lbGetCertificate && !this.bMainLoopStopped )
                                {
                                    if ( lbSingleSessionEnabled )
                                        lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3s", @"
Using Module ""{AcmePsPath}""

$challenge = Get-ACMEChallenge $global:state $global:authZ[$global:SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmeChallengePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge $global:state
                                                ")
                                                .Replace("{AcmePsPath}", lsAcmePsPath)
                                                .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                .Replace("{AcmeChallengePath}", lsAcmePath)
                                                );
                                    else
                                        lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3", @"
Import-Module ""{AcmePsPath}""

$SanList = {SanPsStringArray}
$order   = Find-ACMEOrder        ""{AcmeWorkPath}"" -Identifiers $SanList
$authZ   = Get-ACMEAuthorization ""{AcmeWorkPath}"" -Order $order

[int[]] $SanMap = $null; foreach ($SAN in $SanList) { for ($i=0; $i -lt $authZ.Length; $i++) { if ( $authZ[$i].Identifier.value -eq $SAN ) { $SanMap += $i }}}

$challenge = Get-ACMEChallenge ""{AcmeWorkPath}"" $authZ[$SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmeChallengePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge ""{AcmeWorkPath}""
                                                ")
                                                .Replace("{AcmePsPath}", lsAcmePsPath)
                                                .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                .Replace("{AcmeChallengePath}", lsAcmePath)
                                                );
                                }
                            }
                        }

                        int liSubmissionRetries = moProfile.iValue("-SubmissionRetries", 42);

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(1000 * moProfile.iValue("-SubmissionWaitSecs", 10));

                                this.LogStage(String.Format("4 - update challenge{0} from certificate provider", 1 == lsSanArray.Length ? "" : "s"));

                                string  lsSubmissionPending = ": pending";
                                string  lsOutput = null;

                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage4s", @"
Using Module ""{AcmePsPath}""

$global:order | Update-ACMEOrder $global:state -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage4", @"
Import-Module ""{AcmePsPath}""

$order = Find-ACMEOrder   ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$order | Update-ACMEOrder ""{AcmeWorkPath}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );

                                if ( !lsOutput.Contains(lsSubmissionPending) || this.bMainLoopStopped )
                                    break;
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            this.LogStage("5 - generate certificate request and submit");

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage5s", @"
Using Module ""{AcmePsPath}""

$global:certKey = New-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Complete-ACMEOrder $global:state -Order $global:order -CertificateKey $global:certKey
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage5", @"
Import-Module ""{AcmePsPath}""

$order = Find-ACMEOrder ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$certKey = New-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Complete-ACMEOrder      ""{AcmeWorkPath}"" -Order $order -CertificateKey $certKey
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(1000 * moProfile.iValue("-SubmissionWaitSecs", 10));

                                this.LogStage("6 - update certificate request");

                                string  lsSubmissionPending = "CertificateRequest       : \r\nCrtPemFile               :";
                                string  lsOutput = null;

                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage6s", @"
Using Module ""{AcmePsPath}""

$global:order | Update-ACMEOrder $global:state -PassThru;
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage6", @"
Import-Module ""{AcmePsPath}""

$order = Find-ACMEOrder   ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$order | Update-ACMEOrder ""{AcmeWorkPath}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );

                                if ( !lsOutput.Contains(lsSubmissionPending) || this.bMainLoopStopped )
                                    break;
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            this.LogStage("7 - get certificate");

                            File.Delete(lsCertPathFile);

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage7s", @"
Using Module ""{AcmePsPath}""

Export-ACMECertificate $global:state -Order $global:order -CertificateKey $global:certKey -Path ""{CertificatePathFile}""
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{CertificatePathFile}", lsCertPathFile)
                                        .Replace("{CertificatePassword}", lsCertificatePassword)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage7", @"
Import-Module ""{AcmePsPath}""

$order = Find-ACMEOrder ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$certKey = Import-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Export-ACMECertificate  ""{AcmeWorkPath}"" -Order $order -CertificateKey $certKey -Path ""{CertificatePathFile}""
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{CertificatePathFile}", lsCertPathFile)
                                        .Replace("{CertificatePassword}", lsCertificatePassword)
                                        );
                        }
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped )
                {
                    // At this point there must be a new certificate file ready on disk (ie. for upload, the local cert store, port binding, etc).
                    lbGetCertificate = File.Exists(lsCertPathFile);

                    if ( !lbGetCertificate )
                        DoGetCert.LogIt(DoGetCert.sExceptionMessage("The new certificate could not be found!"));

                    if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && !moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // Upload new certificate to the load balancer.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            if ( lbCertOverrideApplies )
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPathFile
                                                                                    , HashClass.sDecrypted(goSetupCertificate, this.oDomainProfile.sValue("-CertOverridePfxPassword", "")));
                            else
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPathFile, lsCertificatePassword);
                        }

                        // Upload new certificate to the repository.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                        {
                            lbGetCertificate = loGetCertServiceClient.bNewCertificateUploaded(
                                    lsHash, lbtArrayMinProfile, File.ReadAllBytes(lsCertPathFile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            DoGetCert.LogIt("");
                            if ( lbGetCertificate )
                                DoGetCert.LogIt("New certificate successfully uploaded to the repository.");
                            else
                                DoGetCert.LogIt(DoGetCert.sExceptionMessage("Failed uploading new certificate to repository."));
                        }
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped && null != lbtArrayNewCertificate || moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // A non-null "lbtArrayNewCertificate" means the certificate was just downloaded from the
                        // repository. Load it to be installed and bound locally (same if running stand-alone).
                        if ( !moProfile.bValue("-CertificatePrivateKeyExportable", false)
                                || ( !moProfile.bValue("-UseStandAloneMode", true) && !this.oDomainProfile.bValue("-CertPrivateKeyExportAllowed", false)) )
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
                            DoGetCert.LogIt("");
                            DoGetCert.LogIt("The new certificate private key is exportable (\"-CertificatePrivateKeyExportable=True\"). This is not recommended.");

                            if ( lbCertificatePassword )
                                loNewCertificate = new X509Certificate2(lsCertPathFile, lsCertificatePassword
                                        , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
                            else
                                loNewCertificate = new X509Certificate2(lsCertPathFile, (string)null
                                        , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
                        }
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped && (null != lbtArrayNewCertificate || moProfile.bValue("-UseStandAloneMode", true)) )
                {
                    // Install and bind the new certificate locally, if: 1) the certificate was just downloaded from the repository; or 2) we're running stand-alone.
                    // Bottom line: for server farms, never get a new cert (from override file or certifcate network) and bind it during the same maintenance window.

                    DoGetCert.LogIt("");
                    DoGetCert.LogIt(String.Format("Certificate thumbprint: {0} (\"{1}\")", loNewCertificate.Thumbprint, DoGetCert.sCertName(loNewCertificate)));

                    DoGetCert.LogIt("");
                    DoGetCert.LogIt("Install and bind certificate ...");

                    string lsMachineKeyPathFile = null;

                    // Select the local machine certificate store (ie "Local Computer / Personal / Certificates").
                    loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                    // Add the new cert to the certifcate store.
                    loStore.Add(loNewCertificate);

                    // Do SSO stuff (while both old and new certs are still in the certifcate store).
                    if ( null != loOldCertificate && loOldCertificate.Thumbprint != loNewCertificate.Thumbprint )
                        lbGetCertificate = this.bApplyNewCertToSSO(loNewCertificate, loOldCertificate, lsHash, lbtArrayMinProfile);

                    if ( lbGetCertificate && !this.bMainLoopStopped )
                    {
                        if ( this.oDomainProfile.bValue("-NoIIS", false) )
                        {
                            DoGetCert.LogIt("");
                            DoGetCert.LogIt("Using non-IIS certificate binding ...");

                            string lsDiscard = null;
                            lbGetCertificate = DoGetCert.bRunPowerScript(out lsDiscard, moProfile.sValue("-NoIISBindingScript", ""), false, true);
                        }
                        else
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            DoGetCert.LogIt("");
                            DoGetCert.LogIt("Bind certificate in IIS ...");

                            // Apply new site-to-cert bindings.
                            foreach(string lsSanItem in DoGetCert.sSanArrayAmended(loServerManager, lsSanArray))
                            {
                                if ( !lbGetCertificate || this.bMainLoopStopped )
                                    break;

                                DoGetCert.LogIt(String.Format("Applying new certificate for \"{0}\" ...", lsSanItem));

                                Site    loSite = null;
                                        DoGetCert.oCurrentCertificate(lsSanItem, lsSanArray, out loServerManager, out loSite);

                                if ( null == loSite && lsSanItem == lsCertName )
                                {
                                    // No site found. Use the default site, if it exists (for the primary domain only).
                                    if ( 0 != loServerManager.Sites.Count )
                                        loSite = loServerManager.Sites[0];   // Default website.

                                    DoGetCert.LogIt(String.Format("No SAN website match could be found for \"{0}\". The default will be used (ie. \"{1}\").", lsSanItem, loSite.Name));
                                }

                                if ( null == loSite )
                                {
                                    // Still no site found (not even the default). This is an error.

                                    lbGetCertificate = false;

                                    DoGetCert.LogIt(String.Format("No website could be found to bind the new certificate for \"{0}\" (no default website either; FYI, -CreateSanSites is \"{1}\").\r\nCan't continue.", lsSanItem, lbCreateSanSites));
                                }
                                else
                                if ( null == DoGetCert.oSslBinding(loSite) )
                                {
                                    string  lsNewBindingMsg1 = "No SSL binding found.";
                                    string  lsNewBindingInformation = "*:443:{0}";
                                    Binding loNewBinding = null;

                                    try
                                    {
                                        if ( !lsSanArray.Contains(lsSanItem) )
                                        {
                                            // The site found is not SAN specific (ie. don't use SNI flag).
                                            lsNewBindingInformation = String.Format(lsNewBindingInformation, "");
                                            loNewBinding = loSite.Bindings.Add(lsNewBindingInformation, null == loOldCertificate ? null : loOldCertificate.GetCertHash(), loStore.Name);

                                            DoGetCert.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with no SNI) using \"{0}\".", lsNewBindingInformation));
                                        }
                                        else
                                        {
                                            // Set the SNI (Server Name Indication) flag.
                                            lsNewBindingInformation = String.Format(lsNewBindingInformation, lsSanItem);
                                            loNewBinding = loSite.Bindings.Add(lsNewBindingInformation, null == loOldCertificate ? null : loOldCertificate.GetCertHash(), loStore.Name);
                                            loNewBinding.SetAttributeValue("SslFlags", 1 /* SslFlags.Sni */);

                                            DoGetCert.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with SNI) using \"{0}\".", lsNewBindingInformation));
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        loSite.Bindings.Remove(loNewBinding);

                                        // Reverting to the default binding usage (must be an older OS).
                                        lsNewBindingInformation = "*:443:";
                                        loSite.Bindings.Add(lsNewBindingInformation, null == loOldCertificate ? null : loOldCertificate.GetCertHash(), loStore.Name);

                                        DoGetCert.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (on older OS) using \"{0}\".", lsNewBindingInformation));
                                    }

                                    loServerManager.CommitChanges();

                                    DoGetCert.LogIt(String.Format("Default SSL binding added to site: \"{0}\".", loSite.Name));
                                }

                                if ( lbGetCertificate && !this.bMainLoopStopped )
                                {
                                    bool lbBindingFound = false;

                                    foreach(Binding loBinding in loSite.Bindings)
                                    {
                                        if ( !lbGetCertificate || this.bMainLoopStopped )
                                            break;

                                        // Only try to bind the new certificate if a certificate binding already exists (ie. only replace what's already there).
                                        bool    lbDoBinding = null != loBinding.CertificateHash;
                                                if ( lbDoBinding )
                                                {
                                                    // Only replace an existing certificate binding if it matches the old certificate (if an old certificate exists).
                                                    lbDoBinding = null != loOldCertificate && loOldCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash);

                                                    // Finally, only replace an existing certificate binding if it matches the certificate used for communications.
                                                    if ( !lbDoBinding )
                                                        lbDoBinding = (null != goSetupCertificate && goSetupCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash));
                                                }
                                        if ( lbDoBinding )
                                        {
                                            string[]    lsBindingInfoArray  = loBinding.BindingInformation.Split(':');
                                            string      lsBindingAddress    = lsBindingInfoArray[0];
                                            string      lsBindingPort       = lsBindingInfoArray[1];
                                            string      lsBindingHost       = lsBindingInfoArray[2];

                                            loBinding.CertificateHash = loNewCertificate.GetCertHash();

                                            loServerManager.CommitChanges();

                                            if ( "*" == lsBindingAddress )
                                                DoGetCert.LogIt(String.Format("New certificate (\"{0}\") bound to port {1} in IIS for \"{2}\"."
                                                                , loNewCertificate.Thumbprint, lsBindingPort, lsSanItem));
                                            else
                                                DoGetCert.LogIt(String.Format("New certificate (\"{0}\") bound to address {1} and port {2} in IIS for \"{3}\"."
                                                                , loNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, lsSanItem));

                                            lbBindingFound = true;
                                        }
                                    }

                                    if ( !lbBindingFound )
                                    {
                                        // No binding was found matching the old certificate. Assign the new certificate to the default SSL binding.

                                        Binding loDefaultSslBinding = DoGetCert.oSslBinding(loSite);

                                        if ( null == loDefaultSslBinding )
                                        {
                                            // Still no binding found (not even the default). This is an error.

                                            lbGetCertificate = false;

                                            DoGetCert.LogIt(String.Format("No SSL binding could be found for \"{0}\" (no default SSL binding either).\r\nCan't continue.", lsSanItem));
                                        }
                                        else
                                        {
                                            if ( null != loDefaultSslBinding.CertificateHash && loNewCertificate.GetCertHash().SequenceEqual(loDefaultSslBinding.CertificateHash) )
                                            {
                                                DoGetCert.LogIt(String.Format("Binding already applied for  \"{0}\" (via \"{1}\").", lsSanItem, loSite.Name));
                                            }
                                            else
                                            {
                                                string[]    lsBindingInfoArray  = loDefaultSslBinding.BindingInformation.Split(':');
                                                string      lsBindingAddress    = lsBindingInfoArray[0];
                                                string      lsBindingPort       = lsBindingInfoArray[1];
                                                string      lsBindingHost       = lsBindingInfoArray[2];

                                                loDefaultSslBinding.CertificateHash = loNewCertificate.GetCertHash();

                                                loServerManager.CommitChanges();

                                                if ( "*" == lsBindingAddress )
                                                    DoGetCert.LogIt(String.Format("New certificate (\"{0}\") bound to port {1} in IIS for \"{2}\"."
                                                                    , loNewCertificate.Thumbprint, lsBindingPort, lsSanItem));
                                                else
                                                    DoGetCert.LogIt(String.Format("New certificate (\"{0}\") bound to address {1} and port {2} in IIS for \"{3}\"."
                                                                    , loNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, lsSanItem));
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if ( lbGetCertificate )
                        {
                            DoGetCert.LogSuccess();
                            DoGetCert.LogIt("");
                            DoGetCert.LogIt("A new certificate was successfully installed and bound.");
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && !this.bCertificateSetupDone )
                            this.bCertificateSetupDone = true;

                        if ( lbGetCertificate && !this.bMainLoopStopped && null != loOldCertificate )
                        {
                            // Next to last step, remove old certificate from the local store.

                            // Note: this MUST be done before the last step since multiple
                            //       identifying certificates can't exist during SCS calls.

                            if ( loOldCertificate.Thumbprint == loNewCertificate.Thumbprint )
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

                            if ( !moProfile.bValue("-UseStandAloneMode", true)
                                    || (moProfile.bValue("-UseStandAloneMode", true)
                                        && moProfile.bValue("-RemoveReplacedCert", false))
                                    )
                            {
                                if ( loNewCertificate.Thumbprint == loOldCertificate.Thumbprint )
                                {
                                    DoGetCert.LogIt(String.Format("New certificate (\"{0}\") same as old. Not removed from local store.", loOldCertificate.Thumbprint));
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

                                    DoGetCert.LogIt(String.Format("Old certificate (\"{0}\") removed from the local store.", loOldCertificate.Thumbprint));
                                }
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && !moProfile.bValue("-UseStandAloneMode", true) )
                        {
                            // At this point we need to load the new certificate into the service factory object.
                            goSetupCertificate = null;
                            moGetCertServiceFactory = new ChannelFactory<GetCertService.IGetCertServiceChannel>("WSHttpBinding_IGetCertService");
                            DoGetCert.SetCertificate(moGetCertServiceFactory);

                            DoGetCert.LogIt("Requesting removal of the old certificate from the repository.");

                            // Last step, remove old certificate from the repository (it can't be removed until the new certificate is in place everywhere).
                            if ( lbGetCertificate && !this.bMainLoopStopped )
                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                            {
                                bool    bOldCertificateRemoved = loGetCertServiceClient.bOldCertificateRemoved(lsHash, lbtArrayMinProfile);
                                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                            loGetCertServiceClient.Abort();
                                        else
                                            loGetCertServiceClient.Close();

                                if ( !bOldCertificateRemoved )
                                    DoGetCert.LogIt(DoGetCert.sExceptionMessage("GetCertService.bOldCertificateRemoved: Failed removing old certificate."));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DoGetCert.LogIt(DoGetCert.sExceptionMessage(ex));
                lbGetCertificate = false;
            }
            finally
            {
                if ( null != loStore )
                    loStore.Close();

                if ( "" == this.oDomainProfile.sValue("-CertOverridePfxComputer", "") + this.oDomainProfile.sValue("-CertOverridePfxPathFile", "") )
                {
                    // Remove the cert (ie. PFX file from cert provider).
                    File.Delete(lsCertPathFile);
                }
                else
                if ( lbGetCertificate && !this.bMainLoopStopped )
                {
                    // Remove the cert (ie. the override PFX file - only upon success).
                    File.Delete(lsCertPathFile);
                }

                DoGetCert.LogIt("");

                this.bRunPowerScript(moProfile.sValue("-ScriptSessionClose", @"
$session = Connect-PSSession -ComputerName localhost -Name GetCert
$session | Remove-PSSession
Get-WSManInstance -ResourceURI Shell -Enumerate
                        "), true);

                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = moGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.bUnlockCertificateRenewal(lsHash, lbtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }

            if ( this.bMainLoopStopped )
            {
                DoGetCert.LogIt("Stopped.");
                lbGetCertificate = false;
            }

            if ( lbGetCertificate )
            {
                moProfile.Remove("-CertOverrideErrorCount");
                moProfile.Remove("-LoadBalancerPendingErrorCount");
                moProfile.Remove("-RepositoryCertsOnlyErrorCount");
                moProfile.Remove("-CertificateRenewalDateOverride");
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = true;
                moProfile.Save();

                DoGetCert.LogIt("The get certificate process completed successfully.");
            }
            else
            {
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = false;
                moProfile.Save();

                DoGetCert.LogIt("At least one stage failed (or the process was stopped). Check log for errors.");
            }

            DoGetCert.LogIt("");

            return lbGetCertificate;
        }
    }
}
