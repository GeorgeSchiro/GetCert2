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
            DoGetCert       loMain    = null;
            Mutex           loMutex   = null;
            tvProfile       loProfile = null;

            try
            {
                loProfile = tvProfile.oGlobal(DoGetCert.oHandleUpdates(ref args, out loMutex));

                if ( !loProfile.bExit )
                {
                    loProfile.GetAdd("-Help",
                            @"
Introduction


This utility gets a digital certificate from the FREE ""Let's Encrypt"" certificate
provider network (see 'LetsEncrypt.org'). It installs the certificate in your
server's local computer certificate store and binds it to port 443 in IIS (it's 443
by default - any other port can be used instead).

If the current time is not within a given number of days prior to expiration
of the current digital certificate (eg. 30 days, see -RenewalDaysBeforeExpiration
below), this software does nothing. Otherwise, the retrieval process begins.

It's as simple as that when the software runs in 'stand-alone' mode (the default).

If 'stand-alone' mode is disabled (see -UseStandAloneMode below), the certificate
retrieval process is used in concert with the secure certificate service (SCS),
see 'SafeTrust.org'.

If the software is not running in 'stand-alone' mode, it also copies any new cert
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
usually located in the same folder as '{EXE}' and has the same name
with '.config' added (see '{INI}').

Profile file options can be overridden with command-line arguments. The
keys for any '-key=value' pairs passed on the command-line must match
those that appear in the profile (with the exception of the '-ini' key).

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

-AcmeAccountFile='Account.xml'

    This is the name of the ACME account file returned by the certificate
    provider. If it's deleted, it will be replaced during the next run.

-AcmeAccountKeyFile='AccountKey.xml'

    This is the name of the ACME account key file returned by the
    certificate provider. If it's deleted, it will be replaced during
    the next run.

-AcmePsModuleUseGallery=False

    Set this switch True and the 'PowerShell Gallery' version of 'ACME-PS'
    will be used in lieu of the version embedded in the EXE (see -AcmePsPath
    below).

-AcmePsPath='ACME-PS'

    'ACME-PS' is the tool used by this utility to communicate with the
    ""Let's Encrypt"" certificate network. By default, this key is set to the 'ACME-PS'
    folder which, with no absolute path given, will be expected to be found within
    the folder containing '{INI}'. Set -AcmePsModuleUseGallery=True
    (see above) and the OS will look to find 'ACME-PS' in its usual place as a
    module from the PowerShell gallery.

-AcmePsWorkPath='AcmeState'

    This is where temporary ACME state data is cached. It may be useful
    while troubleshooting issues. This data is automatically cleaned-up
    after each run, unless there is an error or unless -CleanupAcmeWork
    is set False (see below).

-AcmeSystemWide=''

    This is a script snippet that is prepended to every ACME script
    (see -ScriptStage1 below). This snippet may be useful to resolve
    certain issues that can arise during ACME script execution.

    Here's an example:

        -AcmeSystemWide='[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12'

        The above forces the use of the TLS 1.2 protocol with every ACME
        call. This is sometimes necessary when the network enforces rules
        that are not supported, by default, in older operating systems.

-AllowSsoThumbprintUpdatesAnytime=True

    Set this switch False to limit SSO thumbprint replacements to occur only
    during the domain defined maintenance window.

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

-CertificateRequestWaitSecs=10

    These are the seconds of wait time before a certificate request
    is submitted. This allows time for the order to be completed by
    the certificate provider network before a request can be accepted.

-CleanupAcmeWork=True

    Set this switch False to preserve temporary ACME state data, even after
    a successful run.

-ContactEmailAddress= NO DEFAULT VALUE

    This is the contact email address the certificate network uses to send
    certificate expiration notices.

-CreateSanSitesForCertGet=True

    If a SAN specific website does not yet exist in IIS, it will be created
    automatically during the first run of the 'get certificate' process for
    that SAN value. Set this switch False to have all SAN challenges routed
    through the IIS default website (such challenges will typically fail). If
    they do fail, you will need to create your SAN specific sites manually.

    Note: each SAN value must be challenged and therefore SAN challenges
          must be routed through your web server just like your primary domain.
          That means every SAN value must have a corresponding entry in the
          global internet DNS database.

          When a new website is created in IIS for a new SAN value, by default
          it is setup to use the same physical path as the primary domain.

-CreateSanSitesForBinding=False

    Set this switch True to have SAN sites created (if they don't already exist)
    during the certificate binding phase. The primary certificate name on the
    SAN list (see -CertificateDomainName above) will be bound to the default
    website as a last resort (if this value is False).

-DefaultPhysicalPath='%SystemDrive%\inetpub\wwwroot'

    This is the default physical storage path used during the creation of
    new IIS websites that can't otherwise be associated with an existing
    site (see -CreateSanSitesFor* above).

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

    This is the maximum number of seconds given to a process after a 'close'
    command is given before the process is forcibly terminated.

-LoadBalancerReleaseCert

    This switch indicates the new certificate has been released by the load
    balancer administrator or process.

    This switch would typically never appear in a profile file. It's meant to be
    used on the command-line only (in a script or a shortcut).

    Note: this switch is ignored when -UseStandAloneMode is True.

-LogCertificateStatus=False

    Set this switch True to see in detail how the {EXE} app's identifying
    certificate is selected. This may be useful to troubleshoot setup issues.

    Note: this switch is ignored when -UseStandAloneMode is True.

-LogEntryDateTimeFormatPrefix='yyyy-MM-dd hh:mm:ss:fff tt  '

    This format string is used to prepend a timestamp prefix to each log entry
    in the process log file (see -LogPathFile below).    

-LogFileDateFormat='-yyyy-MM-dd'

    This format string is used to form the variable part of each log file output 
    filename (see -LogPathFile below). It is inserted between the filename and 
    the extension.

-LogPathFile='Logs\Log.txt'

    This is the output path\file that will contain the process log. The profile
    filename will be prepended to the default filename (see above) and the current
    date (see -LogFileDateFormat above) will be inserted between the filename and
    the extension.

-MaxCertRenewalLockDelaySecs=300

    Wait a random period each cycle (at most this many seconds) to allow different
    clients the opportunity to lock the certificate renewal (ie. only one client
    at a time per domain can communicate with the certificate provider network).

-NonIISBindingScript= SEE PROFILE FOR DEFAULT VALUE

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

-RegexDnsNamePrimary='^([a-zA-Z0-9\-]+\.)+[a-zA-Z]{2,18}$'

    This regular expression is used to validate -CertificateDomainName (see above).

-RegexDnsNameSanList='^([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)([a-zA-Z]{2,18}|)$'

    This regular expression is used to validate -SanList names (see below).

    Note: -RegexDnsNameSanList and -RegexDnsNamePrimary are both used to validate
          SAN list names. A match of either pattern will pass validation.

-RegexEmailAddress='^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9\-]+\.)*([a-zA-Z0-9\-]+\.)[a-zA-Z]{2,18}$'

    This regular expression is used to validate -ContactEmailAddress (see above).

-RemoveReplacedCert=False

    Set this switch True and the old (ie. previously bound) certificate will be
    removed whenever a newly retrieved certificate is bound to replace it.

    Note: this switch is ignored when -UseStandAloneMode is False.

-RenewalDateOverride= NO DEFAULT VALUE

    Set this date value to override the date calculation that subtracts
    -RenewalDaysBeforeExpiration days (see below) from the current certificate
    expiration date to know when to start fetching a new certificate.

    Note: this parameter is ignored when -UseStandAloneMode is False. It will
          be removed from the profile after a certificate has been successfully
          retrieved from the certificate provider network.

-RenewalDaysBeforeExpiration=30

    This is the number of days until certificate expiration before automated
    gets of the next new certificate are attempted.

-ResetStagingLogs=True

    Having to wade through several log sessions during testing can be cumbersome.
    So the default behavior is to clear the log file after it is uploaded to the SCS
    server (following each test). Setting this switch False will retain all previous
    log sessions on the client during testing.

    Note: this switch is ignored when -UseStandAloneMode is True.

-RetainNewCertAfterError=False

    Set this switch True to prevent removal of new certificates after errors. A new
    certificate should typically not persist after a failure. This prevents multiple
    versions of essentially the 'same' certificate from piling up.

    Note: this parameter will be removed from the profile after every run
          (whether used or not). If you need several uses, it might make more
          sense to pass this parameter as a command-line argument instead.

-SanList= SEE PROFILE FOR DEFAULT VALUE

    This is the SAN list (subject alternative names) to appear on the certificate
    when it is generated. It will consist of -CertificateDomainName by default (see
    above). This list can be edited here directly or through the 'SAN List' button
    in the {EXE} UI. Click the 'SAN List' button to see the proper format here.

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

    Set this switch False to allow merged command-lines to be written to
    the profile file (ie. '{INI}'). When True, everything
    but command-line keys will be saved.

-ScriptBindingDone=''

    This is the PowerShell script snippet that handles any final processing after
    the successful binding of a new certificate.

-ScriptCertAcquired=''

    This is the PowerShell script snippet that handles any final processing after
    the successful acquisition of a new certificate from the certificate provider.

-ScriptSSO= SEE PROFILE FOR DEFAULT VALUE

    This is the PowerShell script that updates SSO servers with new certificates.

-ScriptStage1= SEE PROFILE FOR DEFAULT VALUE

    There are multiple stages involved with the process of getting a certificate
    from the certificate provider network. Each stage has an associated PowerShell
    script snippet. The stages are represented in this profile by -ScriptStage1
    through -ScriptStage7.

-ServiceNameLive='LetsEncrypt'

    This is the name mapped to the live production certificate network service URL.

-ServiceNameStaging='LetsEncrypt-Staging'

    This is the name mapped to the non-production (ie. 'staging') certificate
    network service URL.

-SetDefaultBinding=True

    Set this switch False to prevent even a default binding from being applied
    after no binding is found referencing the 'old' certificate.

-Setup=False

    Set this switch True to ignore the usual constraints during the acquisition
    of new certificates. For example, you can get a new certificate even if the
    old one is not about to expire, like when the software is run in interactive
    mode (eg. during initial setup).

    This switch is typically passed as a command-line argument:

        {EXE} -Auto -Setup

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the profile
    file at startup in command-line format. This may be helpful as a diagnostic.

-SingleSessionEnabled=False

    Set this switch True to run all PowerShell scripts in a single global session.

-SkipPreviousStore=False

    The 'GetCertPrevious' certificate store is used to house the previously bound
    'old' certificate (ie. one most recently removed from the 'Personal' store).
    Set this switch False to simply delete old certificates directly.

-SkipSsoServer=False

    When a client is on an SSO domain (ie. it's a member of an SSO server farm),
    it will automatically attempt to update SSO services with each new certificate
    retrieved. Set this switch True to disable SSO updates (for this client only).

-SkipSsoThumbprintUpdates=False

    When a client's domain references an SSO domain, the client will automatically
    attempt to update configuration files with each new SSO certificate thumbprint.
    Set this switch True to disable SSO thumbprint configuration updates (for this
    client only). See {SsoKey} below.

-SsoMaxRenewalSleepSecs=60

    This is the maximum seconds an SSO server will wait for another
    SSO server (in the same farm) to renew the SSO certificates.

-SsoProxySleepSecs=60

    An SSO proxy server will wait this many seconds before
    polling the SSO server again for a certificate change.

-SsoProxyTimeoutMins=30

    An SSO proxy server will wait this many minutes for the SSO server
    to change its certificates before throwing a timeout exception.

{SsoKey}='C:\inetpub\wwwroot\web.config'

    This is the path and filename of files that will have their SSO certificate
    thumbprint replaced whenever the related SSO certificate changes. Each file
    with the same name at all levels of the directory hierarchy will be updated,
    starting with the given base path, if the old SSO certificate thumbprint is
    found there. See -SkipSsoThumbprintUpdates above.

    Note: This key may appear any number of times in the profile and wildcards
          can be used in the filename.

-SsoThumbprintReplacementArgs=''

    Whatever is provided here will be passed to the SSO thumbprint replacement
    sub-process during each run. This makes it easier to adjust the SSO thumbprint
    replacement sub-process without having to edit its configuration separately.

-SubmissionRetries=42

    Pending submissions to the certificate provider network will be retried until
    they succeed or fail, at most this many times. By default, the  process will
    retry for 7 minutes (-SubmissionRetries times -SubmissionWaitSecs, see below)
    for challenge status updates as well as certificate request status updates.

-SubmissionWaitSecs=5

    These are the seconds of wait time after an ACME challenge request has been
    submitted to the certificate network as well as after a certificate request
    has been submitted. This is the amount of time during which the request should
    transition from a 'pending' state to anything other than 'pending'.

-UseNonIISBindingAlso=False

    Set this switch True to use the typical IIS binding procedures
    on this machine as well as the -NonIISBindingScript (see above).

-UseNonIISBindingOnly=False

    Set this switch True to disable the usual IIS binding procedures on
    this machine and use the -NonIISBindingScript instead (see above).

-UseStandAloneMode=True

    Set this switch False and the software will use the SafeTrust Secure Certificate
    Service (see 'SafeTrust.org') to manage certificates between several servers
    in a server farm, on SSO servers, SSO integrated application servers and load
    balancers.


Notes:

    There may be various other settings that can be adjusted also (user
    interface settings, etc). See the profile file ('{INI}')
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added 'on the fly'
    (in order of execution) to '{INI}' as the software runs.

"
                            .Replace("{EXE}", Path.GetFileName(ResourceAssembly.Location))
                            .Replace("{INI}", Path.GetFileName(loProfile.sActualPathFile))
                            .Replace("{SsoKey}", gsSsoThumbprintFilesKey)
                            .Replace("{{", "{")
                            .Replace("}}", "}")
                            );

                    string lsFetchName = null;

                    // Fetch simple setup.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="Setup Application Folder.exe")
                            , loProfile.sRelativeToExePathFile(lsFetchName));

                    // Fetch source code.
                    if ( loProfile.bValue("-FetchSource", false) )
                        tvFetchResource.ToDisk(lsFetchName=System.Windows.Application.ResourceAssembly.GetName().Name, lsFetchName + ".zip", null);

                    // Dependency updates start here.
                    if ( loProfile.bFileJustCreated )
                    {
                        loProfile["-Update2022-10-15"] = true;
                        loProfile.Save();
                    }
                    else
                    {
                        if ( !loProfile.bValue("-Update2022-10-15", false) )
                        {
                            // Force use of the latest version of the replacement module.
                            string  lsPath = tvProfile.oGlobal().sRelativeToExePathFile("ReplaceText");
                            if ( Directory.Exists(lsPath) )
                                Directory.Delete(lsPath, true);

                            loProfile["-Update2022-10-15"] = true;
                            loProfile.Save();
                        }
                    }
                    // Dependency updates end here.

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
                                    loUI.oStartupWaitMsg = goStartupWaitMsg;
                                    loMain.oUI = loUI;
                                    loMain.MainWindow = loUI;
                                    loMain.ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose;

                                loMain.Run(loUI);
                            }
                            catch (ObjectDisposedException) {}

                            GC.KeepAlive(loMutex);
                        }
                        else
                        {
                            // Run in batch mode.

                            DoGetCert   loDoDa = new DoGetCert(loProfile);
                                        if ( null != goStartupWaitMsg )
                                            goStartupWaitMsg.Close();

                            tvProfile   loMinProfile = Env.oMinProfile(loProfile);
                            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                            string      lsHash = HashClass.sHashIt(loMinProfile);
                            string      lsLogFileTextReported = null;
                            bool        lbGetClientCertificates = loDoDa.bGetClientCertificates(lsHash, lbtArrayMinProfile);
                            bool        lbReplaceSsoThumbprint = loDoDa.bReplaceSsoThumbprint(lsHash, lbtArrayMinProfile);
                            bool        lbGetCertificate = loDoDa.bGetCertificate();

                                        // Reinitialize the communications channel in case the certificate was just replaced.
                                        Env.oGetCertServiceFactory = null;

                            if ( !lbGetCertificate || !lbGetClientCertificates || !lbReplaceSsoThumbprint )
                            {
                                Environment.ExitCode = 1;
                                DoGetCert.ReportErrors(out lsLogFileTextReported);
                            }
                            else
                            {
                                DoGetCert.ReportEverything(out lsLogFileTextReported);
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
                Environment.ExitCode = 1;
                Env.LogIt(Env.sExceptionMessage(ex));

                try
                {
                    string  lsLogFileTextReported = null;

                    DoGetCert.ReportErrors(ex, out lsLogFileTextReported);
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

        public bool bCertificateSetupDone
        {
            get
            {
                return moProfile.bValue("-CertificateSetupDone", false);
            }
            set
            {
                if ( value )
                    moProfile["-LogCertificateStatus"] = false;

                moProfile["-CertificateSetupDone"] = value;
                moProfile.Save();
            }
        }


        private static tvMessageBox goStartupWaitMsg        = null;
        private static string gsAcmeAccountFileKey          = "-AcmeAccountFile";
        private static string gsAcmeAccountKeyFileKey       = "-AcmeAccountKeyFile";
        private static string gsBindingFailedKey            = "-BindingFailed";
        private static string gsCertificateKeyPrefix        = "-Cert_";
        private static string gsCertificateHashKey          = "-Hash";
        private static string gsCertStorePreviousName       = "GetCertPrevious";
        private static string gsContentKey                  = "-Content";
        private static string gsDashboardDataFile           = "Data.txt";
        private static string gsDashboardVersionFile        = "Version.txt";
        private static string gsDashboardVersionKey         = "-DashboardVersion";
        private static string gsDashboardZipFile            = "ggc.zip";
        private static string gsInAndOut1LbReleaseCertKey   = "-LoadBalancerReleaseCert";
        private static string gsInAndOut2ServiceCallKey     = "-SCS";
        private static string gsMaintWindBegKey             = "-MaintenanceWindowBeginTime";
        private static string gsMaintWindEndKey             = "-MaintenanceWindowEndTime";
        private static string gsSsoThumbprintFilesKey       = "-SsoThumbprintFiles";
        private static string gsUpdateProfileServerKey      = "-ServerUpdate";
        private static string gsUpdateProfileServerFmt      = "{0}={1}";


        /// <summary>
        /// This switch is used within the main loop of the UI (see "this.oUI")
        /// to stop any loops running within this object as well.
        /// </summary>
        public bool bMainLoopStopped
        {
            get
            {
                return Env.bMainLoopStopped;
            }
            set
            {
                Env.bMainLoopStopped = value;
            }
        }

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
        public static tvProfile oDomainProfile
        {
            get
            {
                if ( null == goDomainProfile )
                {
                    if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                    {
                        goDomainProfile = new tvProfile();
                    }
                    else
                    {
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            tvProfile   loMinProfile = Env.oMinProfile(tvProfile.oGlobal());
                            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                            string      lsHash = HashClass.sHashIt(loMinProfile);

                            goDomainProfile = new tvProfile(loGetCertServiceClient.sDomainProfile(lsHash, lbtArrayMinProfile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            if ( 0 == goDomainProfile.Count )
                                throw new Exception("The domain profile is empty. Can't continue.");
                        }
                    }
                }

                return goDomainProfile;
            }
        }
        private static tvProfile goDomainProfile;

        /// <summary>
        /// This is the profile preserved keys (keys retained after a configuration reset).
        /// </summary>
        public static tvProfile oPreservedKeys
        {
            get
            {
                if ( null == goPreservedKeys )
                {
                    goPreservedKeys = new tvProfile(tvProfile.oGlobal().sValue("-KeysAfterReset", @"
        -AcmeSystemWide
        -AllConfigWizardStepsCompleted
        -CertificateDomainName
        -CertificateSetupDone
        -CfgVersion
        -ContactEmailAddress
        -HostsEntryVersion
        -HostIniVersion
        -KeysAfterReset
        -LicenseAccepted
        -NonIISBindingScript
        -SanList
        -ScriptBindingDone
        -ScriptCertAcquired
        -ScriptSSO
        -SkipSsoServer
        -SkipSsoThumbprintUpdates
        -SsoThumbprintFiles
        -SsoThumbprintReplacementArgs
        -UseNonIISBindingAlso
        -UseNonIISBindingOnly
        -XML_Profile
        "));
                }

                return goPreservedKeys;
            }
        }
        private static tvProfile goPreservedKeys;

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

        public bool bGetClientCertificates(string asHash, byte[] abtArrayMinProfile)
        {
            bool lbGetClientCertificates = false;

            // Do nothing if the domain doesn't use client certificates.
            if ( !DoGetCert.oDomainProfile.bValue("-GetClientCertificates", false) )
                return true;

            Env.LogIt("");
            Env.LogIt("Checking for client certificate changes ...");

            UInt64  liPreviousCertificatesListHashUInt64 = UInt64.Parse(moProfile.sValue("-PreviousCertificatesListHashUInt64", "0"));
            byte[]  lbtArrayCertificatesProfile = null;

            // Get client certificates.
            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                lbtArrayCertificatesProfile = loGetCertServiceClient.btArrayClientCertificateList(asHash, abtArrayMinProfile, liPreviousCertificatesListHashUInt64);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( null == lbtArrayCertificatesProfile )
            {
                // No certificate changes.
                Env.LogIt("No client certificate changes found.");
                return true;
            }
            else
            {
                tvProfile loClientCertificatesWithHash = tvProfile.oProfile(lbtArrayCertificatesProfile);
                tvProfile loCertsAddedRemoved = new tvProfile();
                X509Store loStore = null;

                try
                {
                    loStore = new X509Store(StoreName.TrustedPeople, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                    tvProfile   loClientCertificates = loClientCertificatesWithHash.oOneKeyProfile(gsCertificateKeyPrefix + "*", true);
                    string      lsCertDetails = null;

                    foreach (X509Certificate2 loCertificate in loStore.Certificates)
                    {
                        if ( loClientCertificates.ContainsKey(loCertificate.Thumbprint) )
                        {
                            // The certificate's already in the store, remove it from the download list.
                            loClientCertificates.Remove(loCertificate.Thumbprint);
                        }
                        else
                        if ( !moProfile.bValue("-OnlyAddClientCertificates", false) )
                        {
                            // The certificate is not in the download list, remove it from the store.
                            loStore.Remove(loCertificate);

                            lsCertDetails = String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                    , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint);

                            loCertsAddedRemoved.Add("-Removed", lsCertDetails);

                            Env.LogIt(String.Format("    Removed: \"{0}\"", lsCertDetails));
                        }
                    }

                    foreach (DictionaryEntry loEntry in loClientCertificates)
                    {
                        // Add new certificate to the store.
                        using (X509Certificate2 loCertificate = new X509Certificate2())
                        {
                            loCertificate.Import(Convert.FromBase64String(new tvProfile(loEntry.Value.ToString()).sValue("-PublicKey", "")));

                            loStore.Add(loCertificate);

                            lsCertDetails = String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                    , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint);

                            loCertsAddedRemoved.Add("-Added", lsCertDetails);

                            Env.LogIt(String.Format("    Added: \"{0}\"", lsCertDetails));
                        }
                    }

                    foreach (X509Certificate2 loCertificate in loStore.Certificates)
                    {
                        loCertsAddedRemoved.Add("-All", String.Format("-Name='{0}' -NotAfter='{1}' -Thumbprint={2}"
                                , Env.sCertName(loCertificate), loCertificate.NotAfter, loCertificate.Thumbprint));
                    }
                }
                catch (Exception ex)
                {
                    Environment.ExitCode = 1;
                    Env.LogIt(Env.sExceptionMessage(ex));

                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                    {
                        loGetCertServiceClient.NotifyClientCertificateListFailure(asHash, abtArrayMinProfile, loCertsAddedRemoved.btArrayZipped());
                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                            loGetCertServiceClient.Abort();
                        else
                            loGetCertServiceClient.Close();
                    }
                }
                finally
                {
                    if ( null != loStore )
                        loStore.Close();
                }

                moProfile["PreviousCertificatesListHashUInt64"] = loClientCertificatesWithHash.sValue(gsCertificateHashKey, "0");
                moProfile.Save();

                loCertsAddedRemoved.SortByValueAsString();

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyClientCertificateListInstalled(asHash, abtArrayMinProfile, loCertsAddedRemoved.btArrayZipped());
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

                Env.LogSuccess();

                lbGetClientCertificates = true;
            }

            return lbGetClientCertificates;
        }

        public bool bReplaceSsoThumbprint(string asHash, byte[] abtArrayMinProfile)
        {
            bool lbReplaceSsoThumbprint = false;

            // Do nothing if this is an SSO server.
            if ( DoGetCert.oDomainProfile.bValue("-IsSsoDomain", false) )
                return true;

            // Do nothing if no SSO domain is defined for the current domain.
            string  lsSsoDnsName = DoGetCert.oDomainProfile.sValue("-SsoDnsName", "");
                    if ( "" == lsSsoDnsName )
                        return true;

            Env.LogIt("");
            Env.LogIt(String.Format("Checking for SSO thumbprint change (\"{0}\") ...", lsSsoDnsName));
            
            if ( !moProfile.bValue("-AllowSsoThumbprintUpdatesAnytime", true) && moProfile.bValue("-Auto", false) && !moProfile.bValue("-Setup", false) )
            {
                DateTime    ldtMaintenanceWindowBeginTime = DoGetCert.oDomainProfile.dtValue(gsMaintWindBegKey, DateTime.MinValue);
                DateTime    ldtMaintenanceWindowEndTime   = DoGetCert.oDomainProfile.dtValue(gsMaintWindEndKey, DateTime.MaxValue);
                            if ( DateTime.Now < ldtMaintenanceWindowBeginTime || DateTime.Now >= ldtMaintenanceWindowEndTime )
                            {
                                Env.LogIt(String.Format("    Can't continue. The \"{0}\" maintenance window is \"{1}\" to \"{2}\"."
                                                        , moProfile.sValue("-CertificateDomainName" ,""), ldtMaintenanceWindowBeginTime, ldtMaintenanceWindowEndTime));

                                return true;
                            }
            }

            // At this point it is an error if no SSO thumbprint exists for the current domain.
            string      lsSsoThumbprint = DoGetCert.oDomainProfile.sValue("-SsoThumbprint", "");
                        if ( "" == lsSsoThumbprint )
                        {
                            Env.LogIt("");
                            throw new Exception(String.Format("The SSO certificate thumbprint for the \"{0}\" domain has not yet been set!", lsSsoDnsName));
                        }
            string      lsSsoPreviousThumbprint = moProfile.sValue("-SsoThumbprint", "");
                        if ( !moProfile.ContainsKey(gsSsoThumbprintFilesKey) )
                        {
                            string  lsDefaultPhysicalPathFiles = Path.Combine(moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot"), "web.config");
                            Site[]  loSetupCertBoundSiteArray = DoGetCert.oSetupCertBoundSiteArray(new ServerManager());
                                    moProfile.Remove(gsSsoThumbprintFilesKey + "Note");
                                    if ( null == loSetupCertBoundSiteArray || 0 == loSetupCertBoundSiteArray.Length )
                                    {
                                        Env.LogIt(String.Format("{0} will be set to the IIS default since no sites could be found bound to the current certificate.", gsSsoThumbprintFilesKey));

                                        if ( null != loSetupCertBoundSiteArray )
                                        {
                                            moProfile.Add(gsSsoThumbprintFilesKey, Environment.ExpandEnvironmentVariables(lsDefaultPhysicalPathFiles));
                                            moProfile.Add(gsSsoThumbprintFilesKey + "Note", "No IIS websites could be found bound to the current certificate.");
                                        }
                                        else
                                        {
                                            Env.LogIt("");
                                            throw new Exception("The current setup certificate is null. This is an error.");
                                        }
                                    }
                                    else
                                    {
                                        foreach (Site loSite in loSetupCertBoundSiteArray)
                                            moProfile.Add(gsSsoThumbprintFilesKey
                                                    , Path.Combine(Environment.ExpandEnvironmentVariables(loSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                                                    , Path.GetFileName(lsDefaultPhysicalPathFiles)));
                                    }
                                    moProfile.Save();
                        }
                        if ( String.IsNullOrEmpty(lsSsoPreviousThumbprint) )
                        {
                            // Define the previous SSO thumbprint (only if this is the first time through).
                            moProfile["-SsoThumbprint"] = lsSsoThumbprint;
                            moProfile.Save();

                            Env.LogIt(String.Format("No new SSO thumbprint found (caching old: \"{0}\").", lsSsoThumbprint));

                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
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
                            Env.LogIt("No new SSO thumbprint found.");
                            return true;
                        }

            if ( moProfile.bValue("-SkipSsoThumbprintUpdates", false) )
            {
                Env.LogIt("Noting SSO thumbprint change, but skipping SSO thumbprint updates on this server.");

                lbReplaceSsoThumbprint = true;
            }
            else
            {
                tvProfile   loArgs = new tvProfile();
                            loArgs.Add("-OldSubValue", lsSsoPreviousThumbprint);
                            loArgs.Add("-OldSubValue", lsSsoThumbprint);
                            foreach (DictionaryEntry loEntry in moProfile.oOneKeyProfile(gsSsoThumbprintFilesKey))
                            {
                                // Change gsSsoThumbprintFilesKey to "-Files".
                                loArgs.Add("-Files", loEntry.Value);
                            }
                            loArgs.LoadFromCommandLine(moProfile.sValue("-SsoThumbprintReplacementArgs", ""), tvProfileLoadActions.Merge);

                Env.LogIt("");
                Env.LogIt(String.Format("Replacing SSO thumbprint (\"{0}\") ...", lsSsoDnsName));

                lbReplaceSsoThumbprint = Env.bReplaceInFiles(ResourceAssembly.GetName().Name, loArgs);
            }


            if ( !lbReplaceSsoThumbprint )
            {
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifySsoThumbprintReplacementFailure(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }
            else
            {
                moProfile["-SsoThumbprint"] = lsSsoThumbprint;
                moProfile.Save();

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifySsoThumbprintReplacementSuccess(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            }

            return lbReplaceSsoThumbprint;
        }

        public bool bUploadCertToLoadBalancer(string asHash, byte[] abtArrayMinProfile, string asCertName, string asCertPfxPathFile, string asCertificatePassword)
        {
            if (       DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "") != Env.sComputerName
                    || DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxPending", false)
                    || DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxReady", false)
                    )
                return true;

            bool    lbUploadCertToLoadBalancer = false;
            string  lsLoadBalancerPfxPathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "");
            string  lsLoadBalancerPassword = HashClass.sDecrypted(Env.oChannelCertificate, DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPassword", ""));

            Env.LogIt("");
            Env.LogIt("Checking for load balancer ...");

            // Do nothing if a load balancer PFX file location is not defined.
            if ( String.IsNullOrEmpty(lsLoadBalancerPfxPathFile) )
            {
                Env.LogIt("    No load balancer file location has been defined!");
                return false;
            }

            if ( DateTime.Now < DoGetCert.oDomainProfile.dtValue("-RenewalCertificateBindingDate", DateTime.MinValue).AddDays(-1) )
            {
                Env.LogIt("    The load balancer can't accept the new certificate until the day before certificate binding.");
                return true;
            }

            lsLoadBalancerPfxPathFile = moProfile.sRelativeToProfilePathFile(lsLoadBalancerPfxPathFile.Replace("*", Path.GetFileNameWithoutExtension(asCertPfxPathFile)));

            if ( lsLoadBalancerPfxPathFile == asCertPfxPathFile )
                lsLoadBalancerPfxPathFile += "2";

            Env.LogIt(String.Format("Copying new certificate for use on the \"{0}\" load balancer ...", asCertName));

            if ( DoGetCert.oDomainProfile.bValue("-LoadBalancerPemFormat", false) || ".pem" == Path.GetExtension(lsLoadBalancerPfxPathFile).ToLower() )
            {
                lbUploadCertToLoadBalancer = Env.bPfxToPem(
                        ResourceAssembly.GetName().Name, asCertPfxPathFile, asCertificatePassword, lsLoadBalancerPfxPathFile, lsLoadBalancerPassword);

                if ( lbUploadCertToLoadBalancer )
                    Env.LogSuccess();
            }
            else
            {
                using (X509Certificate2 loLbCertificate = new X509Certificate2(asCertPfxPathFile, asCertificatePassword
                        , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(lsLoadBalancerPfxPathFile));
                    File.WriteAllBytes(lsLoadBalancerPfxPathFile, loLbCertificate.Export(X509ContentType.Pfx, lsLoadBalancerPassword));
                }

                lbUploadCertToLoadBalancer = true;

                Env.LogSuccess();
            }

            string  lsLoadBalancerExePathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerExePathFile", "");
                    if (  "" == lsLoadBalancerExePathFile && lbUploadCertToLoadBalancer && !this.bMainLoopStopped )
                    {
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            loGetCertServiceClient.NotifyLoadBalancerCertificatePending(asHash, abtArrayMinProfile);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

                        return true;
                    }

            if ( !lbUploadCertToLoadBalancer )
                return false;

            if ( "*" != lsLoadBalancerExePathFile )
            {
                lbUploadCertToLoadBalancer = this.bUploadCertToLoadBalancerEXE(lsLoadBalancerExePathFile, lsLoadBalancerPfxPathFile, asCertName);
                if ( lbUploadCertToLoadBalancer )
                    Env.LogSuccess();
            }
            else
            {
                Env.LogIt("");
                Env.LogIt("Uploading certificate to load balancer directly ...");

                lbUploadCertToLoadBalancer = this.bRunPowerScript(moProfile.sValue("-ScriptLoadBalancer", @"

echo ""Command-line to send new certificate PFX file ('{LoadBalancerPfxPathFile}') for domain ('{CertificateDomainName}') to the load balancer.""
                        ")
                        .Replace("{LoadBalancerPfxPathFile}", lsLoadBalancerPfxPathFile)
                        .Replace("{CertificateDomainName}", asCertName)
                        );
            }

            if ( lbUploadCertToLoadBalancer )
                File.Delete(lsLoadBalancerPfxPathFile);

            if ( lbUploadCertToLoadBalancer && !this.bMainLoopStopped )
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateExeSuccess(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }
            else
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateExeFailure(asHash, abtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

            return lbUploadCertToLoadBalancer;
        }

        private bool bUploadCertToLoadBalancerEXE(string asLoadBalancerExePathFile, string asLoadBalancerPfxPathFile, string asCertName)
        {
            bool    lbUploadCertToLoadBalancerEXE = false;

            try
            {
                Process     loProcess = new Process();
                            loProcess.ErrorDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessErrorHandler);
                            loProcess.OutputDataReceived += new DataReceivedEventHandler(Env.PowerScriptProcessOutputHandler);

                            if ( ".ps1" != Path.GetExtension(asLoadBalancerExePathFile).ToLower() )
                            {
                                loProcess.StartInfo.FileName = asLoadBalancerExePathFile;
                                loProcess.StartInfo.Arguments = asLoadBalancerPfxPathFile + " " + asCertName;
                            }
                            else
                            {
                                loProcess.StartInfo.FileName = moProfile.sValue("-PowerShellExePathFile", @"C:\Windows\System32\WindowsPowerShell\v1.0\PowerShell.exe");
                                loProcess.StartInfo.Arguments = String.Format(moProfile.sValue("-PowerShellExeArgsLB", @"-NoProfile -ExecutionPolicy unrestricted -File ""{0}"" ""{1}"" ""{2}""")
                                                                                , asLoadBalancerExePathFile, asLoadBalancerPfxPathFile, asCertName);
                            }

                            loProcess.StartInfo.WorkingDirectory = Path.GetDirectoryName(asLoadBalancerExePathFile);
                            loProcess.StartInfo.UseShellExecute = false;
                            loProcess.StartInfo.RedirectStandardError = true;
                            loProcess.StartInfo.RedirectStandardInput = true;
                            loProcess.StartInfo.RedirectStandardOutput = true;
                            loProcess.StartInfo.CreateNoWindow = true;

                System.Windows.Forms.Application.DoEvents();

                Env.bPowerScriptError = false;
                Env.sPowerScriptOutput = null;

                if ( !this.bMainLoopStopped )
                {
                    Env.LogIt("");
                    Env.LogIt("Uploading certificate to load balancer directly ...");

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
                        Env.LogIt(moProfile.sValue("-ReplaceLBcertTimeoutMsg", "*** load balancer certificate replacement sub-process timed-out ***\r\n\r\n"));

                    int liExitCode = -1;
                        try { liExitCode = loProcess.ExitCode; } catch {}
                }

                loProcess.CancelErrorRead();
                loProcess.CancelOutputRead();

                Env.bKillProcess(loProcess);

                lbUploadCertToLoadBalancerEXE = true;
            }
            catch
            {
                Environment.ExitCode = 1;
            }

            return lbUploadCertToLoadBalancerEXE;
        }

        public static void ClearCache()
        {
            // Clear the domain profile cache.
            goDomainProfile = null;
        }

        public bool bDoInAndOut()
        {
            if ( this.bLoadBalancerReleaseCert() )
                return true;
            else
            if ( this.bCallAnySCS() )
                return true;
            else
                return false;
        }

        public static bool bDoInAndOutContainsKey()
        {
            if ( tvProfile.oGlobal().ContainsKey(gsInAndOut1LbReleaseCertKey)
                    && tvProfile.oGlobal().bValue(gsInAndOut1LbReleaseCertKey, false) )
                return true;
            else
            if ( tvProfile.oGlobal().ContainsKey(gsInAndOut2ServiceCallKey)
                    && "" != tvProfile.oGlobal().sValue(gsInAndOut2ServiceCallKey, "") )
                return true;
            else
                return false;
        }

        public bool bLoadBalancerReleaseCert()
        {
            // This is a command-line only switch. If it's not already there, don't add it.
            if ( !moProfile.ContainsKey(gsInAndOut1LbReleaseCertKey) || !moProfile.bValue(gsInAndOut1LbReleaseCertKey, false) )
                return false;

            if ( moProfile.bValue("-UseStandAloneMode", true) )
                return true;

            bool        lbNoPromptsBackup = moProfile.bValue("-NoPrompts", false);
            tvProfile   loCmdLineProfile = new tvProfile();
                        loCmdLineProfile.LoadFromCommandLineArray(Environment.GetCommandLineArgs(), tvProfileLoadActions.Overwrite);
                        moProfile["-NoPrompts"] = loCmdLineProfile.bValue("-NoPrompts",  false);
            string      lsDnsName = DoGetCert.oDomainProfile.sValue("-DnsName", "-DnsName not defined");
            string      lsCaption = String.Format("Release \"{0}\" Cert", lsDnsName);

            if ( DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxReady", false) )
            {
                this.Show(String.Format("The \"{0}\" load balancer certificate has already been released!", lsDnsName)
                            , lsCaption, tvMessageBoxButtons.OK, tvMessageBoxIcons.Alert, "LoadBalancerCertReleasedAlready");

                moProfile["-NoPrompts"] = lbNoPromptsBackup;
                return true;
            }

            Env.LogIt("");
            Env.LogIt(String.Format("The load balancer admin asserts the new \"{0}\" certificate is ready for general release ...", lsDnsName));

            string      lsLoadBalancerComputerName = DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "-LoadBalancerComputerName not defined");
            string      lsLoadBalancerPfxPathFile = DoGetCert.oDomainProfile.sValue("-LoadBalancerPfxPathFile", "-LoadBalancerPfxPathFile not defined");
            bool        lbWrongServer = "" != lsLoadBalancerComputerName && lsLoadBalancerComputerName != Env.sComputerName;
                        if ( lbWrongServer )
                            this.Show(String.Format("The \"{0}\" load balancer certificate (\"{1}\") can be found on another server (ie. \"{2}\")."
                                            , lsDnsName, lsLoadBalancerPfxPathFile, lsLoadBalancerComputerName)
                                    , lsCaption
                                    , tvMessageBoxButtons.OK, tvMessageBoxIcons.Alert, "LoadBalancerCertWrongServer");
            bool        lbReleased = File.Exists(lsLoadBalancerPfxPathFile);
                        if ( !lbReleased && !lbWrongServer )
                            this.Show(String.Format("The \"{0}\" load balancer certificate (\"{1}\") does not yet exist on this server.", lsDnsName, lsLoadBalancerPfxPathFile)
                                    , lsCaption
                                    , tvMessageBoxButtons.OK, tvMessageBoxIcons.Alert, "LoadBalancerCertDoesNotExist");
            string      lsLogFileTextReported = null;
            tvProfile   loMinProfile = Env.oMinProfile(moProfile);
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);

            try
            {
                if ( lbReleased )
                    File.Delete(lsLoadBalancerPfxPathFile);
            }
            catch (Exception ex)
            {
                Environment.ExitCode = 1;
                lbReleased = false;

                this.ShowError(String.Format("The \"{0}\" load balancer certificate (\"{1}\") on \"{2}\" can't be removed! ({3})"
                                                , lsDnsName, lsLoadBalancerPfxPathFile, Env.sComputerName, ex.Message), lsCaption);

                DoGetCert.ReportErrors(ex, out lsLogFileTextReported);
            }

            if ( lbReleased || lbWrongServer )
            {

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.NotifyLoadBalancerCertificateReady(lsHash, lbtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();

                    if ( !lbWrongServer )
                        this.Show("Load balancer \"Certificate Ready\" notifications have been sent.", lsCaption
                                , tvMessageBoxButtons.OK, tvMessageBoxIcons.Done, "LoadBalancerCertReleased");
                }
            }

            moProfile["-NoPrompts"] = lbNoPromptsBackup;
            moProfile.Remove(gsInAndOut1LbReleaseCertKey);
            moProfile.Save();

            return true;
        }

        public bool bCallAnySCS()
        {
            // This is a command-line only parameter. If it's not already there, don't add it.
            if ( !moProfile.ContainsKey(gsInAndOut2ServiceCallKey) )
                return false;


            tvProfile   loReturnProfile = null;
            tvProfile   loMethodProfile = new tvProfile(moProfile.sValue(gsInAndOut2ServiceCallKey, ""));
            string      lsMethodKey     = "-Method";
            string      lsMethodName    = loMethodProfile.sValue(lsMethodKey, "");
            string      lsSuccessKey    = "-Success";

            if ( String.IsNullOrEmpty(lsMethodName) )
            {
                Env.LogIt("");
                Env.LogIt("SCS calls must include -Method (at least). Can't continue.");
            }
            else
            {
                Env.LogIt("");
                Env.LogIt(String.Format("Calling method \"{0}\" on the SCS ...", loMethodProfile.sValue(lsMethodKey, "")));

                tvProfile   loMinProfile = Env.oMinProfile(moProfile);
                byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                string      lsHash = HashClass.sHashIt(loMinProfile);

                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loReturnProfile = new tvProfile(loGetCertServiceClient.sCallAnyMethod(lsHash, lbtArrayMinProfile, loMethodProfile.btArrayZipped()));
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

                if ( !loReturnProfile.bValue(lsSuccessKey, false) )
                {
                    Env.LogIt(String.Format("    Method \"{0}\" failed.", loMethodProfile.sValue(lsMethodKey, "")));
                }
                else
                {
                    string lsOutputFile = loMethodProfile.sValue("-OutputFile", "");

                    if ( !String.IsNullOrEmpty(lsOutputFile) )
                    {
                        loReturnProfile.bUseXmlFiles = loMethodProfile.bValue("-UseXmlOutput", true);
                        loReturnProfile.Save(lsOutputFile);
                    }

                    if ( "resetconfig" != lsMethodName.ToLower() )
                    {
                        moProfile.Remove(gsInAndOut2ServiceCallKey);
                        moProfile.Save();
                    }
                    else
                    {
                        tvProfile loNewProfile = new tvProfile();

                        foreach (DictionaryEntry loEntry in moProfile)
                            if ( DoGetCert.oPreservedKeys.ContainsKey(loEntry.Key.ToString()) )
                                loNewProfile.Add(loEntry);

                        string lsBackupPathFile = moProfile.sLoadedPathFile + ".backup.txt";

                        Env.LogIt(String.Format("    Backing up \"{0}\" to \"{1}\" ...", Path.GetFileName(moProfile.sLoadedPathFile), Path.GetFileName(lsBackupPathFile)));

                        File.Delete(lsBackupPathFile);
                        File.Copy(moProfile.sLoadedPathFile, lsBackupPathFile);

                        Env.LogIt(String.Format("    Writing profile reset to \"{0}\" ...", Path.GetFileName(moProfile.sLoadedPathFile)));
                        Env.LogIt(              "    Done.");

                        loNewProfile.Save(moProfile.sLoadedPathFile);
                    }
                }

                if ( !loMethodProfile.bValue("-ShowIt", true) )
                {
                    Env.LogIt(loReturnProfile.sCommandLine());
                }
                else
                {
                    bool    lbNoPromptsBackup = moProfile.bValue("-NoPrompts", false);
                            moProfile["-NoPrompts"] = false;
                    string  lsCaption = "SCS Result";
                            if ( loReturnProfile.bValue(lsSuccessKey, false) )
                                this.Show(loReturnProfile.sCommandBlock(), lsCaption);
                            else
                                this.ShowError(loReturnProfile.sCommandBlock(), lsCaption);

                    moProfile["-NoPrompts"] = lbNoPromptsBackup;
                }
            }

            return true;
        }

        public void LogStage(string asStageId)
        {
            Env.LogIt("");
            Env.LogIt("");
            Env.LogIt(String.Format("Stage {0} ...", asStageId));
            Env.LogIt("");
        }

        public static void ReportErrors(out string asLogFileTextReported)
        {
            DoGetCert.ReportErrors(null, out asLogFileTextReported);
        }
        public static void ReportErrors(Exception ex, out string asLogFileTextReported)
        {
            try
            {
                Env.LogIt("");
                Env.LogIt("Reading system performance counters ...");

                PerformanceCounter  loCpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                                    loCpuCounter.NextValue();
                PerformanceCounter  loRamCounter = new PerformanceCounter("Memory", "% Committed Bytes In Use");
                                    loRamCounter.NextValue();
                                    Thread.Sleep(1000);

                                    Env.LogIt(String.Format(@"CPU: {0,3}%   RAM: {1,2}%", (int)loCpuCounter.NextValue(), 100 - (int)loRamCounter.NextValue()));
            }
            catch (Exception ex2)
            {
                Env.LogIt(Env.sExceptionMessage(ex2));
            }

            tvProfile   loMinProfile = Env.oMinProfile(tvProfile.oGlobal());
                        loMinProfile.Add("-Env.iNonErrorBounceSecs", Env.iNonErrorBounceSecs);
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);

            asLogFileTextReported = File.ReadAllText(Env.sLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return;

            string  lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                    lsPreviousErrorsLogPathFile = Path.Combine(Path.GetDirectoryName(lsPreviousErrorsLogPathFile), String.Format("{0}({1}){2}"
                            , Path.GetFileNameWithoutExtension(lsPreviousErrorsLogPathFile), tvProfile.oGlobal().sValue("-CertificateDomainName" ,""), Path.GetExtension(lsPreviousErrorsLogPathFile)));
                    if ( File.Exists(lsPreviousErrorsLogPathFile) )
                        File.AppendAllText(lsPreviousErrorsLogPathFile, asLogFileTextReported);
                    else
                        File.Copy(Env.sLogPathFile, lsPreviousErrorsLogPathFile, false);

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            loCompressContentProfile.Add(gsContentKey, File.ReadAllText(lsPreviousErrorsLogPathFile) + Env.sHostLogText());

                loGetCertServiceClient.ReportErrors(lsHash, lbtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                File.Delete(lsPreviousErrorsLogPathFile);
            }

            if (       tvProfile.oGlobal().bValue("-Auto", false)
                    && 0 == Env.iNonErrorBounceSecs
                    && (null == ex || !(ex is CommunicationException))
                    )
                Env.ScheduleOnErrorBounce(DoGetCert.oDomainProfile);
        }

        public static void ReportEverything(out string asLogFileTextReported)
        {
            tvProfile   loMinProfile = Env.oMinProfile(tvProfile.oGlobal());
                        loMinProfile.Add("-Env.iNonErrorBounceSecs", Env.iNonErrorBounceSecs);
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);

            asLogFileTextReported = File.ReadAllText(Env.sLogPathFile);

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return;

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                tvProfile   loCompressContentProfile = new tvProfile();
                            if ( DoGetCert.oDomainProfile.bValue("-AllLogs", false) )
                                loCompressContentProfile.Add(gsContentKey, asLogFileTextReported + Env.sHostLogText());

                loGetCertServiceClient.ReportEverything(lsHash, lbtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( tvProfile.oGlobal().bValue("-DoStagingTests", true) && tvProfile.oGlobal().bValue("-ResetStagingLogs", true) )
                File.Delete(Env.sLogPathFile);

            if (       tvProfile.oGlobal().bValue("-Auto", false)
                    && Env.iNonErrorBounceSecs > 0
                    )
                Env.ScheduleNonErrorBounce(DoGetCert.oDomainProfile);
        }

        public void Show(string asMessageText, string asMessageCaption)
        {
            Env.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.Show(this.oUI, asMessageText, asMessageCaption);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

        public void ShowError(string asMessageText, string asMessageCaption)
        {
            Env.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

            if ( null != this.oUI && !this.bNoPrompts )
            this.oUI.Dispatcher.BeginInvoke((Action)(() =>
            {
                tvMessageBox.ShowError(this.oUI, asMessageText, asMessageCaption);
            }
            ));
            System.Windows.Forms.Application.DoEvents();
        }

        private bool bDoIisBinding(
                  X509Certificate2 aoOldCertificate
                , X509Certificate2 aoNewCertificate
                , string asCertName
                , string[] asSanArray
                , ServerManager aoServerManager
                , X509Store aoStore
                )
        {
            bool    lbDoIisBinding = true;
            bool    lbCreateSanSitesForBinding = moProfile.bValue("-CreateSanSitesForBinding", false);
            string  lsWellKnownBasePath = moProfile.sValue("-WellKnownBasePath", ".well-known");

            bool lbDidBinding = false;

            // Apply new site-to-cert bindings.
            foreach (string lsSanItem in DoGetCert.sSanArrayAmended(aoServerManager, asSanArray))
            {
                if ( !lbDoIisBinding || this.bMainLoopStopped )
                    break;

                Env.LogIt(String.Format("Applying certificate (\"{0}\") for \"{1}\" ...", aoNewCertificate.Thumbprint, lsSanItem));

                Site    loSite = null;
                        Env.oCurrentCertificate(lsSanItem, asSanArray, out aoServerManager, out loSite);

                if ( null == loSite && lsSanItem == asCertName && !lbCreateSanSitesForBinding )
                {
                    // No site found. Use the default site, if it exists (for the primary domain only).
                    if ( 0 != aoServerManager.Sites.Count )
                        loSite = aoServerManager.Sites[0];   // Default website.

                    Env.LogIt(String.Format("No SAN website match could be found for \"{0}\" (and -CreateSanSitesForBinding is \"False\"). The default will be used (ie. \"{1}\").", lsSanItem, loSite.Name));
                }

                if ( null == loSite && lbCreateSanSitesForBinding )
                    loSite = this.oSanSiteCreated(lsSanItem, aoServerManager, null);

                if ( null == loSite )
                {
                    // Still no site found (not even the default).

                    Env.LogIt(String.Format("No website could be found to bind the certificate (\"{0}\") for \"{1}\"\r\n(no default website either; FYI, -CreateSanSitesForBinding is \"{2}\")."
                                                , aoNewCertificate.Thumbprint, lsSanItem, lbCreateSanSitesForBinding));
                }

                if ( null != loSite && null == DoGetCert.oSslBinding(loSite) )
                {
                    string  lsNewBindingMsg1 = "No SSL binding found.";
                    string  lsNewBindingInformation = "*:443:{0}";
                    Binding loNewBinding = null;

                    try
                    {
                        if ( !asSanArray.Contains(lsSanItem) )
                        {
                            // The site found is not SAN specific (ie. don't use SNI flag).
                            lsNewBindingInformation = String.Format(lsNewBindingInformation, "");
                            loNewBinding = loSite.Bindings.Add(lsNewBindingInformation, null == aoOldCertificate ? aoNewCertificate.GetCertHash() : aoOldCertificate.GetCertHash(), aoStore.Name);
                            lbDidBinding = true;

                            Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with no SNI) using \"{0}\".", lsNewBindingInformation));
                        }
                        else
                        {
                            // Set the SNI (Server Name Indication) flag.
                            lsNewBindingInformation = String.Format(lsNewBindingInformation, lsSanItem);
                            loNewBinding = loSite.Bindings.Add(lsNewBindingInformation, null == aoOldCertificate ? aoNewCertificate.GetCertHash() : aoOldCertificate.GetCertHash(), aoStore.Name);
                            loNewBinding.SetAttributeValue("SslFlags", 1 /* SslFlags.Sni */);
                            lbDidBinding = true;

                            Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied (with SNI) using \"{0}\".", lsNewBindingInformation));
                        }
                    }
                    catch (Exception)
                    {
                        if ( null != loNewBinding )
                            loSite.Bindings.Remove(loNewBinding);

                        // Reverting to the default binding usage (an older OS perhaps).
                        lsNewBindingInformation = "*:443:";
                        loSite.Bindings.Add(lsNewBindingInformation, null == aoOldCertificate ? aoNewCertificate.GetCertHash() : aoOldCertificate.GetCertHash(), aoStore.Name);
                        lbDidBinding = true;

                        Env.LogIt(String.Format(lsNewBindingMsg1 + " Binding applied using \"{0}\".", lsNewBindingInformation));
                    }

                    aoServerManager.CommitChanges();

                    Env.LogIt(String.Format("Default SSL binding added to site: \"{0}\".", loSite.Name));
                }

                if ( null != loSite && lbDoIisBinding && !this.bMainLoopStopped )
                {
                    bool lbBindingFound = false;

                    foreach (Binding loBinding in loSite.Bindings)
                    {
                        if ( !lbDoIisBinding || this.bMainLoopStopped )
                            break;

                        // Only try to bind the new certificate if a certificate binding already exists (ie. only replace what's already there).
                        bool    lbDoBinding = null != loBinding.CertificateHash;
                                if ( lbDoBinding )
                                {
                                    // Only replace an existing certificate binding if it matches the old certificate (if an old certificate exists).
                                    lbDoBinding = null != aoOldCertificate && aoOldCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash);

                                    // Finally, only replace an existing certificate binding if it matches the certificate used for communications.
                                    if ( !lbDoBinding )
                                        lbDoBinding = (null != Env.oChannelCertificate && Env.oChannelCertificate.GetCertHash().SequenceEqual(loBinding.CertificateHash));
                                }
                        if ( lbDoBinding )
                        {
                            string[]    lsBindingInfoArray  = loBinding.BindingInformation.Split(':');
                            string      lsBindingAddress    = lsBindingInfoArray[0];
                            string      lsBindingPort       = lsBindingInfoArray[1];
                            string      lsBindingHost       = lsBindingInfoArray[2];

                            loBinding.CertificateHash = aoNewCertificate.GetCertHash();
                            lbDidBinding = this.bIisBindingCommit(aoServerManager, aoNewCertificate);

                            if ( "*" == lsBindingAddress )
                                Env.LogIt(String.Format("Certificate (\"{0}\") bound to port {1} in IIS for site \"{2}\"."
                                                , aoNewCertificate.Thumbprint, lsBindingPort, loSite.Name));
                            else
                                Env.LogIt(String.Format("Certificate (\"{0}\") bound to address {1} and port {2} in IIS for site \"{3}\"."
                                                , aoNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, loSite.Name));

                            lbBindingFound = true;

                            DoGetCert.WriteFailSafeFile(lsSanItem, loSite, aoNewCertificate, lsWellKnownBasePath);
                        }
                    }

                    if ( !lbBindingFound && moProfile.bValue("-SetDefaultBinding", true) )
                    {
                        // No binding was found matching the old certificate. Assign the new certificate to the default SSL binding.

                        Binding loDefaultSslBinding = DoGetCert.oSslBinding(loSite);

                        if ( null == loDefaultSslBinding )
                        {
                            // Still no binding found (not even the default). This is an error.

                            lbDoIisBinding = false;
                            Env.iNonErrorBounceSecs = -1;

                            Env.LogIt(String.Format("No SSL binding could be found for \"{0}\" (no default SSL binding either).\r\nCan't continue.", lsSanItem));
                        }
                        else
                        {
                            if ( null != loDefaultSslBinding.CertificateHash && aoNewCertificate.GetCertHash().SequenceEqual(loDefaultSslBinding.CertificateHash) )
                            {
                                lbDidBinding = true;

                                Env.LogIt(String.Format("Binding already applied for \"{0}\"{1}", lsSanItem
                                            , loSite.Name == lsSanItem ? "." : String.Format(" (via \"{0}\").", loSite.Name)));
                            }
                            else
                            {
                                string[]    lsBindingInfoArray  = loDefaultSslBinding.BindingInformation.Split(':');
                                string      lsBindingAddress    = lsBindingInfoArray[0];
                                string      lsBindingPort       = lsBindingInfoArray[1];
                                string      lsBindingHost       = lsBindingInfoArray[2];

                                loDefaultSslBinding.CertificateHash = aoNewCertificate.GetCertHash();
                                lbDidBinding = this.bIisBindingCommit(aoServerManager, aoNewCertificate);

                                if ( "*" == lsBindingAddress )
                                    Env.LogIt(String.Format("Certificate (\"{0}\") bound to port {1} in IIS for site \"{2}\"."
                                                    , aoNewCertificate.Thumbprint, lsBindingPort, loSite.Name));
                                else
                                    Env.LogIt(String.Format("Certificate (\"{0}\") bound to address {1} and port {2} in IIS for site \"{3}\"."
                                                    , aoNewCertificate.Thumbprint, lsBindingAddress, lsBindingPort, loSite.Name));

                                DoGetCert.WriteFailSafeFile(lsSanItem, loSite, aoNewCertificate, lsWellKnownBasePath);
                            }
                        }
                    }
                }
            }

            if ( !lbDidBinding )
            {
                lbDoIisBinding = false;

                Env.LogIt(String.Format("The certificate (\"{0}\") was NOT successfully bound to any site. This is an error.", aoNewCertificate.Thumbprint));
            }

            return lbDoIisBinding;
        }

        private bool bIisBindingCommit(ServerManager aoServerManager, X509Certificate2 aoNewCertificate)
        {
            bool lbIisBindingCommit = false;

            try
            {
                aoServerManager.CommitChanges();
                lbIisBindingCommit = true;
            }
            catch (Exception ex)
            {
                if ( !ex.Message.Contains("A specified logon session does not exist.") )
                {
                    throw ex;
                }
                else
                {
                    Env.LogIt(Env.sExceptionMessage(ex));
                    Env.LogIt("");
                    Env.LogIt("Retrying ...");

                    lbIisBindingCommit = this.bRunPowerScript(moProfile.sValue("-ScriptBindingCommitFix", @"

certutil -repairstore my {NewCertificateThumbprint}
                            ")
                            .Replace("{NewCertificateThumbprint}", aoNewCertificate.Thumbprint)
                            );

                    if ( lbIisBindingCommit )
                    {
                        lbIisBindingCommit = false;
                        aoServerManager.CommitChanges();
                        lbIisBindingCommit = true;
                    }
                }
            }

            return lbIisBindingCommit;
        }
        
        private Site oSanSiteCreated(string asSanItem, ServerManager aoServerManager, Site aoPrimarySiteForDefaults)
        {
            Site    loSanSiteCreated = null;
            string  lsPhysicalPath = null;

            Env.LogIt("");
            Env.LogIt(String.Format("No website found for \"{0}\". Creating new website ...", asSanItem));

            if ( 0 == aoServerManager.Sites.Count )
            {
                lsPhysicalPath = moProfile.sValue("-DefaultPhysicalPath", @"%SystemDrive%\inetpub\wwwroot");

                Env.LogIt(String.Format("    No default website could be found. -DefaultPhysicalPath used instead (\"{0}\").", lsPhysicalPath));
            }
            else
            {
                if ( null != aoPrimarySiteForDefaults )
                    lsPhysicalPath = aoPrimarySiteForDefaults.Applications["/"].VirtualDirectories["/"].PhysicalPath;
                else
                    lsPhysicalPath = aoServerManager.Sites[0].Applications["/"].VirtualDirectories["/"].PhysicalPath;
            }

            // Leave hostname blank in the default website to allow for old browsers (ie. no SNI support).
            loSanSiteCreated = aoServerManager.Sites.Add(
                      asSanItem
                    , "http"
                    , String.Format("*:80:{0}", 0 == aoServerManager.Sites.Count ? "" : asSanItem)
                    , lsPhysicalPath
                    );

            aoServerManager.CommitChanges();

            foreach (Site loSite in aoServerManager.Sites)
            {
                if ( loSite.Id == loSanSiteCreated.Id )
                    loSanSiteCreated = loSite;          // Allows subsequent writes to the new site.
            }

            Env.LogIt(String.Format("Successfully created new website for \"{0}\".", asSanItem));
            Env.LogIt("");

            return loSanSiteCreated;
        }

        private string sSingleSessionScriptPathFile
        {
            get
            {
                if ( String.IsNullOrEmpty(msSingleSessionScriptPathFile) && moProfile.bValue("-SingleSessionEnabled", false) )
                {
                    msSingleSessionScriptPathFile = moProfile.sRelativeToExePathFile(moProfile.sValue("-PowerScriptSessionPathFile", "InGetCertSession.ps1"));

                    // Fetch global session script.
                    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                            , String.Format("{0}{1}", Env.sFetchPrefix, Path.GetFileName(msSingleSessionScriptPathFile)), msSingleSessionScriptPathFile);
                }

                return msSingleSessionScriptPathFile;
            }
        }
        private string msSingleSessionScriptPathFile = null;

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

        private static bool bDoCheckIn(string asHash, byte[] abtArrayMinProfile)
        {
            bool lbDoCheckIn = false;

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return true;

            string lsCertName = tvProfile.oGlobal().sValue("-CertificateDomainName" ,"");

            if ( tvProfile.oGlobal().bValue("-UseCheckInDelays", true) )
            {
                // Wait a random period each cycle to allow this client to check-in at times other
                // than the beginning of the maintenance window (skip if manually running tests).

                DateTime    ldtMaintenanceWindowEndTime = tvProfile.oGlobal().dtValue(gsMaintWindEndKey, DateTime.MinValue);
                            ldtMaintenanceWindowEndTime = DateTime.Now.Date.AddMinutes((ldtMaintenanceWindowEndTime - ldtMaintenanceWindowEndTime.Date).TotalMinutes);
                int         liMinsToMaintenanceWindowEndTime = (int)(ldtMaintenanceWindowEndTime - DateTime.Now).TotalMinutes
                                                                - tvProfile.oGlobal().iValue("-MinCheckInMinsToMaintenanceWindowEndTime", 5) + 1;
                            if ( liMinsToMaintenanceWindowEndTime > 0
                                    && !tvProfile.oGlobal().bValue("-DoStagingTests", true)
                                    && tvProfile.oGlobal().bValue("-Auto", false) && !tvProfile.oGlobal().bValue("-Setup", false)
                                    )
                            {
                                int liCheckInDelayMins = new Random().Next(liMinsToMaintenanceWindowEndTime - 1) + 1;

                                Env.LogIt("");
                                Env.LogIt(String.Format("Waiting {0} minute{1} ({2} minute{3} max) before the \"{4}\" domain check-in ..."
                                                        , liCheckInDelayMins, 1 == liCheckInDelayMins ? "" : "s"
                                                        , liMinsToMaintenanceWindowEndTime, 1 == liMinsToMaintenanceWindowEndTime ? "" : "s", lsCertName));

                                Thread.Sleep(60000 * liCheckInDelayMins);
                            }
            }

            Env.ScheduleBounceReset();

            Env.LogIt("");
            Env.LogIt(String.Format("Attempting check-in with certificate repository (\"{0}\") ...", lsCertName));

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                if ( null != goStartupWaitMsg )
                    goStartupWaitMsg.Hide();

                if ( null != goStartupWaitMsg && !tvProfile.oGlobal().bExit )
                    goStartupWaitMsg.Show();

                string      lsPreviousErrorsLogPathFile = tvProfile.oGlobal().sRelativeToProfilePathFile(tvProfile.oGlobal().sValue("-PreviousErrorsLogPathFile", "PreviousErrorsLog.txt"));
                            lsPreviousErrorsLogPathFile = Path.Combine(Path.GetDirectoryName(lsPreviousErrorsLogPathFile), String.Format("{0}({1}){2}"
                                    , Path.GetFileNameWithoutExtension(lsPreviousErrorsLogPathFile), lsCertName, Path.GetExtension(lsPreviousErrorsLogPathFile)));
                string      lsErrorLog = null;
                tvProfile   loCompressContentProfile = new tvProfile();
                            if ( File.Exists(lsPreviousErrorsLogPathFile) )
                            {
                                lsErrorLog = File.ReadAllText(lsPreviousErrorsLogPathFile);
                                loCompressContentProfile.Add(gsContentKey, lsErrorLog);
                            }

                lbDoCheckIn = loGetCertServiceClient.bClientCheckIn(asHash, abtArrayMinProfile, loCompressContentProfile.btArrayZipped());
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                if ( !lbDoCheckIn )
                {
                    throw new Exception("Client FAILED to check-in with the certificate repository. Can't continue.");
                }
                else
                {
                    Env.LogIt(String.Format("Client successfully checked-in{0}.", String.IsNullOrEmpty(lsErrorLog) ? "" : " (and reported previous errors)"));

                    File.Delete(lsPreviousErrorsLogPathFile);

                    // Cache maintenance window details.
                    DoGetCert.ClearCache();
                    tvProfile.oGlobal()[gsMaintWindEndKey] = DoGetCert.oDomainProfile.dtValue(gsMaintWindEndKey, DateTime.MinValue);
                    tvProfile.oGlobal().Save();
                }
            }

            return lbDoCheckIn;
        }

        private static bool bDoCheckOut(bool abPendingResult)
        {
            bool lbDoCheckOut = false;

            if ( tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
                return true;

            Env.LogIt("Checking-out of certificate repository ...");

            tvProfile   loMinProfile = Env.oMinProfile(tvProfile.oGlobal());
            byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
            string      lsHash = HashClass.sHashIt(loMinProfile);

            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                lbDoCheckOut = loGetCertServiceClient.bClientCheckOut(lsHash, lbtArrayMinProfile);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();

                if ( !lbDoCheckOut )
                {
                    throw new Exception("Client FAILED to check-out of the certificate repository. Can't continue.");
                }
                else
                {
                    Env.LogIt("Client successfully checked-out.");
                }
            }

            if ( !abPendingResult )
                lbDoCheckOut = abPendingResult;

            return lbDoCheckOut;
        }

        private static void RemoveCertificate(X509Certificate2 aoCertificate, X509Store aoStore)
        {
            DoGetCert.RemoveCertificate(aoCertificate, aoStore, true);
        }
        private static void RemoveCertificate(X509Certificate2 aoCertificate, X509Store aoStore, bool abDoCleanup)
        {
            if ( null == aoCertificate )
                return;

            aoStore.Remove(aoCertificate);

            if ( tvProfile.oGlobal().bValue("-SkipPreviousStore", false) )
            {
                // Get cert's private key file - and remove it (sadly, the OS typically let's these things accumulate forever).
                string  lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(tvProfile.oGlobal(), aoCertificate);
                        if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) )
                            File.Delete(lsMachineKeyPathFile);
                return;
            }

            X509Store loStore = null;

            try
            {
                try
                {
                    loStore = new X509Store(gsCertStorePreviousName, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.ReadWrite);
                }
                catch
                {
                    loStore = new X509Store(StoreName.Disallowed, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                    Env.LogIt(String.Format("The \"{0}\" certificate store can't be opened. Moving certificate (\"{1}\") to \"Untrusted\" instead."
                                            , gsCertStorePreviousName, aoCertificate.Thumbprint));
                }

                if ( abDoCleanup )
                {
                    X509Certificate2Collection loCertCollection = loStore.Certificates.Find(X509FindType.FindBySubjectName, Env.sCertName(aoCertificate), false);

                    foreach (X509Certificate2 loCertificate in loCertCollection)
                    {
                        // Remove all previous certificates (with the same subject) from "GetCertPrevious" (or "Untrusted") before adding the given one.
                        if ( loCertificate.Subject == aoCertificate.Subject )
                        {
                            loStore.Remove(loCertificate);

                            // Get cert's private key file - and remove it (sadly, the OS typically let's these things accumulate forever).
                            string  lsMachineKeyPathFile = HashClass.sMachineKeyPathFile(tvProfile.oGlobal(), loCertificate);
                                    if ( !String.IsNullOrEmpty(lsMachineKeyPathFile) )
                                        File.Delete(lsMachineKeyPathFile);
                        }
                    }
                }

                // Add the given certificate to "GetCertPrevious" (or "Untrusted") before removing it from wherever it was.
                loStore.Add(aoCertificate);
            }
            finally
            {
                if ( null != loStore )
                    loStore.Close();
            }
        }

        private void RemoveNewCertificate(X509Store aoStore, X509Certificate2 aoOldCertificate, X509Certificate2 aoNewCertificate, string asNewCertPfxPathFile)
        {
            X509Certificate2Collection  loCertCollection = null;
                                        if ( null != aoStore && null != aoStore.Certificates && null != aoOldCertificate )
                                            loCertCollection = aoStore.Certificates.Find(X509FindType.FindByThumbprint, aoOldCertificate.Thumbprint, false);
                                        else
                                        if ( null == aoStore || null == aoStore.Certificates )
                                            throw new Exception("The certificate store does not exist! Can't validate old certificate for new certificate removal.");

            // Only remove the new cert if the old cert is still there.
            if ( null == loCertCollection || 1 != loCertCollection.Count )
            {
                throw new Exception("The old certificate appears not to exist. The new certificate can't be removed!");
            }
            else
            {
                if ( null == aoStore )
                {
                    throw new Exception("Certificate store does not exist. The new certificate can't be removed!");
                }
                else
                if ( null == aoNewCertificate )
                {
                    throw new Exception("The new certificate appears not to exist. It can't be removed!");
                }
                else
                {
                    // Remove the new cert PFX file.
                    File.Delete(asNewCertPfxPathFile);

                    // Remove the new cert from the store.
                    DoGetCert.RemoveCertificate(aoNewCertificate, aoStore, false);

                    Env.LogIt(String.Format("The new certificate (\"{0}\") has been removed.", aoNewCertificate.Thumbprint));
                }
            }
        }


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
            Env.LogIt(asMessageText);

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
            Env.LogIt(asMessageText);

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
            Env.LogIt(asMessageText);

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
            Env.LogIt(String.Format("{1}; {0}", asMessageText, asMessageCaption));

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
            Env.LogIt("");
            Env.LogIt(String.Format("{0} exiting due to log-off or system shutdown.", Path.GetFileName(ResourceAssembly.Location)));

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

            return Env.bRunPowerScript(out lsOutput, null, asScript, abOpenOrCloseSingleSession, false, false);
        }
        private bool bRunPowerScript(out string asOutput, string asScript)
        {
            return Env.bRunPowerScript(out asOutput, this.sSingleSessionScriptPathFile, asScript, false, false, false);
        }

        private bool bApplyNewCertToSSO(X509Certificate2 aoOldCertificate, X509Certificate2 aoNewCertificate, string asHash, byte[] abtArrayMinProfile)
        {
            // Do nothing if running stand-alone or this isn't an SSO server (or we are not doing SSO certificate updates on this server).
            if ( moProfile.bValue("-UseStandAloneMode", true) || !DoGetCert.oDomainProfile.bValue("-IsSsoDomain", false)
                        || null == aoOldCertificate || moProfile.bValue("-SkipSsoServer", false) )
                return true;

            bool lbApplyNewCertToSSO = false;

            Env.LogIt("");
            Env.LogIt("But first, replace SSO certificates ...");

            string  lsOutput = null;
            bool    lbIsOlder = false;
            bool    lbIsProxy = false;
                    if ( !moProfile.ContainsKey("-IsSsoOlder") || !moProfile.ContainsKey("-IsSsoProxy") )
                    {
                        if ( !Env.bRunPowerScript(out lsOutput, null, moProfile.sValue("-ScriptSsoVerChk1", "try {(Get-AdfsProperties).HostName} catch {}"), false, true, false) )
                        {
                            Env.LogIt("Failed checking SSO vs. proxy.");
                            return false;
                        }

                        lbIsProxy = String.IsNullOrEmpty(lsOutput);

                        if ( lbIsProxy )
                        {
                            if ( !Env.bRunPowerScript(out lsOutput, null, moProfile.sValue("-ScriptSsoProxyVerChk", "try {(Get-WebApplicationProxySslCertificate).HostName} catch {}"), false, true, false) )
                            {
                                Env.LogIt("Failed checking SSO proxy version.");
                                return false;
                            }

                            lbIsOlder = String.IsNullOrEmpty(lsOutput);
                        }
                        else
                        {
                            if ( !Env.bRunPowerScript(out lsOutput, null, moProfile.sValue("-ScriptSsoVerChk2", "try {(Get-AdfsProperties).AuditLevel} catch {}"), false, true, false) )
                            {
                                Env.LogIt("Failed checking SSO version.");
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
            bool    lbSsoDone = DoGetCert.oDomainProfile.bValue("-SsoDone", false);
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

if ( (Get-AdfsCertificate -CertificateType Token-Signing).Thumbprint -eq ""{NewCertificateThumbprint}"" )
{
    echo ""SSO certificates already replaced (by another server in the farm).""
    echo ""The ADFS service will be restarted.""
}
else
{
    Set-AdfsCertificate    -CertificateType Service-Communications -Thumbprint {NewCertificateThumbprint}

    Add-AdfsCertificate    -CertificateType Token-Decrypting       -Thumbprint {NewCertificateThumbprint} -IsPrimary
    Remove-AdfsCertificate -CertificateType Token-Decrypting       -Thumbprint {OldCertificateThumbprint}

    Add-AdfsCertificate    -CertificateType Token-Signing          -Thumbprint {NewCertificateThumbprint} -IsPrimary
    Remove-AdfsCertificate -CertificateType Token-Signing          -Thumbprint {OldCertificateThumbprint}
}

netsh http delete sslcert HostnamePort=localhost:443
netsh http add    sslcert HostnamePort=localhost:443                                                    CertHash={NewCertificateThumbprint} CertStoreName=MY AppId=""{4f3277dc-d733-46d4-9cc9-bf33c81a88c1}""
netsh http delete sslcert HostnamePort={CertificateDomainName}:443
netsh http add    sslcert HostnamePort={CertificateDomainName}:443   SslCtlStoreName=AdfsTrustedDevices CertHash={NewCertificateThumbprint} CertStoreName=MY AppId=""{4f3277dc-d733-46d4-9cc9-bf33c81a88c1}""
netsh http delete sslcert HostnamePort={CertificateDomainName}:49443
netsh http add    sslcert HostnamePort={CertificateDomainName}:49443 ClientCertNegotiation=Enable       CertHash={NewCertificateThumbprint} CertStoreName=MY AppId=""{4f3277dc-d733-46d4-9cc9-bf33c81a88c1}""

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
            string  lsNewSsoThumbprint = DoGetCert.oDomainProfile.sValue("-SsoThumbprint", "");
                    if ( "" == lsNewSsoThumbprint )
                    {
                        Env.LogIt("");
                        throw new Exception(String.Format("The SSO certificate thumbprint for the \"{0}\" domain has not yet been set!", DoGetCert.oDomainProfile.sValue("-DnsName", "")));
                    }
            string  lsCurrentThumbprint = Env.oCurrentCertificate(DoGetCert.oDomainProfile.sValue("-DnsName", "")).Thumbprint;

            if ( !lbSsoDone )
            {
                // This means the certificates in the SSO configuration haven't changed yet.

                if ( lbIsProxy )
                {
                    bool    lbRetry = Env.bCanScheduleNonErrorBounce;
                            Env.LogIt("SSO proxy certificate renewal can't be attempted until SSO renewal completes. " + (!lbRetry ? "This is an error." : "Will try again in several minutes."));

                            if ( !lbRetry )
                            {
                                Env.iNonErrorBounceSecs = -1;
                                return false;
                            }
                            else
                            {
                                Env.iNonErrorBounceSecs = moProfile.iValue("-SsoProxyRenewalBounceSecs", 900);
                                return true;
                            }
                }
                else
                {
                    bool    lbLockCertificateRenewal = false;
                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                lbLockCertificateRenewal = loGetCertServiceClient.bLockCertificateRenewal(asHash, abtArrayMinProfile);
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }
                            // If the renewal can't be locked, another client must be doing the
                            // certificate renewal already.
                            if ( lbLockCertificateRenewal )
                            {
                                Env.LogIt("SSO certificate renewal has been locked for this domain.");
                            }
                            else
                            {
                                bool    lbRetry = Env.bCanScheduleNonErrorBounce;
                                        Env.LogIt("SSO certificate renewal can't be locked. " + (!lbRetry ? "This is an error." : "Will try again in several minutes."));

                                        if ( !lbRetry )
                                        {
                                            Env.iNonErrorBounceSecs = -1;
                                            return false;
                                        }
                                        else
                                        {
                                            Env.iNonErrorBounceSecs = moProfile.iValue("-SsoRenewalBounceSecs", 600);
                                            return true;
                                        }
                            }
                }
            }

            string  lsScriptSSO = moProfile.sValue("-ScriptSSO", lsDefaultScript);
                    if ( String.IsNullOrEmpty(lsScriptSSO) )
                    {
                        Env.iNonErrorBounceSecs = -1;
                        throw new Exception("-ScriptSSO is empty. This is an error.");
                    }
                    else
                    {
                        lbApplyNewCertToSSO = this.bRunPowerScript(lsScriptSSO
                                .Replace("{CertificateDomainName}", DoGetCert.oDomainProfile.sValue("-DnsName", ""))
                                .Replace("{NewCertificateThumbprint}", aoNewCertificate.Thumbprint)
                                .Replace("{OldCertificateThumbprint}", aoOldCertificate.Thumbprint)
                                );

                        if ( lbApplyNewCertToSSO && !lbSsoDone )
                        {
                            Env.LogIt("");
                            Env.LogIt(String.Format("Notifying repository: SSO configuration successfully updated with the new SSO certificate (\"{0}\") ...", aoNewCertificate.Thumbprint));

                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                lbApplyNewCertToSSO = loGetCertServiceClient.bSetSsoThumbprint(asHash, abtArrayMinProfile, aoNewCertificate.Thumbprint);
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }

                            if ( lbApplyNewCertToSSO )
                                Env.LogSuccess();
                            else
                                Env.LogIt(Env.sExceptionMessage("The SSO configuration update notification experienced a critical failure."));
                        }
                    }

            return lbApplyNewCertToSSO;
        }

        private bool bCertificateNotExpiring(string asHash, byte[] abtArrayMinProfile, X509Certificate2 aoOldCertificate)
        {
            bool lbCertificateNotExpiring = false;

            if ( moProfile.bValue("-DoStagingTests", true) )
            {
                Env.LogIt("( staging mode is in effect (-DoStagingTests=True) )");
                Env.LogIt("");
            }

            if ( null != aoOldCertificate )
            {
                bool    lbCheckExpiration = true;
                string  lsCertOverrideComputerName = DoGetCert.oDomainProfile.sValue("-CertOverrideComputerName", "");
                string  lsCertOverridePfxPathFile = DoGetCert.oDomainProfile.sValue("-CertOverridePfxPathFile", "");

                if ( "" != lsCertOverrideComputerName + lsCertOverridePfxPathFile && ("" == lsCertOverrideComputerName || "" == lsCertOverridePfxPathFile) )
                {
                    throw new Exception("Both CertOverrideComputerName AND CertOverridePfxPathFile must be defined for certificate overrides to work properly.");
                }
                else
                if ( !DoGetCert.oDomainProfile.bValue("-CertOverridePfxReady", false) && Env.sComputerName == lsCertOverrideComputerName && File.Exists(lsCertOverridePfxPathFile) )
                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                    {
                        loGetCertServiceClient.NotifyCertOverrideCertificateReady(asHash, abtArrayMinProfile);
                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                            loGetCertServiceClient.Abort();
                        else
                            loGetCertServiceClient.Close();
                    }

                // Allow checking staging certificates for expiration only in non-interactive (or non-setup) mode.

                if ( lbCheckExpiration &&  moProfile.bValue("-DoStagingTests", true) && (!moProfile.bValue("-Auto", false) || moProfile.bValue("-Setup", false)) )
                    lbCheckExpiration = false;

                // Always check production certificates for expiration - when not running interactively and not during setup.

                if ( lbCheckExpiration && !moProfile.bValue("-DoStagingTests", true) && (!moProfile.bValue("-Auto", false) || moProfile.bValue("-Setup", false)) )
                {
                    // We are running in production mode interactively (or during setup).
                    // So if the current certificate in invalid, skip the expiration check.
                    lbCheckExpiration = aoOldCertificate.Verify();

                    if ( !lbCheckExpiration )
                    {
                        Env.LogIt(String.Format("The current certificate ({0}) appears to be invalid.", aoOldCertificate.Subject));
                        Env.LogIt("  (If you know otherwise, make sure internet connectivity is available and the system clock is accurate.)");
                        Env.LogIt("Expiration status will therefore be ignored and the process will now run.");
                    }
                }

                if ( lbCheckExpiration )
                {
                    bool        lbClientDownloadsEnabled = DoGetCert.oDomainProfile.bValue("-ClientDownloadsEnabled", true);
                    DateTime    ldtExpiration = TimeZoneInfo.Local.IsDaylightSavingTime(aoOldCertificate.NotAfter) ? aoOldCertificate.NotAfter : aoOldCertificate.NotAfter.AddHours(1);
                    DateTime    ldtBeforeExpirationDate = DateTime.MinValue;
                                if ( moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    if ( moProfile.ContainsKey("-RenewalDateOverride") )
                                    {
                                        ldtBeforeExpirationDate = moProfile.dtValue("-RenewalDateOverride", DateTime.MaxValue.Date);
                                    }
                                    else
                                    {
                                        ldtBeforeExpirationDate = ldtExpiration.AddDays(-moProfile.iValue("-RenewalDaysBeforeExpiration", 30)).Date;
                                    }
                                }
                                else
                                if ( 1 == DoGetCert.oDomainProfile.iValue("-CertCount", 0) )
                                {
                                    int liRenewalDaysBeforeExpiration = DoGetCert.oDomainProfile.iValue("-RenewalDaysBeforeExpiration", 30);
                                        if ( moProfile.iValue("-RenewalDaysBeforeExpiration", 30) != liRenewalDaysBeforeExpiration )
                                        {
                                            moProfile["-RenewalDaysBeforeExpiration"] = liRenewalDaysBeforeExpiration;
                                            moProfile.Save();
                                        }

                                    ldtBeforeExpirationDate = ldtExpiration.AddDays(-liRenewalDaysBeforeExpiration).Date;

                                    DateTime    ldtRenewalDate = DoGetCert.oDomainProfile.dtValue("-RenewalDate", DateTime.MaxValue.Date);
                                                if ( ldtBeforeExpirationDate != ldtRenewalDate && ldtRenewalDate != DateTime.MaxValue.Date )
                                                {
                                                    Env.LogIt(String.Format("The calculated renewal date (\"{0}\") has been overridden. The certificate renewal date override is \"{1}\"."
                                                                                , ldtBeforeExpirationDate.ToShortDateString()
                                                                                , ldtRenewalDate.ToShortDateString()
                                                                                ));

                                                    ldtBeforeExpirationDate = ldtRenewalDate;
                                                }
                                }

                    if ( DateTime.Now < ldtBeforeExpirationDate || (!moProfile.bValue("-UseStandAloneMode", true) && !lbClientDownloadsEnabled) )
                    {
                        Env.LogIt(String.Format("Nothing to do until {0}. The current \"{1}\" certificate (\"{2}\") doesn't expire until {3}."
                                , lbClientDownloadsEnabled ? ldtBeforeExpirationDate.ToShortDateString() : "certificate downloads are enabled"
                                , aoOldCertificate.Subject
                                , aoOldCertificate.Thumbprint
                                , ldtExpiration.ToString()));

                        lbCertificateNotExpiring = true;
                    }
                    else
                    if ( !moProfile.bValue("-UseStandAloneMode", true) && DoGetCert.oDomainProfile.bValue("-Renewed", false) )
                    {
                        Env.LogIt(String.Format("Nothing to do until all servers on the domain have renewed. The current \"{0}\" certificate (\"{1}\") doesn't expire until {2}."
                                , aoOldCertificate.Subject
                                , aoOldCertificate.Thumbprint
                                , ldtExpiration.ToString()));

                        lbCertificateNotExpiring = true;
                    }
                }
            }

            return lbCertificateNotExpiring;
        }

        private static tvProfile oHandleUpdates(ref string[] args, out Mutex aoMutex)
        {
            tvProfile   loProfile = null;
            tvProfile   loCmdLine = new tvProfile(args, tvProfileDefaultFileActions.NoDefaultFile, tvProfileFileCreateActions.NoFileCreate, true);
            tvProfile   loMinProfile = null;
            string      lsRunKey = "-UpdateRunExePathFile";
            string      lsRplKey = "-UpdateRplIniPathFile";
            string      lsDelKey = "-UpdateDeletePathFile";
            string      lsUpdateRunExePathFile = null;
            string      lsUpdateRplIniPathFile = null;
            string      lsUpdateDeletePathFile = null;
            byte[]      lbtArrayMinProfile = null;
            string      lsHash = null;

            aoMutex = null;


            try
            {
                if ( 0 == args.Length )
                {
                    loProfile = tvProfile.oGlobal(new tvProfile(args, true));
                }
                else
                {
                    lsUpdateRunExePathFile = (string)loCmdLine[lsRunKey];
                    lsUpdateRplIniPathFile = (string)loCmdLine[lsRplKey];
                    lsUpdateDeletePathFile = (string)loCmdLine[lsDelKey];
                    loCmdLine.Remove(lsRunKey);
                    loCmdLine.Remove(lsRplKey);
                    loCmdLine.Remove(lsDelKey);

                    if ( !loCmdLine.bDefaultFileReplaced )
                    {
                        loProfile = tvProfile.oGlobal(new tvProfile(loCmdLine.sCommandLineArray(), true));
                    }
                    else
                    {
                        if ( loCmdLine.Count == args.Length )
                        {
                            // Use the given (empty) replacement profile and populate it.
                            loProfile = tvProfile.oGlobal(new tvProfile(loCmdLine.sCommandLineArray(), tvProfileDefaultFileActions.NoDefaultFile, tvProfileFileCreateActions.NoFileCreate, true));
                            loProfile.sActualPathFile = loCmdLine.sActualPathFile;
                            loProfile.sLoadedPathFile = loCmdLine.sActualPathFile;
                        }
                        else
                        {
                            // The replacement profile has already been populated.
                            loProfile = tvProfile.oGlobal(loCmdLine);

                            // Set the replacement profile parameter (in case there's an update).
                            lsUpdateRplIniPathFile = loCmdLine.sLoadedPathFile;

                            // Replace the full profile (loaded during profile replacement) with just the given "args" array - sans the "-ini" value.
                            loCmdLine = new tvProfile();
                            loCmdLine.LoadFromCommandLineArray(args, tvProfileLoadActions.Append);
                            loCmdLine.Remove("-ini");
                        }
                    }
                }

                if ( loProfile.bExit || DoGetCert.bDoInAndOutContainsKey() )
                    return loProfile;

                DoGetCert.VerifyDependenciesInit();
                DoGetCert.VerifyDependencies();
                if ( loProfile.bExit )
                    return loProfile;

                string lsCertName = loProfile.sValue("-CertificateDomainName", "");

                // Only do the following fetches before initial setup and if the containing application folder has been created.
                if ( Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).Contains(Path.GetFileNameWithoutExtension(ResourceAssembly.Location)) )
                {
                    if ( "" == loProfile.sValue("-ContactEmailAddress", "") && "" == lsCertName )
                    {
                        // Discard the default profile. Replace it with the WCF version fetched below.
                        File.Delete(loProfile.sLoadedPathFile);

                        // Fetch WCF config.
                        tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                , String.Format("{0}{1}", Env.sFetchPrefix, "GetCert2.exe.config"), loProfile.sLoadedPathFile);
                    }

                    //// Fetch host (if it's not already running).
                    //if ( String.IsNullOrEmpty(Env.sProcessExePathFile(Env.sHostProcess)) && !loProfile.bValue("-HostFetched", false) )
                    //{
                    //    tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    //            , String.Format("{0}{1}", Env.sFetchPrefix, Env.sHostProcess), loProfile.sRelativeToExePathFile(Env.sHostProcess));
                    //
                    //    loProfile["-HostFetched"] = true;
                    //    loProfile.Save();
                    //}
                    //
                    //string lsFetchName = null;
                    //
                    //// Fetch security context task definition (for SSO servers).
                    //tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    //        , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="GetCert2Task.xml")
                    //        , loProfile.sRelativeToExePathFile(lsFetchName));
                }

                //if ( !loProfile.bExit && null != Env.sProcessExePathFile(Env.sHostProcess) )
                //{
                //    // Remove host setup (if host is already installed).
                //    File.Delete(Path.Combine(Path.GetDirectoryName(loProfile.sLoadedPathFile), Env.sHostProcess));
                //}

                if ( !loProfile.bExit )
                {
                    bool        lbFirstInstance;
                    string      lsMutexName = "Global\\" + ResourceAssembly.GetName().Name + lsCertName;
                                aoMutex = new Mutex(false, lsMutexName, out lbFirstInstance);
                                if ( !lbFirstInstance )
                                {
                                    // Let's give the other instance time to finish - before trying again.
                                    Env.ProcessExitWait();

                                    aoMutex = new Mutex(false, lsMutexName, out lbFirstInstance);
                                    if ( !lbFirstInstance )
                                    {
                                        Env.LogIt(String.Format("Another \"{0}\" instance{1} is already running. Exiting ..."
                                                    , Path.GetFileName(loProfile.sExePathFile), "" == lsCertName ? "" : String.Format(" (\"{0}\")", lsCertName)));

                                        loProfile.bExit = true;
                                    }
                                }
                }

                if ( loProfile.bExit || loProfile.bValue("-UseStandAloneMode", true) )
                    return loProfile;

                if ( !loProfile.bValue("-NoPrompts", false) )
                {
                    goStartupWaitMsg = new tvMessageBox();
                    goStartupWaitMsg.ShowWait(
                            null, Path.GetFileNameWithoutExtension(loProfile.sExePathFile) + " loading, please wait ...", 250);
                }

                Env.ResetConfigMechanism(loProfile);

                loMinProfile = Env.oMinProfile(loProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);

                if ( String.IsNullOrEmpty(lsUpdateRunExePathFile) && String.IsNullOrEmpty(lsUpdateDeletePathFile) )
                {
                    Env.LogIt("");
                    Env.LogIt(String.Format("Version {0} command-line: {1} {2}"
                                                , FileVersionInfo.GetVersionInfo(loProfile.sExePathFile).FileVersion
                                                , Path.GetFileName(loProfile.sExePathFile)
                                                , String.Join(" ", loProfile.sInputCommandLineArray)));

                    // Check-in with the GetCert service.
                    loProfile.bExit = !DoGetCert.bDoCheckIn(lsHash, lbtArrayMinProfile);

                    // Short-circuit updates until basic setup is complete.
                    if ( "" == loProfile.sValue("-ContactEmailAddress", "") )
                        return loProfile;

                    // Does the service reference need updating? (must be handled first)
                    if ( !loProfile.bExit )
                        DoGetCert.HandleSvcUpdate(loProfile, lsHash, lbtArrayMinProfile);

                    // A not-yet-updated EXE is running. Does it need updating?
                    if ( !loProfile.bExit )
                        DoGetCert.HandleExeUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile, lsRunKey, lsRplKey, lsUpdateRplIniPathFile, loCmdLine);
                }
                else
                {
                    if ( String.IsNullOrEmpty(lsUpdateDeletePathFile) )
                    {
                        // An updated EXE is running (from update folder). It must be copied to parent folder and relaunched - then deleted (see "else" below).

                        bool lbUpdateCopied = false;

                        for (int i=0; i < loProfile.iValue("-UpdateFileWriteMaxRetries", 42); i++)
                        {
                            try
                            {
                                File.Copy(loProfile.sExePathFile, lsUpdateRunExePathFile, true);
                                lbUpdateCopied = true;
                                break;
                            }
                            catch
                            {
                                System.Windows.Forms.Application.DoEvents();
                                Thread.Sleep(loProfile.iValue("-UpdateFileWriteMinRetryMS", 500) + new Random().Next(loProfile.iValue("-UpdateFileWriteMaxRetryMS", 500)));
                            }
                        }

                        if ( !lbUpdateCopied )
                        {
                            Env.LogIt("Software update could not be copied into place. Will try again.", true);
                        }
                        else
                        {
                            Process loProcess = new Process();
                                    loProcess.StartInfo.FileName = lsUpdateRunExePathFile;
                                    loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2} {3}"
                                            , lsDelKey, loProfile.sExePathFile
                                            , String.IsNullOrEmpty(lsUpdateRplIniPathFile) ? "" : String.Format("-ini=\"{0}\"", lsUpdateRplIniPathFile)
                                            , loCmdLine.sCommandLine());
                                    loProcess.Start();
                        }

                        System.Windows.Forms.Application.DoEvents();
                        loProfile.bExit = true;
                    }
                    else
                    {
                        string lsCurrentVersion = null;

                        for (int i=0; i < loProfile.iValue("-UpdateFileWriteMaxRetries", 42); i++)
                        {
                            try
                            {
                                lsCurrentVersion = FileVersionInfo.GetVersionInfo(loProfile.sExePathFile).FileVersion;

                                if ( lsCurrentVersion != FileVersionInfo.GetVersionInfo(lsUpdateDeletePathFile).FileVersion )
                                    lsCurrentVersion = null;

                                Directory.Delete(Path.GetDirectoryName(lsUpdateDeletePathFile), true);
                                break;
                            }
                            catch (Exception ex)
                            {
                                if ( i == loProfile.iValue("-UpdateFileWriteMaxRetries", 42) - 1 )
                                {
                                    throw ex;
                                }
                                else
                                {
                                    System.Windows.Forms.Application.DoEvents();
                                    Thread.Sleep(loProfile.iValue("-UpdateFileWriteMinRetryMS", 500) + new Random().Next(loProfile.iValue("-UpdateFileWriteMaxRetryMS", 500)));
                                }    
                            }
                        }

                        if ( null == lsCurrentVersion )
                        {
                            Env.LogIt("Software update DID NOT complete successfully. Will try again.");
                        }
                        else
                        {
                            Env.LogIt(String.Format("Software update successfully completed. Version {0} is now running.", lsCurrentVersion));

                            // The update did not crap out, so cancel the previously scheduled "on error" bounce.
                            Env.ScheduleBounceOnErrorCancel(true);
                        }
                    }
                }

                if ( !loProfile.bExit )
                {
                    loProfile = DoGetCert.oHandleIniUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile, args);

                    DoGetCert.HandleCfgUpdate(UpdatedEXEs.Client, loProfile, lsHash, lbtArrayMinProfile);
                    DoGetCert.HandleHostUpdates(loProfile, lsHash, lbtArrayMinProfile);
                    DoGetCert.HandleSetupCertUpdate(loProfile, loMinProfile, lsHash, lbtArrayMinProfile);

                    DoGetCert.HandleExeUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile, null, null, null, null);
                    DoGetCert.oHandleIniUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile, null);
                    DoGetCert.HandleCfgUpdate(UpdatedEXEs.GcFailSafe, loProfile, lsHash, lbtArrayMinProfile);

                    DoGetCert.HandleDashboardUpdate(loProfile, lsHash, lbtArrayMinProfile);
                }
            }
            catch (Exception ex)
            {
                if ( null != loProfile )
                    loProfile.bExit = true;

                Environment.ExitCode = 1;
                Env.LogIt(Env.sExceptionMessage(ex));

                if ( !loProfile.bValue("-UseStandAloneMode", true) && !loProfile.bValue("-CertificateSetupDone", false) )
                {
                    Env.LogIt(@"

Checklist for {EXE} setup:

    On this local server:

        .) Once the initial setup details have been provided (after the initial use
           of the ""{CertificateDomainName}"" certificate), the SCS will expect to use a
           domain specific certificate for further communications.

        .) ""Host process can't be located on disk. Can't continue."" means the host
           software (ie. the task scheduler) must also be installed to continue.
           Look in the ""{EXE}"" folder for the host EXE and run it.

        .) If you suspect an issue locating the proper identifying certificate,
           set -LogCertificateStatus=True to log certificate selection status.

    On the SCS server:

        .) The ""-StagingTestsEnabled"" switch must be set to ""True"" for the staging
           tests (it gets reset automatically every night).

        .) The ""{CertificateDomainName}"" certificate must be added
           to the trusted people store.

        .) The ""{CertificateDomainName}"" certificate must be removed 
           from the ""untrusted"" certificate store (if it's there).
"
                            .Replace("{EXE}", Path.GetFileName(ResourceAssembly.Location))
                            .Replace("{CertificateDomainName}", Env.sNewClientSetupCertName)
                            )
                            ;
                }

                try
                {
                    string lsDiscard = null;

                    DoGetCert.ReportErrors(ex, out lsDiscard);
                }
                catch {}
            }

            return loProfile;
        }

        private static tvProfile oPathFileSetsNotInUpdate(tvProfile aoBaseProfile, tvProfile aoUpdateProfile, string asMainKey, string asPathFileKey)
        {
            // Look for sets in the base profile that are NOT in the update. If a path\file
            // specification in any base set exists in any update set, exclude the entire base set.
            tvProfile   loPathFileSetsNotInUpdate = new tvProfile();
                        foreach (DictionaryEntry loEntry in aoBaseProfile.oOneKeyProfile(asMainKey))
                        {
                            bool        lbAddSet = true;
                            tvProfile   loBaseSetPathFiles = (new tvProfile(loEntry.Value.ToString())).oOneKeyProfile(asPathFileKey);
                                        foreach (DictionaryEntry loPathFile in loBaseSetPathFiles)
                                        {
                                            string lsBaseSetPathFile = loPathFile.Value.ToString().ToLower();

                                            foreach (DictionaryEntry loEntry2 in aoUpdateProfile.oOneKeyProfile(asMainKey))
                                            {
                                                tvProfile   loSetUpdate = new tvProfile(loEntry2.Value.ToString());

                                                            foreach (DictionaryEntry loSetUpdatePathFile in loSetUpdate.oOneKeyProfile(asPathFileKey))
                                                            {
                                                                if ( lsBaseSetPathFile == loSetUpdatePathFile.Value.ToString().ToLower() )
                                                                {
                                                                    lbAddSet = false;
                                                                    break;
                                                                }
                                                            }

                                                if ( !lbAddSet )
                                                    break;
                                            }

                                            if ( !lbAddSet )
                                                break;
                                        }

                            if ( lbAddSet )
                                loPathFileSetsNotInUpdate.Add(asMainKey, loEntry.Value);
                        }

            return loPathFileSetsNotInUpdate;
        }

        // Return an array of every site using the current setup certificate.
        private static Site[] oSetupCertBoundSiteArray(ServerManager aoServerManager)
        {
            if ( null == Env.oChannelCertificate )
                return null;

            List<Site>  loList = new List<Site>();
            byte[]      lbtSetupCertHashArray = Env.oChannelCertificate.GetCertHash();

            // Walk thru the site list looking for all sites bound to the setup certificate.
            foreach (Site loSite in aoServerManager.Sites)
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

            foreach (Site loSite in loSetupCertBoundSiteArray)
                loList.Add(loSite.Name);

            // Append the given SAN array.
            foreach (string lsSanItem in asSanArray)
            {
                if ( !loList.Contains(lsSanItem) )
                    loList.Add(lsSanItem);
            }

            return loList.ToArray();
        }

        private static string sStopHostExePathFile()
        {
            Process loProcessFound = null;
            string  lsHostExePathFile = Env.sHostExePathFile(out loProcessFound, true);

            // Stop the EXE (can't update it or the INI while it's running).

            if ( !Env.bKillProcess(loProcessFound) )
            {
                Env.LogIt("Host process can't be stopped. Update can't be applied.");
                throw new Exception("Exiting ...");
            }
            else
            {
                if ( null != loProcessFound )
                {
                    // Let's give the host time to finish - before doing anything else.
                    Env.ProcessExitWait();
                    Env.LogIt("Host process stopped successfully.");
                }
            }

            return lsHostExePathFile;
        }

        private static tvProfile oHandleIniUpdate(UpdatedEXEs aeExeName, tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile, string[] args)
        {
            tvProfile   loProfile = null;
            string      lsExeName = aeExeName.ToString();
            string      lsKeyName = "-IniVersion";
            string      lsVersion = "0";
                        switch (aeExeName)
                        {
                            case UpdatedEXEs.Client:
                                // Reload the client profile without merging the command-line.
                                loProfile = new tvProfile(aoProfile.sLoadedPathFile, true);
                                lsVersion = loProfile.sValue(lsKeyName, "1");
                                break;
                            case UpdatedEXEs.GcFailSafe:
                                loProfile = new tvProfile(aoProfile.sRelativeToExePathFile(lsExeName + ".exe.config"), true);
                                if ( loProfile.ContainsKey(lsKeyName) )
                                    lsVersion = loProfile.sValue(lsKeyName, "0");
                                break;
                        }
            string      lsNewVersion = null;
            tvProfile   loIniUpdateProfile = new tvProfile();

                        // Look for profile update.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            switch (aeExeName)
                            {
                                case UpdatedEXEs.Client:
                                    string      lsDomainUpdate = loGetCertServiceClient.sIniUpdate(asHash, abtArrayMinProfile, lsVersion);
                                                if ( !String.IsNullOrEmpty(lsDomainUpdate) )
                                                {
                                                    loIniUpdateProfile.LoadFromCommandLine(lsDomainUpdate, tvProfileLoadActions.Overwrite);
                                                    lsNewVersion = loIniUpdateProfile.sValue("-IniVersion", "1");
                                                }
                                    string      lsServerUpdateVersion = String.Format(gsUpdateProfileServerFmt, gsUpdateProfileServerKey, aoProfile.sValue(gsUpdateProfileServerKey, DateTime.MinValue.ToShortDateString()));
                                    string      lsServerUpdate = loGetCertServiceClient.sIniUpdate(asHash, abtArrayMinProfile, lsServerUpdateVersion);
                                                if ( !String.IsNullOrEmpty(lsServerUpdate) )
                                                {
                                                    loIniUpdateProfile.LoadFromCommandLine(lsServerUpdate, tvProfileLoadActions.Merge);

                                                    if ( String.IsNullOrEmpty(lsDomainUpdate) )
                                                    {
                                                        lsVersion = lsServerUpdateVersion;
                                                        lsNewVersion = String.Format(gsUpdateProfileServerFmt, gsUpdateProfileServerKey, loIniUpdateProfile.sValue(gsUpdateProfileServerKey, DateTime.MinValue.ToShortDateString()));
                                                    }
                                                }
                                    break;
                                case UpdatedEXEs.GcFailSafe:
                                    loIniUpdateProfile.LoadFromCommandLine(loGetCertServiceClient.sFailSafeIniUpdate(asHash, abtArrayMinProfile, lsVersion), tvProfileLoadActions.Overwrite);
                                    break;
                            }
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

            float       lfDiscard = 0;
                        if ( !float.TryParse(lsVersion, out lfDiscard) )
                        {
                            string  lsFormat = "\"{0}\"";
                                    lsVersion = String.Format(lsFormat, lsVersion);
                                    lsNewVersion = String.Format(lsFormat, lsNewVersion);
                        }

            if ( 0 == loIniUpdateProfile.Count )
            {
                Env.LogIt(String.Format("{0} profile version {1} is in use. No update.", lsExeName, lsVersion));
            }
            else
            {
                Env.LogIt(String.Format("{0} profile version {1} is in use. Update found ...", lsExeName, lsVersion));

                // Remove keys slated to be removed (singular - supports wildcards).
                foreach (DictionaryEntry loEntry in loIniUpdateProfile.oOneKeyProfile("-RemoveKey"))
                    loProfile.Remove(loEntry.Value);
                // Remove keys slated to be removed (plural - a nested profile of keys - does NOT support wildcards).
                foreach (DictionaryEntry loEntry in loIniUpdateProfile.oOneKeyProfile("-RemoveKeys"))
                    foreach (DictionaryEntry loEntry2 in (new tvProfile(loEntry.Value.ToString())))
                        loProfile.Remove(loEntry2.Key);
                // Remove the "removals" prior to merge of the actual updated keys.
                loIniUpdateProfile.Remove("-RemoveKey");
                loIniUpdateProfile.Remove("-RemoveKeys");
                // Merge in the updated keys.
                loProfile.LoadFromCommandLine(loIniUpdateProfile.ToString(), tvProfileLoadActions.Merge);
                // Allow for saving everything (including merged keys).
                loProfile.bSaveSansCmdLine = false;
                loProfile.bSaveEnabled = true;
                // Save the updated profile.
                loProfile.Save();

                // Only the running client needs this special handling.
                if ( UpdatedEXEs.Client == aeExeName )
                {
                    // Finally, reload the updated profile (and merge in the command-line).
                    aoProfile = tvProfile.oGlobal(new tvProfile(args, true));
                }

                Env.LogIt(String.Format(
                        "{0} profile update successfully completed. Version {1} is now in use.", lsExeName, lsNewVersion));
            }

            return aoProfile;
        }

        private static void HandleCfgUpdate(UpdatedEXEs aeExeName, tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            tvProfile   loProfile = null;
            string      lsExeName = aeExeName.ToString();
            string      lsKeyName = "-CfgVersion";
            string      lsVersion = "0";
                        switch (aeExeName)
                        {
                            case UpdatedEXEs.Client:
                                loProfile = aoProfile;
                                lsVersion = loProfile.sValue(lsKeyName, "1");
                                break;
                            case UpdatedEXEs.GcFailSafe:
                                loProfile = new tvProfile(aoProfile.sRelativeToExePathFile(lsExeName + ".exe.config"), true);
                                if ( loProfile.ContainsKey(lsKeyName) )
                                    lsVersion = loProfile.sValue(lsKeyName, "0");
                                break;
                        }
            string      lsCfgUpdate = null;

                        // Look for WCF configuration update.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            switch (aeExeName)
                            {
                                case UpdatedEXEs.Client:
                                    lsCfgUpdate = loGetCertServiceClient.sCfgUpdate(asHash, abtArrayMinProfile, lsVersion);
                                    break;
                                case UpdatedEXEs.GcFailSafe:
                                    lsCfgUpdate = loGetCertServiceClient.sFailSafeCfgUpdate(asHash, abtArrayMinProfile, lsVersion);
                                    break;
                            }
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }

            if ( String.IsNullOrEmpty(lsCfgUpdate) )
            {
                Env.LogIt(String.Format("{0} WCF conf version {1} is in use. No update.", lsExeName, lsVersion));
            }
            else
            {
                Env.LogIt(String.Format("{0} WCF conf version {1} is in use. Update found ...", lsExeName, lsVersion));

                // Overwrite WCF config with the current update.
                File.WriteAllText(loProfile.sLoadedPathFile, lsCfgUpdate.Replace("{ServiceEndpoint}", Env.oGetCertServiceFactory.Endpoint.Address.Uri.ToString()));
                Env.ResetConfigMechanism(loProfile);
                Env.oGetCertServiceFactory = null;

                // Overwrite WCF config version number in the current profile.
                tvProfile   loWcfCfg = new tvProfile(loProfile.sLoadedPathFile, true);
                            loProfile["-CfgVersion"] = loWcfCfg.sValue("-CfgVersion", "1");
                            loProfile.Save();

                Env.LogIt(String.Format("{0} WCF conf update successfully completed. Version {1} is now in use.", lsExeName, loProfile.sValue("-CfgVersion", "1")));
            }
        }

        private static void HandleDashboardUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            string lsDashboardBasePath = DoGetCert.oDomainProfile.sValue("-DashboardBasePath", "");

            if ( "" != lsDashboardBasePath )
            {
                bool        lbUpdateSuccess = false;
                string      lsDashboardPath = Path.Combine(lsDashboardBasePath, Path.GetFileNameWithoutExtension(gsDashboardZipFile));
                tvProfile   loVersionProfile = new tvProfile(Path.Combine(lsDashboardPath, gsDashboardVersionFile), false);
                            loVersionProfile.bEnableFileLock = false;   // This is necessary to allow the profile file's deletion below. 
                string      lsVersion = loVersionProfile.sValue(gsDashboardVersionKey, "0");
                byte[]      lbtArrayDashboardUpdate = null;

                            // Look for dashboard update.
                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                lbtArrayDashboardUpdate = loGetCertServiceClient.btArrayDashboardZipUpdate(asHash, abtArrayMinProfile, lsVersion);
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }

                if ( null == lbtArrayDashboardUpdate )
                {
                    Env.LogIt(String.Format("Dashboard version {0} is in use. No update.", lsVersion));

                    lbUpdateSuccess = true;
                }
                else
                {
                    Env.LogIt("");
                    Env.LogIt(String.Format("Dashboard version {0} is in use. Update found ...", lsVersion));

                    string          lsDashboardBasePathToLower = lsDashboardBasePath.ToLower().Trim();
                    List<Site>      loBaseSiteList  = new List<Site>();
                    ServerManager   loServerManager = new ServerManager();
                                    foreach (Site loSite in loServerManager.Sites)
                                        foreach (Application loApp in loSite.Applications)
                                            if ( lsDashboardBasePathToLower == Environment.ExpandEnvironmentVariables(loApp.VirtualDirectories["/"].PhysicalPath).ToLower().Trim() )
                                                loBaseSiteList.Add(loSite);

                    if ( 0 == loBaseSiteList.Count )
                    {
                        Env.LogIt(String.Format("No website could be found matching the given location for the dashboard application (ie. \"{0}\").", lsDashboardBasePath));
                    }
                    else
                    {
                        foreach(Site loSite in loBaseSiteList)
                        {
                            loSite.Stop();

                            foreach (Application loApp in loSite.Applications)
                                if ( lsDashboardBasePathToLower == Environment.ExpandEnvironmentVariables(loApp.VirtualDirectories["/"].PhysicalPath).ToLower().Trim() )
                                    loServerManager.ApplicationPools[loApp.ApplicationPoolName].Recycle();
                        }

                        Thread.Sleep(aoProfile.iValue("-AfterSitesStoppedSleepMS", 1000));

                        try
                        {
                            try
                            {
                                // The deletion may fail on the app root directory.
                                // The goal is to cleanup everything else.
                                Directory.Delete(lsDashboardPath, true);
                            }
                            catch {}

                            string  lsZipPathFile = aoProfile.sRelativeToExePathFile(gsDashboardZipFile);
                                    File.WriteAllBytes(lsZipPathFile, lbtArrayDashboardUpdate);
                                    ZipFile.ExtractToDirectory(lsZipPathFile, lsDashboardBasePath);

                            lsVersion = (new tvProfile(Path.Combine(lsDashboardPath, gsDashboardVersionFile), false)).sValue(gsDashboardVersionKey, "0");

                            /*
                                We are intentionally not adding the application here.
                                That will be up to the dashboard admin to handle manually.
                            */

                            // Add support for "default.aspx".
                            foreach(Site loSite in loBaseSiteList)
                            {
                                try
                                {
                                    Microsoft.Web.Administration.Configuration                  loConfiguration = loServerManager.GetWebConfiguration(loSite.Name);
                                    Microsoft.Web.Administration.ConfigurationSection           loDefaultDocumentSection = loConfiguration.GetSection("system.webServer/defaultDocument");
                                    Microsoft.Web.Administration.ConfigurationElementCollection loFilesCollection = loDefaultDocumentSection.GetCollection("files");
                                    Microsoft.Web.Administration.ConfigurationElement           loNewFileElement = loFilesCollection.CreateElement("add");
                                                                                                loNewFileElement.Attributes["value"].Value = "default.aspx";
                                                                                                loFilesCollection.Add(loNewFileElement);
                                    loServerManager.CommitChanges();
                                }
                                catch
                                {
                                    loServerManager = new ServerManager();
                                }
                            }

                            lbUpdateSuccess = true;
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            foreach(Site loSite in loBaseSiteList)
                                loSite.Start();
                        }
                    }

                    if ( lbUpdateSuccess )
                        Env.LogIt(String.Format("Dashboard update successfully completed. Version {0} is now in use.", lsVersion));
                }

                string  lsDashboardDataPathFile = Path.Combine(lsDashboardPath, gsDashboardDataFile);

                if ( lbUpdateSuccess )
                {
                    byte[]  lbtArrayDashboardData = null;

                            // Look for dashboard data.
                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                lbtArrayDashboardData = loGetCertServiceClient.btArrayDashboardData(asHash, abtArrayMinProfile);
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }

                    if ( null == lbtArrayDashboardData )
                    {
                        Env.LogIt("");
                        Env.LogIt("Dashboard data retrieved but is empty. This is an error.");
                    }
                    else
                    {
                        tvProfile   loDashboardDataProfile = tvProfile.oProfile(lbtArrayDashboardData);
                                    loDashboardDataProfile.Save(lsDashboardDataPathFile);

                        Env.LogIt("Dashboard data received successfully.");
                    }
                }
            }
        }

        private static void HandleExeUpdate(
                  UpdatedEXEs aeExeName
                , tvProfile aoProfile
                , string asHash
                , byte[] abtArrayMinProfile
                , string asRunKey
                , string asRplKey
                , string asRplIniPathFile
                , tvProfile aoCmdLineProfile
                )
        {
            string  lsExeName = aeExeName.ToString();
            string  lsExePathFile = null;
            string  lsIniPathFile = null;
            string  lsUpdatePath = null;
                    switch (aeExeName)
                    {
                        case UpdatedEXEs.Client:
                            lsExePathFile = aoProfile.sExePathFile;
                            lsIniPathFile = aoProfile.sLoadedPathFile;
                            lsUpdatePath = Path.Combine(Path.GetDirectoryName(lsExePathFile), aoProfile.sValue("-UpdateFolder", "Update_" + Path.GetFileName(lsExePathFile)));
                            break;
                        case UpdatedEXEs.GcFailSafe:
                            lsExePathFile = aoProfile.sRelativeToExePathFile(lsExeName + ".exe");
                            lsIniPathFile = lsExePathFile + ".config";
                            lsUpdatePath = Path.GetDirectoryName(lsExePathFile);
                            break;
                    }
            string  lsCurrentVersion = !File.Exists(lsExePathFile) ? "0" : FileVersionInfo.GetVersionInfo(lsExePathFile).FileVersion;
            byte[]  lbtArrayExeUpdate = null;

                    // Look for an updated EXE (unless it was just updated).
                    if ( !File.Exists(lsExePathFile) || DateTime.Now > new FileInfo(lsExePathFile).LastWriteTime.AddMinutes(DoGetCert.oDomainProfile.iValue("-MinExeUpdateAgeMins", 0)) )
                    using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                    {
                        switch (aeExeName)
                        {
                            case UpdatedEXEs.Client:
                                lbtArrayExeUpdate = loGetCertServiceClient.btArrayGetCertExeUpdate(asHash, abtArrayMinProfile, lsCurrentVersion);
                                break;
                            case UpdatedEXEs.GcFailSafe:
                                lbtArrayExeUpdate = loGetCertServiceClient.btArrayFailSafeExeUpdate(asHash, abtArrayMinProfile, lsCurrentVersion);
                                break;
                        }
                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                            loGetCertServiceClient.Abort();
                        else
                            loGetCertServiceClient.Close();
                    }

            bool    lbUpdatedAlready = File.Exists(lsExePathFile) && lsCurrentVersion != FileVersionInfo.GetVersionInfo(lsExePathFile).FileVersion;
                    if ( lbUpdatedAlready )
                        lsCurrentVersion = FileVersionInfo.GetVersionInfo(lsExePathFile).FileVersion;

            if ( null == lbtArrayExeUpdate || lbUpdatedAlready )
            {
                Env.LogIt(String.Format("{0} version {1} is installed. No update.", lsExeName, lsCurrentVersion));

                if ( lbUpdatedAlready && UpdatedEXEs.Client == aeExeName )
                {
                    Env.LogIt("Update completed by another instance. Will try again next cycle.");
                    DoGetCert.bDoCheckOut(true);
                    aoProfile.bExit = true;
                }
            }
            else
            {
                Env.LogIt("");
                Env.LogIt(String.Format("{0} version {1} is installed. Update found ...", lsExeName, lsCurrentVersion));

                bool    lbSkipUpdate = false;
                string  lsNewExePathFile = Path.Combine(lsUpdatePath, Path.GetFileName(lsExePathFile));
                string  lsNewIniPathFile = Path.Combine(lsUpdatePath, Path.GetFileName(aoProfile.sDefaultPathFile));

                if ( File.Exists(lsNewExePathFile) )
                {
                    FileInfo    loFileInfo = new FileInfo(lsNewExePathFile);
                                lbSkipUpdate = loFileInfo.LastWriteTime > DateTime.Now.AddSeconds(-aoProfile.iValue("-UpdateSkipSecs" , 30));
                }

                if ( lbSkipUpdate )
                {
                    Env.LogIt("Update already in progress by another instance. Will try again next cycle.");

                    if ( UpdatedEXEs.Client == aeExeName )
                    {
                        DoGetCert.bDoCheckOut(true);
                        aoProfile.bExit = true;
                    }
                }
                else
                {
                    Env.LogIt(String.Format("Writing update to \"{0}\".", lsNewExePathFile));

                    bool lbUpdateWritten = false;

                    for (int i=0; i < aoProfile.iValue("-UpdateFileWriteMaxRetries", 42); i++)
                    {
                        try
                        {
                            if ( UpdatedEXEs.Client == aeExeName && Directory.Exists(lsUpdatePath) )
                                Directory.Delete(lsUpdatePath, true);

                            Directory.CreateDirectory(lsUpdatePath);
                            File.WriteAllBytes(lsNewExePathFile, lbtArrayExeUpdate);

                            lbUpdateWritten = true;
                            break;
                        }
                        catch
                        {
                            System.Windows.Forms.Application.DoEvents();
                            Thread.Sleep(aoProfile.iValue("-UpdateFileWriteMinRetryMS", 500) + new Random().Next(aoProfile.iValue("-UpdateFileWriteMaxRetryMS", 500)));
                        }
                    }

                    if ( !lbUpdateWritten )
                    {
                        Env.LogIt(String.Format("The update to \"{0}\" could not be written to disk. Will try again next cycle.", lsNewExePathFile));

                        if ( UpdatedEXEs.Client == aeExeName )
                            aoProfile.bExit = true;
                    }
                    else
                    {
                        Env.LogIt(String.Format("The new file version of \"{0}\" is {1}.", lsNewExePathFile, FileVersionInfo.GetVersionInfo(lsNewExePathFile).FileVersion));

                        // Only the client needs this special handling (since it's running).
                        if ( UpdatedEXEs.Client == aeExeName )
                        {
                            Env.LogIt(String.Format("Writing update profile to \"{0}\".", lsNewIniPathFile));
                            File.Copy(lsIniPathFile, lsNewIniPathFile, true);

                            // Schedule a bounce in case the update craps out for whatever reason.
                            // The bounce will be cancelled, assuming the update proceeds normally.
                            Env.ScheduleOnErrorBounce(DoGetCert.oDomainProfile, true);

                            Env.LogIt(String.Format("Starting update \"{0}\" ...", lsNewExePathFile));
                            Process loProcess = new Process();
                                    loProcess.StartInfo.FileName = lsNewExePathFile;
                                    loProcess.StartInfo.Arguments = String.Format("{0}=\"{1}\" {2}=\"{3}\" {4}"
                                            , asRunKey, lsExePathFile, asRplKey, asRplIniPathFile, aoCmdLineProfile.sCommandLine());
                                    loProcess.Start();

                            System.Windows.Forms.Application.DoEvents();
                            aoProfile.bExit = true;
                        }
                    }
                }
            }
        }

        private static void HandleHostUpdates(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            bool        lbRestartHost = false;
            Process     loHostProcess = null;
            string      lsHostExePathFile= Env.sHostExePathFile(out loHostProcess);
            string      lsHostExeVersion = FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion;
            byte[]      lbtArrayGpcExeUpdate = null;

                        // Look for an updated host EXE (unless it was just updated).
                        if ( DateTime.Now > new FileInfo(lsHostExePathFile).LastWriteTime.AddMinutes(DoGetCert.oDomainProfile.iValue("-MinExeUpdateAgeMins", 0)) )
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lbtArrayGpcExeUpdate = loGetCertServiceClient.btArrayGoPcBackupExeUpdate(asHash, abtArrayMinProfile, lsHostExeVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( null != lbtArrayGpcExeUpdate && DateTime.Now > new FileInfo(lsHostExePathFile).LastWriteTime.AddMinutes(DoGetCert.oDomainProfile.iValue("-MinExeUpdateAgeMins", 0)) )
                            Env.LogIt(String.Format("Host version {0} is in use. Update found ...", lsHostExeVersion));

            string      lsHostIniVersion = aoProfile.sValue("-HostIniVersion", "1");
            string      lsHostIniUpdate  = null;

                        // Look for an updated host INI.
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lsHostIniUpdate = loGetCertServiceClient.sGpcIniUpdate(asHash, abtArrayMinProfile, lsHostIniVersion);
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();
                        }
                        if ( !String.IsNullOrEmpty(lsHostIniUpdate) )
                        {
                            Env.LogIt(String.Format("Host profile version {0} is in use. Update found ...", lsHostIniVersion));

                            try
                            {
                                // Validate -AddTasks to (hopefully) avoid inadvertently crippling the host's task manager.
                                tvProfile loTasksProfile = new tvProfile(lsHostIniUpdate).oProfile("-AddTasks").oOneKeyProfile("-Task");

                                foreach (DictionaryEntry loEntry in loTasksProfile)
                                {
                                    tvProfile   loTask = new tvProfile(loEntry.Value.ToString());
                                                loTask.bValue("-CreateNoWindow", false);
                                                loTask.bValue("-IsRerun", false);
                                                loTask.bValue("-OnStartup",false);
                                                loTask.bValue("-RedirectStandardError", false);
                                                loTask.bValue("-RedirectStandardInput", false);
                                                loTask.bValue("-RedirectStandardOutput", false);
                                                loTask.bValue("-TaskDisabled", false);
                                                loTask.bValue("-UnloadOnExit", false);
                                                loTask.bValue("-UseShellExecute", false);

                                                loTask.dtValue("-StartTime", DateTime.MinValue);

                                                loTask.iValue("-DelaySecs", 0);
                                                loTask.iValue("-TimeoutMinutes", 0);
                                }
                            }
                            catch (Exception ex)
                            {
                                Env.LogIt("Host profile update to -AddTasks appears to be corrupt. Can't continue.");
                                throw ex;
                            }
                        }

            if ( null == lbtArrayGpcExeUpdate && String.IsNullOrEmpty(lsHostIniUpdate) && null == loHostProcess )
                lbRestartHost = true;

            if ( null != lbtArrayGpcExeUpdate || !String.IsNullOrEmpty(lsHostIniUpdate) )
            {
                lbRestartHost = true;

                DoGetCert.sStopHostExePathFile();

                // Write the updated EXE (if any).
                if ( null != lbtArrayGpcExeUpdate )
                {
                    // Write the EXE.
                    File.WriteAllBytes(lsHostExePathFile, lbtArrayGpcExeUpdate);

                    Env.LogIt(String.Format("Host update successfully completed. Version {0} is now in use."
                                                            , FileVersionInfo.GetVersionInfo(lsHostExePathFile).FileVersion));
                }

                // Write the updated INI (if any).
                if ( !String.IsNullOrEmpty(lsHostIniUpdate) )
                {
                    tvProfile   loUpdateProfile = new tvProfile(lsHostIniUpdate);
                    tvProfile   loHostProfile = new tvProfile(lsHostExePathFile + ".txt", false);

                                // Look for update task references in the existing list of added tasks. If found, skip them. In other words, within
                                // the new loUpdateProfile, leave out all tasks from loHostProfile that make any references to what's in the update.
                                tvProfile   loTasks = new tvProfile();
                                            // Place tasks in the update on top of the list.
                                            foreach (DictionaryEntry loEntry in loUpdateProfile.oProfile("-AddTasks"))
                                            {
                                                loTasks.Add("-Task", loEntry.Value);
                                            }
                                            // Look for tasks in the host profile that are NOT in the update and NOT in the -RemoveTasks profile.
                                            foreach (DictionaryEntry loEntry in loHostProfile.oProfile("-AddTasks"))
                                            {
                                                tvProfile   loTask = new tvProfile(loEntry.Value.ToString());
                                                string      lsHostTaskExeArgs = (loTask.sValue("-CommandEXE", "") + loTask.sValue("-CommandArgs", "")).ToLower();
                                                string      lsHostTaskExeArgsTime = lsHostTaskExeArgs + loTask.sValue("-StartTime", "").ToLower();
                                                bool        lbAddTask = true;

                                                foreach (DictionaryEntry loEntry2 in loUpdateProfile.oProfile("-AddTasks"))
                                                {
                                                    tvProfile   loTaskUpdate = new tvProfile(loEntry2.Value.ToString());
                                                                if ( lsHostTaskExeArgs == (loTaskUpdate.sValue("-CommandEXE", "") + loTaskUpdate.sValue("-CommandArgs", "")).ToLower() )
                                                                {
                                                                    lbAddTask = false;
                                                                    break;
                                                                }
                                                }
                                                if ( lbAddTask )
                                                    foreach (DictionaryEntry loEntry2 in loUpdateProfile.oProfile("-RemoveTasks"))
                                                    {
                                                        tvProfile   loTaskUpdate = new tvProfile(loEntry2.Value.ToString());
                                                                    if ( lsHostTaskExeArgsTime == (loTaskUpdate.sValue("-CommandEXE", "")
                                                                                                +  loTaskUpdate.sValue("-CommandArgs", "")
                                                                                                +  loTaskUpdate.sValue("-StartTime", "")
                                                                                                ).ToLower()
                                                                            )
                                                                    {
                                                                        lbAddTask = false;
                                                                        break;
                                                                    }
                                                    }

                                                if ( lbAddTask )
                                                    loTasks.Add("-Task", loEntry.Value);
                                            }
                                // Look for update path\file references in the existing backup and cleanup sets. If found, skip them. In other
                                // words, leave out all sets from loHostProfile that make any path\file references to what's in loUpdateProfile.
                                tvProfile   loBackupSets = DoGetCert.oPathFileSetsNotInUpdate(loHostProfile, loUpdateProfile, "-BackupSet", "-FolderToBackup");
                                            // Place backup sets in the update on the bottom of the list.
                                            foreach (DictionaryEntry loEntry in loUpdateProfile.oOneKeyProfile("-BackupSet"))
                                            {
                                                loBackupSets.Add("-BackupSet", loEntry.Value);
                                            }
                                tvProfile   loCleanupSets = DoGetCert.oPathFileSetsNotInUpdate(loHostProfile, loUpdateProfile, "-CleanupSet", "-FilesToDelete");
                                            // Place cleanup sets in the update on the bottom of the list.
                                            foreach (DictionaryEntry loEntry in loUpdateProfile.oOneKeyProfile("-CleanupSet"))
                                            {
                                                loCleanupSets.Add("-CleanupSet", loEntry.Value);
                                            }

                                // The following steps place these 3 keys on top of
                                // the profile for easier access (via text editor).
                                tvProfile   loNewProfile = new tvProfile();
                                tvProfile   loAddTasksProfile = new tvProfile();
                                            foreach (DictionaryEntry loEntry in loTasks)
                                                loAddTasksProfile.Add(loEntry);

                                loNewProfile["-AddTasks"] = loAddTasksProfile.sCommandBlock();
                                foreach (DictionaryEntry loEntry in loBackupSets)
                                    loNewProfile.Add(loEntry);
                                foreach (DictionaryEntry loEntry in loCleanupSets)
                                    loNewProfile.Add(loEntry);
                            
                                loHostProfile.Remove("-AddTasks");
                                loHostProfile.Remove("-BackupSet");
                                loHostProfile.Remove("-CleanupSet");

                                foreach (DictionaryEntry loEntry in loHostProfile)
                                    loNewProfile.Add(loEntry);

                                loHostProfile.Remove("*");

                                foreach (DictionaryEntry loEntry in loNewProfile)
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

                    Env.LogIt(String.Format(
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

                Env.LogIt("Host process restarted successfully.");
            }
        }

        private static void HandleSetupCertUpdate(tvProfile aoProfile, tvProfile aoMinProfile, string asHash, byte[] abtArrayMinProfile)
        {
            X509Certificate2    loOldSetupCertificate = null;
            X509Store           loStore = null;
                                try
                                {
                                    loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                                    loStore.Open(OpenFlags.ReadOnly);

                                    X509Certificate2Collection  loCertCollection = loStore.Certificates.Find(X509FindType.FindBySubjectName, Env.sNewClientSetupCertName, false);
                                                                if ( null != loCertCollection && 0 != loCertCollection.Count )
                                                                    loOldSetupCertificate = loCertCollection[0];
                                }
                                finally
                                {
                                    if ( null != loStore )
                                        loStore.Close();
                                }
            byte[]              lbtArraySetupCertUpdate = null;

                                // Look for setup certificate update.
                                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                                {
                                    lbtArraySetupCertUpdate = loGetCertServiceClient.btArraySetupCertificate(asHash, abtArrayMinProfile
                                                                                            , null == loOldSetupCertificate ? "[none]" : loOldSetupCertificate.Thumbprint);
                                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                        loGetCertServiceClient.Abort();
                                    else
                                        loGetCertServiceClient.Close();
                                }

            if ( null != lbtArraySetupCertUpdate )
            {
                string              lsCertPfxPathFile = aoProfile.sRelativeToProfilePathFile(Guid.NewGuid().ToString("N"));
                string              lsCertificatePassword = HashClass.sHashPw(aoMinProfile);
                                    File.WriteAllBytes(lsCertPfxPathFile, lbtArraySetupCertUpdate);

                X509Certificate2 loNewSetupCertificate = new X509Certificate2(lsCertPfxPathFile, lsCertificatePassword, X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);

                Env.LogIt(String.Format("Setup certificate \"{0}\"=\"{1}\" is in use. Update found ..."
                            , null == loOldSetupCertificate ? Env.sNewClientSetupCertName : loOldSetupCertificate.Subject, null == loOldSetupCertificate ? "[none]" : loOldSetupCertificate.Thumbprint));

                File.Delete(lsCertPfxPathFile);

                try
                {
                    loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.ReadWrite);
                    loStore.Add(loNewSetupCertificate);

                    DoGetCert.RemoveCertificate(loOldSetupCertificate, loStore);

                    string  lsWcfSetupCertNode = "<clientCertificate findValue=\"GetCertClientSetup\" x509FindType=\"FindBySubjectName\" storeLocation=\"LocalMachine\" storeName=\"My\" />";
                    string  lsWcfConfiguration = File.ReadAllText(aoProfile.sLoadedPathFile);
                            if ( !lsWcfConfiguration.Contains(lsWcfSetupCertNode) )
                            {
                                Regex   loRegex = new Regex("(\\s*)<\\/serviceCertificate>(.*?)(\\s*)<\\/clientCredentials>", RegexOptions.Singleline);
                                        if ( !loRegex.IsMatch(lsWcfConfiguration) )
                                            throw new Exception(String.Format("Failed attempting to locate WCF configuration for the \"{0}\"=\"{1}\" certificate.", Env.sNewClientSetupCertName, loNewSetupCertificate.Thumbprint));

                                File.WriteAllText(aoProfile.sLoadedPathFile, loRegex.Replace(lsWcfConfiguration, "$1</serviceCertificate>$1lsWcfSetupCertNode$3</clientCredentials>".Replace("lsWcfSetupCertNode", lsWcfSetupCertNode)));
                                Env.ResetConfigMechanism(aoProfile);

                                Env.LogIt(String.Format("Setup certificate \"{0}\" update successfully completed. Version \"{1}\" is now in use.", Env.sNewClientSetupCertName, loNewSetupCertificate.Thumbprint));
                            }
                }
                finally
                {
                    if ( null != loStore )
                        loStore.Close();
                }
            }
        }

        private static void HandleSvcUpdate(tvProfile aoProfile, string asHash, byte[] abtArrayMinProfile)
        {
            string lsEndpointVersion = aoProfile.sValue("-EndpointVersion", "1");
            string lsSvcUpdate = null;

            // Look for service endpoint update.
            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
            {
                lsSvcUpdate = loGetCertServiceClient.sHostsEntryUpdate(asHash, abtArrayMinProfile, lsEndpointVersion);
                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                    loGetCertServiceClient.Abort();
                else
                    loGetCertServiceClient.Close();
            }

            if ( String.IsNullOrEmpty(lsSvcUpdate) )
            {
                Env.LogIt(String.Format("Service URI version {0} is in use. No update.", lsEndpointVersion));
            }
            else
            {
                Env.LogIt(String.Format("Service URI version {0} is in use. Update found ...", lsEndpointVersion));

                tvProfile       loDnsUpdateProfile = new tvProfile(lsSvcUpdate);
                string          lsOldServiceEndpoint = Env.oGetCertServiceFactory.Endpoint.Address.Uri.ToString();
                string          lsNewServiceEndpoint = loDnsUpdateProfile.sValue("-ServiceEndpoint", "");

                if ( lsNewServiceEndpoint == lsOldServiceEndpoint )
                {
                    Env.LogIt(String.Format("Service URI update already completed by another instance. Version {0} is in effect.", aoProfile.sValue("-EndpointVersion", "1")));
                }
                else
                {
                    tvProfile   loArgs = new tvProfile();
                                loArgs.Add("-OldText", lsOldServiceEndpoint);
                                loArgs.Add("-NewText", lsNewServiceEndpoint);
                                loArgs.Add("-Files", aoProfile.sLoadedPathFile);
                                if ( Path.GetFileName(aoProfile.sLoadedPathFile) == Path.GetFileName(aoProfile.sDefaultPathFile) )
                                    loArgs.Add("-Files", aoProfile.sRelativeToExePathFile(UpdatedEXEs.GcFailSafe.ToString() + ".exe.config"));
                                loArgs.Add("-UseRegularExpressions", false);
                                loArgs.Add("-CopyResultsToSTDOUT", false);

                    // Update WCF configurations with new service endpoint.

                    if ( !Env.bReplaceInFiles(ResourceAssembly.GetName().Name, loArgs, true) )
                    {
                        throw new Exception("Failed replacing service endpoint!");
                    }
                    else
                    {
                        Env.ResetConfigMechanism(aoProfile);
                        Env.oGetCertServiceFactory = null;

                        // Update GetCert profile with current service endpoint version.
                        aoProfile["-EndpointVersion"] = loDnsUpdateProfile.sValue("-EndpointVersion", "1");
                        aoProfile.Save();

                        Env.LogIt(String.Format("Service URI update successfully completed. Version {0} will now take effect.", aoProfile.sValue("-EndpointVersion", "1")));
                    }

                }
            }
        }

        private static void VerifyDependenciesInit()
        {
            string  lsFetchName = null;

            // Fetch IIS Manager DLL.
            tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                    , String.Format("{0}{1}", Env.sFetchPrefix, lsFetchName="Microsoft.Web.Administration.dll"), tvProfile.oGlobal().sRelativeToExePathFile(lsFetchName));
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
                lsErrors = (String.IsNullOrEmpty(lsErrors) ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
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
                    lsErrors = (String.IsNullOrEmpty(lsErrors) ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
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
                        lsErrors = (String.IsNullOrEmpty(lsErrors) ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                                + "PowerShell 5 or later must be installed.";
                    }
                }
            }

            if ( "" != lsErrors )
            {
                tvProfile.oGlobal().bExit = true;

                lsErrors = (String.IsNullOrEmpty(lsErrors) ? "" : lsErrors + Environment.NewLine + Environment.NewLine)
                        + String.Format("Please install {0} and try again.", liErrors == 1 ? "it" : "them");

                if ( !tvProfile.oGlobal().bValue("-NoPrompts", false) )
                    tvMessageBox.Show(null, lsErrors);

                throw new Exception(lsErrors);
            }
        }

        private static void WriteFailSafeFile(string asSanItem, Site aoSite, X509Certificate2 aoCertificate, string asBasePath)
        {
            if ( !tvProfile.oGlobal().bValue("-UseStandAloneMode", true) )
            {
                // Add a certificate expiration status fail-safe public key file.

                string lsPath = Path.Combine(
                          Environment.ExpandEnvironmentVariables(aoSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                        , asBasePath);

                try
                {
                    if ( Directory.Exists(lsPath) )
                        File.WriteAllText(Path.Combine(lsPath, String.Format(tvProfile.oGlobal().sValue("-FailSafeFile", "{0}.txt"), Env.sCertName(aoCertificate)))
                                , Convert.ToBase64String(aoCertificate.Export(X509ContentType.Cert)));
                }
                catch (Exception ex)
                {
                    Environment.ExitCode = 1;
                    Env.LogIt(String.Format("    Note: a fail-safe file can't be written for \"{0}\" (\"{1}\"). Error: \"{2}\"", asSanItem, aoCertificate.Subject, ex.Message));
                }
            }
        }


        /// <summary>
        /// Get digital certificate, install it and bind it to port 443 in IIS.
        /// </summary>
        public bool bGetCertificate()
        {
            bool                lbGetCertificate = true;
            bool                lbCertOverrideApplies = false;
            byte[]              lbtArrayMinProfile = null;
            tvProfile           loMinProfile = DoGetCert.oPreservedKeys;             // Populates a list of preseved keys for future use (ie. after a configuration reset).
            X509Certificate2    loOldCertificate = null;
            X509Certificate2    loNewCertificate = null;
            X509Store           loStore = null;
            string              lsAcmeAccountPathFile = moProfile.sRelativeToExePathFile(moProfile.sValue(gsAcmeAccountFileKey, "Account.xml"));
            string              lsAcmeAccountKeyPathFile = moProfile.sRelativeToExePathFile(moProfile.sValue(gsAcmeAccountKeyFileKey, "AccountKey.xml"));

                                // Reset account files after a change to the -DoStagingTests switch or when not using a DNS challenge.
                                if ( moProfile.bValue("-DoStagingTests", true) != moProfile.bValue("-DoStagingTestsPrevious", false) || !moProfile.bValue("-UseDnsChallenge", false) )
                                {
                                    File.Delete(lsAcmeAccountPathFile);
                                    File.Delete(lsAcmeAccountKeyPathFile);

                                    moProfile["-DoStagingTestsPrevious"] = moProfile.bValue("-DoStagingTests", true);
                                    moProfile.Save();
                                }

            string              lsAcmeWorkPath = null;
            string              lsCertName = null;
            bool                lbCertificatePassword = false;  // Set this "false" to use password only for load balancer file.
            string              lsCertificatePassword = null;
            string              lsCertificateFile = null;
            string              lsCertPfxPathFile = null;
            string              lsNonIISBindingPfxPathFile = null;
            string              lsWellKnownBasePath = moProfile.sValue("-WellKnownBasePath", ".well-known");
            string              lsHash = null;

            this.bMainLoopStopped = false;

            try
            {
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

                Env.LogIt("");
                Env.LogIt(String.Format("Get certificate process started (\"{0}\") ...", lsCertName));

                lsCertificateFile = moProfile.sValue("-InstanceGuid", Guid.NewGuid().ToString("N"));
                lsCertPfxPathFile = moProfile.sRelativeToProfilePathFile(lsCertificateFile);
                moProfile.Remove(gsBindingFailedKey);
                loMinProfile = Env.oMinProfile(moProfile);
                lbtArrayMinProfile = loMinProfile.btArrayZipped();
                lsHash = HashClass.sHashIt(loMinProfile);
                lsCertificatePassword = HashClass.sHashPw(loMinProfile);

                bool                lbAuto = moProfile.bValue("-Auto", false);
                bool                lbSetup = moProfile.bValue("-Setup", false);
                bool                lbCreateSanSitesForCertGet = moProfile.bValue("-CreateSanSitesForCertGet", true);
                bool                lbCreateSanSitesForBinding = moProfile.bValue("-CreateSanSitesForBinding", false);
                bool                lbLoadBalancerCertPending = DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxPending", false);
                bool                lbLoadBalancerCertReady = DoGetCert.oDomainProfile.bValue("-LoadBalancerPfxReady", false);
                bool                lbProceedAnyway = DoGetCert.oDomainProfile.bValue("-Renewing", false) && !DoGetCert.oDomainProfile.bValue("-Renewed", false);
                bool                lbRenewASAP = DoGetCert.oDomainProfile.bValue("-RenewASAP", false);
                bool                lbRepositoryCertsOnly = DoGetCert.oDomainProfile.bValue("-RepositoryCertsOnly", false);
                bool                lbSingleSessionEnabled = moProfile.bValue("-SingleSessionEnabled", false);
                byte[]              lbtArrayNewCertificate = null;
                int                 liCertCount = DoGetCert.oDomainProfile.iValue("-CertCount", 0);
                ServerManager       loServerManager = null;
                                    loOldCertificate = Env.oCurrentCertificate(lsCertName, out loServerManager);
                                    if ( this.bCertificateNotExpiring(lsHash, lbtArrayMinProfile, loOldCertificate) )
                                    {
                                        // The certificate is not ready to expire soon. There is nothing to do.
                                        
                                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                                    }
                DateTime            ldtMaintenanceWindowBeginTime = DoGetCert.oDomainProfile.dtValue(gsMaintWindBegKey, DateTime.MinValue);
                DateTime            ldtMaintenanceWindowEndTime   = DoGetCert.oDomainProfile.dtValue(gsMaintWindEndKey, DateTime.MaxValue);
                string              lsCertOverrideComputerName = DoGetCert.oDomainProfile.sValue("-CertOverrideComputerName", "");
                string              lsCertOverridePfxPathFile = DoGetCert.oDomainProfile.sValue("-CertOverridePfxPathFile", "");
                                    lbCertOverrideApplies = !String.IsNullOrEmpty(lsCertOverrideComputerName) && !String.IsNullOrEmpty(lsCertOverridePfxPathFile);

                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                {
                    if ( lbAuto && !lbSetup && null != loOldCertificate && !String.IsNullOrEmpty(DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "")) )
                    {
                        if ( lbLoadBalancerCertReady )
                        {
                            if ( DateTime.Now < DoGetCert.oDomainProfile.dtValue("-RenewalCertificateBindingDate", DateTime.MinValue) )
                            {
                                Env.LogIt(String.Format("Nothing to do until the certificate binding date (\"{0}\").{1}"
                                        , DoGetCert.oDomainProfile.dtValue("-RenewalCertificateBindingDate", DateTime.MinValue).ToShortDateString()
                                        , !lbProceedAnyway ? "" : " Proceeding anyway since renewal has already commenced elsewhere."
                                        ));

                                if ( !lbProceedAnyway )
                                    return DoGetCert.bDoCheckOut(lbGetCertificate);
                            }
                        }
                        else
                        {
                            if ( !lbLoadBalancerCertPending )
                            {
                                bool    lbRetry = false;
                                        DoGetCert.ClearCache();
                                        if ( Env.sComputerName == DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "") )
                                        {
                                            lbRetry = lbRepositoryCertsOnly && liCertCount < 2;
                                        }
                                        else
                                        {
                                            lbRetry =   liCertCount > 1
                                                    || (liCertCount < 2
                                                        && (lbRepositoryCertsOnly || !DoGetCert.oDomainProfile.bValue("-LoadBalancerRepositoryCertsOnly", false)));
                                        }

                                if ( lbRetry )
                                {
                                    Env.LogIt("The load balancer administrator or process has not yet received the new certificate.");

                                    if ( Env.sComputerName != DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "") )
                                    {
                                        Env.LogIt("Will try again next cycle.");
                                        Env.LogIt("");

                                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                                    }
                                    else
                                    {
                                        lbRetry = Env.bCanScheduleNonErrorBounce;

                                        if ( lbRetry )
                                        {
                                            Env.LogIt("Will try again in several minutes.");

                                            Env.iNonErrorBounceSecs = moProfile.iValue("-LoadBalancerBounceSecs", 900);
                                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                                        }
                                        else
                                        {
                                            Env.iNonErrorBounceSecs = -1;
                                            lbGetCertificate = false;

                                            Env.LogIt("");
                                            Env.LogIt("This is an error.");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Env.LogIt("The load balancer administrator or process has not yet released the new certificate.");

                                int liErrorCount = 1  + moProfile.iValue("-LoadBalancerPendingErrorCount", 0);
                                    if ( liErrorCount > moProfile.iValue("-LoadBalancerPendingMaxErrors",  1) )
                                    {
                                        lbGetCertificate = false;
                                        Env.iNonErrorBounceSecs = -1;

                                        Env.LogIt("This is an error.");
                                    }
                                    else
                                    {
                                        Env.LogIt("Will try again next cycle.");
                                        Env.LogIt("");

                                        moProfile["-LoadBalancerPendingErrorCount"] = liErrorCount;
                                        moProfile.Save();

                                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                                    }
                            }
                        }
                    }

                    if (       lbGetCertificate && !this.bMainLoopStopped
                            && lbAuto && !lbSetup
                            && (lbRepositoryCertsOnly || liCertCount > 1)
                            && (lbLoadBalancerCertReady || Env.sComputerName != DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", ""))
                            && DateTime.Now < DoGetCert.oDomainProfile.dtValue("-RenewalCertificateBindingDate", DateTime.MinValue)
                            && null != loOldCertificate
                            )
                    {
                        Env.LogIt(String.Format("Nothing to do until the certificate binding date (\"{0}\").{1}"
                                , DoGetCert.oDomainProfile.dtValue("-RenewalCertificateBindingDate", DateTime.MinValue).ToShortDateString()
                                , !lbProceedAnyway ? "" : " Proceeding anyway since renewal has already commenced elsewhere."
                                ));

                        if ( !lbProceedAnyway )
                            return DoGetCert.bDoCheckOut(lbGetCertificate);
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped )
                    {
                        if (       lbAuto && !lbSetup
                                && (DateTime.Now < ldtMaintenanceWindowBeginTime || DateTime.Now >= ldtMaintenanceWindowEndTime)
                                && (lbLoadBalancerCertReady || Env.sComputerName != DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", ""))
                                && null != loOldCertificate
                                )
                        {
                            bool lbProceedAnywayRenewASAP = !lbRepositoryCertsOnly && lbRenewASAP;

                            Env.LogIt(String.Format("{0}The \"{1}\" maintenance window is \"{2}\" to \"{3}\".{4}", lbProceedAnyway || lbProceedAnywayRenewASAP ? "" : "Can't continue. "
                                                    , lsCertName, ldtMaintenanceWindowBeginTime, ldtMaintenanceWindowEndTime
                                                    , lbProceedAnyway ? " Proceeding anyway since renewal has already commenced elsewhere."
                                                            : !lbProceedAnywayRenewASAP ? "" : " Proceeding anyway since the domain's 'RenewASAP' flag has been set."
                                                    ));
                            Env.LogIt("");

                            if ( !lbProceedAnyway && !lbProceedAnywayRenewASAP )
                                return DoGetCert.bDoCheckOut(lbGetCertificate);
                        }

                        using (GetCertService.IGetCertServiceChannel moGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lbtArrayNewCertificate = moGetCertServiceClient.btArrayNewCertificate(lsHash, lbtArrayMinProfile);
                            if ( CommunicationState.Faulted == moGetCertServiceClient.State )
                                moGetCertServiceClient.Abort();
                            else
                                moGetCertServiceClient.Close();

                            if ( null == lbtArrayNewCertificate )
                            {
                                Env.LogIt("No new certificate was found in the repository.");
                            }
                            else
                            {
                                Env.LogIt("New certificate downloaded from the repository.");

                                // Buffer certificate to disk (same as what's done below).

                                string  lsCertPfxPath = Path.GetDirectoryName(lsCertPfxPathFile);
                                        if ( !String.IsNullOrEmpty(lsCertPfxPath) )
                                            Directory.CreateDirectory(lsCertPfxPath);

                                File.WriteAllBytes(lsCertPfxPathFile, lbtArrayNewCertificate);

                                // Upload new certificate to the load balancer.
                                if ( lbGetCertificate && !this.bMainLoopStopped && !lbLoadBalancerCertPending && !lbLoadBalancerCertReady )
                                {
                                    if ( lbCertOverrideApplies )
                                        lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPfxPathFile
                                                                                            , HashClass.sDecrypted(Env.oChannelCertificate, DoGetCert.oDomainProfile.sValue("-CertOverridePfxPassword", "")));
                                    else
                                        lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPfxPathFile, lsCertificatePassword);

                                    if ( lbGetCertificate && !this.bMainLoopStopped && Env.sComputerName == DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "") )
                                    {
                                        Env.LogIt("");

                                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                                    }
                                }
                            }
                        }
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && lbRepositoryCertsOnly && !lbCertOverrideApplies )
                    {
                        int liErrorCount = 1  + moProfile.iValue("-RepositoryCertsOnlyErrorCount", 0);
                            if ( liErrorCount > moProfile.iValue("-RepositoryCertsOnlyMaxErrors",  6) )
                            {
                                Env.iNonErrorBounceSecs = -1;
                                lbGetCertificate = false;

                                Env.LogIt("This is an error since -RepositoryCertsOnly=True.");
                            }
                            else
                            {
                                if ( lbLoadBalancerCertReady || Env.sComputerName != DoGetCert.oDomainProfile.sValue("-LoadBalancerComputerName", "") )
                                {
                                    Env.LogIt("-RepositoryCertsOnly=True. Will try again next cycle.");
                                    Env.LogIt("");

                                    moProfile["-RepositoryCertsOnlyErrorCount"] = liErrorCount;
                                    moProfile.Save();

                                    return DoGetCert.bDoCheckOut(lbGetCertificate);
                                }
                                else
                                if ( !lbLoadBalancerCertPending && !lbLoadBalancerCertReady )
                                {
                                    bool    lbRetry = Env.bCanScheduleNonErrorBounce;
                                            Env.LogIt("");
                                            Env.LogIt("This is the -LoadBalancerComputerName and -RepositoryCertsOnly=True. " + (!lbRetry ? "This is an error." : "Will try again in several minutes."));

                                            if ( !lbRetry )
                                            {
                                                Env.iNonErrorBounceSecs = -1;
                                                lbGetCertificate = false;
                                            }
                                            else
                                            {
                                                Env.iNonErrorBounceSecs = moProfile.iValue("-LoadBalancerBounceSecs", 900);
                                                return DoGetCert.bDoCheckOut(lbGetCertificate);
                                            }
                                }
                            }
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && (!lbRepositoryCertsOnly || lbCertOverrideApplies) )
                {
                    if ( lbCertOverrideApplies )
                    {
                        Env.LogIt("");

                        if ( Env.sComputerName == lsCertOverrideComputerName )
                        {
                            Env.LogIt(String.Format("A new certificate for \"{0}\" will come from the certificate override PFX file.", lsCertName));

                            lsCertPfxPathFile = lsCertOverridePfxPathFile;
                        }
                        else
                        {
                            Env.LogIt(String.Format("The certificate override for \"{0}\" will be found on another server (ie. \"{1}\").", lsCertName, lsCertOverrideComputerName));

                            int liErrorCount = 1  + moProfile.iValue("-CertOverrideErrorCount", 0);
                                if ( liErrorCount > moProfile.iValue("-CertOverrideMaxErrors",  3) )
                                {
                                    lbGetCertificate = false;
                                    Env.iNonErrorBounceSecs = -1;

                                    Env.LogIt("No more retries. This is an error.");
                                    Env.LogIt("");
                                }
                                else
                                {
                                    Env.LogIt("Will try again next cycle.");
                                    Env.LogIt("");

                                    moProfile["-CertOverrideErrorCount"] = liErrorCount;
                                    moProfile.Save();

                                    return DoGetCert.bDoCheckOut(lbGetCertificate);
                                }
                        }
                    }
                    else
                    if ( !moProfile.bValue("-UseStandAloneMode", true) && !DoGetCert.oDomainProfile.bValue("-ClientDownloadsEnabled", true) )
                    {
                        Env.LogIt("Nothing to do until certificate downloads are enabled.");

                        return DoGetCert.bDoCheckOut(lbGetCertificate);
                    }
                    else
                    {
                        // There is no certificate override file. So proceed to get a new certificate from the certificate provider network.

                        bool    lbLockCertificateRenewal = false;
                                DoGetCert.ClearCache();
                                if ( liCertCount < 2 && !moProfile.bValue("-UseStandAloneMode", true) )
                                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                                {
                                    lbLockCertificateRenewal = loGetCertServiceClient.bLockCertificateRenewal(lsHash, lbtArrayMinProfile);
                                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                        loGetCertServiceClient.Abort();
                                    else
                                        loGetCertServiceClient.Close();
                                }
                                // If the renewal can't be locked, another client must be doing the
                                // certificate renewal already (or not all clients renewed previously or
                                // certificate overrides are reserved to another computer on the domain
                                // or the new certificate has already been acquired by another client).
                                // Try again next time.
                                if ( lbLockCertificateRenewal )
                                {
                                    Env.LogIt("Certificate renewal has been locked for this domain.");
                                }
                                else
                                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                                {
                                    Env.LogIt("Certificate renewal can't be locked. Will try again next cycle.");
                                    Env.LogIt("");

                                    return DoGetCert.bDoCheckOut(lbGetCertificate);
                                }

                        Env.LogIt("");
                        Env.LogIt(String.Format("Retrieving new certificate for \"{0}\" from the certificate provider network ...", lsCertName));

                        string  lsAcmePsFile = "ACME-PS.zip";
                        string  lsAcmePsPath = Path.GetFileNameWithoutExtension(lsAcmePsFile);
                        string  lsAcmePsPathFull = moProfile.sRelativeToExePathFile(lsAcmePsPath);
                        string  lsAcmePsPathFile = moProfile.sRelativeToExePathFile(lsAcmePsFile);
                                // Fetch ACME-PS module.
                                tvFetchResource.ToDisk(ResourceAssembly.GetName().Name
                                        , String.Format("{0}{1}", Env.sFetchPrefix, lsAcmePsFile), lsAcmePsPathFile);
                                if ( !Directory.Exists(lsAcmePsPathFull) )
                                {
                                    string  lsZipDir = Path.GetDirectoryName(lsAcmePsPathFile);
                                            ZipFile.ExtractToDirectory(lsAcmePsPathFile, String.IsNullOrEmpty(lsZipDir) ? "." : lsZipDir);
                                }
                        string  lsAcmePsVer = Path.GetFileName(Directory.GetDirectories(lsAcmePsPathFull)[0]);

                        // Don't use the PowerShell gallery installed ACME-PS module (by default, use what's embedded instead).
                        if ( !moProfile.bValue("-AcmePsModuleUseGallery", false) )
                        {
                            string lsAcmePsPathFolder = moProfile.sValue("-AcmePsPath", lsAcmePsPath);

                            moProfile.sValue("-AcmePsPathHelp", String.Format(
                                    "Any alternative -AcmePsPath must include a subfolder named \"{0}\" that contains the files.", lsAcmePsPathFolder));

                            lsAcmePsPath = moProfile.sRelativeToExePathFile(moProfile.sValue("-AcmePsPath", lsAcmePsPath));

                            if ( String.IsNullOrEmpty(Path.GetDirectoryName(lsAcmePsPath)) )
                                lsAcmePsPath = Path.Combine(".", lsAcmePsPath);
                        }

                        string      lsSanPsStringArray = "(";
                                    foreach (string lsSanItem in lsSanArray)
                                    {
                                        lsSanPsStringArray += String.Format(",\"{0}\"",  lsSanItem);
                                    }
                                    lsSanPsStringArray = lsSanPsStringArray.Replace("(,", "(") + ")";

                        // Append primary certificate name to prevent conflict with multiple instances.
                        lsAcmeWorkPath = moProfile.sRelativeToProfilePathFile(
                                String.Format("{0}_{1}", moProfile.sValue("-AcmePsWorkPath", "AcmeState"), lsCertName));

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            this.LogStage("1 - init ACME workspace");

                            // Delete ACME workspace.
                            if ( moProfile.bValue("-CleanupAcmeWork", true) && Directory.Exists(lsAcmeWorkPath) )
                                Directory.Delete(lsAcmeWorkPath, true);

                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptSessionOpen", @"
New-PSSession -ComputerName localhost -Name GetCert
$session = Disconnect-PSSession -Name GetCert
                                    "), true);

                            Env.LogIt("");

                            if ( lbGetCertificate )
                            {
                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage1s", @"
Using Module ""{AcmePsPath}""   # Version: {AcmePsVer}
{AcmeSystemWide}
$global:state = New-ACMEState -Path ""{AcmeWorkPath}""
Get-ACMEServiceDirectory $global:state -ServiceName ""{AcmeServiceName}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmePsVer}", lsAcmePsVer)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            .Replace("{AcmeServiceName}", moProfile.bValue("-DoStagingTests", true) ? moProfile.sValue("-ServiceNameStaging", "LetsEncrypt-Staging")
                                                                                                                    : moProfile.sValue("-ServiceNameLive", "LetsEncrypt"))
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage1", @"
Import-Module ""{AcmePsPath}""   # Version: {AcmePsVer}
{AcmeSystemWide}
New-ACMEState      -Path ""{AcmeWorkPath}""
Get-ACMEServiceDirectory ""{AcmeWorkPath}"" -ServiceName ""{AcmeServiceName}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmePsVer}", lsAcmePsVer)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
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

                            bool    lbUseNewAccount = !File.Exists(lsAcmeAccountPathFile) || !File.Exists(lsAcmeAccountKeyPathFile);
                            string  lsAcmeWorkAccountPathFile = Path.Combine(lsAcmeWorkPath, moProfile.sValue(gsAcmeAccountFileKey, "Account.xml"));
                            string  lsAcmeWorkAccountKeyPathFile = Path.Combine(lsAcmeWorkPath, moProfile.sValue(gsAcmeAccountKeyFileKey, "AccountKey.xml"));
                                    if ( !lbUseNewAccount )
                                    {
                                        File.Copy(lsAcmeAccountPathFile, lsAcmeWorkAccountPathFile, true);
                                        File.Copy(lsAcmeAccountKeyPathFile, lsAcmeWorkAccountKeyPathFile, true);
                                    }

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$SanList = {SanPsStringArray}
if ( {UseNewAccount} )
{
    New-ACMENonce      $global:state
    New-ACMEAccountKey $global:state
    New-ACMEAccount    $global:state -EmailAddresses ""{-ContactEmailAddress}"" -AcceptTOS -PassThru
}
$global:order = New-ACMEOrder $global:state -Identifiers $SanList

$global:authZ = Get-ACMEAuthorization $global:state -Order $global:order

[int[]] $global:SanMap = $null
        foreach ($SAN in $SanList) { for ($i=0; $i -lt $global:authZ.Length; $i++) { if ( $global:authZ[$i].Identifier.value -eq $SAN ) { $global:SanMap += $i }}}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{UseNewAccount}", lbUseNewAccount ? "$true" : "$false")
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage2", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
if ( {UseNewAccount} )
{
    New-ACMENonce      ""{AcmeWorkPath}""
    New-ACMEAccountKey ""{AcmeWorkPath}""
    New-ACMEAccount    ""{AcmeWorkPath}"" -EmailAddresses ""{-ContactEmailAddress}"" -AcceptTOS -PassThru
}
New-ACMEOrder ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{UseNewAccount}", lbUseNewAccount ? "$true" : "$false")
                                        );

                            if ( lbGetCertificate && lbUseNewAccount )
                            {
                                File.Copy(lsAcmeWorkAccountPathFile, lsAcmeAccountPathFile, true);
                                File.Copy(lsAcmeWorkAccountKeyPathFile, lsAcmeAccountKeyPathFile, true);
                            }
                        }

                        ArrayList loAcmeCleanedList = new ArrayList();

                        for (int liSanArrayIndex=0; liSanArrayIndex < lsSanArray.Length; liSanArrayIndex++)
                        {
                            string lsSanItem = lsSanArray[liSanArrayIndex];

                            if ( lbGetCertificate && !this.bMainLoopStopped )
                            {
                                this.LogStage(String.Format("3 - Define DNS name to be challenged (\"{0}\"), setup domain challenge and submit it to certificate provider", lsSanItem));

                                string  lsAcmePath = null;

                                if ( !moProfile.bValue("-UseDnsChallenge", false) )
                                {
                                    Site    loSanSite = null;
                                    Site    loPrimarySiteForDefaults = null;
                                            Env.oCurrentCertificate(lsSanItem, lsSanArray, out loServerManager, out loSanSite, out loPrimarySiteForDefaults);

                                    // If -CreateSanSitesForCertGet is false, only create a default
                                    // website (as needed) and use it for all SAN values.
                                    if ( null == loSanSite && !lbCreateSanSitesForCertGet && 0 != loServerManager.Sites.Count )
                                    {
                                        Env.LogIt(String.Format("No website found for \"{0}\". -CreateSanSitesForCertGet is \"False\". No website created.\r\n", lsSanItem));

                                        loSanSite = loServerManager.Sites[0];
                                    }

                                    if ( null == loSanSite && lbCreateSanSitesForCertGet )
                                        loSanSite = this.oSanSiteCreated(lsSanItem, loServerManager, loPrimarySiteForDefaults);

                                    string  lsAcmeBasePath = moProfile.sValue("-AcmeBasePath", Path.Combine(lsWellKnownBasePath, "acme-challenge"));
                                            lsAcmePath = Path.Combine(
                                              Environment.ExpandEnvironmentVariables(loSanSite.Applications["/"].VirtualDirectories["/"].PhysicalPath)
                                            , lsAcmeBasePath);

                                    // Cleanup ACME folder.
                                    if ( !loAcmeCleanedList.Contains(lsAcmePath) )
                                    {
                                        if ( Directory.Exists(lsAcmePath) )
                                        {
                                            Directory.Delete(lsAcmePath, true);
                                            Thread.Sleep(moProfile.iValue("-AcmeCleanupSleepMS", 200));
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
                                }

                                if ( lbGetCertificate && !this.bMainLoopStopped )
                                {
                                    if ( !moProfile.bValue("-UseDnsChallenge", false) )
                                    {
                                        if ( lbSingleSessionEnabled )
                                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$challenge = Get-ACMEChallenge $global:state $global:authZ[$global:SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmeChallengePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine
Start-Sleep -Seconds {HttpChallengeSleepSecs}

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge $global:state
                                                    ")
                                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                                    .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                    .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                    .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                    .Replace("{AcmeChallengePath}", lsAcmePath)
                                                    .Replace("{HttpChallengeSleepSecs}", moProfile.iValue("-HttpChallengeSleepSecs", 0).ToString())
                                                    );
                                        else
                                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$SanList = {SanPsStringArray}
$order   = Find-ACMEOrder        ""{AcmeWorkPath}"" -Identifiers $SanList
$authZ   = Get-ACMEAuthorization ""{AcmeWorkPath}"" -Order $order

[int[]] $SanMap = $null; foreach ($SAN in $SanList) { for ($i=0; $i -lt $authZ.Length; $i++) { if ( $authZ[$i].Identifier.value -eq $SAN ) { $SanMap += $i }}}

$challenge = Get-ACMEChallenge ""{AcmeWorkPath}"" $authZ[$SanMap[{SanArrayIndex}]] ""http-01""

$challengePath = ""{AcmeChallengePath}""
$fileName = $challengePath + ""/"" + $challenge.Data.Filename
if(-not (Test-Path $challengePath)) { New-Item -Path $challengePath -ItemType Directory }
Set-Content -Path $fileName -Value $challenge.Data.Content -NoNewLine
Start-Sleep -Seconds {HttpChallengeSleepSecs}

$challenge.Data.AbsoluteUrl
$challenge | Complete-ACMEChallenge ""{AcmeWorkPath}""
                                                    ")
                                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                                    .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                    .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                    .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                    .Replace("{AcmeChallengePath}", lsAcmePath)
                                                    .Replace("{HttpChallengeSleepSecs}", moProfile.iValue("-HttpChallengeSleepSecs", 0).ToString())
                                                    );
                                    }
                                    else
                                    {
                                        if ( lbSingleSessionEnabled )
                                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3sDns", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$challenge = Get-ACMEChallenge $global:state $global:authZ[$global:SanMap[{SanArrayIndex}]] ""dns-01""

echo DNS challenge command-line goes here (-Name $challenge.Data.TxtRecordName -Value $challenge.Data.Content).
Start-Sleep -Seconds {DnsChallengeSleepSecs}

$challenge | Complete-ACMEChallenge $global:state
                                                    ")
                                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                                    .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                    .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                    .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                    .Replace("{DnsChallengeSleepSecs}", moProfile.iValue("-DnsChallengeSleepSecs", 15).ToString())
                                                    );
                                        else
                                            lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage3Dns", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$SanList = {SanPsStringArray}
$order   = Find-ACMEOrder        ""{AcmeWorkPath}"" -Identifiers $SanList
$authZ   = Get-ACMEAuthorization ""{AcmeWorkPath}"" -Order $order

[int[]] $SanMap = $null; foreach ($SAN in $SanList) { for ($i=0; $i -lt $authZ.Length; $i++) { if ( $authZ[$i].Identifier.value -eq $SAN ) { $SanMap += $i }}}

$challenge = Get-ACMEChallenge ""{AcmeWorkPath}"" $authZ[$SanMap[{SanArrayIndex}]] ""dns-01""

echo ""Command-line to add a 'DNS Challenge' TXT record goes here (-Name $challenge.Data.TxtRecordName -Value $challenge.Data.Content).""
Start-Sleep -Seconds {DnsChallengeSleepSecs}

$challenge | Complete-ACMEChallenge ""{AcmeWorkPath}""
                                                    ")
                                                    .Replace("{AcmePsPath}", lsAcmePsPath)
                                                    .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                                    .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                                    .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                                    .Replace("{SanArrayIndex}", liSanArrayIndex.ToString())
                                                    .Replace("{DnsChallengeSleepSecs}", moProfile.iValue("-DnsChallengeSleepSecs", 15).ToString())
                                                    );
                                    }
                                }
                            }
                        }

                        int liSubmissionRetries = moProfile.iValue("-SubmissionRetries", 42);
                        int liSubmissionWaitSecs = moProfile.iValue("-SubmissionWaitSecs", 5);

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                if ( 0 != liSubmissionWaitSecs )
                                {
                                    DateTime ltdTimeout = DateTime.Now.AddSeconds(liSubmissionWaitSecs);

                                    Env.LogIt("");
                                    Env.LogIt(String.Format("Waiting {0} second{1} before requesting challenge update ..."
                                                            , liSubmissionWaitSecs, 1 == liSubmissionWaitSecs ? "" : "s"));

                                    while ( !this.bMainLoopStopped && DateTime.Now < ltdTimeout )
                                    {
                                        System.Windows.Forms.Application.DoEvents();
                                        System.Threading.Thread.Sleep(moProfile.iValue("-UpdateChallengeSleepMS", 200));
                                    }
                                }

                                this.LogStage(String.Format("4 - update challenge{0} from certificate provider", 1 == lsSanArray.Length ? "" : "s"));

                                string  lsSubmissionPending = ": pending";
                                string  lsOutput = null;

                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage4s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$global:order | Update-ACMEOrder $global:state -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage4", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$order = Find-ACMEOrder   ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$order | Update-ACMEOrder ""{AcmeWorkPath}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );

                                if ( !lsOutput.Contains(lsSubmissionPending) || this.bMainLoopStopped )
                                    break;
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && moProfile.bValue("-UseDnsChallenge", false) )
                        {
                            Env.LogIt("");
                            Env.LogIt("Cleaning up ACME DNS records ...");

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptDnsCleanupS", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}

for ($i=0; $i -lt $authZ.Length; $i++)
{
    $challenge = Get-ACMEChallenge $global:state $global:$authZ[$i] ""dns-01""
    echo ""Command-line to remove a 'DNS Challenge' TXT record goes here (-Name $challenge.Data.TxtRecordName -Value $challenge.Data.Content -Remove -ErrorAction SilentlyContinue).""
}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptDnsCleanup", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$SanList = {SanPsStringArray}
$order   = Find-ACMEOrder        ""{AcmeWorkPath}"" -Identifiers $SanList
$authZ   = Get-ACMEAuthorization ""{AcmeWorkPath}"" -Order $order

for ($i=0; $i -lt $authZ.Length; $i++)
{
    $challenge = Get-ACMEChallenge ""{AcmeWorkPath}"" $authZ[$i] ""dns-01""
    echo Remove DNS challenge command-line goes here (-Name $challenge.Data.TxtRecordName -Value $challenge.Data.Content -Remove).
}
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            int liCertificateRequestWaitSecs = moProfile.iValue("-CertificateRequestWaitSecs", 10);
                                if ( 0 != liCertificateRequestWaitSecs )
                                {
                                    Env.LogIt("");
                                    Env.LogIt(String.Format("Waiting {0} second{1} before generating certificate request ..."
                                                    , liCertificateRequestWaitSecs, 1 == liCertificateRequestWaitSecs ? "" : "s"));

                                    Thread.Sleep(1000 * liCertificateRequestWaitSecs);
                                }

                            this.LogStage("5 - generate certificate request and submit");

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage5s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$global:certKey = New-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Complete-ACMEOrder $global:state -Order $global:order -CertificateKey $global:certKey
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage5", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$order = Find-ACMEOrder ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$certKey = New-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Complete-ACMEOrder      ""{AcmeWorkPath}"" -Order $order -CertificateKey $certKey
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        );
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            for (int i=0; i < liSubmissionRetries; ++i)
                            {
                                if ( 0 != liSubmissionWaitSecs )
                                {
                                    DateTime ltdTimeout = DateTime.Now.AddSeconds(liSubmissionWaitSecs);

                                    Env.LogIt("");
                                    Env.LogIt(String.Format("Waiting {0} second{1} before updating certificate request ..."
                                                            , liSubmissionWaitSecs, 1 == liSubmissionWaitSecs ? "" : "s"));

                                    while ( !this.bMainLoopStopped && DateTime.Now < ltdTimeout )
                                    {
                                        System.Windows.Forms.Application.DoEvents();
                                        System.Threading.Thread.Sleep(moProfile.iValue("-UpdateChallengeSleepMS", 200));
                                    }
                                }

                                this.LogStage("6 - update certificate request");

                                string  lsSubmissionPending = "CertificateRequest       : \r\nCrtPemFile               :";
                                string  lsOutput = null;

                                if ( lbSingleSessionEnabled )
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage6s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
$global:order | Update-ACMEOrder $global:state -PassThru;
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                            .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                            .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                            );
                                else
                                    lbGetCertificate = this.bRunPowerScript(out lsOutput, moProfile.sValue("-ScriptStage6", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$order = Find-ACMEOrder   ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$order | Update-ACMEOrder ""{AcmeWorkPath}"" -PassThru
                                            ")
                                            .Replace("{AcmePsPath}", lsAcmePsPath)
                                            .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
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

                            File.Delete(lsCertPfxPathFile);

                            if ( lbSingleSessionEnabled )
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage7s", @"
Using Module ""{AcmePsPath}""
{AcmeSystemWide}
Export-ACMECertificate $global:state -Order $global:order -CertificateKey $global:certKey -Path ""{CertificatePathFile}""
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{CertificateFile}", lsCertificateFile)
                                        .Replace("{CertificatePathFile}", lsCertPfxPathFile)
                                        .Replace("{CertificatePassword}", lsCertificatePassword)
                                        );
                            else
                                lbGetCertificate = this.bRunPowerScript(moProfile.sValue("-ScriptStage7", @"
Import-Module ""{AcmePsPath}""
{AcmeSystemWide}
$order = Find-ACMEOrder ""{AcmeWorkPath}"" -Identifiers {SanPsStringArray}
$certKey = Import-ACMECertificateKey -Path ""{AcmeWorkPath}\cert.key.xml""
Export-ACMECertificate  ""{AcmeWorkPath}"" -Order $order -CertificateKey $certKey -Path ""{CertificatePathFile}""
                                        ")
                                        .Replace("{AcmePsPath}", lsAcmePsPath)
                                        .Replace("{AcmeSystemWide}", moProfile.sValue("-AcmeSystemWide", ""))
                                        .Replace("{AcmeWorkPath}", lsAcmeWorkPath)
                                        .Replace("{SanPsStringArray}", lsSanPsStringArray)
                                        .Replace("{CertificateFile}", lsCertificateFile)
                                        .Replace("{CertificatePathFile}", lsCertPfxPathFile)
                                        .Replace("{CertificatePassword}", lsCertificatePassword)
                                        );
                        }
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped )
                {
                    // At this point there must be a new certificate file ready on disk (ie. for: uploads, local cert store, port binding, etc).
                    lbGetCertificate = File.Exists(lsCertPfxPathFile);

                    if ( !lbGetCertificate )
                        Env.LogIt(Env.sExceptionMessage("The new certificate could not be found!"));

                    if ( lbGetCertificate && !this.bMainLoopStopped && null == lbtArrayNewCertificate && !moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // Upload new certificate to the load balancer.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            if ( lbCertOverrideApplies )
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPfxPathFile
                                                                                    , HashClass.sDecrypted(Env.oChannelCertificate, DoGetCert.oDomainProfile.sValue("-CertOverridePfxPassword", "")));
                            else
                                lbGetCertificate = this.bUploadCertToLoadBalancer(lsHash, lbtArrayMinProfile, lsCertName, lsCertPfxPathFile, lsCertificatePassword);
                        }

                        // Upload new certificate to the repository.
                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                        {
                            lbGetCertificate = loGetCertServiceClient.bNewCertificateUploaded(lsHash, lbtArrayMinProfile, File.ReadAllBytes(lsCertPfxPathFile));
                            if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                loGetCertServiceClient.Abort();
                            else
                                loGetCertServiceClient.Close();

                            Env.LogIt("");
                            if ( lbGetCertificate )
                                Env.LogIt("Certificate successfully uploaded to the repository.");
                            else
                                Env.LogIt(Env.sExceptionMessage("Failed uploading new certificate to repository."));
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped )
                        {
                            X509Certificate2 loCertAcquired = new X509Certificate2(lsCertPfxPathFile, lbCertificatePassword ? lsCertificatePassword : (string)null);

                            Env.LogIt("");
                            Env.LogIt(String.Format("New certificate thumbprint: {0} (\"{1}\")", loCertAcquired.Thumbprint, Env.sCertName(loCertAcquired)));

                            string  lsScriptCertAcquired  = moProfile.sValue("-ScriptCertAcquired", "");
                                    if ( !String.IsNullOrEmpty(lsScriptCertAcquired) )
                                    {
                                        Env.LogIt("");
                                        this.bRunPowerScript(lsScriptCertAcquired
                                                .Replace("{CertificateDomainName}", lsCertName)
                                                .Replace("{OldCertificateThumbprint}", null == loOldCertificate ? "" : loOldCertificate.Thumbprint)
                                                .Replace("{NewCertificateThumbprint}", loCertAcquired.Thumbprint)
                                                );
                                    }
                        }
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped && null != lbtArrayNewCertificate || moProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // A non-null "lbtArrayNewCertificate" means the certificate was just downloaded from the
                        // repository. Load it to be installed and bound locally (same if running stand-alone).
                        if ( !moProfile.bValue("-CertificatePrivateKeyExportable", false)
                                || ( !moProfile.bValue("-UseStandAloneMode", true) && !DoGetCert.oDomainProfile.bValue("-CertPrivateKeyExportAllowed", false)) )
                        {
                            loNewCertificate = new X509Certificate2(lsCertPfxPathFile, lbCertificatePassword ? lsCertificatePassword : (string)null
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);
                        }
                        else 
                        {
                            Env.LogIt("");
                            Env.LogIt("The new certificate private key is exportable (\"-CertificatePrivateKeyExportable=True\"). This is not recommended.");

                            loNewCertificate = new X509Certificate2(lsCertPfxPathFile, lbCertificatePassword ? lsCertificatePassword : (string)null
                                    , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);
                        }
                    }
                }

                if ( lbGetCertificate && !this.bMainLoopStopped && (null != lbtArrayNewCertificate || moProfile.bValue("-UseStandAloneMode", true)) )
                {
                    // Install and bind the new certificate locally, if: 1) the certificate was just downloaded from the repository; or 2) we're running stand-alone.
                    // Bottom line: for server farms, never get a new cert (from override file or certifcate network) and bind it during the same maintenance window.

                    Env.LogIt("");
                    Env.LogIt(String.Format("Certificate thumbprint: {0} (\"{1}\")", loNewCertificate.Thumbprint, Env.sCertName(loNewCertificate)));

                    Env.LogIt("");
                    Env.LogIt("Install and bind certificate ...");

                    // Select the local machine certificate store (ie "Local Computer / Personal / Certificates").
                    loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                    loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                    // Add the new cert to the certifcate store.
                    loStore.Add(loNewCertificate);

                    // Do SSO stuff (while both old and new certs are still in the certifcate store).
                    if ( null != loOldCertificate && loOldCertificate.Thumbprint != loNewCertificate.Thumbprint )
                    {
                        lbGetCertificate = this.bApplyNewCertToSSO(loOldCertificate, loNewCertificate, lsHash, lbtArrayMinProfile);

                        if ( !lbGetCertificate || (lbGetCertificate && !this.bMainLoopStopped && 0 != Env.iNonErrorBounceSecs) )
                        {
                            this.RemoveNewCertificate(loStore, loOldCertificate, loNewCertificate, lsCertPfxPathFile);

                            return DoGetCert.bDoCheckOut(lbGetCertificate);
                        }
                    }

                    if ( lbGetCertificate && !this.bMainLoopStopped )
                    {
                        if ( !moProfile.bValue("-UseNonIISBindingOnly", false) )
                        {
                            Env.LogIt("");
                            Env.LogIt("Bind certificate in IIS ...");

                            lbGetCertificate = this.bDoIisBinding(
                                      loOldCertificate
                                    , loNewCertificate
                                    , lsCertName
                                    , lsSanArray
                                    , loServerManager
                                    , loStore
                                    );

                            moProfile[gsBindingFailedKey] = !lbGetCertificate;
                            moProfile.Save();
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped
                                && (moProfile.bValue("-UseNonIISBindingOnly", false) || moProfile.bValue("-UseNonIISBindingAlso", false)) )
                        {
                            Env.LogIt("");
                            Env.LogIt("Using non-IIS certificate binding ...");

                            if ( moProfile.bValue("-UseNonIISBindingPfxFile", false) )
                            {
                                lsNonIISBindingPfxPathFile = DoGetCert.oDomainProfile.sValue("-NonIISBindingPfxPathFile", "");

                                if ( String.IsNullOrEmpty(lsNonIISBindingPfxPathFile) )
                                {
                                    Env.LogIt(String.Format("    No PFX file has been defined for non-IIS binding of \"{0}\". Proceeding anyway ...", lsCertName));
                                }
                                else
                                {
                                    string  lsNewPassword = DoGetCert.oDomainProfile.sValue("-NonIISBindingPfxPassword", "");
                                            if ( "" == lsNewPassword )
                                                lsNewPassword = lsCertificatePassword;
                                            else
                                                lsNewPassword = HashClass.sDecrypted(Env.oChannelCertificate, lsNewPassword);

                                    lsNonIISBindingPfxPathFile = moProfile.sRelativeToProfilePathFile(lsNonIISBindingPfxPathFile.Replace("*", Path.GetFileNameWithoutExtension(lsCertPfxPathFile)));

                                    if ( lsNonIISBindingPfxPathFile == lsCertPfxPathFile )
                                        lsNonIISBindingPfxPathFile += "3";

                                    Env.LogIt("");
                                    Env.LogIt(String.Format("Copying new certificate to the non-IIS binding file for \"{0}\" ...", lsCertName));

                                    Directory.CreateDirectory(Path.GetDirectoryName(lsNonIISBindingPfxPathFile));

                                    if ( DoGetCert.oDomainProfile.bValue("-NonIISBindingPemFormat", false) || ".pem" == Path.GetExtension(lsNonIISBindingPfxPathFile).ToLower() )
                                    {
                                        Env.bPfxToPem(ResourceAssembly.GetName().Name, lsCertPfxPathFile, lsCertificatePassword, lsNonIISBindingPfxPathFile, lsNewPassword);
                                    }
                                    else
                                    {
                                        using (X509Certificate2 loCertificate = new X509Certificate2(lsCertPfxPathFile, lsCertificatePassword
                                                , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable))
                                        {
                                            File.WriteAllBytes(lsNonIISBindingPfxPathFile, loCertificate.Export(X509ContentType.Pfx, lsNewPassword));
                                        }

                                    }

                                    Env.LogSuccess();
                                    Env.LogIt("");
                                }
                            }

                            lbGetCertificate = Env.bRunPowerScript(moProfile.sValue("-NonIISBindingScript"
                                                                        , "echo \"'{NewCertificateThumbprint}' is the new thumbprint from the '{NewCertificatePfxPathFile}' file to bind for '{CertificateDomainName}' with an old thumbprint of '{OldCertificateThumbprint}'.\"")
                                                                        .Replace("{CertificateDomainName}", lsCertName)
                                                                        .Replace("{OldCertificateThumbprint}", null == loOldCertificate ? "" : loOldCertificate.Thumbprint)
                                                                        .Replace("{NewCertificateThumbprint}", loNewCertificate.Thumbprint)
                                                                        .Replace("{NewCertificatePfxPathFile}", lsNonIISBindingPfxPathFile)
                                                                        );

                            moProfile[gsBindingFailedKey] = !lbGetCertificate;
                            moProfile.Save();

                            if ( !lbGetCertificate && !moProfile.bValue("-UseNonIISBindingOnly", false) )
                            {
                                Env.LogIt("");
                                Env.LogIt("Undoing new certificate bindings in IIS ...");

                                bool lbBindingUndone = this.bDoIisBinding(
                                          loNewCertificate
                                        , loOldCertificate
                                        , lsCertName
                                        , lsSanArray
                                        , loServerManager
                                        , loStore
                                        );

                                if ( lbBindingUndone )
                                    Env.LogIt("    The new certificate bindings were successfully undone.");
                                else
                                    Env.LogIt("    The new certificate bindings were NOT successfully undone. MANUAL INTERVENTION IS REQUIRED!");

                                Env.LogIt("");
                            }
                        }

                        if ( lbGetCertificate )
                        {
                            Env.LogIt("");
                            Env.LogIt("The new certificate was successfully installed and bound.");

                            string  lsScriptBindingDone = moProfile.sValue("-ScriptBindingDone", "");
                                    if ( !String.IsNullOrEmpty(lsScriptBindingDone) )
                                    {
                                        Env.LogIt("");
                                        this.bRunPowerScript(lsScriptBindingDone
                                                .Replace("{CertificateDomainName}", lsCertName)
                                                .Replace("{OldCertificateThumbprint}", null == loOldCertificate ? "" : loOldCertificate.Thumbprint)
                                                .Replace("{NewCertificateThumbprint}", loNewCertificate.Thumbprint)
                                                );
                                    }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && !this.bCertificateSetupDone )
                            this.bCertificateSetupDone = true;

                        if ( lbGetCertificate && !this.bMainLoopStopped && null == loOldCertificate )
                        {
                            Env.LogIt("Old certificate (ie. setup certificate) NOT removed from the local store.");
                        }

                        if ( !moProfile.bValue(gsBindingFailedKey, false)  && !this.bMainLoopStopped )
                        {
                            if ( Env.sComputerName == DoGetCert.oDomainProfile.sValue("-CertArchiveComputerName", "") )
                            {
                                string lsNewPfxPathFile = DoGetCert.oDomainProfile.sValue("-CertArchivePfxPathFile", "");

                                if ( String.IsNullOrEmpty(lsNewPfxPathFile) )
                                {
                                    throw new Exception("The archive PFX file location is not defined! Can't continue.");
                                }
                                else
                                {
                                    string  lsNewPassword = DoGetCert.oDomainProfile.sValue("-CertArchivePfxPassword", "");
                                            if ( "" == lsNewPassword )
                                                lsNewPassword = lsCertificatePassword;
                                            else
                                                lsNewPassword = HashClass.sDecrypted(Env.oChannelCertificate, lsNewPassword);

                                    lsNewPfxPathFile = moProfile.sRelativeToProfilePathFile(lsNewPfxPathFile.Replace("*", Path.GetFileNameWithoutExtension(lsCertPfxPathFile)));

                                    if ( lsNewPfxPathFile == lsCertPfxPathFile )
                                        lsNewPfxPathFile += "4";

                                    Env.LogIt("");
                                    Env.LogIt(String.Format("Copying new certificate to the \"{0}\" certificate archive ...", lsCertName));

                                    if ( DoGetCert.oDomainProfile.bValue("-CertArchivePemFormat", false) || ".pem" == Path.GetExtension(lsNewPfxPathFile).ToLower() )
                                    {
                                        Env.bPfxToPem(ResourceAssembly.GetName().Name, lsCertPfxPathFile, lsCertificatePassword, lsNewPfxPathFile, lsNewPassword);
                                    }
                                    else
                                    {
                                        using (X509Certificate2 loCertificate = new X509Certificate2(lsCertPfxPathFile, lsCertificatePassword
                                                , X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.Exportable))
                                        {
                                            Directory.CreateDirectory(Path.GetDirectoryName(lsNewPfxPathFile));
                                            File.WriteAllBytes(lsNewPfxPathFile, loCertificate.Export(X509ContentType.Pfx, lsNewPassword));
                                        }

                                    }

                                    Env.LogSuccess();
                                    Env.LogIt("");
                                }
                            }

                            // Next to last step, remove old certificate from the local store.

                            // Note: this MUST be done before the last step since multiple
                            //       identifying certificates can't exist during SCS calls.

                            if ( !moProfile.bValue("-UseStandAloneMode", true)
                                    || (moProfile.bValue("-UseStandAloneMode", true)
                                        && moProfile.bValue("-RemoveReplacedCert", false))
                                    )
                            {
                                if ( null != loOldCertificate && loNewCertificate.Thumbprint == loOldCertificate.Thumbprint )
                                {
                                    Env.LogIt(String.Format("New certificate (\"{0}\") same as old. Not removed from local store.", loOldCertificate.Thumbprint));
                                }
                                else
                                {
                                    if ( null != loStore )
                                        loStore.Close();

                                    loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                                    loStore.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

                                    X509Certificate2Collection  loCertCollection = loStore.Certificates.Find(X509FindType.FindByThumbprint, loNewCertificate.Thumbprint, false);
                                                                lbGetCertificate = null != loCertCollection && 1 == loCertCollection.Count;

                                    if ( !lbGetCertificate )
                                    {
                                        Env.LogIt(String.Format("New certificate (\"{0}\") can't actually be found in the local store! Old certificate not removed.", loNewCertificate.Thumbprint));
                                    }
                                    else
                                    if ( null != loOldCertificate )
                                    {
                                        // Remove the old cert.
                                        DoGetCert.RemoveCertificate(loOldCertificate, loStore);

                                        Env.LogIt(String.Format("Old certificate (\"{0}\") removed from the local store.", loOldCertificate.Thumbprint));
                                    }
                                }
                            }
                        }

                        if ( lbGetCertificate && !this.bMainLoopStopped && !moProfile.bValue("-UseStandAloneMode", true) )
                        {
                            // At this point we need to load the new certificate into the service factory object.
                            Env.oChannelCertificate = null;
                            Env.oGetCertServiceFactory = null;

                            // Last step, remove old certificate from the repository (if there - it won't be removed until the new certificate is in place everywhere).

                            if ( null != loOldCertificate )
                                Env.LogIt("Requesting removal of the old certificate from the repository.");
                            else
                                Env.LogIt("Finalizing setup certificate replacement with the repository.");

                            if ( lbGetCertificate && !this.bMainLoopStopped )
                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                bool    bOldCertificateRemoved = loGetCertServiceClient.bOldCertificateRemoved(lsHash, lbtArrayMinProfile);
                                        if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                            loGetCertServiceClient.Abort();
                                        else
                                            loGetCertServiceClient.Close();

                                if ( !bOldCertificateRemoved )
                                    Env.LogIt(Env.sExceptionMessage("GetCertService.bOldCertificateRemoved: Failed removing old certificate (does not impact overall process success)."));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Env.LogIt(Env.sExceptionMessage(ex));
                lbGetCertificate = false;
            }
            finally
            {
                if ( lbGetCertificate && !this.bMainLoopStopped )
                {
                    // Delete ACME workspace.
                    if ( moProfile.bValue("-CleanupAcmeWork", true) && Directory.Exists(lsAcmeWorkPath) )
                        Directory.Delete(lsAcmeWorkPath, true);
                }

                if ( !lbCertOverrideApplies )
                {
                    // Remove the non-IIS binding cert (if any).
                    if ( !String.IsNullOrEmpty(lsNonIISBindingPfxPathFile) )
                        File.Delete(lsNonIISBindingPfxPathFile);

                    // Remove the cert (ie. PFX file from cert provider).
                    if ( !DoGetCert.oDomainProfile.bValue("-RetainLocalPfxFile", false) || !lbGetCertificate )
                        File.Delete(lsCertPfxPathFile);
                }
                else
                if ( lbGetCertificate && !this.bMainLoopStopped )
                {
                    // Remove the non-IIS binding cert (if any).
                    if ( !String.IsNullOrEmpty(lsNonIISBindingPfxPathFile) )
                        File.Delete(lsNonIISBindingPfxPathFile);

                    // Remove the cert (ie. the override PFX file - only upon success).
                    File.Delete(lsCertPfxPathFile);
                }

                if (       moProfile.bValue(gsBindingFailedKey, false)
                        && null != loNewCertificate
                        && !moProfile.bValue("-RetainNewCertAfterError", false)
                        )
                    this.RemoveNewCertificate(loStore, loNewCertificate, loNewCertificate, lsCertPfxPathFile);
                    // Referencing loNewCertificate as aoOldCertificate guarantees removal.

                moProfile.Remove("-RetainNewCertAfterError");
                moProfile.Save();                    

                Env.LogIt("");

                this.bRunPowerScript(moProfile.sValue("-ScriptSessionClose", @"
$session = Connect-PSSession -ComputerName localhost -Name GetCert
$session | Remove-PSSession
Get-WSManInstance -ResourceURI Shell -Enumerate
                        "), true);

                if ( !moProfile.bValue("-UseStandAloneMode", true) )
                using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                {
                    loGetCertServiceClient.bUnlockCertificateRenewal(lsHash, lbtArrayMinProfile);
                    if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                        loGetCertServiceClient.Abort();
                    else
                        loGetCertServiceClient.Close();
                }

                if ( null != loStore )
                    loStore.Close();
            }

            if ( this.bMainLoopStopped )
            {
                Env.LogIt("Stopped.");
                lbGetCertificate = false;
            }

            if ( lbGetCertificate )
            {
                moProfile.Remove("-CertOverrideErrorCount");
                moProfile.Remove("-LoadBalancerPendingErrorCount");
                moProfile.Remove("-RenewalDateOverride");
                moProfile.Remove("-RepositoryCertsOnlyErrorCount");
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = true;
                moProfile.Save();

                Env.LogIt(String.Format("The get certificate process completed successfully (\"{0}\").", lsCertName));
            }
            else
            {
                moProfile["-PreviousProcessTime"] = DateTime.Now;
                moProfile["-PreviousProcessOk"] = false;
                moProfile.Save();

                Env.LogIt(String.Format("At least one stage failed or the process was stopped (\"{0}\"). Check log for errors.", lsCertName));
            }

            Env.LogIt("");

            lbGetCertificate = DoGetCert.bDoCheckOut(lbGetCertificate);

            return lbGetCertificate;
        }
    }

    public enum UpdatedEXEs
    {
         Client = 1
        ,GcFailSafe = 2
    }
}
